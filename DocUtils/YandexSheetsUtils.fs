namespace DocUtils.YandexSheets

open System
open DocUtils.Xlsx
open System.Net
open System.Net.Http
open System.Text
open System.Net.Http.Headers
open FSharp.Json
open System.IO

/// Record for deserialization of authorization token info from Yandex.OAuth.
type HttpAuthToken =
    { access_token: string
      expires_in: int
      refresh_token: string
      token_type: string }

/// Record for storing authorization token info in a file.
type SerializedAuthToken =
    { accessToken: string
      expiresAt: DateTime
      refreshToken: string }

/// Record for providing Yandex client secrets to an application.
type SerializedClientSecrets =
    { clientId: string
      clientSecret: string }

/// Record for deserialization of Yandex API Link object.
type HttpLink =
    { href: string
      method: string
      templated: bool }

/// Exception thrown when something is wrong with server communication.
exception ServerCommunicationException of string

/// Specialized .xlsx spreadsheet that allows convenient work with Yandex service.
type YandexSpreadsheet internal (service: YandexService, path: string, data: Stream) =
    inherit Spreadsheet(data)

    /// Creates a spreadsheed from downloaded document. Supposed to be called from YandexService.
    static member internal FromByteArray(service: YandexService, path: string, data: byte array) =
        new YandexSpreadsheet(service, path, new MemoryStream(data))

    /// Saves and uploads the spreadsheet to Yandex.Cloud using its original path.
    member this.SaveAsync() =
        task {
            use stream = new MemoryStream()
            do! (this :> Spreadsheet).SaveTo(stream)
            do! service.UploadAsync(stream, path)
        }

/// Service for Yandex Disk, typically one for the entire application.
and YandexService(clientId: string, clientSecret: string) =

    /// Authorization token. Gotten from Yandex OAuth or loaded from file (if fresh enough).
    let mutable authToken = ""

    /// Authenticates a service with Yandex.OAuth and fills authToken.
    member private _.AuthenticateAsync() =
        let getAuthCode () =
            task {
                let codeUrl =
                    $"https://oauth.yandex.ru/authorize?response_type=code^&client_id={clientId}^&redirect_uri=http://localhost:8888/"

                use listener = new HttpListener()
                listener.Prefixes.Add "http://localhost:8888/"
                listener.Start()

                if Runtime.InteropServices.RuntimeInformation.IsOSPlatform(Runtime.InteropServices.OSPlatform.Windows) then
                    System.Diagnostics.Process.Start("cmd.exe", $"/C start {codeUrl}") |> ignore
                elif Runtime.InteropServices.RuntimeInformation.IsOSPlatform(Runtime.InteropServices.OSPlatform.Linux) then
                    let codeUrl = codeUrl.Replace("^", "")
                    System.Diagnostics.Process.Start("yandex-browser-stable", $"{codeUrl}") |> ignore
                else
                    failwith "Unsupported OS"

                let! context = listener.GetContextAsync()
                let request = context.Request
                let code = request.QueryString["code"]

                let responseString =
                    $"<HTML><BODY>Access code {code} received, you can close this window</BODY></HTML>"

                let buffer = System.Text.Encoding.UTF8.GetBytes(responseString)

                let response = context.Response
                response.ContentLength64 <- buffer.Length
                let output = response.OutputStream
                do! output.WriteAsync(buffer, 0, buffer.Length)
                output.Close()
                listener.Stop()
                return code
            }

        let getAuthToken code =
            task {
                let tokenUrl = $"https://oauth.yandex.ru/token"

                let credentialsString =
                    Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes($"{clientId}:{clientSecret}"))

                use httpClient = new HttpClient()
                use message = new HttpRequestMessage(HttpMethod.Post, tokenUrl)

                message.Content <-
                    new StringContent(
                        $"grant_type=authorization_code&code={code}",
                        Encoding.UTF8,
                        "application/x-www-form-urlencoded"
                    )

                message.Headers.Authorization <- AuthenticationHeaderValue("Basic", credentialsString)

                let! response = httpClient.SendAsync(message)
                let! responseContent = response.Content.ReadAsStringAsync()

                let authToken = Json.deserialize<HttpAuthToken> (responseContent)
                return authToken
            }

        task {
            if File.Exists("yandexToken.json") then
                let contents = File.ReadAllText("yandexToken.json")
                let token = Json.deserialize<SerializedAuthToken> (contents)

                if token.expiresAt > DateTime.Now then
                    authToken <- token.accessToken

            if authToken = "" then
                let! code = getAuthCode ()
                let! token = getAuthToken code

                let serializedToken =
                    { accessToken = token.access_token
                      expiresAt = DateTime.Now + TimeSpan(0, 0, token.expires_in - 10)
                      refreshToken = token.refresh_token }

                File.WriteAllText("yandexToken.json", Json.serialize serializedToken)

                authToken <- token.access_token
        }

    /// Initializes the service from "clientSecrets.json" that shall contain JSON with "clientId" and "clientSecret" fields.
    /// Client secrets are provided by https://oauth.yandex.ru/ for a registered application.
    static member FromClientSecretsFile() =
        let contents = File.ReadAllText("clientSecrets.json")
        let secrets = Json.deserialize<SerializedClientSecrets> (contents)
        YandexService(secrets.clientId, secrets.clientSecret)

    /// Downloads and returns a spreadsheet by a given absolute path on Yandex.Disk.
    member this.GetSpreadsheetAsync(path: string) =
        task {
            if authToken = "" then
                do! this.AuthenticateAsync()

            let encodedPath = Uri.EscapeDataString(path)

            let requestUri =
                $"https://cloud-api.yandex.net/v1/disk/resources/download?path={encodedPath}"

            use httpClient = new HttpClient()
            httpClient.DefaultRequestHeaders.Authorization <- new AuthenticationHeaderValue("OAuth", authToken)
            let! response = httpClient.GetAsync requestUri
            let! responseContent = response.Content.ReadAsStringAsync()
            let linkObject = Json.deserialize<HttpLink> (responseContent)
            let downloadLink = linkObject.href

            let! fileResponse = httpClient.GetAsync downloadLink
            let! responseContent = fileResponse.Content.ReadAsByteArrayAsync()
            return YandexSpreadsheet.FromByteArray(this, path, responseContent)
        }

    /// Downloads and returns a spreadsheet by a given folder URL and file name without extension.
    member this.GetSpreadsheetByFolderAndFileNameAsync(folderUrl: string, fileName: string) =
        let spreadsheetPath =
            folderUrl.Remove(0, "https://disk.yandex.ru/client/disk/".Length)

        let unencodedSpreadsheetPath = Uri.UnescapeDataString(spreadsheetPath)

        let unencodedFullSpreadsheetPath =
            unencodedSpreadsheetPath + "/" + fileName + ".xlsx"

        this.GetSpreadsheetAsync unencodedFullSpreadsheetPath

    /// Uploads sheet back to Yandex.Disk using given path. Supposed to be used from YandexSpreadsheet.Save.
    member internal this.UploadAsync(stream: Stream, path: string) =
        task {
            if authToken = "" then
                do! this.AuthenticateAsync()

            let path = Uri.EscapeDataString(path)

            let requestUri =
                $"https://cloud-api.yandex.net/v1/disk/resources/upload?path={path}&overwrite=true"

            use httpClient = new HttpClient()
            httpClient.DefaultRequestHeaders.Authorization <- new AuthenticationHeaderValue("OAuth", authToken)
            let! response = httpClient.GetAsync requestUri
            let! responseContent = response.Content.ReadAsStringAsync()

            try
                let linkObject = Json.deserialize<HttpLink> (responseContent)
                let uploadLink = linkObject.href

                use message = new HttpRequestMessage(HttpMethod.Put, uploadLink)
                stream.Seek(0, SeekOrigin.Begin) |> ignore
                use fileStreamContent = new StreamContent(stream)
                fileStreamContent.Headers.ContentType <- MediaTypeHeaderValue("application/octet-stream")
                fileStreamContent.Headers.ContentLength <- stream.Length
                message.Content <- fileStreamContent

                let! response = httpClient.SendAsync(message)
                response.EnsureSuccessStatusCode() |> ignore
            with :? JsonDeserializationError ->
                raise (ServerCommunicationException(responseContent))
        }
