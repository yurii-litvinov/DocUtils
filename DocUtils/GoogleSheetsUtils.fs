namespace DocUtils

open System
open Google.Apis.Sheets.v4
open System.IO
open Google.Apis.Auth.OAuth2
open Google.Apis.Services
open System.Threading
open Google.Apis.Util.Store
open Google.Apis.Sheets.v4.Data
open System.Threading.Tasks

/// Represents a single sheet (or tab) in a spreadsheet, provides methods to read and modify data.
type Sheet internal (service: SheetsService, spreadsheetId: string, sheetId: string) =
    let checkOffset offset =
        if offset <= 0 then 
            raise <| ArgumentOutOfRangeException ("Offset shall be greater than zero.")

    let readAsync startColumn endColumn offset size =
        task {
            let offset = defaultArg offset 1
            let size = defaultArg size 1000
            let range = $"'{sheetId}'!{startColumn}{offset}:{endColumn}{offset + size}"

            let request = service.Spreadsheets.Values.Get(spreadsheetId, range)

            let! response = request.ExecuteAsync()
            let values = response.Values
            let values =
                if not (isNull values) then
                    values |> Seq.map (Seq.map string)
                else
                    Seq.empty

            let maxRowLength = values |> Seq.map Seq.length |> Seq.max

            return values
                |> Seq.map (fun row -> Seq.append row (Seq.replicate (maxRowLength - Seq.length row) ""))
        }

    /// Writes a column to a sheet.
    member _.WriteColumnAsync(column: string, data: #seq<string>, ?offset: int) =
        task {
            let offset = defaultArg offset 1
            checkOffset offset
            let range = sheetId + "!" + column + (string offset) + ":" + column + (string ((data |> Seq.length) + offset))
            let valueRange = ValueRange(Values = [| data |> Seq.cast<obj> |> Seq.toArray |])
            valueRange.MajorDimension <- "COLUMNS"
            let request = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range)
            request.ValueInputOption <- Nullable(SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW)

            request.Execute() |> ignore
        }

    /// Reads given range of a sheet and returns a sequence of rows in that range.
    member _.ReadSheetAsync(startColumn: string, endColumn: string, ?offset: int, ?size: int) =
        readAsync startColumn endColumn offset size

    /// Reads single column of a sheet.
    member _.ReadColumnAsync(column: string, ?offset: int, ?size: int) =
        task {
            let! result = readAsync column column offset size
            return result |> Seq.concat
        }

    /// Assumes that the offset row is a row with headings and reads only those columns. 
    /// Returns a list of maps that map header names to row values.
    member _.ReadByHeadersAsync(columnNames: string list, ?offset: int, ?size: int): Task<seq<Map<string, string>>> =
        task {
            let offset = defaultArg offset 1
            let size = defaultArg size 1000
            let columnNames = Set.ofList columnNames

            let range = $"'{sheetId}'!{offset}:{offset + size}"

            let request = service.Spreadsheets.Values.Get(spreadsheetId, range)

            let! response = request.ExecuteAsync()
            let values = response.Values
            if not (isNull values) then
                let values = values |> Seq.map (Seq.map string)
                let maxRowLength = values |> Seq.map Seq.length |> Seq.max
                let values = values |> Seq.map (fun row -> Seq.append row (Seq.replicate (maxRowLength - Seq.length row) ""))
                let headersRow = values |> Seq.head
                return
                    values 
                    |> Seq.skip 1
                    |> Seq.map (
                        fun row -> 
                            let zippedRow = Seq.zip headersRow row
                            zippedRow
                            |> Seq.filter (fun (h, _) -> columnNames.Contains h)
                            |> Map.ofSeq
                       )
            else
                return Seq.empty
        }

/// Represents a single Google Sheets document.
type Spreadsheet internal (service: SheetsService, spreadsheetId: string) =
    /// Returns sheet object by Id. Does not query Google Sheets servers, just provides proxy.
    member _.Sheet (sheetId: string) = Sheet(service, spreadsheetId, sheetId)

    /// Returns a list of sheet ids (or tabs) in a given document. Does query server. Actual sheets can then be retrieved by id.
    member _.GetSheetsAsync () =
        task {
            let! spreadsheet = service.Spreadsheets.Get(spreadsheetId).ExecuteAsync ()
            return spreadsheet.Sheets |> Seq.map (fun s -> s.Properties.Title)
        }

/// Represents a service for accessing Google Sheets. Typically, one for the entire application.
type GoogleSheetService private (service: SheetsService) =
    static member CreateAsync(credentialsFileName: string, applicationName: string) =
        task {
            use credentialsStream = new FileStream(credentialsFileName, FileMode.Open, FileAccess.Read)

            let! credential = 
                GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(credentialsStream).Secrets,
                    [ SheetsService.Scope.Spreadsheets ],
                    "user",
                    CancellationToken.None,
                    FileDataStore("token.json", true))

            return new GoogleSheetService(new SheetsService(
                    BaseClientService.Initializer(
                        HttpClientInitializer = credential,
                        ApplicationName = applicationName
                    )
                )
            )
        }

    /// Gets a document by its id (hash). Does not actually query servers, returns proxy.
    member _.Spreadsheet (spreadsheetId: string) =
        Spreadsheet(service, spreadsheetId)

    /// Gets sheet (a tab in a document) by document id (hash) and sheet id (tab name). Does not query servers.
    member _.Sheet (spreadsheetId: string, sheetId: string) = 
        Spreadsheet(service, spreadsheetId).Sheet(sheetId)

    interface IDisposable with
        member _.Dispose () = service.Dispose ()
