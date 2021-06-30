namespace DocUtils

open System
open Google.Apis.Sheets.v4
open System.IO
open Google.Apis.Auth.OAuth2
open Google.Apis.Services
open System.Threading
open Google.Apis.Util.Store
open Google.Apis.Sheets.v4.Data

type Sheet(service: SheetsService, spreadsheetId: string, sheetId: string) =
    let checkOffset offset =
        if offset <= 0 then 
            raise <| new ArgumentOutOfRangeException ("Offset shall be greater than zero.")

    let checkBoundaries offset size =
        checkOffset offset
        if size < offset then 
            raise <| new ArgumentOutOfRangeException ("size shall be greater than offset.")

    let read startColumn endColumn offset size =
        let offset = defaultArg offset 1
        let size = defaultArg size 1000
        checkBoundaries offset size
        let range = $"{sheetId}!{startColumn}{offset}:{endColumn}{offset + size}"

        let request = service.Spreadsheets.Values.Get(spreadsheetId, range)

        let values = request.Execute().Values 
        if values <> null then
            values |> Seq.map (Seq.map string)
        else
            Seq.empty

    member _.WriteColumn(column: string, data: #seq<string>, ?offset: int) =
        let offset = defaultArg offset 1
        checkOffset offset
        let range = sheetId + "!" + column + (string offset) + ":" + column + (string ((data |> Seq.length) + offset))
        let valueRange = ValueRange(Values = [| data |> Seq.cast<obj> |> Seq.toArray |])
        valueRange.MajorDimension <- "COLUMNS"
        let request = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range)
        request.ValueInputOption <- Nullable(SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW)

        request.Execute() |> ignore

    member _.ReadSheet(startColumn: string, endColumn: string, ?offset: int, ?size: int) =
        read startColumn endColumn offset size

    member _.ReadColumn(column: string, ?offset: int, ?size: int) =
        read column column offset size |> Seq.concat

type Spreadsheet(service: SheetsService, spreadsheetId: string) =
    member _.Sheet (sheetId: string) = Sheet(service, spreadsheetId, sheetId)

    member _.Sheets () =
        let spreadsheet = service.Spreadsheets.Get(spreadsheetId).Execute ()
        spreadsheet.Sheets |> Seq.map (fun s -> s.Properties.Title)

type GoogleSheetService(credentialsFileName: string, applicationName: string) =
    let service = 
        use credentialsStream = new FileStream(credentialsFileName, FileMode.Open, FileAccess.Read)

        let credential = 
            GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.FromStream(credentialsStream).Secrets,
                [ SheetsService.Scope.Spreadsheets ],
                "user",
                CancellationToken.None,
                new FileDataStore("token.json", true)).Result

        new SheetsService(
            BaseClientService.Initializer(
                HttpClientInitializer = credential,
                ApplicationName = applicationName
            )
        )

    member _.Spreadsheet (spreadsheetId: string) =
        Spreadsheet(service, spreadsheetId)

    member _.Sheet (spreadsheetId: string, sheetId: string) = 
        Spreadsheet(service, spreadsheetId).Sheet(sheetId)

    interface IDisposable with
        member _.Dispose () = service.Dispose ()
