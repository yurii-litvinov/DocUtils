module DocUtils.YandexSheets.Tests

open DocUtils.YandexSheets
open NUnit.Framework
open System.IO

[<Test>]
[<Ignore("Requires manual authentication")>]
let YandexSheetsShallSuccessfullyAuthenticate () =
    task {
        let service = YandexService.FromClientSecretsFile()
        let! service = service.Spreadsheet "Описание.xlsx"
        ()
    }
    |> Async.AwaitTask
    |> Async.RunSynchronously
