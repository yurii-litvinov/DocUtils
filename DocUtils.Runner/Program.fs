open DocUtils.YandexSheets

let (!) task =
    task |> Async.AwaitTask |> Async.RunSynchronously

let yandexService = YandexService.FromClientSecretsFile()

let spreadsheet =
    !(yandexService.Spreadsheet "Курсы/ТРПО/ТРПО, программа курса-копия.xlsx")

let sheet = spreadsheet.Sheet "Лист1"
sheet.Column 0 |> Seq.iter (printfn "%A")

let data = Seq.init 10 id |> Seq.map string
sheet.WriteColumn 1 1 data

!(spreadsheet.SaveTo "test.xlsx")

!(spreadsheet.Save())
