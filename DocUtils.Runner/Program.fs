/// Minimal example of library usage. Used to quickly test library functions.
open DocUtils.Xlsx

let (!) task =
    task |> Async.AwaitTask |> Async.RunSynchronously

let spreadsheet = Spreadsheet.New()

(spreadsheet.Sheets() |> Seq.head).WriteRow [ "1"; "2" ]

!(spreadsheet.SaveTo "test.xlsx")
