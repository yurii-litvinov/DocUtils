module DocUtils.Xlsx.Tests

open DocUtils.Xlsx
open NUnit.Framework
open System.IO
open System

[<Test>]
let EmptySheetIsSuccessfullyCreated () =
    use spreadsheet = Spreadsheet.New()

    try
        spreadsheet.SaveTo("test.xlsx") |> Async.AwaitTask |> Async.RunSynchronously
        Assert.IsTrue <| File.Exists("test.xlsx")
        Assert.That(FileInfo("test.xlsx").Length, Is.GreaterThan 0)
    finally
        File.Delete("test.xlsx")

[<Test>]
let WriteRowActuallyWritesRows () =
    use originalSpreadsheet = Spreadsheet.New("Sheet1")
    let originalSheet = originalSpreadsheet.Sheet("Sheet1")
    originalSheet.WriteRow([ "1"; "2"; "3" ])
    originalSheet.WriteRow([ "4"; "5"; "6" ])

    try
        originalSpreadsheet.SaveTo("test.xlsx")
        |> Async.AwaitTask
        |> Async.RunSynchronously

        (originalSpreadsheet :> IDisposable).Dispose()

        use spreadsheetFromFile = Spreadsheet.FromFile("test.xlsx")
        let sheetFromFile = spreadsheetFromFile.Sheet("Sheet1")
        Assert.That(sheetFromFile.Column "A", Is.EqualTo([ "1"; "4" ]))
        Assert.That(sheetFromFile.Column "B", Is.EqualTo([ "2"; "5" ]))
        Assert.That(sheetFromFile.Column "C", Is.EqualTo([ "3"; "6" ]))
    finally
        File.Delete("test.xlsx")
