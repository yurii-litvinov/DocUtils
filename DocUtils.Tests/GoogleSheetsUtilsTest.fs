module DocUtils.Tests

open NUnit.Framework
open System.IO

let credentials = "credentials.json"

[<Test>]
let SheetServiceShallListCorrectSheetsTest () =
    if File.Exists credentials then
        task {
            use! service = GoogleSheetService.CreateAsync(credentials, "DocUtilsTests")
            let! sheets = service.Spreadsheet("1b1fhGFInVDNXAb_Ok14Nl03V-DviKe-GrE2Geuwsw9o").GetSheetsAsync()
            Assert.AreEqual([ "19.Б07"; "19.Б08"; "19.Б09"; "19.Б10" ], sheets)
        } |> Async.AwaitTask |> Async.RunSynchronously
    else
        Assert.Ignore("No credentials for Google Sheets found.")

[<Test>]
let ReadingSheetShallGetExpectedValuesTest () =
    if File.Exists credentials then
        task {
            use! service = GoogleSheetService.CreateAsync(credentials, "DocUtilsTests")
            let sheet = service.Sheet("1b1fhGFInVDNXAb_Ok14Nl03V-DviKe-GrE2Geuwsw9o", "19.Б07")
            let! column = sheet.ReadColumnAsync("A")
            Assert.AreEqual("ФИО", Seq.head column)
        } |> Async.AwaitTask |> Async.RunSynchronously
    else
        Assert.Ignore("No credentials for Google Sheets found.")

[<Test>]
let ReadingByColumnHeadersShallGetExpectedValues () =
    if File.Exists credentials then
        task {
            use! service = GoogleSheetService.CreateAsync(credentials, "DocUtilsTests")
            let sheet = service.Sheet("1MCVf88nLnYuRdPKURYbX8dNcMLweTWUUrqOCmfcJmvI", "СП")
            let! result = sheet.ReadByHeadersAsync(["ФИО"; "Научный руководитель"; "Тема"; "Зачёт"])
            Assert.AreEqual("Израилев Андрей Дмитриевич", (Seq.head result)["ФИО"])
            Assert.AreEqual("да", (Seq.head result)["Зачёт"])
        } |> Async.AwaitTask |> Async.RunSynchronously
    else
        Assert.Ignore("No credentials for Google Sheets found.")
