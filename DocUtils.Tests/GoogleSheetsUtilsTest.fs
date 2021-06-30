module DocUtils.Tests

open NUnit.Framework

[<Test>]
let SheetServiceShallListCorrectSheetsTest () =
    use service = new GoogleSheetService("../../../../../credentials.json", "DocUtilsTests")
    let sheets = service.Spreadsheet("1b1fhGFInVDNXAb_Ok14Nl03V-DviKe-GrE2Geuwsw9o").Sheets()
    Assert.AreEqual([ "19.Б07"; "19.Б08"; "19.Б09"; "19.Б10" ], sheets)


[<Test>]
let ReadingTest () =
    use service = new GoogleSheetService("../../../../../credentials.json", "AssignmentMatcher")
    let sheet = service.Sheet("1b1fhGFInVDNXAb_Ok14Nl03V-DviKe-GrE2Geuwsw9o", "19.Б07")
    let column = sheet.ReadColumn "A" 0
    Assert.AreEqual("ФИО", Seq.head column)
