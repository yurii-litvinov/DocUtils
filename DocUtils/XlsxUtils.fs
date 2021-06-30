module LessPainfulXlsx

open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml

let openXlsxSheet (fileName: string) sheetName =
    let dataDocument = SpreadsheetDocument.Open(fileName, false)
    let workbookPart = dataDocument.WorkbookPart

    let sheets = workbookPart.Workbook.Sheets |> Seq.cast<Spreadsheet.Sheet>
    let sheet = sheets |> Seq.find(fun s -> s.Name.Value = sheetName)
    let sheetId = sheet.Id.Value
    let worksheet = (workbookPart.GetPartById(sheetId) :?> WorksheetPart).Worksheet
    let sheetData = worksheet.Elements<Spreadsheet.SheetData>() |> Seq.head

    (workbookPart, sheetData)

let private cellValue (workbookPart: WorkbookPart) (cell: Spreadsheet.Cell) =
    let sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>() |> Seq.head
    let sst = sstPart.SharedStringTable
    if cell.DataType <> null && cell.DataType = EnumValue(Spreadsheet.CellValues.SharedString) then
        let ssid = cell.CellValue.Text |> int
        sst.ChildElements.[ssid].InnerText
    elif cell.CellValue = null then
        ""
    else
        cell.CellValue.Text

let private cellValueByColumn (workbookPart: WorkbookPart) (row: Spreadsheet.Row) column = 
    let cell = row.Elements<Spreadsheet.Cell>() |> Seq.skip column |> Seq.head
    cellValue workbookPart cell

let readColumn (workbookPart: WorkbookPart, sheet: Spreadsheet.SheetData) columnNumber =
    seq {
        for row in sheet.Elements<Spreadsheet.Row>() do
            if row.Elements<Spreadsheet.Cell>() |> Seq.length > columnNumber then
                yield cellValueByColumn workbookPart row columnNumber
    }

let readColumnByName (workbookPart: WorkbookPart, sheet: Spreadsheet.SheetData) columnName =
    let header = sheet.Elements<Spreadsheet.Row>() |> Seq.head
    let mutable column = 0
    seq {
        for cell in header.Elements<Spreadsheet.Cell>() do
            if cellValue workbookPart cell = columnName then
                yield! readColumn (workbookPart, sheet) column
            column <- column + 1
    } |> Seq.skip 1