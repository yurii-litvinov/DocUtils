module DocUtils.Xlsx

open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml
open System.IO
open System
open DocumentFormat.OpenXml.Spreadsheet

/// Represents a single sheet (or tab) in a spreadsheet, provides methods to read and modify data.
type Sheet internal (workbookPart: WorkbookPart, sheet: SheetData) =
    let cellValue (cell: Cell) =
        let sharedStringTablePart =
            workbookPart.GetPartsOfType<SharedStringTablePart>() |> Seq.head

        let sharedStringTable = sharedStringTablePart.SharedStringTable

        if not (isNull cell.DataType) && cell.DataType = EnumValue(CellValues.SharedString) then
            let ssid = cell.CellValue.Text |> int
            sharedStringTable.ChildElements.[ssid].InnerText
        elif isNull cell.CellValue then
            ""
        else
            cell.CellValue.Text

    let cellValueByColumn (row: Row) column =
        let cell = row.Elements<Cell>() |> Seq.skip column |> Seq.head
        cellValue cell

    let readColumn columnNumber =
        seq {
            for row in sheet.Elements<Row>() do
                if row.Elements<Cell>() |> Seq.length > columnNumber then
                    yield cellValueByColumn row columnNumber
        }

    let readColumnByName columnName =
        let header = sheet.Elements<Row>() |> Seq.head
        let mutable column = 0

        seq {
            for cell in header.Elements<Cell>() do
                if cellValue cell = columnName then
                    yield! readColumn column

                column <- column + 1
        }
        |> Seq.skip 1

    /// Returns contents of a column with given header (first row value) as a string sequence.
    member _.Column(columnName: string) = readColumnByName columnName

    /// Returns contents of a column with given number as a string sequence.
    member _.Column(columnNumber: int) = readColumn columnNumber

    /// Assumes that the the first row is a row with headings and reads only those columns.
    /// Returns a list of maps that map header names to row values.
    member _.ReadByHeaders(columnNames: string list) : seq<Map<string, string>> =
        let valuesByColumn = columnNames |> List.map readColumnByName |> List.map Seq.toList

        let maxColumnLength = valuesByColumn |> List.map List.length |> List.max

        let valuesByColumn =
            valuesByColumn
            |> List.map (fun column ->
                if List.length column < maxColumnLength then
                    column @ (List.replicate (maxColumnLength - List.length column) "")
                else
                    column)

        let valuesByRow =
            valuesByColumn
            |> List.fold
                (fun (acc: list<list<string>>) column ->
                    List.zip column acc |> List.map (fun (item, list) -> item :: list))
                (List.replicate maxColumnLength [])
            |> List.map List.rev

        let valuesWithColumnNames =
            valuesByRow |> List.map (List.zip columnNames) |> List.map Map.ofList

        valuesWithColumnNames

    /// Writes values to a given column starting from given offset as string values.
    member _.WriteColumn (columnNumber: int) (offset: int) (data: string seq) =
        let mutable rowNumber = 0
        let dataWithOffset = Seq.append (Seq.replicate offset "") data
        let dataAndRow = Seq.zip dataWithOffset (sheet.Elements<Row>())

        for data, row in dataAndRow do
            if rowNumber >= offset then
                let cell = row.Elements<Cell>() |> Seq.skip columnNumber |> Seq.head
                cell.CellValue <- new CellValue(data)
                cell.DataType <- new EnumValue<_>(CellValues.String)

            rowNumber <- rowNumber + 1

        workbookPart.Workbook.Save()

/// Represents a .xlsx document.
type Spreadsheet internal (dataStream: Stream) =
    let openXlsxSheetFromStream (stream: Stream) =
        let document = SpreadsheetDocument.Open(stream, true)
        let workbookPart = document.WorkbookPart

        let sheets = workbookPart.Workbook.Sheets |> Seq.cast<Spreadsheet.Sheet>

        let sheets =
            sheets
            |> Seq.map (fun sheet ->
                let sheetId = sheet.Id.Value
                let worksheet = (workbookPart.GetPartById(sheetId) :?> WorksheetPart).Worksheet
                let sheetData = worksheet.Elements<SheetData>() |> Seq.head
                sheet.Name, Sheet(workbookPart, sheetData))
            |> Map.ofSeq

        (sheets, document)

    let sheets, document = openXlsxSheetFromStream (dataStream)

    /// Opens .xlsx spreadsheet from file.
    static member FromFile(fileName: string) =
        new Spreadsheet(new FileStream(fileName, FileMode.Open))

    /// Opens .xlsx spreadsheet from memory.
    static member FromByteArray(data: byte array) = new Spreadsheet(new MemoryStream(data))

    /// Returns all sheets in a spreadsheet.
    member _.Sheets() : Sheet seq = sheets.Values

    /// Returns a sheet (tab in a spreadsheet) by name.
    member _.Sheet(sheetName: string) : Sheet = sheets[sheetName]

    /// Saves entire spreadsheet to a given stream.
    member _.SaveTo(stream: Stream) =
        task {
            let clone = document.Clone(stream)
            clone.Close()
        }

    /// Saves entire spreadsheet to a given file.
    member this.SaveTo(path: string) =
        task {
            use stream = new FileStream(path, FileMode.Create)
            do! this.SaveTo stream
        }

    interface IDisposable with
        member _.Dispose() = dataStream.Dispose()
