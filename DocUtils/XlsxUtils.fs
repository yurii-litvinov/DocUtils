module DocUtils.Xlsx

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open System.IO
open System

/// Represents a single sheet (or tab) in a spreadsheet, provides methods to read and modify data.
type Sheet internal (workbookPart: WorkbookPart, sheet: SheetData) =
    let alphabet =
        seq {
            let letters = [| 'A' .. 'Z' |] |> Seq.map string

            for letter in letters do
                yield letter

            for letter1 in letters do
                for letter2 in letters do
                    yield letter1 + letter2

            for letter1 in letters do
                for letter2 in letters do
                    for letter3 in letters do
                        yield letter1 + letter2 + letter3
        }
        |> Seq.toArray

    let tryFindCell (row: Row) (column: string) =
        row.Elements<Cell>()
        |> Seq.tryFind (fun c -> c.CellReference.ToString() = $"{column}{row.RowIndex}")

    let cellValue (cell: Cell) =
        let sharedStringTableSeq = workbookPart.GetPartsOfType<SharedStringTablePart>()

        if Seq.isEmpty sharedStringTableSeq then
            if isNull cell.CellValue then "" else cell.CellValue.Text
        else
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
        match tryFindCell row column with
        | Some c -> cellValue c
        | None -> ""

    let createCell column row (value: string) =
        new Cell(CellReference = $"{column}{row}", CellValue = new CellValue(value), DataType = CellValues.String)

    let readColumn columnIndex =
        seq {
            for row in sheet.Elements<Row>() do
                yield cellValueByColumn row columnIndex
        }

    let readColumnByName columnName =
        let header = sheet.Elements<Row>() |> Seq.head
        let mutable column = 0

        seq {
            for cell in header.Elements<Cell>() do
                if cellValue cell = columnName then
                    yield! readColumn alphabet[column]

                column <- column + 1
        }
        |> Seq.skip 1

    /// Returns contents of a column with given header (first row value) as a string sequence.
    member _.ColumnByName(columnName: string) = readColumnByName columnName

    /// Returns contents of a column with given letter index as a string sequence.
    member _.Column(columnIndex: string) = readColumn columnIndex

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
    member _.WriteColumn (columnIndex: string) (offset: int) (data: string seq) =
        let mutable rowNumber = 0

        // rowNumber starts with 0, but actual row numbers in .xlsx --- from 1.
        let offset = offset - 1
        let dataWithOffset = Seq.append (Seq.replicate offset "") data
        let dataAndRow = Seq.zip dataWithOffset (sheet.Elements<Row>())

        for data, row in dataAndRow do
            if rowNumber >= offset then
                let cell = tryFindCell row columnIndex

                match cell with
                | Some c ->
                    c.CellValue <- new CellValue(data)
                    c.DataType <- new EnumValue<_>(CellValues.String)
                | None ->
                    let c = createCell columnIndex row.RowIndex data
                    row.AppendChild(c) |> ignore

            rowNumber <- rowNumber + 1

        workbookPart.Workbook.Save()

    /// Appends a new row with given string values to the end of the sheet.
    member _.WriteRow(data: string seq) =
        let nextRowIndex = sheet.Elements<Row>() |> Seq.length |> uint |> (+) 1u
        let row = new Row(RowIndex = nextRowIndex)
        let mutable columnIndex = 0

        for cellValue in data do
            let cell = createCell alphabet[columnIndex] nextRowIndex cellValue

            row.AppendChild(cell) |> ignore
            columnIndex <- columnIndex + 1

        sheet.AppendChild(row) |> ignore
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

    /// Creates new empty spreadsheet
    /// <param> sheetName - spreadshet is created with one empty sheet, this parameter allows to specify a name for it, "Лист 1" by default </param>
    static member New(?sheetName: string) =
        let memoryStream = new MemoryStream()

        use document =
            SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook)

        let workbookPart = document.AddWorkbookPart()
        workbookPart.Workbook <- new Workbook()

        let worksheetPart = workbookPart.AddNewPart<WorksheetPart>()
        worksheetPart.Worksheet <- new Worksheet()

        let sheetData = new SheetData()

        worksheetPart.Worksheet.AppendChild(sheetData) |> ignore
        worksheetPart.Worksheet.Save()

        let sheets = workbookPart.Workbook.AppendChild(new Sheets())

        let sheet =
            new DocumentFormat.OpenXml.Spreadsheet.Sheet(
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1u,
                Name = new StringValue(defaultArg sheetName "Лист 1")
            )

        sheets.AppendChild(sheet) |> ignore

        workbookPart.Workbook.Save()
        document.Dispose()

        memoryStream.Seek(0, SeekOrigin.Begin) |> ignore

        new Spreadsheet(memoryStream)

    /// Opens .xlsx spreadsheet from file.
    static member FromFile(fileName: string) =
        new Spreadsheet(new FileStream(fileName, FileMode.Open))

    /// Opens .xlsx spreadsheet from memory.
    static member FromByteArray(data: byte array) = new Spreadsheet(new MemoryStream(data))

    /// Returns all sheets in a spreadsheet.
    member _.Sheets() : Sheet seq = sheets.Values

    /// Returns a sheet (tab in a spreadsheet) by name.
    member _.Sheet(sheetName: string) : Sheet = 
        if sheets.ContainsKey sheetName then
            sheets[sheetName]
        else
            failwithf "Spreadsheet does not contain sheet %s" sheetName

    /// Saves entire spreadsheet to a given stream.
    member _.SaveTo(stream: Stream) =
        task {
            let clone = document.Clone(stream)
            clone.Dispose()
        }

    /// Saves entire spreadsheet to a given file.
    member this.SaveTo(path: string) =
        task {
            use stream = new FileStream(path, FileMode.Create)
            do! this.SaveTo stream
        }

    interface IDisposable with
        member _.Dispose() = dataStream.Dispose()
