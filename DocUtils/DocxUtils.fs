module LessPainfulDocx

open DocumentFormat.OpenXml.Wordprocessing
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging

type Justification =
    | Center
    | Stretch
    | Left

let createParagraph text (justification: Justification) (bold: bool) (italic: bool) (indent: bool) =
    let text = Text(text)
    let run = Run()

    let runProperties = RunProperties()
    let fonts = 
        RunFonts
            (Ascii = StringValue("Times New Roman"), 
             HighAnsi = StringValue("Times New Roman"), 
             ComplexScript = StringValue("Times New Roman"))

    let size = FontSize(Val = StringValue("28"))
    let sizeCs = FontSizeComplexScript(Val = StringValue("28"))
    
    runProperties.AppendChild fonts |> ignore
    runProperties.AppendChild size |> ignore
    runProperties.AppendChild sizeCs |> ignore

    if bold then 
        runProperties.AppendChild(Bold()) |> ignore
    if italic then 
        runProperties.AppendChild(Italic()) |> ignore

    run.AppendChild(runProperties) |> ignore
    run.AppendChild(text) |> ignore

    let paragraphProperties = ParagraphProperties()
    let justificationValue = 
        match justification with
        | Center -> JustificationValues.Center 
        | Stretch -> JustificationValues.Both
        | Left -> JustificationValues.Left

    if justification <> Justification.Center && indent then
        // Do not indent paragraphs with centering.
        let indent = Indentation(FirstLine = StringValue("720"))
        paragraphProperties.Indentation <- indent
    let justification = Justification(Val = EnumValue<_>(justificationValue))
    paragraphProperties.Justification <- justification
    paragraphProperties.SpacingBetweenLines <- SpacingBetweenLines(After = StringValue("0"))

    let paragraph = Paragraph()
    paragraph.AppendChild paragraphProperties |> ignore
    paragraph.AppendChild run |> ignore
    paragraph

let createNumbering (document: WordprocessingDocument) =
    let numberingDefinitionsPart = document.MainDocumentPart.NumberingDefinitionsPart
    if numberingDefinitionsPart = null then
        document.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>() |> ignore
        document.MainDocumentPart.NumberingDefinitionsPart.Numbering <- Numbering()
    let numberingPart = document.MainDocumentPart.NumberingDefinitionsPart.Numbering
    let lastAbstractNumberingId = 
        if numberingPart.Descendants<AbstractNum>() |> Seq.isEmpty then
            0
        else
            numberingPart.Descendants<AbstractNum>() 
            |> Seq.map (fun an -> an.AbstractNumberId.Value) 
            |> Seq.sortDescending 
            |> Seq.head

    let abstractNum = AbstractNum()
    abstractNum.AbstractNumberId <- Int32Value(lastAbstractNumberingId + 1)
    let lvl0 = Level()
    lvl0.LevelIndex <- Int32Value 0
    lvl0.StartNumberingValue <- StartNumberingValue(Val = Int32Value 1)
    lvl0.NumberingFormat <- NumberingFormat(Val = EnumValue<NumberFormatValues>(NumberFormatValues.Decimal))
    lvl0.LevelText <- LevelText(Val = StringValue "%1.")
    lvl0.LevelJustification <- LevelJustification(Val = EnumValue<LevelJustificationValues>(LevelJustificationValues.Left))

    let lvl0PPr = ParagraphProperties()
    lvl0PPr.Indentation <- Indentation(Left = StringValue("360"), FirstLine = StringValue("0"), Hanging = StringValue("360"))
    lvl0.AppendChild lvl0PPr |> ignore

    let lvl0RunPr = RunProperties()
    let fonts = 
        RunFonts
            (Ascii = StringValue("Times New Roman"), 
             HighAnsi = StringValue("Times New Roman"), 
             ComplexScript = StringValue("Times New Roman"))

    let size = FontSize(Val = StringValue("28"))
    let sizeCs = FontSizeComplexScript(Val = StringValue("28"))
    
    lvl0RunPr.AppendChild fonts |> ignore
    lvl0RunPr.AppendChild size |> ignore
    lvl0RunPr.AppendChild sizeCs |> ignore
    lvl0.AppendChild lvl0RunPr |> ignore

    abstractNum.AppendChild lvl0 |> ignore

    if numberingPart.Descendants<AbstractNum>() |> Seq.isEmpty then
        numberingPart.AppendChild abstractNum |> ignore
    else
        numberingPart.Descendants<AbstractNum>() |> Seq.last |> (fun n -> n.InsertAfterSelf abstractNum |> ignore)

    let num = NumberingInstance()
    let lastNumberingId = 
        if numberingPart.Descendants<NumberingInstance>() |> Seq.isEmpty then
            0
        else
            numberingPart.Descendants<NumberingInstance>() 
            |> Seq.map (fun n -> n.NumberID.Value) 
            |> Seq.sortDescending
            |> Seq.head

    num.NumberID <- Int32Value(lastNumberingId + 1)

    let abstractNumberingRef = AbstractNumId(Val = Int32Value(lastAbstractNumberingId + 1))
    num.AppendChild abstractNumberingRef |> ignore

    numberingPart.AppendChild num |> ignore

    num.NumberID

let createHyperlink link (document: WordprocessingDocument) =
    let text = Text(link)
    let run = Run()

    let runProperties = RunProperties()
    let fonts = 
        RunFonts
            (Ascii = StringValue("Times New Roman"), 
             HighAnsi = StringValue("Times New Roman"), 
             ComplexScript = StringValue("Times New Roman"))

    let size = FontSize(Val = StringValue("24"))
    let sizeCs = FontSizeComplexScript(Val = StringValue("24"))
    let underline = Underline(Val = EnumValue<_>(UnderlineValues.Single))
    let color = Color(ThemeColor = EnumValue<_>(ThemeColorValues.Hyperlink))

    runProperties.AppendChild fonts |> ignore
    runProperties.AppendChild size |> ignore
    runProperties.AppendChild sizeCs |> ignore
    runProperties.AppendChild underline|> ignore
    runProperties.AppendChild color|> ignore

    run.AppendChild(runProperties) |> ignore
    run.AppendChild(text) |> ignore

    let paragraphProperties = ParagraphProperties()

    let indent = Indentation(Left = StringValue("360"))
    paragraphProperties.Indentation <- indent
    let justification = Justification(Val = EnumValue<_>(JustificationValues.Left))
    paragraphProperties.Justification <- justification
    paragraphProperties.SpacingBetweenLines <- SpacingBetweenLines(After = StringValue("120"))

    let linkParagraph = Paragraph()
    linkParagraph.AppendChild paragraphProperties |> ignore

    let rel = document.MainDocumentPart.AddHyperlinkRelationship(System.Uri(link), true)

    let hyperlink = Hyperlink()
    hyperlink.AppendChild run |> ignore
    hyperlink.Id <- StringValue(rel.Id)
    hyperlink.History <- OnOffValue.FromBoolean(true)            

    linkParagraph.AppendChild hyperlink |> ignore
    linkParagraph

let findParagraph (body: Body) (text: string) =
    let paragraphs = body.Descendants<Paragraph>()
    paragraphs |> Seq.tryFind (fun p -> p.InnerText.Contains(text))

let findPreviousParagraph (body: Body) (paragraph: Paragraph) =
    body.Descendants<Paragraph>() 
    |> Seq.pairwise 
    |> Seq.find (fun (_, next) -> next = paragraph) 
    |> fst

let findNextParagraph (body: Body) (paragraph: Paragraph) =
    body.Descendants<Paragraph>() 
    |> Seq.pairwise 
    |> Seq.find (fun (prev, _) -> prev = paragraph) 
    |> snd

let rec removeNextParagraphIfEmpty (body: Body) (paragraph: Paragraph) (recursive: bool) =
    let nextParagraph = findNextParagraph body paragraph

    if nextParagraph.InnerText = "" then
        nextParagraph.Remove()
        if recursive then
            removeNextParagraphIfEmpty body paragraph recursive

let rec removePreviousParagraphIfEmpty (body: Body) (paragraph: Paragraph) (recursive: bool) =
    let previousParagraph = findPreviousParagraph body paragraph

    if previousParagraph.InnerText = "" then
        previousParagraph.Remove()
        if recursive then
            removePreviousParagraphIfEmpty body paragraph recursive

let fixTextEverywhere (body: Body) (from: string) (replaceTo: string) =
    body.Descendants<Text>()
    |> Seq.iter (fun t -> t.Text <- t.Text.Replace(from, replaceTo))
