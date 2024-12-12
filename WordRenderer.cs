using GemBox.Document;
using GemBox.Document.Drawing;
using GemBox.Document.Tables;
using GemBoxTableColumn = GemBox.Document.Tables.TableColumn;

public partial class WordRenderer
{
    private readonly DocumentModel _doc = new();
    private readonly ParagraphStyle _h2;
    private readonly ParagraphStyle _h3;
    private readonly ParagraphStyle _h4;
    private readonly ParagraphStyle _h5;
    private readonly ParagraphStyle _h6;
    private readonly Color _lightGrey = new(240, 240, 240);
    private readonly Table _headerTable;
    private readonly HeaderFooter _footer;

    public const string WORD_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    public const string PDF_CONTENT_TYPE = "application/pdf";
    
    public WordRenderer() {

        var baseFontSize = 11;
        
        FontSettings.FontsBaseDirectory = ".";
        _doc.DefaultCharacterFormat.FontName = "Montserrat";
        _doc.DefaultCharacterFormat.Size = baseFontSize;

        var section = new Section(_doc);
        _doc.Sections.Add(section);
        section.PageSetup.PaperType = PaperType.A4;

        _h2 = new ParagraphStyle("H2") { CharacterFormat = new CharacterFormat { Size = baseFontSize * 2.00 } };
        _doc.Styles.Add(_h2);
        _h3 = new ParagraphStyle("H3") {  CharacterFormat = new CharacterFormat { Size = baseFontSize * 1.75 } };
        _doc.Styles.Add(_h3);
        _h4 = new ParagraphStyle("H4") { CharacterFormat = new CharacterFormat { Size = baseFontSize * 1.5, Bold = true } };
        _doc.Styles.Add(_h4);
        _h5 = new ParagraphStyle("H5") { CharacterFormat = new CharacterFormat { Size = baseFontSize * 1.25, Bold = true } };
        _doc.Styles.Add(_h5);
        _h6 = new ParagraphStyle("H6") { CharacterFormat = new CharacterFormat { Size = baseFontSize * 1.1, Bold = true } };
        _doc.Styles.Add(_h6);

        _headerTable = new Table(_doc);
        _headerTable.TableFormat.AutomaticallyResizeToFitContents = true;
        _headerTable.TableFormat.PreferredWidth = new TableWidth(ParagraphWidth(section), TableWidthUnit.Point);
        _headerTable.TableFormat.Borders.SetBorders(MultipleBorderTypes.All, BorderStyle.None, Color.Empty, 0);
        _headerTable.Rows.Add(new TableRow(_doc));
        var headerTableParagraphStyle = new TableStyle("Header table paragraph style") { ParagraphFormat = new ParagraphFormat { SpaceAfter = 24 } };
        _doc.Styles.Add(headerTableParagraphStyle);
        _headerTable.TableFormat.Style = headerTableParagraphStyle;
        var header = new HeaderFooter(_doc, HeaderFooterType.HeaderDefault);
        header.Blocks.Add(_headerTable);
        section.HeadersFooters.Add(header);

        _footer = new HeaderFooter(_doc, HeaderFooterType.FooterDefault);
        section.HeadersFooters.Add(_footer);
        var footerParagraph = new Paragraph(
            _doc,
            new Field(_doc, FieldType.Page),
            new Run(_doc, " of "),
            new Field(_doc, FieldType.NumPages)
        );
        footerParagraph.ParagraphFormat.Alignment = HorizontalAlignment.Right;
        _footer.Blocks.Add(footerParagraph);

    }

    public void AddContentHeading2(string text) => AddContentHeading(text, _h2);

    public void AddContentHeading3(string text) => AddContentHeading(text, _h3);

    public void AddContentHeading4(string text) => AddContentHeading(text, _h4);

    public void AddContentHeading5(string text) => AddContentHeading(text, _h5);

    public void AddContentHeading6(string text) => AddContentHeading(text, _h6);

    private void AddContentHeading(string text, ParagraphStyle style)
    {
        var p = new Paragraph(_doc, text);
        p.ParagraphFormat.Style = style;
        p.ParagraphFormat.KeepWithNext = true;
        _doc.Sections[0].Blocks.Add(p);
    }

    public void AddContentImage(string name, string contentType, byte[] bytes)
    {
        var p = new Paragraph(_doc);
        var picture = new Picture(_doc, new MemoryStream(bytes));
        SizeToParagraph(picture);
        p.Inlines.Add(picture);
        _doc.Sections[0].Blocks.Add(p);
    }


    public void AddContentSectionBreak()
    {
        var p = new Paragraph(_doc);
        var line = new Shape(_doc, ShapeType.Rectangle,
            Layout.Inline(new Size(ParagraphWidth(_doc.Sections[0]), 0.1))
        );
        line.Fill.SetSolid(Color.LightGray);
        line.Outline.Fill.SetEmpty();
        line.Layout!.EffectPadding = new Padding(0,10,LengthUnit.Point);
        p.Inlines.Add(line);
        p.ParagraphFormat.KeepWithNext = true;
        _doc.Sections[0].Blocks.Add(p);
    }

    public void AddContentTable(TableColumn[] tableColumns, bool UseBandedRows = false)
    {
        var firstColumn = tableColumns.FirstOrDefault();
        if(firstColumn == null) return;
        var table = new Table(_doc);
        table.TableFormat.Borders.SetBorders(MultipleBorderTypes.All, BorderStyle.None, Color.Empty, 0);

        var fixedWidths = tableColumns.Where(x => !x.IsRelative).Sum(x => x.Width);
        var totalRelativeWidths = tableColumns.Where(x => x.IsRelative).Sum(x => x.Width);

        var paragraphWidth = ParagraphWidth(_doc.Sections[0]);
        var relativeWidth = paragraphWidth - fixedWidths;
        var ratio = relativeWidth / totalRelativeWidths;

        table.TableFormat.AutomaticallyResizeToFitContents = false;
        foreach(var col in tableColumns) {
            if(col.IsRelative)
                table.Columns.Add(new GemBoxTableColumn((double)col.Width * ratio));
            else
                table.Columns.Add(new GemBoxTableColumn((double)col.Width!));
        }
        for(var i = 0; i < firstColumn.Rows.Length; i++)
        {
            var item = new TableRow(_doc, tableColumns.Select(c => TextCell(c.Rows[i])));
            if(UseBandedRows && i % 2 == 1) {
                foreach(var cell in item.Cells) {
                    cell.CellFormat.BackgroundColor = _lightGrey;
                }
            }
            table.Rows.Add(item);
        }
        _doc.Sections[0].Blocks.Add(table);
        _doc.Sections[0].Blocks.Add(new Paragraph(_doc));
    }

    public void AddContentText(ContentText[] texts)
    {
        var p = new Paragraph(_doc);
        foreach(var text in texts) {
            AddSafeInlineParagraphText(text.Text, text.IsBold, text.linkTarget, p);
        }
        _doc.Sections[0].Blocks.Add(p);
    }

    public void AddHeaderTitleImage(int width, byte[] imageData)
    {
        var picture = new Picture(_doc, new MemoryStream(imageData));
        var ratio = width / picture.Layout!.Size.Width;
        picture.Layout.Size = new Size(picture.Layout.Size.Width * ratio, picture.Layout.Size.Height * ratio);
        var cell = new TableCell(_doc, new Paragraph(_doc, picture));
        _headerTable.Rows[0].Cells.Add(cell);
    }

    public void AddHeaderSubTitleText(string text)
    {
        var p = new Paragraph(_doc, text);
        p.ParagraphFormat.Alignment = HorizontalAlignment.Right;
        p.ParagraphFormat.Style = _h3;
        var cell = new TableCell(_doc, p);
        _headerTable.Rows[0].Cells.Add(cell);
    }

    internal MemoryStream GenerateDoc()
    {
        var ret = new MemoryStream();
        _doc.Save(ret, SaveOptions.DocxDefault);
        ret.Position = 0;
        return ret;
    }

    public MemoryStream GenerateDocAsPDF()
    {
        var ret = new MemoryStream();
        _doc.Save(ret, SaveOptions.PdfDefault);
        ret.Position = 0;
        return ret;
    }

    private static double ParagraphWidth(Section section)
    {
        var pageSetup = section.PageSetup;
        return pageSetup.PageWidth - pageSetup.PageMargins.Left - pageSetup.PageMargins.Right;
    }

    private TableCell TextCell(TableContentText tableContentText)
    {
        var cell = new TableCell(_doc, [
            SafeParagraphText(
                tableContentText.Text,
                tableContentText.IsBold || tableContentText.IsBoldAndShaded
            )
        ]);
        if (tableContentText.IsBoldAndShaded)
        {
            cell.CellFormat.BackgroundColor = _lightGrey;
        }
        cell.CellFormat.Padding = new Padding(4, LengthUnit.Point);
        return cell;
    }

    private Paragraph SafeParagraphText(string text, bool isBold)
    {
        var p = new Paragraph(_doc);
        AddSafeInlineParagraphText(text, isBold, "", p);

        return p;
    }

    private void AddSafeInlineParagraphText(string text, bool isBold, string linkTarget, Paragraph p)
    {
        var bits = text.Split("\n");
        var lastBit = bits[^1];
        foreach (var bit in bits)
        {
            Inline item;
            if (linkTarget != "")
            {
                item = new Hyperlink(_doc, linkTarget, bit);
            }
            else
            {
                var r = new Run(_doc, bit);
                r.CharacterFormat.Bold = isBold;
                item = r;
            }
            p.Inlines.Add(item);
            if (bit != lastBit)
                p.Inlines.Add(new SpecialCharacter(_doc, SpecialCharacterType.LineBreak));
        }
    }

    private void SizeToParagraph(Picture picture)
    {
        var paragraphWidth = ParagraphWidth(_doc.Sections[0]);
        var originalPictureSize = picture.Layout!.Size;
        var ratio = paragraphWidth / originalPictureSize.Width;
        picture.Layout.Size = new Size(picture.Layout.Size.Width * ratio, picture.Layout.Size.Height * ratio);
    }

    public IDisposable HeaderSection()
    {
        // no-op
        return new CompleteOnDispose(() => {});
    }

    class CompleteOnDispose(Action completion) : IDisposable
    {
        public void Dispose() => completion();
    }
}