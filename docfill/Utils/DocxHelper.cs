using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;




namespace docfill.Utils;

/// <summary>
/// Helper methods for manipulating DOCX documents (text replacement, table filling).
/// </summary>
public static class DocxHelper
{
    /// <summary>
    /// Sets the text content of a table cell.
    /// </summary>
    private static void SetCellText(TableCell cell, string value)
    {
        cell.RemoveAllChildren<Paragraph>();

        var p = new Paragraph();
        var r = new Run();
        var t = new Text(value ?? string.Empty);

        r.Append(t);
        p.Append(r);
        cell.Append(p);
    }

    /// <summary>
    /// Replaces text placeholders in the document body.
    /// </summary>
    public static void ReplaceText(MainDocumentPart main, Dictionary<string, string> replacements)
    {
        if (main?.Document?.Body == null)
            throw new InvalidOperationException("Document body not found.");

        // Replace in body
        ReplaceTextInElement(main.Document.Body, replacements);

        // Replace in headers
        foreach (var headerPart in main.HeaderParts)
        {
            if (headerPart.Header != null)
                ReplaceTextInElement(headerPart.Header, replacements);
        }

        // Replace in footers
        foreach (var footerPart in main.FooterParts)
        {
            if (footerPart.Footer != null)
                ReplaceTextInElement(footerPart.Footer, replacements);
        }

        main.Document.Save();
    }

    // Shared internal method for any OpenXmlElement (Body, Header, Footer)
    private static void ReplaceTextInElement(OpenXmlElement root, Dictionary<string, string> replacements)
    {
        foreach (var text in root.Descendants<Text>())
        {
            foreach (var kvp in replacements)
            {
                if (text.Text.Contains(kvp.Key))
                {
                    text.Text = text.Text.Replace(kvp.Key, kvp.Value ?? string.Empty);
                }
            }
        }
    }


    /// <summary>
    /// Finds a table by its alternative text (caption).
    /// </summary>
    public static Table? FindTable(string altText, MainDocumentPart main)
    {
        if (main?.Document?.Body == null)
            throw new InvalidOperationException("Document body not found.");
        Body body = main.Document.Body;

        var targetTable = body.Descendants<Table>()
            .FirstOrDefault(t =>
            {
                var props = t.GetFirstChild<TableProperties>();
                if (props == null) return false;

                var title = props.GetFirstChild<TableCaption>()?.Val;
                return title != null && title.Value == altText;
            });

        if (targetTable is null)
        {
            throw new InvalidOperationException($"Table with alt text '{altText}' not found.");
        }

        return targetTable;
    }

    /// <summary>
    /// Fills a table with data, adjusting row count as needed.
    /// </summary>
    public static void FillTable(MainDocumentPart main, string tableName, List<List<string>> data)
    {
        if (main?.Document?.Body == null)
            throw new InvalidOperationException("Document body not found.");
        Body body = main.Document.Body;

        var targetTable = FindTable(tableName, main);
        if (targetTable == null)
        {
            throw new InvalidOperationException($"Table with alt text '{tableName}' not found.");
        }

        var rows = targetTable.Elements<TableRow>().ToList();
        if (rows.Count == 0)
            throw new InvalidOperationException("Table has no rows.");

        var templateRow = rows[0];
        int existingDataRows = rows.Count - 1;
        int requiredDataRows = data.Count;

        // Remove excess rows
        if (existingDataRows > requiredDataRows)
        {
            for (int i = existingDataRows; i > requiredDataRows; i--)
            {
                rows[i].Remove();
            }
        }
        // Add missing rows
        else if (existingDataRows < requiredDataRows)
        {
            for (int i = existingDataRows; i < requiredDataRows; i++)
            {
                var newRow = (TableRow)templateRow.CloneNode(true);
                targetTable.Append(newRow);
            }
        }

        // Update rows reference after modifications
        rows = targetTable.Elements<TableRow>().ToList();

        // Fill data rows (skip header row at index 0)
        for (int rowIndex = 0; rowIndex < data.Count; rowIndex++)
        {
            var cells = rows[rowIndex + 1].Elements<TableCell>().ToList();
            for (int cellIndex = 0; cellIndex < data[rowIndex].Count && cellIndex < cells.Count; cellIndex++)
            {
                SetCellText(cells[cellIndex], data[rowIndex][cellIndex]);
            }
        }
    }

    /// <summary>
    /// Inserts an inline image into the document body.
    /// Width and height must be provided in EMUs.
    /// 1 cm = 360000 EMU, 1 inch = 914400 EMU.
    /// </summary>
    public static void AddImage(MainDocumentPart main, string placeholder, string imagePath, long widthEmu, long heightEmu)
    {
        if (main?.Document?.Body == null)
            throw new InvalidOperationException("Document body not found.");
        Body body = main.Document.Body;

        var relId = AddImagePart(main, imagePath);
        var drawing = CreateInline(relId, widthEmu, heightEmu);

        foreach (var text in body.Descendants<Text>())
        {
            if (text.Text.Contains(placeholder))
            {
                text.Text = text.Text.Replace(placeholder, string.Empty);

                var run = text.Parent as Run;
                if (run != null)
                    run.Append(drawing);

                break;
            }
        }
    }

    private static string AddImagePart(MainDocumentPart main, string imagePath)
    {
        if (main == null) throw new ArgumentNullException(nameof(main));
        if (!File.Exists(imagePath)) throw new FileNotFoundException(imagePath);

        var ext = Path.GetExtension(imagePath).ToLowerInvariant();
        ImagePartType partType = ext switch
        {
            ".png" => ImagePartType.Png,
            ".jpg" => ImagePartType.Jpeg,
            ".jpeg" => ImagePartType.Jpeg,
            ".gif" => ImagePartType.Gif,
            ".bmp" => ImagePartType.Bmp,
            ".tiff" => ImagePartType.Tiff,
            _ => throw new NotSupportedException($"Unsupported image extension: {ext}")
        };

        var part = main.AddImagePart(partType);
        using (var stream = File.OpenRead(imagePath))
            part.FeedData(stream);

        return main.GetIdOfPart(part);
    }


    public static Drawing CreateFloatingImage(string relationshipId, long widthEmu, long heightEmu)
    {
        var anchor = new WP.Anchor(
            // Required element order according to schema
            new WP.SimplePosition { X = 0, Y = 0 },

            // Horizontal positioning
            new WP.HorizontalPosition(
                new WP.HorizontalAlignment("right")
            )
            {
                RelativeFrom = WP.HorizontalRelativePositionValues.Page
            },

            // Vertical positioning
            new WP.VerticalPosition(
                new WP.PositionOffset("0")
            )
            {
                RelativeFrom = WP.VerticalRelativePositionValues.Page
            },

            new WP.Extent { Cx = widthEmu, Cy = heightEmu },
            new WP.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },

            // No text wrapping
            new WP.WrapNone(),

            // Picture properties
            new WP.DocProperties { Id = 1U, Name = "FloatingImage" },

            // Graphic container
            new A.Graphic(
                new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = 0U, Name = "Picture" },
                            new PIC.NonVisualPictureDrawingProperties()
                        ),
                        new PIC.BlipFill(
                            new A.Blip { Embed = relationshipId },
                            new A.Stretch(new A.FillRectangle())
                        ),
                        new PIC.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0, Y = 0 },
                                new A.Extents { Cx = widthEmu, Cy = heightEmu }
                            ),
                            new A.PresetGeometry(new A.AdjustValueList())
                            {
                                Preset = A.ShapeTypeValues.Rectangle
                            }
                        )
                    )
                )
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                }
            )
        )
        {
            BehindDoc = true,
            Locked = false,
            LayoutInCell = false,
            AllowOverlap = true,
            SimplePos = false,
            RelativeHeight = 0U
        };

        return new Drawing(anchor);
    }




    private static Drawing CreateInline(string relationshipId, long widthEmu, long heightEmu)
    {
        return new Drawing(
            new WP.Inline(
                new WP.Extent { Cx = widthEmu, Cy = heightEmu },
                new WP.DocProperties { Id = 1, Name = "Image" },
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0, Name = "Picture" },
                                new PIC.NonVisualPictureDrawingProperties()
                            ),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle())
                            ),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0, Y = 0 },
                                    new A.Extents { Cx = widthEmu, Cy = heightEmu }
                                ),
                                new A.PresetGeometry(new A.AdjustValueList())
                                { Preset = A.ShapeTypeValues.Rectangle }
                            )
                        )
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                )
            )
            {
                DistanceFromTop = 0,
                DistanceFromBottom = 0,
                DistanceFromLeft = 0,
                DistanceFromRight = 0
            }
        );
    }

}

