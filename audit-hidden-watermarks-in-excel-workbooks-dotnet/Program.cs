// Topic: Find, watermark and clear the four places a mark hides in an Excel workbook.
// Uses GroupDocs.Watermark for .NET. A watermark in .xlsx is never one thing: it may be
// header/footer text, a header picture, a worksheet background, a shape on the sheet, or
// something inside an embedded attachment. Each lives in a different part of the object
// model, so each needs its own pass.

using GroupDocs.Watermark;
using GroupDocs.Watermark.Common;
using GroupDocs.Watermark.Contents.Spreadsheet;
using GroupDocs.Watermark.Options.Spreadsheet;
using GroupDocs.Watermark.Watermarks;

namespace Demo.FindWatermarksInExcel;

internal static class Program
{
    private const string DocsFolder = "documents";
    private const string ResultFolder = "Result";

    private static readonly string SourceWorkbook = Path.Combine(DocsFolder, "workbook.xlsx");

    private static int Main()
    {
        Directory.CreateDirectory(DocsFolder);
        Directory.CreateDirectory(ResultFolder);
        ApplyLicense();

        if (!File.Exists(SourceWorkbook))
        {
            Console.Error.WriteLine($"Missing source workbook: {Path.GetFullPath(SourceWorkbook)}");
            return 1;
        }

        InventoryWorkbook(SourceWorkbook);

        string withAttachment = Path.Combine(ResultFolder, "with-attachment.xlsx");
        AddAttachment(SourceWorkbook, withAttachment, "distribution-terms.txt");
        Console.WriteLine($"Attachment embedded: {Path.GetFullPath(withAttachment)}");

        int attachments = ListAttachments(withAttachment);
        Console.WriteLine($"Attachments found after embedding: {attachments}");

        string headerStamped = Path.Combine(ResultFolder, "header-images-stamped.xlsx");
        int headerImages = WatermarkHeaderFooterImages(SourceWorkbook, headerStamped, "Protected");
        Console.WriteLine($"Header and footer images watermarked: {headerImages}");

        string backgroundStamped = Path.Combine(ResultFolder, "backgrounds-stamped.xlsx");
        int backgrounds = WatermarkWorksheetBackgrounds(SourceWorkbook, backgroundStamped, "Protected");
        Console.WriteLine($"Worksheet backgrounds watermarked: {backgrounds}");

        string shapesStamped = Path.Combine(ResultFolder, "shape-images-stamped.xlsx");
        int shapeImages = WatermarkShapeImages(SourceWorkbook, shapesStamped, "Protected");
        Console.WriteLine($"Shape images watermarked: {shapeImages}");

        string cleaned = Path.Combine(ResultFolder, "headers-cleared.xlsx");
        int cleared = ClearHeaderFooterText(SourceWorkbook, cleaned);
        Console.WriteLine($"Header and footer sections cleared: {cleared}");

        string noBackground = Path.Combine(ResultFolder, "background-removed.xlsx");
        int removed = RemoveWorksheetBackgrounds(SourceWorkbook, noBackground);
        Console.WriteLine($"Worksheet backgrounds removed: {removed}");

        return 0;
    }

    private static void ApplyLicense()
    {
        // Point this at your .lic file to remove evaluation limits.
        // Get a free temporary licence: https://purchase.groupdocs.com/temporary-license
        const string licensePath = "REPLACE_WITH_YOUR_LICENSE_PATH";
        if (File.Exists(licensePath))
        {
            new License().SetLicense(licensePath);
            Console.WriteLine("[license] applied");
        }
        else
        {
            Console.WriteLine("[license] no licence set - running in evaluation mode");
        }
    }

    /// <summary>
    /// Reports every place in a workbook where a watermark can hide.
    /// </summary>
    /// <remarks>
    /// Walks <see cref="SpreadsheetContent.Worksheets"/> and counts the four carriers
    /// separately: header and footer sections (text in <c>Script</c>, pictures in
    /// <c>Image</c>), the worksheet <c>BackgroundImage</c>, the shapes on the sheet, and
    /// embedded attachments. Run this before changing anything - a workbook that looks
    /// unmarked on screen frequently carries a logo in a header or a tiled background
    /// that only appears in print preview.
    /// </remarks>
    private static void InventoryWorkbook(string sourcePath)
    {
        using var watermarker = new Watermarker(sourcePath, new SpreadsheetLoadOptions());
        SpreadsheetContent content = watermarker.GetContent<SpreadsheetContent>();

        for (int i = 0; i < content.Worksheets.Count; i++)
        {
            SpreadsheetWorksheet worksheet = content.Worksheets[i];

            int headerTexts = 0;
            int headerImages = 0;
            foreach (SpreadsheetHeaderFooter headerFooter in worksheet.HeadersFooters)
            {
                foreach (SpreadsheetHeaderFooterSection section in headerFooter.Sections)
                {
                    if (!string.IsNullOrEmpty(section.Script))
                    {
                        headerTexts++;
                    }

                    if (section.Image != null)
                    {
                        headerImages++;
                    }
                }
            }

            Console.WriteLine($"Worksheet {i}: header texts {headerTexts}, header images "
                + $"{headerImages}, background {(worksheet.BackgroundImage != null ? "yes" : "no")}, "
                + $"shapes {worksheet.Shapes.Count}, attachments {worksheet.Attachments.Count}");
        }
    }

    /// <summary>
    /// Embeds a file into the first worksheet as an OLE attachment.
    /// </summary>
    /// <remarks>
    /// <c>Attachments.AddAttachment</c> takes the file bytes, a source file name whose
    /// extension decides which application opens it, an optional preview image, and the
    /// frame rectangle in points. Attachments matter here because a watermark applied to
    /// the workbook does not reach inside them - the attached file keeps whatever marking
    /// it already had, which is exactly how an unmarked copy leaves a marked workbook.
    /// </remarks>
    private static void AddAttachment(string sourcePath, string outputPath, string attachmentName)
    {
        using var watermarker = new Watermarker(sourcePath, new SpreadsheetLoadOptions());
        SpreadsheetContent content = watermarker.GetContent<SpreadsheetContent>();

        byte[] payload = System.Text.Encoding.UTF8.GetBytes(
            "Distribution terms: internal use only. Do not forward.");

        content.Worksheets[0].Attachments.AddAttachment(
            payload,
            attachmentName,
            null,
            50,
            100,
            200,
            100);

        watermarker.Save(outputPath);
    }

    /// <summary>
    /// Lists the attachments embedded in a workbook and returns how many were found.
    /// </summary>
    /// <remarks>
    /// Each <see cref="SpreadsheetAttachment"/> exposes the originating file name and the
    /// raw <c>Content</c> bytes, so an audit can report what travels inside the workbook
    /// without extracting anything to disk. A non-zero count is the signal that marking
    /// the workbook alone is not enough.
    /// </remarks>
    private static int ListAttachments(string sourcePath)
    {
        using var watermarker = new Watermarker(sourcePath, new SpreadsheetLoadOptions());
        SpreadsheetContent content = watermarker.GetContent<SpreadsheetContent>();

        int found = 0;
        foreach (SpreadsheetWorksheet worksheet in content.Worksheets)
        {
            foreach (SpreadsheetAttachment attachment in worksheet.Attachments)
            {
                Console.WriteLine($"  attachment: {attachment.SourceFullName} "
                    + $"({attachment.Content.Length} bytes)");
                found++;
            }
        }

        return found;
    }

    /// <summary>
    /// Stamps a watermark into every picture used in a header or footer.
    /// </summary>
    /// <remarks>
    /// A header picture is a company logo more often than not, and it prints on every
    /// page while being invisible in the normal editing view. The image is reached through
    /// <see cref="SpreadsheetHeaderFooterSection.Image"/>, and calling <c>Add</c> on that
    /// <see cref="SpreadsheetWatermarkableImage"/> rasterises the mark into the picture
    /// data. Returns the number of images stamped.
    /// </remarks>
    private static int WatermarkHeaderFooterImages(string sourcePath, string outputPath, string text)
    {
        using var watermarker = new Watermarker(sourcePath, new SpreadsheetLoadOptions());
        SpreadsheetContent content = watermarker.GetContent<SpreadsheetContent>();

        var watermark = new TextWatermark(text, new Font("Arial", 8))
        {
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            RotateAngle = 45,
            SizingType = SizingType.ScaleToParentDimensions,
            ScaleFactor = 1,
        };

        int stamped = 0;
        foreach (SpreadsheetWorksheet worksheet in content.Worksheets)
        {
            foreach (SpreadsheetHeaderFooter headerFooter in worksheet.HeadersFooters)
            {
                foreach (SpreadsheetHeaderFooterSection section in headerFooter.Sections)
                {
                    if (section.Image != null)
                    {
                        section.Image.Add(watermark);
                        stamped++;
                    }
                }
            }
        }

        watermarker.Save(outputPath);
        return stamped;
    }

    /// <summary>
    /// Stamps a watermark into each worksheet background image.
    /// </summary>
    /// <remarks>
    /// A worksheet background is a tiled picture behind the cells, set through Page Layout
    /// in Excel. It is a frequent carrier for an "INTERNAL" or "DRAFT" pattern and, unlike
    /// a shape, cannot be clicked and deleted by a reader. <c>BackgroundImage</c> is null
    /// on sheets without one, so the null check is what keeps the pass honest. Returns the
    /// number of backgrounds marked.
    /// </remarks>
    private static int WatermarkWorksheetBackgrounds(string sourcePath, string outputPath, string text)
    {
        using var watermarker = new Watermarker(sourcePath, new SpreadsheetLoadOptions());
        SpreadsheetContent content = watermarker.GetContent<SpreadsheetContent>();

        var watermark = new TextWatermark(text, new Font("Arial", 12))
        {
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            RotateAngle = 45,
            SizingType = SizingType.ScaleToParentDimensions,
            ScaleFactor = 1,
        };

        int stamped = 0;
        foreach (SpreadsheetWorksheet worksheet in content.Worksheets)
        {
            if (worksheet.BackgroundImage != null)
            {
                worksheet.BackgroundImage.Add(watermark);
                stamped++;
            }
        }

        watermarker.Save(outputPath);
        return stamped;
    }

    /// <summary>
    /// Stamps a watermark into the pictures carried by shapes on the sheets.
    /// </summary>
    /// <remarks>
    /// Shapes are the visible layer: logos, callouts and pasted charts all arrive as
    /// <see cref="SpreadsheetShape"/> entries with an <c>Image</c> when they carry a
    /// picture. Marking the image rather than the sheet means the mark survives the
    /// picture being copied into another document, where deleting a shape would not help.
    /// Returns how many shape images were stamped.
    /// </remarks>
    private static int WatermarkShapeImages(string sourcePath, string outputPath, string text)
    {
        using var watermarker = new Watermarker(sourcePath, new SpreadsheetLoadOptions());
        SpreadsheetContent content = watermarker.GetContent<SpreadsheetContent>();

        var watermark = new TextWatermark(text, new Font("Arial", 10))
        {
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            RotateAngle = 45,
            SizingType = SizingType.ScaleToParentDimensions,
            ScaleFactor = 1,
        };

        int stamped = 0;
        foreach (SpreadsheetWorksheet worksheet in content.Worksheets)
        {
            foreach (SpreadsheetShape shape in worksheet.Shapes)
            {
                if (shape.Image != null)
                {
                    shape.Image.Add(watermark);
                    stamped++;
                }
            }
        }

        watermarker.Save(outputPath);
        return stamped;
    }

    /// <summary>
    /// Clears the text and pictures from every header and footer section.
    /// </summary>
    /// <remarks>
    /// Setting <c>Script</c> and <c>Image</c> to null empties a section. This is the
    /// cleanup half of an audit: a workbook received from outside may carry another
    /// organisation's confidentiality notice or logo in a header, which should not survive
    /// into your own copy. Counts only sections that actually held something, so the
    /// return value reports real changes rather than iterations.
    /// </remarks>
    private static int ClearHeaderFooterText(string sourcePath, string outputPath)
    {
        using var watermarker = new Watermarker(sourcePath, new SpreadsheetLoadOptions());
        SpreadsheetContent content = watermarker.GetContent<SpreadsheetContent>();

        int cleared = 0;
        foreach (SpreadsheetWorksheet worksheet in content.Worksheets)
        {
            foreach (SpreadsheetHeaderFooter headerFooter in worksheet.HeadersFooters)
            {
                foreach (SpreadsheetHeaderFooterSection section in headerFooter.Sections)
                {
                    if (!string.IsNullOrEmpty(section.Script) || section.Image != null)
                    {
                        section.Script = null;
                        section.Image = null;
                        cleared++;
                    }
                }
            }
        }

        watermarker.Save(outputPath);
        return cleared;
    }

    /// <summary>
    /// Removes the background picture from every worksheet that has one.
    /// </summary>
    /// <remarks>
    /// Assigning null to <see cref="SpreadsheetWorksheet.BackgroundImage"/> drops the
    /// tiled picture entirely. Worth pairing with the header cleanup when preparing a
    /// workbook for external distribution: a background marked INTERNAL is easy to miss on
    /// screen because Excel shows it only behind the used range. Returns the number of
    /// worksheets changed.
    /// </remarks>
    private static int RemoveWorksheetBackgrounds(string sourcePath, string outputPath)
    {
        using var watermarker = new Watermarker(sourcePath, new SpreadsheetLoadOptions());
        SpreadsheetContent content = watermarker.GetContent<SpreadsheetContent>();

        int removed = 0;
        foreach (SpreadsheetWorksheet worksheet in content.Worksheets)
        {
            if (worksheet.BackgroundImage != null)
            {
                worksheet.BackgroundImage = null;
                removed++;
            }
        }

        watermarker.Save(outputPath);
        return removed;
    }
}
