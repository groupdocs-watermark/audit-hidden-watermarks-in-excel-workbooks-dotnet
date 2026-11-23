# Audit Hidden Watermarks in Excel Workbooks

[![Product Page](https://img.shields.io/badge/Product%20Page-2865E0?style=for-the-badge&logo=appveyor&logoColor=white)](https://products.groupdocs.com/watermark/net/)
[![Docs](https://img.shields.io/badge/Docs-2865E0?style=for-the-badge&logo=Hugo&logoColor=white)](https://docs.groupdocs.com/watermark/net/)
[![Blog](https://img.shields.io/badge/Blog-2865E0?style=for-the-badge&logo=WordPress&logoColor=white)](https://blog.groupdocs.com/categories/groupdocs.watermark-product-family/)
[![Free Support](https://img.shields.io/badge/Free%20Support-2865E0?style=for-the-badge&logo=Discourse&logoColor=white)](https://forum.groupdocs.com/c/watermark/19)
[![Temporary License](https://img.shields.io/badge/Temporary%20License-2865E0?style=for-the-badge&logo=rocket&logoColor=white)](https://purchase.groupdocs.com/temp-license/100354)

## Overview

A watermark in an `.xlsx` file is not one thing. It can be text in a header
that only appears in print preview, a logo picture bound to that header, a
tiled background behind the cells, a picture shape sitting on the sheet, or a
file embedded as an OLE attachment that carries its own marking. Five places,
five different parts of the object model, and a workbook can use all of them
at once.

This repository works through each of them with GroupDocs.Watermark for .NET:
first an inventory pass that reports what a workbook actually carries, then
stamping marks into the images each location holds, then clearing them again.
Built on .NET 8 with GroupDocs.Watermark 26.9.0, for developers who need to
know what is in a workbook before it leaves the building - or what arrived in
one that came from outside.

## Technology Stack

- **Platform**: .NET 8.0
- **Product**: GroupDocs.Watermark for .NET 26.9.0
- **Language**: C# 12
- **Document format**: Office Open XML workbooks (`.xlsx`)

## Problem Statement

Open a workbook in Excel and most marking is invisible. Header and footer
content shows up in print preview, not in the editing grid. A worksheet
background renders behind the used range and is easy to scroll past. An OLE
attachment is a small icon that nobody clicks. Only shapes are obvious, and
they are the one carrier a reader can delete in a single click.

That asymmetry causes two opposite failures. A workbook goes out believing it
is marked, when the mark lives in a header nobody will read on screen. Or a
workbook arrives from a client, gets reused internally, and quietly carries
their confidentiality notice and logo into your own reporting.

Both need the same first step, and it is not "add a watermark" - it is "list
what is already here". GroupDocs.Watermark exposes each carrier separately, so
that inventory is a short loop rather than an XML-unzipping exercise.

## Solution Overview

The spreadsheet object model mirrors the file format. `Watermarker` loads the
workbook, `GetContent<SpreadsheetContent>()` exposes the worksheets, and each
`SpreadsheetWorksheet` gives direct access to its carriers:

- **Headers and footers**: `HeadersFooters` yields a `SpreadsheetHeaderFooter`
  per type (primary, even, first), each with left, centre and right sections
  holding `Script` text and an optional `Image`.
- **Worksheet background**: `BackgroundImage` is the tiled picture behind the
  cells, or null when the sheet has none.
- **Shapes**: `Shapes` lists everything drawn on the sheet, with `Image` set
  when the shape carries a picture.
- **Attachments**: `Attachments` holds embedded OLE files, each with its
  source name and raw bytes.
- **Marking images in place**: every image above is a
  `SpreadsheetWatermarkableImage`, so `Add` rasterises a watermark into the
  picture data rather than placing a new shape on top.

## Prerequisites

Before running these examples, ensure you have:

- **Runtime** - .NET 8.0 SDK or later
- **Package** - GroupDocs.Watermark 26.9.0 (restored from NuGet)
- **Licence** - a GroupDocs licence or [temporary licence](https://purchase.groupdocs.com/temp-license/100354);
  without one the library runs in evaluation mode and marks the output
- **Input file** - an `.xlsx` workbook; one with all four carriers populated is
  included under `documents/`

## Getting Started

### Installation

**Using Package Manager:**

```bash
dotnet add package GroupDocs.Watermark --version 26.9.0
```

**Manual Installation:**

Clone the repository and run `dotnet restore` in the project folder; the
package reference in the `.csproj` pulls everything else.

### Configuration

1. Open `Program.cs` and find the `ApplyLicense` method.
2. Replace `REPLACE_WITH_YOUR_LICENSE_PATH` with the path to your `.lic` file.
   Without a licence the program still runs and prints a notice.
3. Run `dotnet run`. The inventory prints to the console and each pass writes
   its own `.xlsx` into `Result/`.

## Repository Structure

```
audit-hidden-watermarks-in-excel-workbooks-dotnet/
│
├── FindWatermarksInExcelDemo.csproj
├── Program.cs
└── documents/
    └── workbook.xlsx
```

### File Descriptions

- **FindWatermarksInExcelDemo.csproj** - project descriptor targeting `net8.0`
  with the single GroupDocs.Watermark package reference.
- **Program.cs** - the inventory pass, three watermarking passes, two cleanup
  passes, and attachment handling, plus licence setup.
- **documents/workbook.xlsx** - a two-sheet workbook built for this demo. The
  first sheet carries header text, a header picture, a tiled background and a
  picture shape; the second carries header text only, so the inventory output
  shows a real difference between sheets.

## Code Implementation

### Implementation: Inventory every carrier

The pass to run first. It changes nothing and answers the only question worth
asking up front.

```csharp
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
            if (!string.IsNullOrEmpty(section.Script)) headerTexts++;
            if (section.Image != null) headerImages++;
        }
    }

    Console.WriteLine($"Worksheet {i}: header texts {headerTexts}, ...");
}
```

#### Technical Details

`HeadersFooters` enumerates six entries per worksheet - primary, even and
first, each as header and footer - and every one has three sections. That is
eighteen slots per sheet, which is why a header mark is so easy to miss by
hand and so cheap to find in a loop.

**Key Components:**

- `SpreadsheetContent`: spreadsheet-specific view over the loaded workbook
- `SpreadsheetHeaderFooter.Sections`: left, centre and right slots
- `SpreadsheetHeaderFooterSection.Script`: the header text, `&`-codes included

Output: one line per worksheet with counts for each carrier.

---

### Implementation: Embed an attachment

Attachments are the carrier most audits forget, so the demo creates one to
have something real to find.

```csharp
using var watermarker = new Watermarker(sourcePath, new SpreadsheetLoadOptions());
SpreadsheetContent content = watermarker.GetContent<SpreadsheetContent>();

byte[] payload = System.Text.Encoding.UTF8.GetBytes(
    "Distribution terms: internal use only. Do not forward.");

content.Worksheets[0].Attachments.AddAttachment(
    payload, attachmentName, null, 50, 100, 200, 100);

watermarker.Save(outputPath);
```

#### Technical Details

`AddAttachment` takes the content bytes, a source file name whose extension
decides which application opens it, an optional preview image, and the frame
rectangle in points. The preview argument accepts null, in which case Excel
shows a generic icon.

**Parameters:**

- `payload`: the file content as bytes
- `attachmentName`: source file name, extension included
- `50, 100, 200, 100`: frame X, Y, width and height in points

Output: a workbook with one OLE attachment on the first sheet.

---

### Implementation: List what is embedded

```csharp
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
```

#### Technical Details

`SourceFullName` and `Content` are enough for an audit report: what the file
was called and how big it is, without writing anything to disk. A non-zero
count is the signal that marking the workbook is not sufficient - a watermark
applied to the sheet does not reach inside an attached file.

Output: the number of attachments, with one line each.

---

### Implementation: Watermark header and footer pictures

A header picture is almost always a logo, and it prints on every page.

```csharp
var watermark = new TextWatermark(text, new Font("Arial", 8))
{
    HorizontalAlignment = HorizontalAlignment.Center,
    VerticalAlignment = VerticalAlignment.Center,
    RotateAngle = 45,
    SizingType = SizingType.ScaleToParentDimensions,
    ScaleFactor = 1,
};

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
```

#### Technical Details

`section.Image` is a `SpreadsheetWatermarkableImage`, so `Add` writes into the
picture data. `ScaleToParentDimensions` with `ScaleFactor = 1` scales against
the picture, not the page - a header logo is small, and a fixed point size
would either vanish or overflow.

Output: the number of header and footer images stamped.

---

### Implementation: Watermark worksheet backgrounds

```csharp
foreach (SpreadsheetWorksheet worksheet in content.Worksheets)
{
    if (worksheet.BackgroundImage != null)
    {
        worksheet.BackgroundImage.Add(watermark);
        stamped++;
    }
}

watermarker.Save(outputPath);
```

#### Technical Details

A worksheet background is set through Page Layout in Excel and tiles behind
the cells. Unlike a shape it cannot be selected and deleted by a reader, which
makes it a genuinely sturdy carrier - and correspondingly easy to leave behind
by accident. The null check matters: most sheets have no background, and
`BackgroundImage` is null for those.

Output: the number of backgrounds marked.

---

### Implementation: Watermark shape pictures

```csharp
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
```

#### Technical Details

Shapes cover logos, callouts and pasted charts. Marking `shape.Image` burns
the watermark into the picture, so it survives the picture being copied into
another document - a shape-level mark would not. Shapes without a picture
(text boxes, autoshapes) have a null `Image` and are skipped.

Output: the number of shape images stamped.

---

### Implementation: Clear headers and footers

The cleanup half. Useful when a workbook arrives from outside carrying someone
else's notice.

```csharp
foreach (SpreadsheetHeaderFooterSection section in headerFooter.Sections)
{
    if (!string.IsNullOrEmpty(section.Script) || section.Image != null)
    {
        section.Script = null;
        section.Image = null;
        cleared++;
    }
}
```

#### Technical Details

Assigning null to both properties empties a section. The guard is what makes
the return value meaningful - without it the method would report eighteen
"cleared" sections per sheet regardless of what was there.

Output: the number of sections that actually held something.

---

### Implementation: Remove worksheet backgrounds

```csharp
foreach (SpreadsheetWorksheet worksheet in content.Worksheets)
{
    if (worksheet.BackgroundImage != null)
    {
        worksheet.BackgroundImage = null;
        removed++;
    }
}

watermarker.Save(outputPath);
```

#### Technical Details

Assigning null drops the tiled picture entirely. Pair this with the header
cleanup when preparing a workbook for external distribution; between them they
cover the two carriers a reader is least likely to notice.

Output: the number of worksheets changed.

## Best Practices

- **Inventory before you mark.** The counts tell you whether the workbook is
  already marked, and where - which is often not where you assumed.
- **Treat attachments separately.** A watermark on the workbook never reaches
  inside an embedded file. Check the attachment count explicitly.
- **Expect eighteen header slots per sheet.** Six header/footer types times
  three sections; iterate rather than checking the primary header alone.
- **Null-check every carrier.** `BackgroundImage`, `section.Image` and
  `shape.Image` are null far more often than not.
- **Return counts, not void.** "Stamped zero images" and "worked" look
  identical on disk; a count distinguishes them.
- **Mark the image, not the sheet,** when the mark has to survive the picture
  being copied elsewhere.

## Additional Resources

- [Working with worksheet headers and footers](https://docs.groupdocs.com/watermark/net/working-with-worksheet-headers-and-footers/)
- [Working with worksheet backgrounds](https://docs.groupdocs.com/watermark/net/working-with-worksheet-backgrounds/)
- [Shapes in spreadsheet document](https://docs.groupdocs.com/watermark/net/shapes-in-spreadsheet-document/)
- [API Reference](https://reference.groupdocs.com/watermark/net/)
- [Release Notes](https://releases.groupdocs.com/watermark/net/release-notes/latest/)

## Keywords

`excel watermark`, `xlsx watermark`, `watermark excel c#`, `groupdocs
watermark`, `groupdocs.watermark for .net`, `worksheet background`, `header
footer image`, `spreadsheet shapes`, `ole attachment`, `embedded attachment`,
`remove watermark excel`, `clear header footer`, `document audit`, `workbook
inventory`, `spreadsheetcontent`, `spreadsheetworksheet`,
`watermarkableimage`, `document security`, `data room`, `csharp`, `dotnet`,
`net8`, `office open xml`, `document watermarking api`

## Support

For technical support, visit:

- [Free Support Forum](https://forum.groupdocs.com/c/watermark/19)
- [Product Documentation](https://docs.groupdocs.com/watermark/net/)
- [Get Temporary License](https://purchase.groupdocs.com/temp-license/100354)
