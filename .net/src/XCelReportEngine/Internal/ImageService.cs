using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XCelReportEngine.Internal
{
    internal static class ImageService
    {
        private const long EmusPerPixel = 9525;
        private const double DefaultColumnWidthCharacters = 8.43;
        private const double DefaultRowHeightPoints = 15.0;

        internal static int Append(
            WorkbookSession session,
            string imagePath,
            int rowIndex,
            int columnIndex,
            string cellAddress,
            string operation)
        {
            if (!string.IsNullOrWhiteSpace(cellAddress))
            {
                CellAddress.ToCoordinates(cellAddress, operation, out rowIndex, out columnIndex);
            }
            else
            {
                CellAddress.FromZeroBased(rowIndex, columnIndex, operation);
            }

            var fullPath = ValidateImagePath(imagePath, operation);
            var imageContentType = GetImageContentType(fullPath, operation);
            byte[] content;
            int widthPixels;
            int heightPixels;
            try
            {
                content = File.ReadAllBytes(fullPath);
                using (var stream = new MemoryStream(content, false))
                using (var image = Image.FromStream(stream, false, true))
                {
                    widthPixels = image.Width;
                    heightPixels = image.Height;
                }
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException)
            {
                throw new ReportEngineException(ReportErrorCode.InvalidImage, operation, $"Unable to read image: {fullPath}", ex);
            }

            if (widthPixels <= 0 || heightPixels <= 0)
            {
                throw new ReportEngineException(ReportErrorCode.InvalidImage, operation, $"Image has invalid dimensions: {fullPath}");
            }

            var worksheetPart = session.ActiveWorksheetPart;
            var worksheet = worksheetPart.Worksheet
                ?? throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, operation, "The active worksheet root element is missing.");
            var drawingsPart = worksheetPart.DrawingsPart;
            if (drawingsPart == null)
            {
                drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
                var relationshipId = worksheetPart.GetIdOfPart(drawingsPart);
                worksheet.Append(new Drawing { Id = relationshipId });
            }
            else if (drawingsPart.WorksheetDrawing == null)
            {
                drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
            }

            var imagePart = FindIdenticalImagePart(drawingsPart, content);
            if (imagePart == null)
            {
                imagePart = drawingsPart.AddImagePart(imageContentType);
                using (var destination = imagePart.GetStream(FileMode.Create, FileAccess.Write))
                {
                    destination.Write(content, 0, content.Length);
                }
            }

            var imageRelationshipId = drawingsPart.GetIdOfPart(imagePart);
            var drawing = drawingsPart.WorksheetDrawing;
            var pictureIndex = drawing.Elements<Xdr.TwoCellAnchor>().Count();
            var nextId = drawing.Descendants<Xdr.NonVisualDrawingProperties>()
                .Select(item => item.Id?.Value ?? 0U)
                .DefaultIfEmpty(0U)
                .Max() + 1U;
            var widthEmus = checked((long)widthPixels * EmusPerPixel);
            var heightEmus = checked((long)heightPixels * EmusPerPixel);
            var end = CalculateEndMarker(worksheet, rowIndex, columnIndex, widthEmus, heightEmus);

            var anchor = new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId(columnIndex.ToString(CultureInfo.InvariantCulture)),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId(rowIndex.ToString(CultureInfo.InvariantCulture)),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId(end.ColumnIndex.ToString(CultureInfo.InvariantCulture)),
                    new Xdr.ColumnOffset(end.ColumnOffsetEmus.ToString(CultureInfo.InvariantCulture)),
                    new Xdr.RowId(end.RowIndex.ToString(CultureInfo.InvariantCulture)),
                    new Xdr.RowOffset(end.RowOffsetEmus.ToString(CultureInfo.InvariantCulture))),
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = nextId, Name = $"Picture {nextId}" },
                        new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = imageRelationshipId, CompressionState = A.BlipCompressionValues.Print },
                        new A.Stretch(new A.FillRectangle())),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0, Y = 0 },
                            new A.Extents { Cx = widthEmus, Cy = heightEmus }),
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
                new Xdr.ClientData())
            {
                EditAs = Xdr.EditAsValues.OneCell
            };

            drawing.Append(anchor);
            drawing.Save();
            worksheet.Save();
            return pictureIndex;
        }

        internal static void Format(
            WorkbookSession session,
            int measurementSystem,
            double scaleFactor,
            int pictureIndex,
            double height,
            double width,
            int colorType,
            string operation)
        {
            if (measurementSystem < 0 || measurementSystem > 1)
            {
                throw new ReportEngineException(ReportErrorCode.InvalidArgument, operation, $"Invalid measurement system: {measurementSystem}");
            }

            if (scaleFactor <= 0 || double.IsNaN(scaleFactor) || double.IsInfinity(scaleFactor))
            {
                throw new ReportEngineException(ReportErrorCode.InvalidArgument, operation, "Scale factor must be a finite value greater than zero.");
            }

            if (height == 0 || width == 0 || double.IsNaN(height) || double.IsNaN(width) || double.IsInfinity(height) || double.IsInfinity(width))
            {
                throw new ReportEngineException(ReportErrorCode.InvalidArgument, operation, "Explicit image dimensions must be finite and greater than zero; use a negative value to leave a dimension unchanged.");
            }

            var worksheetPart = session.ActiveWorksheetPart;
            var worksheet = worksheetPart.Worksheet
                ?? throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, operation, "The active worksheet root element is missing.");
            var drawing = worksheetPart.DrawingsPart?.WorksheetDrawing;
            var anchors = drawing?.Elements<Xdr.TwoCellAnchor>().ToArray() ?? new Xdr.TwoCellAnchor[0];
            var resolvedIndex = pictureIndex == -1 ? anchors.Length - 1 : pictureIndex;
            if (pictureIndex < -1 || resolvedIndex < 0 || resolvedIndex >= anchors.Length)
            {
                throw new ReportEngineException(
                    ReportErrorCode.PictureIndexOutOfRange,
                    operation,
                    $"Picture index {pictureIndex} is outside the valid range 0..{anchors.Length - 1}; -1 selects the last picture.");
            }

            var anchor = anchors[resolvedIndex];
            var picture = anchor.GetFirstChild<Xdr.Picture>()
                ?? throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, operation, $"Drawing object {resolvedIndex} is not a picture.");
            var extents = picture.ShapeProperties?.Transform2D?.Extents
                ?? throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, operation, $"Picture {resolvedIndex} has no size information.");
            var widthEmus = checked((long)Math.Round((extents.Cx?.Value ?? 0L) * scaleFactor, MidpointRounding.AwayFromZero));
            var heightEmus = checked((long)Math.Round((extents.Cy?.Value ?? 0L) * scaleFactor, MidpointRounding.AwayFromZero));

            if (width > 0)
            {
                widthEmus = DimensionToEmus(width, measurementSystem);
            }

            if (height > 0)
            {
                heightEmus = DimensionToEmus(height, measurementSystem);
            }

            extents.Cx = widthEmus;
            extents.Cy = heightEmus;
            ApplyColorType(picture, colorType, operation);

            var from = anchor.FromMarker
                ?? throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, operation, $"Picture {resolvedIndex} has no start marker.");
            var startColumn = int.Parse(from.ColumnId?.Text ?? "0", CultureInfo.InvariantCulture);
            var startRow = int.Parse(from.RowId?.Text ?? "0", CultureInfo.InvariantCulture);
            var startColumnOffset = long.Parse(from.ColumnOffset?.Text ?? "0", CultureInfo.InvariantCulture);
            var startRowOffset = long.Parse(from.RowOffset?.Text ?? "0", CultureInfo.InvariantCulture);
            var end = CalculateEndMarker(worksheet, startRow, startColumn, widthEmus + startColumnOffset, heightEmus + startRowOffset);
            var to = anchor.ToMarker
                ?? throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, operation, $"Picture {resolvedIndex} has no end marker.");
            to.ColumnId = new Xdr.ColumnId(end.ColumnIndex.ToString(CultureInfo.InvariantCulture));
            to.ColumnOffset = new Xdr.ColumnOffset(end.ColumnOffsetEmus.ToString(CultureInfo.InvariantCulture));
            to.RowId = new Xdr.RowId(end.RowIndex.ToString(CultureInfo.InvariantCulture));
            to.RowOffset = new Xdr.RowOffset(end.RowOffsetEmus.ToString(CultureInfo.InvariantCulture));
            drawing!.Save();
            worksheet.Save();
        }

        private static long DimensionToEmus(double value, int measurementSystem)
        {
            var inches = measurementSystem == 0 ? value : value / 2.54;
            var points = Math.Round(inches * 72.0, MidpointRounding.AwayFromZero);
            return checked((long)points * 12700L);
        }

        private static void ApplyColorType(Xdr.Picture picture, int colorType, string operation)
        {
            var blip = picture.BlipFill?.Blip
                ?? throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, operation, "Picture has no embedded image reference.");
            blip.RemoveAllChildren<A.Grayscale>();
            if (colorType == 0)
            {
                return;
            }

            if (colorType == 2)
            {
                blip.PrependChild(new A.Grayscale());
                return;
            }

            throw new ReportEngineException(
                ReportErrorCode.UnsupportedPictureColorType,
                operation,
                $"Picture color type {colorType} is not supported. Automatic (0) and Grayscale (2) are supported.");
        }

        private static string ValidateImagePath(string imagePath, string operation)
        {
            if (string.IsNullOrWhiteSpace(imagePath))
            {
                throw new ReportEngineException(ReportErrorCode.InvalidArgument, operation, "Image path is required.");
            }

            string fullPath;
            try
            {
                fullPath = Path.GetFullPath(imagePath);
            }
            catch (Exception ex) when (ex is ArgumentException || ex is NotSupportedException || ex is PathTooLongException)
            {
                throw new ReportEngineException(ReportErrorCode.InvalidArgument, operation, $"Invalid image path: {imagePath}", ex);
            }

            if (!File.Exists(fullPath))
            {
                throw new ReportEngineException(ReportErrorCode.ImageFileNotFound, operation, $"Image file not found: {fullPath}");
            }

            return fullPath;
        }

        private static string GetImageContentType(string path, string operation)
        {
            switch (Path.GetExtension(path).ToLowerInvariant())
            {
                case ".png": return "image/png";
                case ".jpg":
                case ".jpeg": return "image/jpeg";
                case ".gif": return "image/gif";
                case ".bmp": return "image/bmp";
                case ".tif":
                case ".tiff": return "image/tiff";
                default:
                    throw new ReportEngineException(ReportErrorCode.UnsupportedImageType, operation, $"Unsupported image type: {Path.GetExtension(path)}");
            }
        }

        private static ImagePart? FindIdenticalImagePart(DrawingsPart drawingsPart, byte[] content)
        {
            var expectedHash = ComputeHash(content);
            foreach (var part in drawingsPart.ImageParts)
            {
                using (var stream = part.GetStream(FileMode.Open, FileAccess.Read))
                using (var sha256 = SHA256.Create())
                {
                    var hash = sha256.ComputeHash(stream);
                    if (hash.SequenceEqual(expectedHash))
                    {
                        return part;
                    }
                }
            }

            return null;
        }

        private static byte[] ComputeHash(byte[] content)
        {
            using (var sha256 = SHA256.Create())
            {
                return sha256.ComputeHash(content);
            }
        }

        private static EndMarker CalculateEndMarker(Worksheet worksheet, int rowIndex, int columnIndex, long widthEmus, long heightEmus)
        {
            var remainingWidth = widthEmus;
            var endColumn = columnIndex;
            while (remainingWidth >= GetColumnWidthEmus(worksheet, endColumn))
            {
                remainingWidth -= GetColumnWidthEmus(worksheet, endColumn);
                endColumn++;
            }

            var remainingHeight = heightEmus;
            var endRow = rowIndex;
            while (remainingHeight >= GetRowHeightEmus(worksheet, endRow))
            {
                remainingHeight -= GetRowHeightEmus(worksheet, endRow);
                endRow++;
            }

            return new EndMarker(endRow, endColumn, remainingHeight, remainingWidth);
        }

        private static long GetColumnWidthEmus(Worksheet worksheet, int zeroBasedColumnIndex)
        {
            var oneBased = (uint)zeroBasedColumnIndex + 1U;
            var definition = worksheet.GetFirstChild<Columns>()?.Elements<Column>()
                .FirstOrDefault(item => item.Min?.Value <= oneBased && item.Max?.Value >= oneBased);
            if (definition?.Hidden?.Value == true)
            {
                return 1;
            }

            var width = definition?.Width?.Value
                ?? worksheet.SheetFormatProperties?.DefaultColumnWidth?.Value
                ?? DefaultColumnWidthCharacters;
            var pixels = Math.Floor(((256.0 * width + Math.Floor(128.0 / 7.0)) / 256.0) * 7.0);
            return Math.Max(1L, checked((long)pixels * EmusPerPixel));
        }

        private static long GetRowHeightEmus(Worksheet worksheet, int zeroBasedRowIndex)
        {
            var oneBased = (uint)zeroBasedRowIndex + 1U;
            var row = worksheet.GetFirstChild<SheetData>()?.Elements<Row>()
                .FirstOrDefault(item => item.RowIndex?.Value == oneBased);
            if (row?.Hidden?.Value == true)
            {
                return 1;
            }

            var points = row?.Height?.Value
                ?? worksheet.SheetFormatProperties?.DefaultRowHeight?.Value
                ?? DefaultRowHeightPoints;
            return Math.Max(1L, checked((long)Math.Round(points * 12700.0, MidpointRounding.AwayFromZero)));
        }

        private sealed class EndMarker
        {
            internal EndMarker(int rowIndex, int columnIndex, long rowOffsetEmus, long columnOffsetEmus)
            {
                RowIndex = rowIndex;
                ColumnIndex = columnIndex;
                RowOffsetEmus = rowOffsetEmus;
                ColumnOffsetEmus = columnOffsetEmus;
            }

            internal int RowIndex { get; }
            internal int ColumnIndex { get; }
            internal long RowOffsetEmus { get; }
            internal long ColumnOffsetEmus { get; }
        }
    }
}
