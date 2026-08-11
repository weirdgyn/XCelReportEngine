using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XCelReportEngine.Internal
{
    internal static class CellStyleService
    {
        internal static void SetAlignment(
            WorkbookSession session,
            string startAddress,
            string endAddress,
            int horizontalAlignment,
            int verticalAlignment,
            bool wrapText,
            int textRotation,
            string operation)
        {
            var horizontal = GetHorizontalAlignment(horizontalAlignment, operation);
            var vertical = GetVerticalAlignment(verticalAlignment, operation);
            if (textRotation < 0 || textRotation > 180)
            {
                throw new ReportEngineException(ReportErrorCode.InvalidArgument, operation, "Text rotation must be in the range 0..180.");
            }

            var styles = GetOrCreateStylesheet(session.WorkbookPart);
            var styleMap = new Dictionary<uint, uint>();
            foreach (var cell in GetRangeCells(session, startAddress, endAddress, operation))
            {
                var sourceStyle = cell.StyleIndex?.Value ?? 0;
                if (styleMap.TryGetValue(sourceStyle, out var existingStyle))
                {
                    cell.StyleIndex = existingStyle;
                    continue;
                }

                var format = CloneCellFormat(styles, cell);
                format.Alignment = new Alignment
                {
                    Horizontal = horizontal,
                    Vertical = vertical,
                    WrapText = wrapText,
                    TextRotation = (uint)textRotation
                };
                format.ApplyAlignment = true;
                var newStyle = AppendCellFormat(styles, format);
                styleMap[sourceStyle] = newStyle;
                cell.StyleIndex = newStyle;
            }

            Save(session, styles);
        }

        internal static void SetColorAndBorder(
            WorkbookSession session,
            string startAddress,
            string endAddress,
            bool applyFill,
            int fillColor,
            bool applyBorder,
            int borderColor,
            int borderStyle,
            int borderEdges,
            string operation)
        {
            if (!applyFill && !applyBorder)
            {
                return;
            }

            if (applyFill)
            {
                ValidateRgb(fillColor, nameof(fillColor), operation);
            }

            if (applyBorder)
            {
                ValidateRgb(borderColor, nameof(borderColor), operation);
                if (borderEdges < 0 || (borderEdges & ~15) != 0)
                {
                    throw new ReportEngineException(ReportErrorCode.InvalidArgument, operation, "Border edges must be a bit mask in the range 0..15.");
                }
            }

            var style = applyBorder ? GetBorderStyle(borderStyle, operation) : BorderStyleValues.None;
            var styles = GetOrCreateStylesheet(session.WorkbookPart);
            uint? fillId = applyFill ? AppendFill(styles, fillColor) : (uint?)null;
            uint? borderId = applyBorder ? AppendBorder(styles, borderColor, style, borderEdges) : (uint?)null;

            var styleMap = new Dictionary<uint, uint>();
            foreach (var cell in GetRangeCells(session, startAddress, endAddress, operation))
            {
                var sourceStyle = cell.StyleIndex?.Value ?? 0;
                if (styleMap.TryGetValue(sourceStyle, out var existingStyle))
                {
                    cell.StyleIndex = existingStyle;
                    continue;
                }

                var format = CloneCellFormat(styles, cell);
                if (fillId.HasValue)
                {
                    format.FillId = fillId.Value;
                    format.ApplyFill = true;
                }

                if (borderId.HasValue)
                {
                    format.BorderId = borderId.Value;
                    format.ApplyBorder = true;
                }

                var newStyle = AppendCellFormat(styles, format);
                styleMap[sourceStyle] = newStyle;
                cell.StyleIndex = newStyle;
            }

            Save(session, styles);
        }

        private static IEnumerable<Cell> GetRangeCells(WorkbookSession session, string startAddress, string endAddress, string operation)
        {
            var start = CellAddress.Normalize(startAddress, operation);
            var end = string.IsNullOrWhiteSpace(endAddress) ? start : CellAddress.Normalize(endAddress, operation);
            CellAddress.Parse(start, out var startRow, out var startColumn);
            CellAddress.Parse(end, out var endRow, out var endColumn);
            if (endRow < startRow || endColumn < startColumn)
            {
                throw new ReportEngineException(ReportErrorCode.InvalidRangeDimensions, operation, "The end cell must not precede the start cell.");
            }

            for (var row = startRow; row <= endRow; row++)
            {
                for (var column = startColumn; column <= endColumn; column++)
                {
                    yield return CellService.GetOrCreateCell(session, CellAddress.FromOneBased(row, checked((uint)column), operation));
                }
            }
        }

        private static Stylesheet GetOrCreateStylesheet(WorkbookPart workbookPart)
        {
            var part = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
            var stylesheet = part.Stylesheet;
            if (stylesheet == null)
            {
                stylesheet = new Stylesheet();
                part.Stylesheet = stylesheet;
            }
            stylesheet.Fonts ??= new Fonts(new Font()) { Count = 1 };
            stylesheet.Fills ??= new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 2 };
            stylesheet.Borders ??= new Borders(new Border()) { Count = 1 };
            stylesheet.CellStyleFormats ??= new CellStyleFormats(new CellFormat()) { Count = 1 };
            stylesheet.CellFormats ??= new CellFormats(new CellFormat()) { Count = 1 };
            return stylesheet;
        }

        private static CellFormat CloneCellFormat(Stylesheet styles, Cell cell)
        {
            var formats = styles.CellFormats!;
            var index = cell.StyleIndex?.Value ?? 0;
            var source = formats.Elements<CellFormat>().ElementAtOrDefault((int)index) ?? new CellFormat();
            return (CellFormat)source.CloneNode(true);
        }

        private static uint AppendCellFormat(Stylesheet styles, CellFormat format)
        {
            styles.CellFormats!.Append(format);
            styles.CellFormats.Count = (uint)styles.CellFormats.ChildElements.Count;
            return styles.CellFormats.Count.Value - 1;
        }

        private static uint AppendFill(Stylesheet styles, int color)
        {
            var fill = new Fill(new PatternFill(
                new ForegroundColor { Rgb = ToArgb(color) },
                new BackgroundColor { Indexed = 64U }) { PatternType = PatternValues.Solid });
            styles.Fills!.Append(fill);
            styles.Fills.Count = (uint)styles.Fills.ChildElements.Count;
            return styles.Fills.Count.Value - 1;
        }

        private static uint AppendBorder(Stylesheet styles, int color, BorderStyleValues style, int edges)
        {
            var border = new Border
            {
                LeftBorder = CreateBorder<LeftBorder>(color, style, (edges & 1) != 0),
                RightBorder = CreateBorder<RightBorder>(color, style, (edges & 2) != 0),
                TopBorder = CreateBorder<TopBorder>(color, style, (edges & 4) != 0),
                BottomBorder = CreateBorder<BottomBorder>(color, style, (edges & 8) != 0),
                DiagonalBorder = new DiagonalBorder()
            };
            styles.Borders!.Append(border);
            styles.Borders.Count = (uint)styles.Borders.ChildElements.Count;
            return styles.Borders.Count.Value - 1;
        }

        private static T CreateBorder<T>(int color, BorderStyleValues style, bool enabled) where T : BorderPropertiesType, new()
        {
            var border = new T();
            if (enabled)
            {
                border.Style = style;
                border.Color = new Color { Rgb = ToArgb(color) };
            }
            return border;
        }

        private static HorizontalAlignmentValues GetHorizontalAlignment(int value, string operation)
        {
            switch (value)
            {
                case 0: return HorizontalAlignmentValues.General;
                case 1: return HorizontalAlignmentValues.Left;
                case 2: return HorizontalAlignmentValues.Center;
                case 3: return HorizontalAlignmentValues.Right;
                case 4: return HorizontalAlignmentValues.Fill;
                case 5: return HorizontalAlignmentValues.Justify;
                case 6: return HorizontalAlignmentValues.CenterContinuous;
                case 7: return HorizontalAlignmentValues.Distributed;
                default: throw new ReportEngineException(ReportErrorCode.InvalidArgument, operation, $"Invalid horizontal alignment value: {value}");
            }
        }

        private static VerticalAlignmentValues GetVerticalAlignment(int value, string operation)
        {
            switch (value)
            {
                case 0: return VerticalAlignmentValues.Bottom;
                case 1: return VerticalAlignmentValues.Center;
                case 2: return VerticalAlignmentValues.Top;
                case 3: return VerticalAlignmentValues.Justify;
                case 4: return VerticalAlignmentValues.Distributed;
                default: throw new ReportEngineException(ReportErrorCode.InvalidArgument, operation, $"Invalid vertical alignment value: {value}");
            }
        }

        private static BorderStyleValues GetBorderStyle(int value, string operation)
        {
            switch (value)
            {
                case 0: return BorderStyleValues.None;
                case 1: return BorderStyleValues.Thin;
                case 2: return BorderStyleValues.Medium;
                case 3: return BorderStyleValues.Thick;
                case 4: return BorderStyleValues.Double;
                case 5: return BorderStyleValues.Dashed;
                case 6: return BorderStyleValues.Dotted;
                default: throw new ReportEngineException(ReportErrorCode.InvalidArgument, operation, $"Invalid border style value: {value}");
            }
        }

        private static void ValidateRgb(int color, string parameter, string operation)
        {
            if (color < 0 || color > 0xFFFFFF)
            {
                throw new ReportEngineException(ReportErrorCode.InvalidArgument, operation, $"{parameter} must be an RGB value in the range 0x000000..0xFFFFFF.");
            }
        }

        private static HexBinaryValue ToArgb(int color) => new HexBinaryValue($"FF{color:X6}");

        private static void Save(WorkbookSession session, Stylesheet styles)
        {
            styles.Save();
            var worksheet = session.ActiveWorksheetPart.Worksheet
                ?? throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, "ActiveWorksheet", "The worksheet root element is missing.");
            worksheet.Save();
        }
    }
}
