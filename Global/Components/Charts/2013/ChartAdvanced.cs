// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using OpenXMLOffice.Global_2007;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C15 = DocumentFormat.OpenXml.Office2013.Drawing.Chart;

namespace OpenXMLOffice.Global_2013
{
    /// <summary>
    /// TODO: Reorganise to skip the loop back to 2007 namespace
    /// </summary>
    public class ChartAdvance<ApplicationSpecificSetting> : ChartBase<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition
    {
        /// <summary>
        /// 
        /// </summary>
        public ChartAdvance(ChartSetting<ApplicationSpecificSetting> chartSetting) : base(chartSetting) { }

        /// <summary>
        /// Create Data Labels for the chart
        /// </summary>
        internal C.DataLabels CreateDataLabels(ChartDataLabel chartDataLabel, int? dataLabelCount = 0)
        {
            C.Extension extension = new(
                    new C15.ShowDataLabelsRange() { Val = chartDataLabel.showValueFromColumn },
                    new C15.ShowLeaderLines() { Val = false }
                )
            { Uri = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" };
            if (chartDataLabel.showValueFromColumn)
            {
                extension.InsertAt(new C15.DataLabelFieldTable(), 0);
            }
            C.ExtensionList extensionList = new(extension);
            C.DataLabels dataLabels = new();
            if (chartDataLabel.showValueFromColumn)
            {
                for (int i = 0; i < dataLabelCount; i++)
                {
                    A.Paragraph Paragraph = new(CreateField("CELLRANGE", "[CELLRANGE]"));
                    if (chartDataLabel.showSeriesName)
                    {
                        Paragraph.Append(CreateDrawingRun(new() { text = chartDataLabel.separator }));
                        Paragraph.Append(CreateField("SERIESNAME", "[SERIES NAME]"));
                    }
                    if (chartDataLabel.showCategoryName)
                    {
                        Paragraph.Append(CreateDrawingRun(new() { text = chartDataLabel.separator }));
                        Paragraph.Append(CreateField("CATEGORYNAME", "[CATEGORY NAME]"));
                    }
                    if (chartDataLabel.showValue)
                    {
                        Paragraph.Append(CreateDrawingRun(new() { text = chartDataLabel.separator }));
                        Paragraph.Append(CreateField("VALUE", "[VALUE]"));
                    }
                    Paragraph.Append(new A.EndParagraphRunProperties { Language = "en-IN" });
                    dataLabels.Append(new C.DataLabel(
                        new C.Index() { Val = (uint)i },
                        new C.SeriesText(
                            new C.RichText(
                                new A.BodyProperties(),
                                new A.ListStyle(),
                                Paragraph
                            )
                        ),
                        new C.ShowLegendKey { Val = chartDataLabel.showLegendKey },
                        new C.ShowValue { Val = chartDataLabel.showValue },
                        new C.ShowCategoryName { Val = chartDataLabel.showCategoryName },
                        new C.ShowSeriesName { Val = chartDataLabel.showSeriesName },
                        new C.ShowPercent() { Val = true },
                        new C.ShowBubbleSize() { Val = true },
                        new C.Separator(chartDataLabel.separator),
                        (OpenXmlElement)extensionList.Clone()
                    ));
                }
            }
            dataLabels.Append(new C.ShowLegendKey { Val = chartDataLabel.showLegendKey },
                new C.ShowValue { Val = chartDataLabel.showValue },
                new C.ShowCategoryName { Val = chartDataLabel.showCategoryName },
                new C.ShowSeriesName { Val = chartDataLabel.showSeriesName },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize() { Val = true },
                new C.Separator(chartDataLabel.separator),
                new C.ShowLeaderLines() { Val = false },
                (OpenXmlElement)extensionList.Clone());
            dataLabels.Append(CreateChartShapeProperties());
            SolidFillModel solidFillModel = new()
            {
                schemeColorModel = new()
                {
                    themeColorValues = ThemeColorValues.TEXT_1,
                    luminanceModulation = 65000,
                    luminanceOffset = 35000
                }
            };
            if (chartDataLabel.fontColor != null)
            {
                solidFillModel.hexColor = chartDataLabel.fontColor;
                solidFillModel.schemeColorModel = null;
            }
            dataLabels.Append(CreateChartTextProperties(new()
            {
                drawingBodyProperties = new()
                {
                    rotation = 0,
                    anchorCenter = true,
                    anchor = TextAnchoringValues.CENTER,
                    bottomInset = 19050,
                    leftInset = 38100,
                    rightInset = 38100,
                    topInset = 19050,
                    useParagraphSpacing = true,
                    vertical = TextVerticalAlignmentValues.HORIZONTAL,
                    verticalOverflow = TextVerticalOverflowValues.ELLIPSIS,
                    wrap = TextWrappingValues.SQUARE,
                },
                drawingParagraph = new()
                {
                    paragraphPropertiesModel = new()
                    {
                        defaultRunProperties = new()
                        {
                            solidFill = solidFillModel,
                            complexScriptFont = "+mn-cs",
                            eastAsianFont = "+mn-ea",
                            latinFont = "+mn-lt",
                            fontSize = ConverterUtils.FontSizeToFontSize(chartDataLabel.fontSize),
                            isBold = chartDataLabel.isBold,
                            isItalic = chartDataLabel.isItalic,
                            underline = chartDataLabel.underLineValues,
                            strike = chartDataLabel.strikeValues,
                            kerning = 1200,
                            baseline = 0,
                        }
                    }
                }
            }));
            return dataLabels;
        }

        /// <summary>
        /// Create Data Labels Range for the chart.Used in value from Column
        /// </summary>
        internal static C15.DataLabelsRange CreateDataLabelsRange(string formula, ChartData[] cells)
        {
            return new(new C15.Formula(formula), AddDataLabelCacheValue(cells));
        }
        private static C15.DataLabelsRangeChache AddDataLabelCacheValue(ChartData[] cells)
        {
            try
            {
                C15.DataLabelsRangeChache dataLabelsRangeChache = new()
                {
                    PointCount = new C.PointCount()
                    {
                        Val = (UInt32Value)(uint)cells.Length
                    },
                };
                int count = 0;
                foreach (ChartData Cell in cells)
                {
                    C.StringPoint stringPoint = new()
                    {
                        Index = (UInt32Value)(uint)count,
                    };
                    stringPoint.AppendChild(new C.NumericValue(Cell.value ?? ""));
                    dataLabelsRangeChache.AppendChild(stringPoint);
                    ++count;
                }
                return dataLabelsRangeChache;
            }
            catch
            {
                throw new Exception("Chart. Data Label Ref Error");
            }
        }
    }
}