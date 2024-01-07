/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    public class AreaFamilyChart : ChartBase
    {
        #region Protected Fields

        protected readonly AreaChartSetting AreaChartSetting;

        #endregion Protected Fields

        #region Public Constructors

        public AreaFamilyChart(AreaChartSetting AreaChartSetting, ChartData[][] DataCols) : base(AreaChartSetting)
        {
            this.AreaChartSetting = AreaChartSetting;
            SetChartPlotArea(CreateChartPlotArea(DataCols));
        }

        #endregion Public Constructors

        #region Private Methods

        private C.AreaChartSeries CreateAreaChartSeries(int seriesIndex, ChartDataGrouping ChartDataGrouping, AreaChartSeriesSetting AreaChartSeriesSetting, A.SolidFill SolidFill, C.DataLabels? DataLabels)
        {
            C.AreaChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                CreateSeriesText(ChartDataGrouping.SeriesHeaderFormula!, new[] { ChartDataGrouping.SeriesHeaderCells! }));
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(new A.Outline(SolidFill, new A.Outline(new A.NoFill())));
            ShapeProperties.Append(new A.EffectList());
            if (DataLabels != null)
            {
                series.Append(DataLabels);
            }
            series.Append(ShapeProperties);
            series.Append(CreateCategoryAxisData(ChartDataGrouping.XaxisFormula!, ChartDataGrouping.XaxisCells!, AreaChartSeriesSetting));
            series.Append(CreateValueAxisData(ChartDataGrouping.YaxisFormula!, ChartDataGrouping.YaxisCells!, AreaChartSeriesSetting));
            return series;
        }

        private C.PlotArea CreateChartPlotArea(ChartData[][] DataCols)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            C.AreaChart AreaChart = new(
                new C.Grouping
                {
                    Val = AreaChartSetting.AreaChartTypes switch
                    {
                        AreaChartTypes.STACKED => C.GroupingValues.Stacked,
                        AreaChartTypes.PERCENT_STACKED => C.GroupingValues.PercentStacked,
                        // Clusted
                        _ => C.GroupingValues.Standard,
                    }
                },
                new C.VaryColors { Val = false });
            int seriesIndex = 0;
            CreateDataSeries(DataCols, AreaChartSetting.ChartDataSetting)
            .ForEach(Series =>
            {
                AreaChart.Append(CreateAreaChartSeries(seriesIndex, Series, AreaChartSetting.AreaChartSeriesSettings.Count > seriesIndex ? AreaChartSetting.AreaChartSeriesSettings[seriesIndex] : new AreaChartSeriesSetting(),
                                CreateSolidFill(AreaChartSetting.AreaChartSeriesSettings
                                        .Where(item => item.FillColor != null)
                                        .Select(item => item.FillColor!)
                                        .ToList(), seriesIndex),
                                GetDataLabels(seriesIndex)));
                seriesIndex++;
            });
            C.DataLabels DataLabels = new(
                new C.ShowLegendKey { Val = false },
                new C.ShowValue { Val = false },
                new C.ShowCategoryName { Val = false },
                new C.ShowSeriesName { Val = false },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false },
                new C.ShowLeaderLines { Val = true });
            AreaChart.Append(DataLabels);
            AreaChart.Append(new C.AxisId { Val = 1362418656 });
            AreaChart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(AreaChart);
            plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
            {
                Id = 1362418656,
                CrossAxisId = 1358349936
            }));
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                Id = 1358349936,
                CrossAxisId = 1362418656
            }));
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        private C.DataLabels? CreateDataLabel(AreaChartDataLabel AreaChartDataLabel)
        {
            if (AreaChartDataLabel.GetType().GetProperties()
                .Where(Prop => Prop.PropertyType == typeof(bool))
                .Any(Prop => (bool)Prop.GetValue(AreaChartDataLabel)!))
            {
                C.DataLabels DataLabels = new(
                                new C.ShowLegendKey { Val = AreaChartDataLabel.ShowLegendKey },
                                new C.ShowValue { Val = AreaChartDataLabel.ShowValue },
                                new C.ShowCategoryName { Val = AreaChartDataLabel.ShowCategoryName },
                                new C.ShowSeriesName { Val = AreaChartDataLabel.ShowSeriesName },
                                new C.ShowPercent { Val = false },
                                new C.ShowBubbleSize { Val = false },
                                new C.ShowLeaderLines() { Val = false });
                DataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = AreaChartDataLabel.DataLabelPosition switch
                    {
                        //Show
                        _ => C.DataLabelPositionValues.Center,
                    }
                }, 0);
                DataLabels.InsertAt(new C.ShapeProperties(new A.NoFill(), new A.Outline(new A.NoFill()), new A.EffectList()), 0);
                A.Paragraph Paragraph = new(new A.ParagraphProperties(new A.DefaultRunProperties(
                    new A.SolidFill(new A.SchemeColor(new A.LuminanceModulation() { Val = 75000 }, new A.LuminanceOffset() { Val = 25000 }) { Val = A.SchemeColorValues.Text1 }),
                    new A.LatinFont() { Typeface = "+mn-lt" }, new A.EastAsianFont() { Typeface = "+mn-ea" }, new A.ComplexScriptFont() { Typeface = "+mn-cs" })
                {
                    FontSize = 1197,
                    Bold = false,
                    Italic = false,
                    Underline = A.TextUnderlineValues.None,
                    Strike = A.TextStrikeValues.NoStrike,
                    Kerning = 1200,
                    Baseline = 0
                }), new A.EndParagraphRunProperties() { Language = "en-US" });
                DataLabels.InsertAt(new C.TextProperties(new A.BodyProperties(new A.ShapeAutoFit())
                {
                    Rotation = 0,
                    UseParagraphSpacing = true,
                    VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
                    Vertical = A.TextVerticalValues.Horizontal,
                    Wrap = A.TextWrappingValues.Square,
                    LeftInset = 38100,
                    TopInset = 19050,
                    RightInset = 38100,
                    BottomInset = 19050,
                    Anchor = A.TextAnchoringTypeValues.Center,
                    AnchorCenter = true
                }, new A.ListStyle(),
               Paragraph), 0);
                return DataLabels;
            }
            return null;
        }

        private C.DataLabels? GetDataLabels(int index)
        {
            if (index < AreaChartSetting.AreaChartSeriesSettings.Count)
            {
                return CreateDataLabel(AreaChartSetting.AreaChartSeriesSettings?[index]?.AreaChartDataLabel ?? new AreaChartDataLabel());
            }
            return null;
        }

        #endregion Private Methods
    }
}