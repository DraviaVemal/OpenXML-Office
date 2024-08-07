// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using OpenXMLOffice.Global_2013;
using C = DocumentFormat.OpenXml.Drawing.Charts;
namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// Represents the types of scatter charts.
	/// </summary>
	public class ScatterChart<ApplicationSpecificSetting> : ChartAdvance<ApplicationSpecificSetting>
		where ApplicationSpecificSetting : class, ISizeAndPosition, new()
	{
		/// <summary>
		/// Scatter Chart Setting
		/// </summary>
		protected ScatterChartSetting<ApplicationSpecificSetting> scatterChartSetting;
		internal ScatterChart(ScatterChartSetting<ApplicationSpecificSetting> scatterChartSetting) : base(scatterChartSetting)
		{
			this.scatterChartSetting = scatterChartSetting;
		}
		/// <summary>
		/// Create Scatter Chart with provided settings
		/// </summary>
		public ScatterChart(ScatterChartSetting<ApplicationSpecificSetting> scatterChartSetting, ChartData[][] dataCols, DataRange dataRange = null) : base(scatterChartSetting)
		{
			this.scatterChartSetting = scatterChartSetting;
			SetChartPlotArea(CreateChartPlotArea(dataCols, dataRange));
		}
		private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols, DataRange dataRange)
		{
			if (scatterChartSetting.scatterChartType == ScatterChartTypes.BUBBLE || scatterChartSetting.scatterChartType == ScatterChartTypes.BUBBLE_3D)
			{
				scatterChartSetting.chartDataSetting.is3dData = true;
			}
			C.PlotArea plotArea = new C.PlotArea();
			plotArea.Append(CreateLayout(scatterChartSetting.plotAreaOptions != null ? scatterChartSetting.plotAreaOptions.manualLayout : null));
			if (scatterChartSetting.scatterChartType == ScatterChartTypes.BUBBLE || scatterChartSetting.scatterChartType == ScatterChartTypes.BUBBLE_3D)
			{
				plotArea.Append(CreateChart<C.BubbleChart>(CreateDataSeries(scatterChartSetting.chartDataSetting, dataCols, dataRange)));
			}
			else
			{
				plotArea.Append(CreateChart<C.ScatterChart>(CreateDataSeries(scatterChartSetting.chartDataSetting, dataCols, dataRange)));
			}
			plotArea.Append(CreateAxis(new AxisSetting<XAxisOptions<ValueAxis>, ValueAxis>()
			{
				id = scatterChartSetting.isSecondaryAxis ? SecondaryCategoryAxisId : CategoryAxisId,
				crossAxisId = scatterChartSetting.isSecondaryAxis ? SecondaryValueAxisId : ValueAxisId,
				axisOptions = scatterChartSetting.chartAxisOptions.xAxisOptions,
				axisPosition = scatterChartSetting.chartAxisOptions.xAxisOptions.chartAxesOptions.inReverseOrder ? AxisPosition.TOP : AxisPosition.BOTTOM
			}));
			plotArea.Append(CreateAxis(new AxisSetting<YAxisOptions<ValueAxis>, ValueAxis>()
			{
				id = scatterChartSetting.isSecondaryAxis ? SecondaryValueAxisId : ValueAxisId,
				crossAxisId = scatterChartSetting.isSecondaryAxis ? SecondaryCategoryAxisId : CategoryAxisId,
				axisOptions = scatterChartSetting.chartAxisOptions.yAxisOptions,
				axisPosition = scatterChartSetting.chartAxisOptions.yAxisOptions.chartAxesOptions.inReverseOrder ? AxisPosition.RIGHT : AxisPosition.LEFT
			}));
			plotArea.Append(CreateChartShapeProperties());
			return plotArea;
		}
		internal ChartType CreateChart<ChartType>(List<ChartDataGrouping> chartDataGroupings) where ChartType : OpenXmlCompositeElement, new()
		{
			ChartType chart = new ChartType();
			C.ScatterStyleValues scatterStyleValue;
			switch (scatterChartSetting.scatterChartType)
			{
				case ScatterChartTypes.SCATTER:
					scatterStyleValue = C.ScatterStyleValues.LineMarker;
					chart.Append(new C.ScatterStyle
					{
						Val = scatterStyleValue
					});
					break;
				case ScatterChartTypes.SCATTER_SMOOTH:
					scatterStyleValue = C.ScatterStyleValues.Smooth;
					chart.Append(new C.ScatterStyle
					{
						Val = scatterStyleValue
					});
					break;
				case ScatterChartTypes.SCATTER_SMOOTH_MARKER:
					scatterStyleValue = C.ScatterStyleValues.SmoothMarker;
					chart.Append(new C.ScatterStyle
					{
						Val = scatterStyleValue
					});
					break;
				case ScatterChartTypes.SCATTER_STRAIGHT:
					scatterStyleValue = C.ScatterStyleValues.Line;
					chart.Append(new C.ScatterStyle
					{
						Val = scatterStyleValue
					});
					break;
				case ScatterChartTypes.SCATTER_STRAIGHT_MARKER:
					scatterStyleValue = C.ScatterStyleValues.LineMarker;
					chart.Append(new C.ScatterStyle
					{
						Val = scatterStyleValue
					});
					break;
				default:
					break;
			}
			chart.Append(new C.VaryColors() { Val = false });
			int seriesIndex = 0;
			chartDataGroupings.ForEach(Series =>
			{
				chart.Append(CreateScatterChartSeries(seriesIndex, Series));
				seriesIndex++;
			});
			C.DataLabels dataLabels = CreateScatterDataLabels(scatterChartSetting.scatterChartDataLabel);
			if (dataLabels != null)
			{
				chart.Append(dataLabels);
			}
			if (scatterChartSetting.scatterChartType == ScatterChartTypes.BUBBLE || scatterChartSetting.scatterChartType == ScatterChartTypes.BUBBLE_3D)
			{
				chart.Append(new C.BubbleScale() { Val = 100 });
				chart.Append(new C.ShowNegativeBubbles() { Val = false });
			}
			chart.Append(new C.AxisId { Val = CategoryAxisId });
			chart.Append(new C.AxisId { Val = ValueAxisId });
			return chart;
		}
		private ColorOptionModel<SolidOptions> GetSeriesBorderColor(int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			ColorOptionModel<SolidOptions> solidFillModel = new ColorOptionModel<SolidOptions>();
			string hexColor = scatterChartSetting.scatterChartSeriesSettings
						.Select(item => item.borderColor)
						.ToList().ElementAtOrDefault(seriesIndex);
			if (hexColor != null)
			{
				solidFillModel.colorOption.hexColor = hexColor;
				return solidFillModel;
			}
			else
			{
				solidFillModel.colorOption.schemeColorModel = new SchemeColorModel()
				{
					themeColorValues = ThemeColorValues.ACCENT_1 + (chartDataGrouping.id % AccentColorCount),
				};
			}
			return solidFillModel;
		}
		private C.ScatterChartSeries CreateScatterChartSeries(int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			C.DataLabels dataLabels = null;
			ScatterChartSeriesSetting scatterChartSeriesSetting = scatterChartSetting.scatterChartSeriesSettings.ElementAtOrDefault(seriesIndex);
			if (seriesIndex < scatterChartSetting.scatterChartSeriesSettings.Count)
			{
				ScatterChartDataLabel scatterChartDataLabel = scatterChartSeriesSetting != null ? scatterChartSeriesSetting.scatterChartDataLabel : null;
				int dataLabelCellsLength = chartDataGrouping.dataLabelCells != null ? chartDataGrouping.dataLabelCells.Length : 0;
				dataLabels = CreateScatterDataLabels(scatterChartDataLabel ?? new ScatterChartDataLabel(), dataLabelCellsLength);
			}
			C.ScatterChartSeries series = new C.ScatterChartSeries(
				new C.Index { Val = new UInt32Value((uint)chartDataGrouping.id) },
				new C.Order { Val = new UInt32Value((uint)chartDataGrouping.id) },
				CreateSeriesText(chartDataGrouping.seriesHeaderFormula, new[] { chartDataGrouping.seriesHeaderCells }));
			if (scatterChartSetting.scatterChartType == ScatterChartTypes.BUBBLE || scatterChartSetting.scatterChartType == ScatterChartTypes.BUBBLE_3D)
			{
				ShapePropertiesModel<SolidOptions, SolidOptions> shapePropertiesModel = new ShapePropertiesModel<SolidOptions, SolidOptions>
				{
					lineColor = new OutlineModel<SolidOptions>()
					{
						lineColor = scatterChartSetting.scatterChartType == ScatterChartTypes.SCATTER ? null : GetSeriesBorderColor(seriesIndex, chartDataGrouping),
					},
					fillColor = new ColorOptionModel<SolidOptions>()
					{
						colorOption = new SolidOptions()
						{
							schemeColorModel = new SchemeColorModel()
							{
								themeColorValues = ThemeColorValues.ACCENT_1 + (seriesIndex % AccentColorCount),
								tint = 75000,
							}
						}
					}
				};
				series.Append(new C.InvertIfNegative() { Val = scatterChartSeriesSetting != null ? scatterChartSeriesSetting.invertIfNegative : scatterChartSetting.invertIfNegative });
				series.Append(CreateChartShapeProperties(shapePropertiesModel));
			}
			if (new[] { ScatterChartTypes.SCATTER, ScatterChartTypes.SCATTER_SMOOTH_MARKER, ScatterChartTypes.SCATTER_STRAIGHT_MARKER }.Contains(scatterChartSetting.scatterChartType))
			{
				MarkerModel<SolidOptions, SolidOptions> markerModel = new MarkerModel<SolidOptions, SolidOptions>
				{
					markerShapeType = scatterChartSetting.scatterChartType == ScatterChartTypes.SCATTER ? MarkerShapeTypes.AUTO : MarkerShapeTypes.CIRCLE,
					shapeProperties = new ShapePropertiesModel<SolidOptions, SolidOptions>()
					{
						fillColor = new ColorOptionModel<SolidOptions>()
						{
							colorOption = new SolidOptions()
							{
								schemeColorModel = new SchemeColorModel()
								{
									themeColorValues = ThemeColorValues.ACCENT_1 + (chartDataGrouping.id % AccentColorCount),
								}
							}
						},
						lineColor = new OutlineModel<SolidOptions>()
						{
							lineColor = new ColorOptionModel<SolidOptions>()
							{
								colorOption = new SolidOptions()
								{
									schemeColorModel = new SchemeColorModel()
									{
										themeColorValues = ThemeColorValues.ACCENT_1 + (chartDataGrouping.id % AccentColorCount),
									}
								}
							}
						}
					}
				};
				if (scatterChartSetting.scatterChartType != ScatterChartTypes.BUBBLE && scatterChartSetting.scatterChartType != ScatterChartTypes.BUBBLE_3D)
				{
					series.Append(CreateMarker(markerModel));
				}
			}
			else
			{
				MarkerModel<NoFillOptions, NoFillOptions> markerModel = new MarkerModel<NoFillOptions, NoFillOptions>();
				if (scatterChartSetting.scatterChartType != ScatterChartTypes.BUBBLE && scatterChartSetting.scatterChartType != ScatterChartTypes.BUBBLE_3D)
				{
					series.Append(CreateMarker(markerModel));
				}
			}

			if (scatterChartSeriesSetting != null)
			{
				scatterChartSeriesSetting.trendLines.ForEach(trendLine =>
				{
					ColorOptionModel<SolidOptions> solidFillModel = new ColorOptionModel<SolidOptions>();
					if (trendLine.hexColor != null)
					{
						solidFillModel.colorOption.hexColor = trendLine.hexColor;
					}
					else
					{
						solidFillModel.colorOption.schemeColorModel = new SchemeColorModel()
						{
							themeColorValues = ThemeColorValues.ACCENT_1 + (seriesIndex % AccentColorCount)
						};
					}
					TrendLineModel trendLineModel = new TrendLineModel
					{
						secondaryValue = trendLine.secondaryValue,
						trendLineType = trendLine.trendLineType,
						trendLineName = trendLine.trendLineName,
						forecastBackward = trendLine.forecastBackward,
						forecastForward = trendLine.forecastForward,
						setIntercept = trendLine.setIntercept,
						showEquation = trendLine.showEquation,
						showRSquareValue = trendLine.showRSquareValue,
						interceptValue = trendLine.interceptValue,
						solidFill = solidFillModel,
						drawingPresetLineDashValues = trendLine.lineStye
					};
					series.Append(CreateTrendLine(trendLineModel));
				});
			}
			if (dataLabels != null)
			{
				series.Append(dataLabels);
			}
			series.Append(CreateXValueAxisData(chartDataGrouping.xAxisFormula, chartDataGrouping.xAxisCells));
			series.Append(CreateYValueAxisData(chartDataGrouping.yAxisFormula, chartDataGrouping.yAxisCells));
			if (scatterChartSetting.scatterChartType == ScatterChartTypes.BUBBLE || scatterChartSetting.scatterChartType == ScatterChartTypes.BUBBLE_3D)
			{
				series.Append(CreateBubbleSizeAxisData(chartDataGrouping.zAxisFormula, chartDataGrouping.zAxisCells));
				series.Append(new C.Bubble3D() { Val = scatterChartSetting.scatterChartType == ScatterChartTypes.BUBBLE_3D });
			}
			else
			{
				series.Append(new C.Smooth() { Val = new[] { ScatterChartTypes.SCATTER_SMOOTH, ScatterChartTypes.SCATTER_SMOOTH_MARKER }.Contains(scatterChartSetting.scatterChartType) });
			}
			if (chartDataGrouping.dataLabelCells != null && chartDataGrouping.dataLabelFormula != null)
			{
				series.Append(new C.ExtensionList(new C.Extension(
					CreateDataLabelsRange(chartDataGrouping.dataLabelFormula, chartDataGrouping.dataLabelCells.Skip(1).ToArray())
				)
				{ Uri = "{02D57815-91ED-43cb-92C2-25804820EDAC}" }));
			}
			return series;
		}
		private C.DataLabels CreateScatterDataLabels(ScatterChartDataLabel scatterChartDataLabel, int? dataLabelCounter = 0)
		{
			if (scatterChartDataLabel.showValue || scatterChartSetting.chartDataSetting.advancedDataLabel.showValueFromColumn || scatterChartDataLabel.showCategoryName || scatterChartDataLabel.showLegendKey || scatterChartDataLabel.showSeriesName || scatterChartDataLabel.showBubbleSize)
			{
				C.DataLabels dataLabels = CreateDataLabels(scatterChartDataLabel, dataLabelCounter);
				dataLabels.Append(new C.ShowBubbleSize { Val = scatterChartDataLabel.showBubbleSize });
				C.DataLabelPositionValues dataLabelPositionValue;
				if (scatterChartDataLabel.dataLabelPosition == ScatterChartDataLabel.DataLabelPositionValues.LEFT)
				{
					dataLabelPositionValue = C.DataLabelPositionValues.Left;
				}
				else if (scatterChartDataLabel.dataLabelPosition == ScatterChartDataLabel.DataLabelPositionValues.RIGHT)
				{
					dataLabelPositionValue = C.DataLabelPositionValues.Right;
				}
				else if (scatterChartDataLabel.dataLabelPosition == ScatterChartDataLabel.DataLabelPositionValues.ABOVE)
				{
					dataLabelPositionValue = C.DataLabelPositionValues.Top;
				}
				else if (scatterChartDataLabel.dataLabelPosition == ScatterChartDataLabel.DataLabelPositionValues.BELOW)
				{
					dataLabelPositionValue = C.DataLabelPositionValues.Bottom;
				}
				else
				{
					dataLabelPositionValue = C.DataLabelPositionValues.Center;
				}
				dataLabels.InsertAt(new C.DataLabelPosition()
				{
					Val = dataLabelPositionValue
				}, 0);
				return dataLabels;
			}
			return null;
		}
	}
}
