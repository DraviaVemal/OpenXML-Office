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
	/// Represents the types of pie charts.
	/// </summary>
	public class PieChart<ApplicationSpecificSetting> : ChartAdvance<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition, new()
	{
		/// <summary>
		/// The settings for the pie chart.
		/// </summary>
		protected PieChartSetting<ApplicationSpecificSetting> pieChartSetting;
		internal PieChart(PieChartSetting<ApplicationSpecificSetting> pieChartSetting) : base(pieChartSetting)
		{
			this.pieChartSetting = pieChartSetting;
		}
		/// <summary>
		/// Create Pie Chart with provided settings
		/// </summary>
		public PieChart(PieChartSetting<ApplicationSpecificSetting> pieChartSetting, ChartData[][] dataCols, DataRange dataRange = null) : base(pieChartSetting)
		{
			this.pieChartSetting = pieChartSetting;
			if (pieChartSetting.pieChartType == PieChartTypes.PIE_3D)
			{
				this.pieChartSetting.is3DChart = true;
				Add3dControl();
			}
			SetChartPlotArea(CreateChartPlotArea(dataCols, dataRange));
		}
		private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols, DataRange dataRange)
		{
			// Set default data column end for pie and doughnut
			pieChartSetting.chartDataSetting.chartDataColumnEnd = pieChartSetting.chartDataSetting.chartDataColumnEnd == 0 ? pieChartSetting.chartDataSetting.chartDataColumnStart : pieChartSetting.chartDataSetting.chartDataColumnEnd;
			C.PlotArea plotArea = new C.PlotArea();
			plotArea.Append(CreateLayout(pieChartSetting.plotAreaOptions != null ? pieChartSetting.plotAreaOptions.manualLayout : null));
			if (pieChartSetting.pieChartType == PieChartTypes.DOUGHNUT)
			{
				plotArea.Append(CreateChart<C.DoughnutChart>(CreateDataSeries(pieChartSetting.chartDataSetting, dataCols, dataRange)));
			}
			else
			{
				if (pieChartSetting.is3DChart)
				{
					plotArea.Append(CreateChart<C.Pie3DChart>(CreateDataSeries(pieChartSetting.chartDataSetting, dataCols, dataRange)));
				}
				else
				{
					plotArea.Append(CreateChart<C.PieChart>(CreateDataSeries(pieChartSetting.chartDataSetting, dataCols, dataRange)));
				}
			}
			plotArea.Append(CreateChartShapeProperties());
			return plotArea;
		}
		internal ChartType CreateChart<ChartType>(List<ChartDataGrouping> chartDataGroupings) where ChartType : OpenXmlCompositeElement, new()
		{
			ChartType chart = new ChartType();
			chart.Append(new C.VaryColors { Val = true });
			int seriesIndex = 0;
			chartDataGroupings.ForEach(Series =>
			{
				chart.Append(CreateChartSeries(seriesIndex, Series));
				seriesIndex++;
			});
			C.DataLabels dataLabels = CreatePieDataLabels(pieChartSetting.pieChartDataLabel);
			if (dataLabels != null)
			{
				chart.Append(dataLabels);
			}
			chart.Append(new C.FirstSliceAngle { Val = (UInt16Value)pieChartSetting.angleOfFirstSlice });
			if (pieChartSetting.pieChartType == PieChartTypes.DOUGHNUT)
			{
				chart.Append(new C.HoleSize { Val = (ByteValue)pieChartSetting.doughnutHoleSize });
			}
			return chart;
		}
		private PieChartDataLabel GetDataPointLabel(uint index, int seriesIndex)
		{
			PieChartDataLabel pieChartDataLabel = pieChartSetting.pieChartSeriesSettings.ElementAtOrDefault(seriesIndex) != null ? pieChartSetting.pieChartSeriesSettings.ElementAtOrDefault(seriesIndex).pieChartDataPointSettings
						.Select(item => item != null ? item.pieChartDataLabel : null)
						.ToList().ElementAtOrDefault((int)index) : null;
			if (pieChartDataLabel != null)
			{
				return pieChartDataLabel;
			}
			return null;
		}
		private ColorOptionModel<SolidOptions> GetDataPointFill(uint index, int seriesIndex)
		{
			ColorOptionModel<SolidOptions> solidFillModel = new ColorOptionModel<SolidOptions>();
			string hexColor = pieChartSetting.pieChartSeriesSettings.ElementAtOrDefault(seriesIndex) != null ? pieChartSetting.pieChartSeriesSettings.ElementAtOrDefault(seriesIndex).pieChartDataPointSettings
						.Select(item => item != null ? item.fillColor : null)
						.ToList().ElementAtOrDefault((int)index) : null;
			if (hexColor != null)
			{
				solidFillModel.colorOption.hexColor = hexColor;
				return solidFillModel;
			}
			else
			{
				solidFillModel.colorOption.schemeColorModel = new SchemeColorModel()
				{
					themeColorValues = ThemeColorValues.ACCENT_1 + ((int)index % AccentColorCount),
				};
			}
			return solidFillModel;
		}
		private ColorOptionModel<SolidOptions> GetDataPointBorder(uint index, int seriesIndex)
		{
			ColorOptionModel<SolidOptions> solidFillModel = new ColorOptionModel<SolidOptions>();
			string hexColor = pieChartSetting.pieChartSeriesSettings.ElementAtOrDefault(seriesIndex) != null ? pieChartSetting.pieChartSeriesSettings.ElementAtOrDefault(seriesIndex).pieChartDataPointSettings
						.Select(item => item != null ? item.borderColor : null)
						.ToList().ElementAtOrDefault((int)index) : null;
			if (hexColor != null)
			{
				solidFillModel.colorOption.hexColor = hexColor;
				return solidFillModel;
			}
			else
			{
				solidFillModel.colorOption.schemeColorModel = new SchemeColorModel()
				{
					themeColorValues = ThemeColorValues.ACCENT_1 + ((int)index % AccentColorCount),
				};
			}
			return solidFillModel;
		}
		private C.PieChartSeries CreateChartSeries(int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			C.PieChartSeries series = new C.PieChartSeries(
				new C.Index { Val = new UInt32Value((uint)chartDataGrouping.id) },
				new C.Order { Val = new UInt32Value((uint)chartDataGrouping.id) },
				CreateSeriesText(chartDataGrouping.seriesHeaderFormula, new[] { chartDataGrouping.seriesHeaderCells }));
			for (uint index = 0; index < chartDataGrouping.xAxisCells.Length; index++)
			{
				C.DataPoint dataPoint = new C.DataPoint(new C.Index { Val = index }, new C.Bubble3D { Val = false });
				ShapePropertiesModel<SolidOptions, SolidOptions> shapePropertiesModel = new ShapePropertiesModel<SolidOptions, SolidOptions>()
				{
					fillColor = GetDataPointFill(index, seriesIndex)
				};
				if (pieChartSetting.pieChartType != PieChartTypes.DOUGHNUT)
				{
					shapePropertiesModel.lineColor = new OutlineModel<SolidOptions>()
					{
						lineColor = GetDataPointBorder(index, seriesIndex)
					};
				}
				dataPoint.Append(CreateChartShapeProperties(shapePropertiesModel));
				series.Append(dataPoint);
			}
			int dataLabelCellsLength = chartDataGrouping.dataLabelCells != null ? chartDataGrouping.dataLabelCells.Length : 0;
			List<C.DataLabel> dataLabelList = new List<C.DataLabel>();
			for (uint index = 0; index < chartDataGrouping.xAxisCells.Length; index++)
			{
				if (seriesIndex < pieChartSetting.pieChartSeriesSettings.Count)
				{
					PieChartDataLabel pieChartDataLabel = GetDataPointLabel(index, seriesIndex);
					if (pieChartDataLabel != null)
					{
						C.DataLabel dataLabel = CreateDataLabel(pieChartDataLabel, index);
						if (pieChartSetting.pieChartType != PieChartTypes.DOUGHNUT)
						{
							C.DataLabelPositionValues dataLabelPositionValues;
							switch (pieChartDataLabel.dataLabelPosition)
							{
								case PieChartDataLabel.DataLabelPositionValues.INSIDE_END:
									dataLabelPositionValues = C.DataLabelPositionValues.InsideEnd;
									break;
								case PieChartDataLabel.DataLabelPositionValues.OUTSIDE_END:
									dataLabelPositionValues = C.DataLabelPositionValues.OutsideEnd;
									break;
								case PieChartDataLabel.DataLabelPositionValues.BEST_FIT:
									dataLabelPositionValues = C.DataLabelPositionValues.BestFit;
									break;
								default:
									dataLabelPositionValues = C.DataLabelPositionValues.Center;
									break;
							}
							dataLabel.InsertBefore(new C.DataLabelPosition() { Val = dataLabelPositionValues }, dataLabel.Descendants<C.ShowLegendKey>().LastOrDefault());
						}
						dataLabelList.Add(dataLabel);
					}
				}
			}
			C.DataLabels dataLabels = CreatePieDataLabels(pieChartSetting.pieChartDataLabel, dataLabelCellsLength, dataLabelList.ToArray());
			series.Append(dataLabels);
			series.Append(CreateCategoryAxisData(chartDataGrouping.xAxisFormula, chartDataGrouping.xAxisCells));
			series.Append(CreateValueAxisData(chartDataGrouping.yAxisFormula, chartDataGrouping.yAxisCells));
			if (chartDataGrouping.dataLabelCells != null && chartDataGrouping.dataLabelFormula != null)
			{
				series.Append(new C.ExtensionList(new C.Extension(
					CreateDataLabelsRange(chartDataGrouping.dataLabelFormula, chartDataGrouping.dataLabelCells.Skip(1).ToArray())
				)
				{ Uri = "{02D57815-91ED-43cb-92C2-25804820EDAC}" }));
			}
			return series;
		}
		private C.DataLabels CreatePieDataLabels(PieChartDataLabel pieChartDataLabel, int? dataLabelCounter = 0, C.DataLabel[] dataLabelArray = null)
		{
			if (pieChartDataLabel.showValue || pieChartSetting.chartDataSetting.advancedDataLabel.showValueFromColumn || pieChartDataLabel.showCategoryName || pieChartDataLabel.showLegendKey || pieChartDataLabel.showSeriesName || pieChartDataLabel.showPercentage)
			{
				C.DataLabels dataLabels = CreateDataLabels(pieChartDataLabel, dataLabelCounter, dataLabelArray);
				if (pieChartSetting.pieChartType == PieChartTypes.DOUGHNUT &&
					new[] { PieChartDataLabel.DataLabelPositionValues.CENTER, PieChartDataLabel.DataLabelPositionValues.INSIDE_END, PieChartDataLabel.DataLabelPositionValues.OUTSIDE_END, PieChartDataLabel.DataLabelPositionValues.BEST_FIT }.Contains(pieChartDataLabel.dataLabelPosition))
				{
					throw new ArgumentException("DataLabelPosition is not supported for Doughnut Chart.");
				}
				if (pieChartSetting.pieChartType != PieChartTypes.DOUGHNUT)
				{
					C.DataLabelPositionValues dataLabelPositionValues;
					switch (pieChartDataLabel.dataLabelPosition)
					{
						case PieChartDataLabel.DataLabelPositionValues.INSIDE_END:
							dataLabelPositionValues = C.DataLabelPositionValues.InsideEnd;
							break;
						case PieChartDataLabel.DataLabelPositionValues.OUTSIDE_END:
							dataLabelPositionValues = C.DataLabelPositionValues.OutsideEnd;
							break;
						case PieChartDataLabel.DataLabelPositionValues.BEST_FIT:
							dataLabelPositionValues = C.DataLabelPositionValues.BestFit;
							break;
						default:
							dataLabelPositionValues = C.DataLabelPositionValues.Center;
							break;
					}
					dataLabels.InsertBefore(new C.DataLabelPosition() { Val = dataLabelPositionValues }, dataLabels.Descendants<C.ShowLegendKey>().LastOrDefault());
				}
				if (pieChartDataLabel.formatCode != null)
				{
					dataLabels.InsertAt(new C.NumberingFormat { FormatCode = pieChartDataLabel.formatCode, SourceLinked = false }, 0);
				}
				return dataLabels;
			}
			return null;
		}
	}
}