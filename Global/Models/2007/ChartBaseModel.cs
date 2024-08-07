// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using C = DocumentFormat.OpenXml.Drawing.Charts;
using OpenXMLOffice.Global_2013;
using System;


namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// 
	/// </summary>
	public enum AxesLabelPosition
	{
		/// <summary>
		/// 
		/// </summary>
		NEXT_TO_AXIS,
		/// <summary>
		/// 
		/// </summary>
		LOW,
		/// <summary>
		/// 
		/// </summary>
		HIGH,
		/// <summary>
		/// 
		/// </summary>
		NONE
	}
	/// <summary>
	/// Represents the position of an axis in a chart.
	/// </summary>
	public enum AxisPosition
	{
		/// <summary>
		/// Top
		/// </summary>
		TOP,
		/// <summary>
		/// Bottom
		/// </summary>
		BOTTOM,
		/// <summary>
		/// Left
		/// </summary>
		LEFT,
		/// <summary>
		/// Right
		/// </summary>
		RIGHT
	}
	/// <summary>
	///
	/// </summary>
	public enum MarkerShapeTypes
	{
		/// <summary>
		///
		/// </summary>
		NONE,
		/// <summary>
		///
		/// </summary>
		AUTO,
		/// <summary>
		///
		/// </summary>
		CIRCLE,
		/// <summary>
		///
		/// </summary>
		DASH,
		/// <summary>
		///
		/// </summary>
		DIAMOND,
		/// <summary>
		///
		/// </summary>
		DOT,
		/// <summary>
		///
		/// </summary>
		PICTURE,
		/// <summary>
		///
		/// </summary>
		PLUS,
		/// <summary>
		///
		/// </summary>
		SQUARE,
		/// <summary>
		///
		/// </summary>
		STAR,
		/// <summary>
		///
		/// </summary>
		TRIANGLE,
		/// <summary>
		///
		/// </summary>
		X
	}
	/// <summary>
	/// Text direction in charts
	/// </summary>
	public enum ChartTextDirectionValues
	{
		/// <summary>
		/// 
		/// </summary>
		HORIZONTAL,
		/// <summary>
		/// 
		/// </summary>
		ROTATE_90,
		/// <summary>
		/// 
		/// </summary>
		ROTATE_270,
		/// <summary>
		/// 
		/// </summary>
		STACKED
	}
	/// <summary>
	/// 
	/// </summary>
	public enum ChartVerticalTextAlignmentValues
	{
		/// <summary>
		/// 
		/// </summary>
		RIGHT,
		/// <summary>
		/// 
		/// </summary>
		CENTER,
		/// <summary>
		/// 
		/// </summary>
		LEFT,
		/// <summary>
		/// 
		/// </summary>
		RIGHT_MIDDLE,
		/// <summary>
		/// 
		/// </summary>
		CENTER_MIDDLE,
		/// <summary>
		/// 
		/// </summary>
		LEFT_MIDDLE,
		/// <summary>
		/// 
		/// </summary>
		TOP,
		/// <summary>
		/// 
		/// </summary>
		MIDDLE,
		/// <summary>
		/// 
		/// </summary>
		BOTTOM,
		/// <summary>
		/// 
		/// </summary>
		TOP_CENTER,
		/// <summary>
		/// 
		/// </summary>
		MIDDLE_CENTER,
		/// <summary>
		/// 
		/// </summary>
		BOTTOM_CENTER,
	}
	/// <summary>
	/// 
	/// </summary>
	public class ChartTextOptions : TextOptions
	{
		/// <summary>
		/// TODO
		/// </summary>
		internal ChartTextDirectionValues textDirectionValue = ChartTextDirectionValues.HORIZONTAL;
		internal ChartVerticalTextAlignmentValues chartVerticalTextAlignmentValue;
		/// <summary>
		/// TODO
		/// </summary>
		internal ChartVerticalTextAlignmentValues ChartVerticalTextAlignmentValue
		{
			get
			{
				return chartVerticalTextAlignmentValue;
			}
			set
			{
				switch (textDirectionValue)
				{
					case ChartTextDirectionValues.HORIZONTAL:
						if (value == ChartVerticalTextAlignmentValues.LEFT ||
							value == ChartVerticalTextAlignmentValues.CENTER ||
							value == ChartVerticalTextAlignmentValues.RIGHT ||
							value == ChartVerticalTextAlignmentValues.LEFT_MIDDLE ||
							value == ChartVerticalTextAlignmentValues.CENTER_MIDDLE ||
							value == ChartVerticalTextAlignmentValues.RIGHT_MIDDLE)
						{
							throw new Exception("Selected Text Configuration is not acceptable");
						}
						break;
					default:
						if (value == ChartVerticalTextAlignmentValues.TOP ||
						value == ChartVerticalTextAlignmentValues.MIDDLE ||
						value == ChartVerticalTextAlignmentValues.BOTTOM ||
						value == ChartVerticalTextAlignmentValues.TOP_CENTER ||
						value == ChartVerticalTextAlignmentValues.MIDDLE_CENTER ||
						value == ChartVerticalTextAlignmentValues.BOTTOM_CENTER)
						{
							throw new Exception("Selected Text Configuration is not acceptable");
						}
						break;
				}
				chartVerticalTextAlignmentValue = value;
			}
		}
		private int textAngle = 0;
		/// <summary>
		/// Set Text Angle between -90 to 90 degree
		/// </summary>
		public int TextAngle
		{
			get
			{
				return textAngle;
			}
			set
			{
				if (value > 90)
				{
					textAngle = 90;
				}
				else if (value < -90)
				{
					textAngle = -90;
				}
				else
				{
					textAngle = value;
				}
			}
		}
	}
	/// <summary>
	/// Represents the options for the axes in a chart.
	/// </summary>
	public class ChartAxesLabel : ChartTextOptions
	{
		/// <summary>
		/// Axis Label Position.
		/// </summary>
		public AxesLabelPosition axesLabelPosition = AxesLabelPosition.NEXT_TO_AXIS;
		/// <summary>
		/// Invert the axis order
		/// </summary>
		public bool inReverseOrder = false;
	}
	/// <summary>
	/// Represents the options for the axis title text
	/// </summary>
	public class ChartAxisTitle : ChartTextOptions { }

	/// <summary>
	/// Common Chart Axis Options
	/// </summary>
	public class AxisOptions<AxisType> where AxisType : class, IAxisTypeOptions, new()
	{
		internal C.CrossesValues crosses = C.CrossesValues.AutoZero;
		internal C.TickMarkValues majorTickMark = C.TickMarkValues.None;
		internal C.TickMarkValues minorTickMark = C.TickMarkValues.None;
		/// <summary>
		/// This Are options Related to Type of axis. CategoryAxis or Value Axis
		/// </summary>
		public AxisType axisTypeOption = new AxisType();
		/// <summary>
		/// Is Horizontal Axes Enabled
		/// </summary>
		public bool isAxesVisible = true;
		/// <summary>
		/// axis line color for the chart
		/// </summary>
		public string axisLineColor;
		/// <summary>
		/// Option for Axis's Axes label options
		/// </summary>
		public ChartAxesLabel chartAxesOptions = new ChartAxesLabel();
		/// <summary>
		/// Option for Axis title options
		/// </summary>
		public ChartAxisTitle chartAxisTitle = new ChartAxisTitle();

		internal static C.TickLabelPositionValues GetLabelAxesPosition(AxesLabelPosition axesLabelPosition)
		{
			switch (axesLabelPosition)
			{
				case AxesLabelPosition.LOW:
					return C.TickLabelPositionValues.Low;
				case AxesLabelPosition.HIGH:
					return C.TickLabelPositionValues.High;
				case AxesLabelPosition.NONE:
					return C.TickLabelPositionValues.None;
				default:
					return C.TickLabelPositionValues.NextTo;
			}
		}
	}
	/// <summary>
	/// 
	/// </summary>
	public interface IAxisTypeOptions { }
	/// <summary>
	/// 
	/// </summary>
	public class ValueAxis : IAxisTypeOptions
	{
		/// <summary>
		/// 
		/// </summary>
		public float? boundsMinimum;
		/// <summary>
		/// 
		/// </summary>
		public float? boundsMaximum;
		/// <summary>
		/// 
		/// </summary>
		public float? unitsMajor;
		/// <summary>
		/// 
		/// </summary>
		public float? unitsMinor;

	}
	/// <summary>
	/// 
	/// </summary>
	public class CategoryAxis : IAxisTypeOptions
	{
		/// <summary>
		/// 
		/// </summary>
		public uint? specificIntervalUnit;
	}
	/// <summary>
	/// X Axis Specific Options
	/// </summary>
	public class XAxisOptions<AxisType> : AxisOptions<AxisType> where AxisType : class, IAxisTypeOptions, new() { }
	/// <summary>
	/// Y Axis Specific Options
	/// </summary>
	public class YAxisOptions<AxisType> : AxisOptions<AxisType> where AxisType : class, IAxisTypeOptions, new() { }
	/// <summary>
	/// Z Axis Specific Options
	/// </summary>
	public class ZAxisOptions<AxisType> : AxisOptions<AxisType> where AxisType : class, IAxisTypeOptions, new() { }
	/// <summary>
	/// Represents the options for a chart axis.
	/// Pass Each Axis Dimension Type for more accurate options
	/// </summary>
	public class ChartAxisOptions<XAxisType, YAxisType, ZAxisType>
		where XAxisType : class, IAxisTypeOptions, new()
	 	where YAxisType : class, IAxisTypeOptions, new()
	  	where ZAxisType : class, IAxisTypeOptions, new()
	{
		/// <summary>
		/// X-Axis and Axes options
		/// </summary>
		public XAxisOptions<XAxisType> xAxisOptions = new XAxisOptions<XAxisType>();
		/// <summary>
		/// Y-Axis and Axes options
		/// </summary>
		public YAxisOptions<YAxisType> yAxisOptions = new YAxisOptions<YAxisType>();
		/// <summary>
		/// Z-Axis and Axes options
		/// Totally optional for secondary action options
		/// TODO : Implementation
		/// </summary>
		public ZAxisOptions<ZAxisType> zAxisOptions = new ZAxisOptions<ZAxisType>();
	}
	/// <summary>
	/// Represents the grouping options for chart data.
	/// </summary>
	public class ChartDataGrouping
	{
		/// <summary>
		/// Gets or sets the data label cells.
		/// </summary>
		public ChartData[] dataLabelCells;
		/// <summary>
		/// Gets or sets the data label formula.
		/// </summary>
		public string dataLabelFormula;
		/// <summary>
		/// Gets or sets the series header cells.
		/// </summary>
		public ChartData seriesHeaderCells;
		/// <summary>
		/// Gets or sets the series header format.
		/// </summary>
		public string seriesHeaderFormat;
		/// <summary>
		/// Gets or sets the series header formula.
		/// </summary>
		public string seriesHeaderFormula;
		/// <summary>
		/// Gets or sets the X-axis cells.
		/// </summary>
		public ChartData[] xAxisCells;
		/// <summary>
		/// Gets or sets the X-axis format.
		/// </summary>
		public string xAxisFormat;
		/// <summary>
		/// Gets or sets the X-axis formula.
		/// </summary>
		public string xAxisFormula;
		/// <summary>
		/// Gets or sets the Y-axis cells.
		/// </summary>
		public ChartData[] yAxisCells;
		/// <summary>
		/// Gets or sets the Y-axis format.
		/// </summary>
		public string yAxisFormat;
		/// <summary>
		/// Gets or sets the Y-axis formula.
		/// </summary>
		public string yAxisFormula;
		/// <summary>
		/// Gets or sets the Z-axis cells.
		/// </summary>
		public ChartData[] zAxisCells;
		/// <summary>
		/// Gets or sets the Z-axis format.
		/// </summary>
		public string zAxisFormat;
		/// <summary>
		/// Gets or sets the Z-axis formula.
		/// </summary>
		public string zAxisFormula;
		/// <summary>
		///
		/// </summary>
		public int id;
	}
	/// <summary>
	/// Represents the options for chart data labels.
	/// </summary>
	public class ChartDataLabel : TextOptions
	{
		/// <summary>
		/// The separator used for displaying multiple values.
		/// </summary>
		public string separator = ", ";
		/// <summary>
		/// Determines whether to show the category name in the chart.
		/// </summary>
		public bool showCategoryName;
		/// <summary>
		/// Determines whether to show the legend key in the chart.
		/// </summary>
		public bool showLegendKey;
		/// <summary>
		/// Determines whether to show the series name in the chart.
		/// </summary>
		public bool showSeriesName;
		/// <summary>
		/// Determines whether to show the value in the chart.
		/// </summary>
		public bool showValue;
		/// <summary>
		/// Determines whether to show the value in percentage the chart.
		/// </summary>
		public bool showPercentage;
		/// <summary>
		/// Determines whether to format the label values in the chart
		/// </summary>
		public string formatCode;
	}
	/// <summary>
	/// Options to select data range and custom row column data placements
	/// </summary>
	public class ChartDataSetting
	{
		/// <summary>
		/// Chart Chart Category column
		/// </summary>
		public uint chartCategoryColumn = 0;
		/// <summary>
		/// Set 0 To Use Till End
		/// </summary>
		public uint chartDataColumnEnd = 0;
		/// <summary>
		/// Chart data Start column 0 based
		/// </summary>
		public uint chartDataColumnStart = 1;
		/// <summary>
		/// Set 0 To Use Till End
		/// </summary>
		public uint chartDataRowEnd = 0;
		/// <summary>
		/// Chart data Start Row 0 based
		/// </summary>
		public uint chartDataRowStart = 0;
		/// <summary>
		/// Is Data is used in 3D chart type
		/// </summary>
		internal bool is3dData;
		/// <summary>
		/// Use 2013 Version Data Label Option
		/// </summary>
		/// <remarks>This Property May get updated in future be in lookout.</remarks>
		public AdvancedDataLabel advancedDataLabel = new AdvancedDataLabel();
	}
	/// <summary>
	/// Represents the options for chart grid lines.
	/// </summary>
	public class ChartGridLinesOptions
	{
		/// <summary>
		/// Is Major Category Lines Enabled
		/// </summary>
		public bool isMajorCategoryLinesEnabled;
		/// <summary>
		/// Is Major Value Lines Enabled
		/// </summary>
		public bool isMajorValueLinesEnabled = true;
		/// <summary>
		/// Is Minor Category Lines Enabled
		/// </summary>
		public bool isMinorCategoryLinesEnabled;
		/// <summary>
		/// Is Minor Value Lines Enabled
		/// </summary>
		public bool isMinorValueLinesEnabled;
	}
	/// <summary>
	/// Represents the options for chart legend.
	/// </summary>
	public class ChartLegendOptions : TextOptions
	{
		/// <summary>
		/// Is Legend Enabled
		/// </summary>
		public bool isEnableLegend = true;
		/// <summary>
		/// Is Legend Chart OverLap
		/// </summary>
		public bool isLegendChartOverLap;
		/// <summary>
		/// Legend Position
		/// </summary>
		public LegendPositionValues legendPosition = LegendPositionValues.BOTTOM;
		/// <summary>
		/// Legend Position Values
		/// </summary>
		public enum LegendPositionValues
		{
			/// <summary>
			/// Bottom
			/// </summary>
			BOTTOM,
			/// <summary>
			/// Top
			/// </summary>
			TOP,
			/// <summary>
			/// Left
			/// </summary>
			LEFT,
			/// <summary>
			/// Right
			/// </summary>
			RIGHT,
			/// <summary>
			/// Top Right
			/// </summary>
			TOP_RIGHT
		}
		/// <summary>
		/// Manual Position legend
		/// </summary>
		public LayoutModel manualLayout;
	}
	/// <summary>
	///
	/// </summary>
	public class ChartDataPointSettings
	{
		/// <summary>
		/// The color of the fill.
		/// </summary>
		public string fillColor;
		/// <summary>
		///
		/// </summary>
		public string borderColor;
	}
	/// <summary>
	/// Represents the settings for a chart series.
	/// </summary>
	public class ChartSeriesSetting
	{
		/// <summary>
		/// The color of the border.
		/// </summary>
		public virtual string borderColor { get; set; }
		internal ChartSeriesSetting() { }
	}
	/// <summary>
	///
	/// </summary>
	public class PlotAreaModel
	{
		/// <summary>
		/// Manual Position Char Graph Area
		/// </summary>
		public LayoutModel manualLayout;
	}
	/// <summary>
	///
	/// </summary>
	public class ChartTitleModel : ChartTextOptions { }
	/// <summary>
	///
	/// </summary>
	public class AnchorPosition
	{
		/// <summary>
		///
		/// </summary>
		public uint column = 1;
		/// <summary>
		///
		/// </summary>
		public uint columnOffset = 0;
		/// <summary>
		///
		/// </summary>
		public uint row = 1;
		/// <summary>
		///
		/// </summary>
		public uint rowOffset = 0;
	}
	/// <summary>
	///
	/// </summary>
	public interface ISizeAndPosition { }
	/// <summary>
	///
	/// </summary>
	public class PresentationSetting : ISizeAndPosition
	{
		/// <summary>
		/// Value in EMU
		/// </summary>
		private int _height = 6858000;
		private int _width = 12192000;
		/// <summary>
		/// Value in EMU
		/// </summary>
		private int _y = 0;
		/// <summary>
		/// Value in EMU
		/// </summary>
		private int _x = 0;
		/// <summary>
		/// Chart Height in Px
		/// </summary>
		public int Height
		{
			get
			{
				return _height;
			}
			set
			{
				// _height = (int)ConverterUtils.PixelsToEmu(value);
				_height = value;
			}
		}
		/// <summary>
		/// Chart Width in Px
		/// </summary>
		public int Width
		{
			get
			{
				return _width;
			}
			set
			{
				// _width = (int)ConverterUtils.PixelsToEmu(value);
				_width = value;
			}
		}
		/// <summary>
		/// Chart X Position in Px
		/// </summary>
		public int X
		{
			get
			{
				return _x;
			}
			set
			{
				// _x = (int)ConverterUtils.PixelsToEmu(value);
				_x = value;
			}
		}
		/// <summary>
		/// Chart Y Position in Px
		/// </summary>
		public int Y
		{
			get
			{
				return _y;
			}
			set
			{
				// _y = (int)ConverterUtils.PixelsToEmu(value);
				_y = value;
			}
		}
	}
	/// <summary>
	///
	/// </summary>
	public class ExcelSetting : ISizeAndPosition
	{
		/// <summary>
		///
		/// </summary>
		public AnchorPosition from = new AnchorPosition();
		/// <summary>
		///
		/// </summary>
		public AnchorPosition to = new AnchorPosition();
	}
	/// <summary>
	/// Represents the settings for a chart.
	/// </summary>
	public class ChartSetting<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition, new()
	{
		internal uint? categoryAxisId;
		internal uint? valueAxisId;
		internal bool is3DChart;
		/// <summary>
		///
		/// </summary>
		public HyperlinkProperties hyperlinkProperties;
		/// <summary>
		/// Only useful when used with Combo chart
		/// </summary>
		public bool isSecondaryAxis;
		/// <summary>
		///
		/// </summary>
		public PlotAreaModel plotAreaOptions;
		/// <summary>
		/// Chart Data Setting
		/// </summary>
		public ChartDataSetting chartDataSetting = new ChartDataSetting();
		/// <summary>
		/// Chart Grid Line Options
		/// </summary>
		public ChartGridLinesOptions chartGridLinesOptions = new ChartGridLinesOptions();
		/// <summary>
		/// Chart Legend Options
		/// </summary>
		public ChartLegendOptions chartLegendOptions = new ChartLegendOptions();
		/// <summary>
		/// Chart Title
		/// </summary>
		public ChartTitleModel titleOptions;
		/// <summary>
		///
		/// </summary>
		public ApplicationSpecificSetting applicationSpecificSetting = new ApplicationSpecificSetting();
		internal ChartSetting() { }
	}

	/// <summary>
	/// Represents the Axis settings for chart.
	/// </summary>
	public class AxisSetting<AxisTypeOption, AxisType>
		where AxisTypeOption : AxisOptions<AxisType>, new()
		where AxisType : class, IAxisTypeOptions, new()
	{
		internal uint id;
		internal uint crossAxisId;
		internal AxisPosition axisPosition;
		internal AxisTypeOption axisOptions = new AxisTypeOption();
	}
	/// <summary>
	///
	/// </summary>
	public class LayoutModel
	{
		/// <summary>
		/// Considered from left to right
		/// Value is between 0 to 1
		/// </summary>
		public float x = 0;
		/// <summary>
		/// Considered from top to bottom
		/// Value is between 0 to 1
		/// </summary>
		public float y = 0;
		/// <summary>
		/// Value is between 0 to 1
		/// </summary>
		public float width = 1;
		/// <summary>
		/// Value is between 0 to 1
		/// </summary>
		public float height = 1;
	}
	/// <summary>
	/// 
	/// </summary>
	public class TrendLineModel
	{
		/// <summary>
		/// Use for Order value if 'Polynomial' type else Period for 'Moving Average'
		/// Default 2
		/// </summary>
		public int secondaryValue = 2;
		/// <summary>
		/// 
		/// </summary>
		public bool setIntercept = false;
		/// <summary>
		/// 
		/// </summary>
		public float interceptValue = 0.0F;
		/// <summary>
		/// 
		/// </summary>
		public bool showEquation = false;
		/// <summary>
		/// 
		/// </summary>
		public bool showRSquareValue = false;
		/// <summary>
		/// Default 0.0
		/// </summary>
		public float forecastForward = 0.0F;
		/// <summary>
		/// Default 0.0
		/// </summary>
		public float forecastBackward = 0.0F;
		/// <summary>
		/// This is to set custom Trending if null it will assume automatic 
		/// </summary>
		public string trendLineName = null;
		/// <summary>
		/// 
		/// </summary>
		public DrawingPresetLineDashValues drawingPresetLineDashValues = DrawingPresetLineDashValues.SYSTEM_DOT;
		/// <summary>
		/// 
		/// </summary>
		public TrendLineTypes trendLineType = TrendLineTypes.NONE;
		/// <summary>
		/// 
		/// </summary>
		public ColorOptionModel<SolidOptions> solidFill = new ColorOptionModel<SolidOptions>();
		internal static C.TrendlineValues GetTrendlineValues(TrendLineTypes trendLineType)
		{
			switch (trendLineType)
			{
				case TrendLineTypes.EXPONENTIAL:
					return C.TrendlineValues.Exponential;
				case TrendLineTypes.LINEAR:
					return C.TrendlineValues.Linear;
				case TrendLineTypes.LOGARITHMIC:
					return C.TrendlineValues.Logarithmic;
				case TrendLineTypes.POLYNOMIAL:
					return C.TrendlineValues.Polynomial;
				case TrendLineTypes.POWER:
					return C.TrendlineValues.Power;
				case TrendLineTypes.MOVING_AVERAGE:
					return C.TrendlineValues.MovingAverage;
				default:
					return C.TrendlineValues.MovingAverage;
			}
		}
	}
	/// <summary>
	///
	/// </summary>
	public class MarkerModel<LineColorOption, FillColorOption>
	where LineColorOption : class, IColorOptions, new()
	where FillColorOption : class, IColorOptions, new()
	{
		/// <summary>
		/// Market Size
		/// </summary>
		public int size = 5;
		/// <summary>
		///
		/// </summary>
		public MarkerShapeTypes markerShapeType = MarkerShapeTypes.NONE;
		/// <summary>
		///
		/// </summary>
		public ShapePropertiesModel<LineColorOption, FillColorOption> shapeProperties = new ShapePropertiesModel<LineColorOption, FillColorOption>();
		internal static C.MarkerStyleValues GetMarkerStyleValues(MarkerShapeTypes markerShapeType)
		{
			switch (markerShapeType)
			{
				case MarkerShapeTypes.AUTO:
					return C.MarkerStyleValues.Auto;
				case MarkerShapeTypes.CIRCLE:
					return C.MarkerStyleValues.Circle;
				case MarkerShapeTypes.DASH:
					return C.MarkerStyleValues.Dash;
				case MarkerShapeTypes.DIAMOND:
					return C.MarkerStyleValues.Diamond;
				case MarkerShapeTypes.DOT:
					return C.MarkerStyleValues.Dot;
				case MarkerShapeTypes.PICTURE:
					return C.MarkerStyleValues.Picture;
				case MarkerShapeTypes.PLUS:
					return C.MarkerStyleValues.Plus;
				case MarkerShapeTypes.SQUARE:
					return C.MarkerStyleValues.Square;
				case MarkerShapeTypes.STAR:
					return C.MarkerStyleValues.Star;
				case MarkerShapeTypes.TRIANGLE:
					return C.MarkerStyleValues.Triangle;
				case MarkerShapeTypes.X:
					return C.MarkerStyleValues.X;
				default:
					return C.MarkerStyleValues.None;
			}
		}
	}
}
