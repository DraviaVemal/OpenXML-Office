// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.Collections.Generic;
namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// Represents the types of area charts.
	/// </summary>
	public enum AreaChartTypes
	{
		/// <summary>
		/// Clustered area chart.
		/// </summary>
		CLUSTERED,
		/// <summary>
		/// Stacked area chart.
		/// </summary>
		STACKED,
		/// <summary>
		/// Percent stacked area chart.
		/// </summary>
		PERCENT_STACKED,
		/// <summary>
		///
		/// </summary>
		CLUSTERED_3D,
		/// <summary>
		///
		/// </summary>
		STACKED_3D,
		/// <summary>
		///
		/// </summary>
		PERCENT_STACKED_3D
	}
	/// <summary>
	/// Represents the data label settings for an area chart.
	/// </summary>
	public class AreaChartDataLabel : ChartDataLabel
	{
		/// <summary>
		/// The position of the data labels.
		/// </summary>
		public DataLabelPositionValues dataLabelPosition = DataLabelPositionValues.SHOW;
		/// <summary>
		/// The possible values for the data label position.
		/// </summary>
		public enum DataLabelPositionValues
		{
			/// <summary>
			/// Data Label option display type
			/// </summary>
			SHOW,
			// /// <summary>
			// /// Select Data Call out as Data label style
			// /// </summary>
			// DATA_CALLOUT
		}
	}
	/// <summary>
	/// Represents the series settings for an area chart.
	/// </summary>
	public class AreaChartSeriesSetting : ChartSeriesSetting
	{
		/// <summary>
		/// Treadline setting for particular series
		/// </summary>
		public List<TrendLineSettings> trendLines = new List<TrendLineSettings>();
		/// <summary>
		/// Option to customize specific data series. This will override the chart level setting.
		/// </summary>
		public AreaChartDataLabel areaChartDataLabel = new AreaChartDataLabel();
		/// <summary>
		/// The color of the fill.
		/// </summary>
		public string fillColor;
	}
	/// <summary>
	/// Represents the settings for an area chart.
	/// </summary>
	public class AreaChartSetting<ApplicationSpecificSetting> : ChartSetting<ApplicationSpecificSetting>
		where ApplicationSpecificSetting : class, ISizeAndPosition, new()
	{
		/// <summary>
		/// The data label settings for the entire chart. This will get overridden by series level setting.
		/// </summary>
		public AreaChartDataLabel areaChartDataLabel = new AreaChartDataLabel();
		/// <summary>
		/// The series settings for the area chart.
		/// </summary>
		public List<AreaChartSeriesSetting> areaChartSeriesSettings = new List<AreaChartSeriesSetting>();
		/// <summary>
		/// The type of the area chart.
		/// </summary>
		public AreaChartTypes areaChartType = AreaChartTypes.CLUSTERED;
		/// <summary>
		/// The options for the axis of the chart.
		/// </summary>
		public ChartAxisOptions<CategoryAxis, ValueAxis, ValueAxis> chartAxisOptions = new ChartAxisOptions<CategoryAxis, ValueAxis, ValueAxis>();
	}
}
