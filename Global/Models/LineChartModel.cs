// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Global {
    /// <summary>
    /// Represents the types of line charts.
    /// </summary>
    public enum LineChartTypes {
        /// <summary>
        /// Clustered Line Chart
        /// </summary>
        CLUSTERED,

        /// <summary>
        /// Stacked Line Chart
        /// </summary>
        STACKED,

        /// <summary>
        /// Percent Stacked Line Chart
        /// </summary>
        PERCENT_STACKED,

        /// <summary>
        /// Clustered Marker Line Chart
        /// </summary>
        CLUSTERED_MARKER,

        /// <summary>
        /// Stacked Marker Line Chart
        /// </summary>
        STACKED_MARKER,

        /// <summary>
        /// Percent Stacked Marker Line Chart
        /// </summary>
        PERCENT_STACKED_MARKER,

        // CLUSTERED_3D
    }

    /// <summary>
    /// Represents the data label settings for a line chart.
    /// </summary>
    public class LineChartDataLabel : ChartDataLabel {
        #region Public Fields

        /// <summary>
        /// The position of the data labels.
        /// </summary>
        public DataLabelPositionValues dataLabelPosition = DataLabelPositionValues.CENTER;

        #endregion Public Fields

        #region Public Enums

        /// <summary>
        /// The possible positions for the data labels.
        /// </summary>
        public enum DataLabelPositionValues {
            /// <summary>
            /// Left Side
            /// </summary>
            LEFT,

            /// <summary>
            /// Right Side
            /// </summary>
            RIGHT,

            /// <summary>
            /// Center
            /// </summary>
            CENTER,

            /// <summary>
            /// Above
            /// </summary>
            ABOVE,

            /// <summary>
            /// Below
            /// </summary>
            BELOW,

            /// <summary>
            /// Data Callout
            /// </summary>
            DATA_CALLOUT
        }

        #endregion Public Enums
    }

    /// <summary>
    /// Represents the series settings for a line chart.
    /// </summary>
    public class LineChartSeriesSetting : ChartSeriesSetting {
        #region Public Fields

        /// <summary>
        /// The color of the series border.
        /// </summary>
        public string? borderColor;

        /// <summary>
        /// The color of the series fill.
        /// </summary>
        public string? fillColor;

        /// <summary>
        /// Option to customize specific data series, which will override the chart level setting.
        /// </summary>
        public LineChartDataLabel lineChartDataLabel = new();

        #endregion Public Fields
    }

    /// <summary>
    /// Represents the settings for a line chart.
    /// </summary>
    public class LineChartSetting : ChartSetting {
        #region Public Fields

        /// <summary>
        /// The options for the chart axes.
        /// </summary>
        public ChartAxesOptions chartAxesOptions = new();

        /// <summary>
        /// The options for the chart axis.
        /// </summary>
        public ChartAxisOptions chartAxisOptions = new();

        /// <summary>
        /// The data label settings for the line chart, which will get overridden by series level setting.
        /// </summary>
        public LineChartDataLabel lineChartDataLabel = new();

        /// <summary>
        /// The series settings for the line chart.
        /// </summary>
        public List<LineChartSeriesSetting?> lineChartSeriesSettings = new();

        /// <summary>
        /// The type of the line chart.
        /// </summary>
        public LineChartTypes lineChartTypes = LineChartTypes.CLUSTERED;

        #endregion Public Fields
    }
}