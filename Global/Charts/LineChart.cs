using C = DocumentFormat.OpenXml.Drawing.Charts;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class LineChart : LineFamilyChart
    {
        #region Public Methods

        public C.ChartSpace GetChartSpace(ChartData[][] DataCols, GlobalConstants.LineChartTypes LineChartType, LineChartSetting chartSetting)
        {
            // Apply Properties
            ChartGridLinesOptions = chartSetting.ChartGridLinesOptions;
            C.Chart Chart = CreateChart(chartSetting);
            Chart.PlotArea = LineChartType switch
            {
                GlobalConstants.LineChartTypes.CLUSTERED_MARKER => CreateChartPlotArea(DataCols, C.GroupingValues.Standard, chartSetting, true),
                GlobalConstants.LineChartTypes.STACKED_MARKER => CreateChartPlotArea(DataCols, C.GroupingValues.Stacked, chartSetting, true),
                GlobalConstants.LineChartTypes.PERCENT_STACKED_MARKER => CreateChartPlotArea(DataCols, C.GroupingValues.PercentStacked, chartSetting, true),
                GlobalConstants.LineChartTypes.STACKED => CreateChartPlotArea(DataCols, C.GroupingValues.Stacked, chartSetting),
                GlobalConstants.LineChartTypes.PERCENT_STACKED => CreateChartPlotArea(DataCols, C.GroupingValues.PercentStacked, chartSetting),
                // Clusted
                _ => CreateChartPlotArea(DataCols, C.GroupingValues.Standard, chartSetting),
            };
            GetChartSpace().Append(Chart);
            return GetChartSpace();
        }

        public CS.ChartStyle GetChartStyle()
        {
            return CreateChartStyles();
        }

        public CS.ColorStyle GetColorStyle()
        {
            return CreateColorStyles();
        }

        #endregion Public Methods
    }
}