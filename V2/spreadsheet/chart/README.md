---
layout:
  title:
    visible: true
  description:
    visible: false
  tableOfContents:
    visible: true
  outline:
    visible: true
  pagination:
    visible: true
---

# Chart

The `Chart` class, a versatile component within the `OpenXMLOffice.Spreadsheet` library, empowers developers to seamlessly integrate various types of charts into Excel spreadsheet. This class supports multiple chart types and configurations, allowing users to add new charts to a sheet with dynamic and data-driven visualizations.

<details>

<summary>List of supported charts</summary>

* [**Area Chart**](area.md) (2007) **:**
  * Cluster
  * Stacked
  * 100% Stacked
  * Cluster 3D
  * Stacket 3D
  * 100% Stacked 3D

<!---->

* [**Bar Chart**](bar.md) (2007) **:**
  * Cluster
  * Stacked
  * 100% Stacked
  * Cluster 3D
  * Stacket 3D
  * 100% Stacked 3D

<!---->

* [**Column Chart**](column.md) (2007) **:**
  * Cluster
  * Stacked
  * 100% Stacked
  * Cluster 3D
  * Stacket 3D
  * 100% Stacked 3D

<!---->

* [**Line Chart**](line.md) (2007) **:**
  * Cluster
  * Stacked
  * 100% Stacked
  * Cluster Marker
  * Stacked Marker
  * 100% Stacked Marker

<!---->

* [**Pie Chart**](pie.md) (2007) **:**
  * Pie
  * Pie 3D
  * Doughnut

<!---->

* [**X Y (Scatter) Chart**](scatter-in-progress.md) (2007 - In Progress **:**
  * Scatter
  * Scatter Smooth Line Marker
  * Scatter Smooth Line
  * Scatter Line Marker
  * Scatter Line
  * Bubble

<!---->

* [Combo Chart](combo.md) (2007) :&#x20;
  * [Area](area.md)
  * [Bar](bar.md)
  * [Column](column.md)
  * [Line](line.md)
  * [Pie](pie.md)
* [Waterfall Chart](../../presentation/chart/waterfall.md) (2016) - In Progress

</details>

### Basic Code Samples

&#x20;For each chart family `ChartSetting<ExcelSetting>` have its releavent options and settings for customization.

{% tabs %}
{% tab title="C#" %}
```csharp
public void ChartSample(Excel excel)
{
	// Default Chart Type
	Excel excel1 = new("./TestFiles/basic_test.xlsx", true);
	Worksheet worksheet = excel1.AddSheet("AreaChart");
	worksheet.AddChart(new()
	{
		cellIdStart = "A1",
		cellIdEnd = "D4"
	}, new AreaChartSetting<ExcelSetting>()
	{
		applicationSpecificSetting = new()
		{
			from = new()
			{
				row = 5,
				column = 5
			},
			to = new()
			{
				row = 20,
				column = 20
			}
		}
	});
}
```
{% endtab %}
{% endtabs %}

### `ChartSetting<ExcelSetting>` Options

This section outlines the options available when configuring charts with `ChartSetting` using `ExcelSetting` parameters.

<table><thead><tr><th width="218">Property</th><th width="205">Type</th><th>Details</th></tr></thead><tbody><tr><td>isSecondaryAxis</td><td>bool</td><td>If combo chart this can be used to indicate secondary axis activation.</td></tr><tr><td>chartDataSetting</td><td><a href="./#chartdatasetting-options">ChartDataSetting</a></td><td>This setting enables users to customize both the input chart data range and value from cell labels with precision.</td></tr><tr><td>chartGridLinesOptions</td><td><a href="./#chartgridlinesoptions-options">ChartGridLinesOptions</a></td><td>This feature offers crisp options for users to finely customize the gridline settings of the chart.</td></tr><tr><td>chartLegendOptions</td><td><a href="./#chartlegendoptions-options">ChartLegendOptions</a></td><td>This feature offers crisp options for users to finely customize the gridline settings of the chart.</td></tr><tr><td>applicationSpecificSetting</td><td>&#x3C;ApplicationSpecificSetting></td><td>This is generic class setting. For Spreadsheet it is <a href="./#excelsetting-options"><code>ExcelSetting</code></a></td></tr></tbody></table>

### `ExcelSetting` Options

| Property | Type                                        | Details                                   |
| -------- | ------------------------------------------- | ----------------------------------------- |
| from     | [AnchorPosition](./#anchorposition-options) | Placement details for from starting point |
| to       | [AnchorPosition](./#anchorposition-options) | Placement details for to Ending point     |

### `AnchorPosition` Options

| Property     | Type | Details   |
| ------------ | ---- | --------- |
| column       | uint | Default:1 |
| columnOffset | uint | Default:0 |
| row          | uint | Default:1 |
| rowOffset    | uint | Default:0 |

### `ChartDataSetting` Options

<table><thead><tr><th width="218">Property</th><th width="192">Type</th><th>Details</th></tr></thead><tbody><tr><td>chartDataColumnEnd</td><td>uint</td><td>Specify the number of columns for chart series; set to 0 for utilizing all columns. <br>Default: 0</td></tr><tr><td>chartDataColumnStart</td><td>uint</td><td>Specify the starting column for chart data.<br>Default: 0</td></tr><tr><td>chartDataRowEnd</td><td>uint</td><td>Specify the number of rows for chart series; set to 0 for utilizing all rows. <br>Default: 0</td></tr><tr><td>chartDataRowStart</td><td>uint</td><td>Specify the starting row for chart data.<br>Default: 0</td></tr><tr><td>valueFromColumn</td><td>Dictionary&#x3C;uint, uint></td><td>This option allows configuring a key map where series corresponds to the key, and the value is mapped to a target column based on cell column configuration.</td></tr></tbody></table>

### `ChartGridLinesOptions` Options

<table><thead><tr><th width="276">Property</th><th width="91">Type</th><th>Details</th></tr></thead><tbody><tr><td>isMajorCategoryLinesEnabled</td><td>bool</td><td>Toggle visibility of major category lines with clarity.</td></tr><tr><td>isMajorValueLinesEnabled</td><td>bool</td><td>Toggle visibility of major value lines with clarity.</td></tr><tr><td>isMinorCategoryLinesEnabled</td><td>bool</td><td>Toggle visibility of minor category lines with clarity.</td></tr><tr><td>isMinorValueLinesEnabled</td><td>bool</td><td>Toggle visibility of minor value lines with clarity.</td></tr></tbody></table>

### `ChartLegendOptions` Options

<table><thead><tr><th width="220">Property</th><th width="196">Type</th><th>Details</th></tr></thead><tbody><tr><td>isEnableLegend</td><td>bool</td><td>Toggle visibility of legend with clarity.</td></tr><tr><td>isLegendChartOverLap</td><td>bool</td><td>Activate the option for a sleek and tidy display by allowing the legends to overlap.</td></tr><tr><td>isBold</td><td>bool</td><td>Provide the option to set text in a bold format with clarity.</td></tr><tr><td>isItalic</td><td>bool</td><td>Provide the option to set text in a italic format with clarity.</td></tr><tr><td>fontSize</td><td>float</td><td>Provide the option to set font size with clarity.</td></tr><tr><td>fontColor</td><td>string?</td><td>Optional font color using hex code (without #).<br>Default : Theme Text 1.</td></tr><tr><td>underLineValues</td><td>UnderLineValues</td><td>Text underline options. Default: None</td></tr><tr><td>strikeValues</td><td>StrikeValues</td><td>Text strike options</td></tr><tr><td>legendPosition</td><td>LegendPositionValues</td><td>Legend position in chart. Default: Bottom</td></tr></tbody></table>

### `ChartDataLabel` Options

This is base data label class extended by each chart type to give more specific/relavent options

<table><thead><tr><th width="226">Property</th><th width="160">Type</th><th>Details</th></tr></thead><tbody><tr><td>separator</td><td>string</td><td>Data lable text separator used if multiple label enabled</td></tr><tr><td>showCategoryName</td><td>bool</td><td>Show category name in label</td></tr><tr><td>showLegendKey</td><td>bool</td><td>Show legend key in label</td></tr><tr><td>showSeriesName</td><td>bool</td><td>Show series name in label</td></tr><tr><td>showValue</td><td>bool</td><td>Show value in label</td></tr><tr><td>showValueFromColumn</td><td>bool</td><td>Show value from different column in label</td></tr><tr><td>isBold</td><td>bool</td><td>Set label bold</td></tr><tr><td>isItalic</td><td>bool</td><td>Set label italic</td></tr><tr><td>fontSize</td><td>float</td><td>Set label font size</td></tr><tr><td>fontColor</td><td>string?</td><td>Set label font color</td></tr><tr><td>underLineValues</td><td>UnderLineValues</td><td>Set label underline type</td></tr><tr><td>strikeValues</td><td>StrikeValues</td><td>Set label strike type</td></tr></tbody></table>

### `ChartAxesOptions` Options

This properties give control over the X and Y axes. (Relate placement based on your chart option)

<table><thead><tr><th width="246">Property</th><th width="162">Type</th><th>Details</th></tr></thead><tbody><tr><td>invertVerticalAxesOrder</td><td>string?</td><td></td></tr><tr><td>invertHorizontalAxesOrder</td><td>string?</td><td></td></tr><tr><td>isHorizontalAxesEnabled</td><td>bool</td><td></td></tr><tr><td>isHorizontalBold</td><td>bool</td><td></td></tr><tr><td>isHorizontalItalic</td><td>bool</td><td></td></tr><tr><td>horizontalFontSize</td><td>float</td><td></td></tr><tr><td>horizontalFontColor</td><td>string?</td><td></td></tr><tr><td>horizontalUnderLineValues</td><td>UnderLineValues</td><td></td></tr><tr><td>horizontalStrikeValues</td><td>StrikeValues</td><td></td></tr><tr><td>isVerticalBold</td><td>bool</td><td></td></tr><tr><td>isVerticalItalic</td><td>bool</td><td></td></tr><tr><td>verticalFontSize</td><td>float</td><td></td></tr><tr><td>verticalFontColor</td><td>string?</td><td></td></tr><tr><td>verticalUnderLineValues</td><td>UnderLineValues</td><td></td></tr><tr><td>verticalStrikeValues</td><td>StrikeValues</td><td></td></tr><tr><td>isVerticalAxesEnabled</td><td>bool</td><td></td></tr></tbody></table>

### `ChartSeriesSetting` Options

| Property    | Type    | Details                                       |
| ----------- | ------- | --------------------------------------------- |
| borderColor | string? | Explicit border color for current data series |

### `ChartDataPointSettings` Options

| Property    | Type    | Details                                                       |
| ----------- | ------- | ------------------------------------------------------------- |
| fillColor   | string? | Explicit fill color for one specific data point in a series   |
| borderColor | string? | Explicit border color for one specific data point in a series |

### Embedded Excel Component

Embedded excel can be accessed using `GetChartWorkBook` return OpenXMLOffice.Excel Worksheet. Refer [Worksheet](../worksheet.md) section for more details

{% tabs %}
{% tab title="C#" %}
```csharp
Chart chart = powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
				.AddChart(CreateDataCellPayload(), new G.LineChartSetting());
Worksheet worksheet = chart.GetChartWorksheet();
worksheet.SetRow(12, 1, new DataCell[] { 
new() {
  cellValue = "Added Additional Data To Chart",
  dataType = CellDataType.STRING
  }
}, new());
```
{% endtab %}
{% endtabs %}
