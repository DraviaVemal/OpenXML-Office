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

The `Chart` class, a versatile component within the `OpenXMLOffice.Presentation` library, empowers developers to seamlessly integrate various types of charts into PowerPoint presentations. This class supports multiple chart types and configurations, allowing users to add new charts to a slide or replace existing shapes with dynamic and data-driven visualizations.

<details>

<summary>List of supported charts</summary>

* [**Area Chart**](area.md) (2013) **:**
  * Cluster
  * Stacked
  * 100% Stacked

<!---->

* [**Bar Chart**](bar.md) (2013) **:**
  * Cluster
  * Stacked
  * 100% Stacked

<!---->

* [**Column Chart**](column.md) (2013) **:**
  * Cluster
  * Stacked
  * 100% Stacked

<!---->

* [**Line Chart**](line.md) (2013) **:**
  * Cluster
  * Stacked
  * 100% Stacked
  * Cluster Marker
  * Stacked Marker
  * 100% Stacked Marker

<!---->

* [**Pie Chart**](pie.md) (2013) **:**
  * Pie
  * Doughnut

<!---->

* [**X Y (Scatter) Chart**](scatter.md) (2013) **:**
  * Scatter
  * Scatter Smooth Line Marker
  * Scatter Smooth Line
  * Scatter Line Marker
  * Scatter Line
  * Bubble

<!---->

* [Combo Chart](combo.md) (2013) :&#x20;
  * [Area](area.md)
  * [Bar](bar.md)
  * [Column](column.md)
  * [Line](line.md)
  * [Pie](pie.md)
* [Waterfall Chart](waterfall.md) (2016)

</details>

### Basic Code Samples

&#x20;For each chart family `ChartSetting` have its releavent options and settings for customization.

{% tabs %}
{% tab title="C#" %}
```csharp
public void ChartSample(PowerPoint powerPoint)
{
    // Default Chart Type
    powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new AreaChartSetting());
    // Customised Chart Type
    powerPoint.GetSlideByIndex(0).AddChart(CreateDataCellPayload(), new AreaChartSetting()
    {
        AreaChartTypes = AreaChartTypes.STACKED
    });
    Slide slide = powerPoint.GetSlideByIndex(1);
    Shape shape = slide.FindShapeByText("shape_id_1");
    shape.ReplaceChart(new Chart(slide, CreateDataCellPayload(),
            new BarChartSetting()
            {
                ChartLegendOptions = new ChartLegendOptions()
                {
                    LegendPosition = ChartLegendOptions.LegendPositionValues.RIGHT
                }
            })
}
```
{% endtab %}
{% endtabs %}

### `ChartSetting` Options

<table><thead><tr><th width="218">Property</th><th width="205">Type</th><th>Details</th></tr></thead><tbody><tr><td>chartDataSetting</td><td><a href="./#chartdatasetting-options">ChartDataSetting</a></td><td>This setting enables users to customize both the input chart data range and value from cell labels with precision.</td></tr><tr><td>chartGridLinesOptions</td><td><a href="./#chartgridlinesoptions-options">ChartGridLinesOptions</a></td><td>This feature offers crisp options for users to finely customize the gridline settings of the chart.</td></tr><tr><td>chartLegendOptions</td><td><a href="./#chartlegendoptions-options">ChartLegendOptions</a></td><td>This feature offers crisp options for users to finely customize the gridline settings of the chart.</td></tr><tr><td>height</td><td>uint</td><td>This parameter precisely determines the height of the entire chart.<br>Default : 6858000</td></tr><tr><td>width</td><td>uint</td><td>This parameter precisely determines the width of the entire chart.<br>Default : 12192000</td></tr><tr><td>x</td><td>uint</td><td>This parameter precisely determines the X position of the entire chart.<br>Default: 0</td></tr><tr><td>y</td><td>uint</td><td>This parameter precisely determines the Y position of the entire chart.<br>Default : 0</td></tr></tbody></table>

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

Embedded excel can be accessed using `GetChartWorkBook` return OpenXMLOffice.Excel Worksheet. Refer [Worksheet](../../excel/worksheet.md) section for more details

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
