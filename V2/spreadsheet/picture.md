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

# Picture

Details about adding and manipulating picture to a worksheet

### Basic Code Sample

```csharp
public void AddPicture()
	{
		Worksheet worksheet = excel.AddSheet("Data4");
		worksheet.AddPicture("./TestFiles/tom_and_jerry.jpg", new()
		{
			imageType = ImageType.JPEG,
			from = new()
			{
				column = 6,
				row = 6
			},
			to = new()
			{
				column = 8,
				row = 8
			}
		});
		Assert.IsTrue(true);
	}
```

### `ExcelPictureSetting` Options

<table><thead><tr><th width="197">Property</th><th width="203">Type</th><th>Details</th></tr></thead><tbody><tr><td>hyperlinkProperties</td><td><a href="picture.md#hyperlinkproperties-options">HyperlinkProperties</a></td><td></td></tr><tr><td>imageType</td><td>ImageType</td><td></td></tr><tr><td>anchorEditType</td><td>AnchorEditType</td><td></td></tr><tr><td>from</td><td><a href="picture.md#anchorposition-options">AnchorPosition</a></td><td></td></tr><tr><td>to</td><td><a href="picture.md#anchorposition-options">AnchorPosition</a></td><td></td></tr></tbody></table>

### `HyperlinkProperties` Options

| Property              | Type                  | Details |
| --------------------- | --------------------- | ------- |
| hyperlinkPropertyType | HyperlinkPropertyType |         |
| value                 | string                |         |
| toolTip               | string                |         |

### `AnchorPosition` Options

| Property     | Type | Details |
| ------------ | ---- | ------- |
| column       | uint |         |
| columnOffset | uint |         |
| row          | uint |         |
| rowOffset    | uint |         |
