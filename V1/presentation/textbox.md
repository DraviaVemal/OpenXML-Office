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

# Textbox

Textbox Control to add and update Text Box

{% tabs %}
{% tab title="C#" %}
```csharp
shapes3[0].ReplaceTextBox(new TextBox(new G.TextBoxSetting()
			{
				text = "This is text box",
				fontSize = 22,
				isBold = true,
				textColor = "AAAAAA"
			}));
```
{% endtab %}
{% endtabs %}

### `TextBoxSetting` Options

<table><thead><tr><th width="204">Property</th><th>Type</th><th>Details</th></tr></thead><tbody><tr><td>horizontalAlignment</td><td>HorizontalAlignmentValues?</td><td></td></tr><tr><td>fontFamily</td><td>string</td><td></td></tr><tr><td>fontSize</td><td>int</td><td></td></tr><tr><td>x</td><td>uint</td><td></td></tr><tr><td>y</td><td>uint</td><td></td></tr><tr><td>height</td><td>uint</td><td></td></tr><tr><td>width</td><td>uint</td><td></td></tr><tr><td>isBold</td><td>bool</td><td></td></tr><tr><td>isItalic</td><td>bool</td><td></td></tr><tr><td>isUnderline</td><td>bool</td><td></td></tr><tr><td>shapeBackground</td><td>string?</td><td></td></tr><tr><td>text</td><td>string</td><td></td></tr><tr><td>textBackground</td><td>string?</td><td></td></tr><tr><td>textColor</td><td>string</td><td></td></tr></tbody></table>
