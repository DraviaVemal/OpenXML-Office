---
description: Details about adding and manipulating picture to a slide
layout:
  title:
    visible: true
  description:
    visible: true
  tableOfContents:
    visible: true
  outline:
    visible: true
  pagination:
    visible: true
---

# Picture

### Basic Code Sample

```csharp
powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
    .AddPicture("./TestFiles/tom_and_jerry.jpg", new G.PictureSetting());
```

### `PictureSetting` Options

| Property            | Type                                                          | Details                     |
| ------------------- | ------------------------------------------------------------- | --------------------------- |
| hyperlinkProperties | [HyperlinkProperties](picture.md#hyperlinkproperties-options) | Hyperlink propertie setting |
| imageType           | ImageType                                                     | Inserted Image Type         |
| height              | uint                                                          | Image Height                |
| width               | uint                                                          | Image Width                 |
| x                   | uint                                                          | Image Top Left X            |
| y                   | uint                                                          | Image Top Left Y            |

### `HyperlinkProperties` Options

| Property              | Type                  | Details                   |
| --------------------- | --------------------- | ------------------------- |
| hyperlinkPropertyType | HyperlinkPropertyType | Type of hyperlink options |
| value                 | string                | Hyperlink address value   |
| toolTip               | string                | Hyperlink tooltip         |
