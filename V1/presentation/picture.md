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

### Basic Code Samples

```csharp
powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
    .AddPicture("./TestFiles/tom_and_jerry.jpg", new G.PictureSetting());
```

`PictureSetting` Options

| Property  | Type      | Details |
| --------- | --------- | ------- |
| imageType | ImageType |         |
| height    | uint      |         |
| width     | uint      |         |
| x         | uint      |         |
| y         | uint      |         |
