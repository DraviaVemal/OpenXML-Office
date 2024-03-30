---
description: >-
  The page furnishes comprehensive installation instructions, guiding you
  through the process of incorporating the package dependency.
---

# Getting Started

## Feel free to start discussion for any new feature requirement. [Discussion Channel](https://github.com/DraviaVemal/OpenXMLOffice/discussions)

{% tabs %}
{% tab title="C#" %}
The library is available on NuGet. You can install it using the following command

```shell
#Using Package Manager
Install-Package OpenXMLOffice
```

<pre class="language-shell"><code class="lang-shell">#Using .NET CLI
<strong>dotnet add package OpenXMLOffice.Presentation
</strong></code></pre>

```bash
# For Pre Release
dotnet add package OpenXMLOffice.Presentation --prerelease
```

```bash
#Using .NET CLI
dotnet add package OpenXMLOffice.Excel
```

```bash
# For Pre Release
dotnet add package OpenXMLOffice.Excel--prerelease
```
{% endtab %}
{% endtabs %}

### Package Version Details

{% tabs %}
{% tab title="C#" %}
The official release NuGet packages for OpenXMLOffice on NuGet.org:

| Package                    | Dev Status | Download                                                                                                                             | Prerelease                                                                                                                              |
| -------------------------- | ---------- | ------------------------------------------------------------------------------------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------------------- |
| OpenXMLOffice.Presentation | Active     | [![NuGet](https://img.shields.io/nuget/v/OpenXMLOffice.Presentation.svg)](https://www.nuget.org/packages/OpenXMLOffice.Presentation) | [![NuGet](https://img.shields.io/nuget/vpre/OpenXMLOffice.Presentation.svg)](https://www.nuget.org/packages/OpenXMLOffice.Presentation) |
| OpenXMLOffice.Excel        | Active     | [![NuGet](https://img.shields.io/nuget/v/OpenXMLOffice.Excel.svg)](https://www.nuget.org/packages/OpenXMLOffice.Excel)               | [![NuGet](https://img.shields.io/nuget/vpre/OpenXMLOffice.Excel.svg)](https://www.nuget.org/packages/OpenXMLOffice.Excel)               |
| OpenXMLOffice.Document     | Not Active | [![NuGet](https://img.shields.io/nuget/v/OpenXMLOffice.Document.svg)](https://www.nuget.org/packages/OpenXMLOffice.Document)         | [![NuGet](https://img.shields.io/nuget/vpre/OpenXMLOffice.Document.svg)](https://www.nuget.org/packages/OpenXMLOffice.Document)         |
{% endtab %}
{% endtabs %}

Once Installed the package should be direct use availabel like below example. More samples can be seen in test project's of the repo or check other parts of the documents

{% tabs %}
{% tab title="C#" %}
```csharp
using OpenXMLOffice.Presentation;

public static main(){
    PowerPoint powerPoint = new(string.Format("../../test-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")), null);
    powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
    powerPoint.save();
}
```
{% endtab %}
{% endtabs %}

{% tabs %}
{% tab title="C#" %}
```csharp
using OpenXMLOffice.Excel;

public static main(){
    Spreadsheet spreadsheet = new(string.Format("../../test-{0}.xlsx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")));
    spreadsheet.AddSheet("Sheet1");
    spreadsheet.save();
}
```
{% endtab %}
{% endtabs %}
