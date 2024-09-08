---
description: >-
  The page furnishes comprehensive installation instructions, guiding you
  through the process of incorporating the package dependency.
---

# Getting Started

Feel free to start discussion for any new feature requirement. [Discussion Channel](https://github.com/DraviaVemal/OpenXMLOffice/discussions)

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
dotnet add package OpenXMLOffice.Spreadsheet
```

```bash
# For Pre Release
dotnet add package OpenXMLOffice.Spreadsheet --prerelease
```
{% endtab %}
{% endtabs %}

### Package Quality

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/5b420a599805426ab8a990a1a741247a)](https://app.codacy.com/gh/DraviaVemal/OpenXML-Office/dashboard?utm\_source=gh\&utm\_medium=referral\&utm\_content=\&utm\_campaign=Badge\_grade) [![Codacy Badge](https://app.codacy.com/project/badge/Coverage/5b420a599805426ab8a990a1a741247a)](https://app.codacy.com/gh/DraviaVemal/OpenXML-Office/dashboard?utm\_source=gh\&utm\_medium=referral\&utm\_content=\&utm\_campaign=Badge\_coverage)

### Package Version Details

{% tabs %}
{% tab title="C#" %}
The official release NuGet packages for OpenXMLOffice on NuGet.org:

<table><thead><tr><th width="272">Package</th><th width="120" align="center">Maintenance</th><th width="145" align="center">Current Stable Release</th><th align="center">Current Alpha Release</th></tr></thead><tbody><tr><td>OpenXMLOffice.Presentation</td><td align="center">Active</td><td align="center"><a href="https://www.nuget.org/packages/OpenXMLOffice.Presentation"><img src="https://img.shields.io/nuget/v/OpenXMLOffice.Presentation.svg" alt="NuGet"></a></td><td align="center"><a href="https://www.nuget.org/packages/OpenXMLOffice.Presentation"><img src="https://img.shields.io/nuget/vpre/OpenXMLOffice.Presentation.svg" alt="NuGet"></a></td></tr><tr><td>OpenXMLOffice.Spreadsheet</td><td align="center">Active</td><td align="center"><a href="https://www.nuget.org/packages/OpenXMLOffice.Presentation"><img src="https://img.shields.io/nuget/v/OpenXMLOffice.Spreadsheet.svg" alt="NuGet"></a></td><td align="center"><a href="https://www.nuget.org/packages/OpenXMLOffice.Presentation"><img src="https://img.shields.io/nuget/vpre/OpenXMLOffice.Spreadsheet.svg" alt="NuGet"></a></td></tr><tr><td>OpenXMLOffice.Document</td><td align="center">NA</td><td align="center"><a href="https://www.nuget.org/packages/OpenXMLOffice.Document"><img src="https://img.shields.io/nuget/v/OpenXMLOffice.Document.svg" alt="NuGet"></a></td><td align="center"><a href="https://www.nuget.org/packages/OpenXMLOffice.Document"><img src="https://img.shields.io/nuget/vpre/OpenXMLOffice.Document.svg" alt="NuGet"></a></td></tr></tbody></table>
{% endtab %}
{% endtabs %}

Once Installed the package should be direct use availabel like below example. More samples can be seen in test project's of the repo or check other parts of the documents

{% tabs %}
{% tab title="C#" %}
```csharp
using OpenXMLOffice.Presentation_2007;

public static main(){
    PowerPoint powerPoint = new();
    powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
    powerPoint.SaveAs(string.Format("../../test-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")));
}
```
{% endtab %}
{% endtabs %}

{% tabs %}
{% tab title="C#" %}
```csharp
using OpenXMLOffice.Spreadsheet_2007;

public static main(){
    Excel excel = new();
    excel.AddSheet("Sheet1");
    excel.SaveAs(string.Format("../../test-{0}.xlsx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")));
}
```
{% endtab %}
{% endtabs %}
