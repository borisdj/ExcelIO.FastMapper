# ExcelIO.FastMapper
**Excel InputOutput Mapper** lib. for Import to and Export from Poco class and .xlsx file using attribute annotation on model Properties.
In addition it has several useful confing options and Export also supports columns formatting.  
Both Reading and Writing are **very fast** while the package itself is lightweight with minimum dependencies.

Logo  
<img src="ExcelIO.png" height=60>

[![NuGet](https://img.shields.io/npm/l/express.svg)](https://github.com/borisdj/ExcelIO.FastMapper/blob/master/LICENSE)  

Also take a look into others packages:</br>
-Open source (MIT or cFOSS) authored [.Net libraries](https://infopedia.io/dot-net-libraries/) (@**Infopedia.io** personal blog post)
| №  | .Net library             | Description                                              |
| -  | ------------------------ | -------------------------------------------------------- |
| 1  | [EFCore.BulkExtensions](https://github.com/borisdj/EFCore.BulkExtensions) | EF Core Bulk CRUD Ops (**Flagship** Lib) |
| 2  | [EFCore.UtilExtensions](https://github.com/borisdj/EFCore.UtilExtensions) | EF Core Custom Annotations and AuditInfo |
| 3  | [EFCore.FluentApiToAnnotation](https://github.com/borisdj/EFCore.FluentApiToAnnotation) | Converting FluentApi configuration to Annotations |
| 4* | [ExcelIO.FastMapper](https://github.com/borisdj/ExcelIO.FastMapper) | Excel Input Output Mapper to-from Poco & .xlsx with attribute |
| 5  | [FixedWidthParserWriter](https://github.com/borisdj/FixedWidthParserWriter) | Reading & Writing fixed-width/flat data files |
| 6  | [CsCodeGenerator](https://github.com/borisdj/CsCodeGenerator) | C# code generation based on Classes and elements |
| 7  | [CsCodeExample](https://github.com/borisdj/CsCodeExample) | Examples of C# code in form of a simple tutorial |

## Support
If you find this project useful you can mark it by leaving a Github **Star** :star:  
And even with community license, if you want help development, you can make a DONATION:  
[<img src="https://www.buymeacoffee.com/assets/img/custom_images/yellow_img.png" alt="Buy Me A Coffee" height=28>](https://www.buymeacoffee.com/boris.dj) _ or _ 
[![Button](https://img.shields.io/badge/donate-Bitcoin-orange.svg?logo=bitcoin):zap:](https://borisdj.net/donation/donate-btc.html)

## Contributing
Please read [CONTRIBUTING](CONTRIBUTING.md) for details on code of conduct, and the process for submitting pull requests.  
When opening issues do write detailed explanation of the problem or feature with reproducible example.  
Want to **Contact** for Development & Consulting: [www.codis.tech](http://www.codis.tech) (*Quality Delivery*) 

## Configuration
**Excel-IO Mapper config**:  
```C#
PROPERTY : DEFAULTvalue
----------------------------------------------
 1 FileName, 
 2 SheetName: "Data",
 3 UseDefaultColumnFormat: true,
 4 AutoFilterVisible: true,
 5 UseDynamicColumnWidth: true,
 6 WrapHeader: false,
 7 DynamicColumnWidthCoefficient: false,
 8 FreezeHeader: true,
 9 FreezeColumnNumber: true,
10 HeaderFont: null,
11 DataFont: null,
12 ExportOnlyPropertiesWithAttribute: null,
13 Dictionary<string, ExcelIOColumnAttribute> DynamicSettings
-----------------------------------------------
```

**ExcelIO Column Attribute** : defaultValue
```C#
bool Ignore : false ................ // Field omitted from Excel
string Header : 0 .................. // Header Name
string Format : null ............... // Column format
int Order : new List<string>() ..... // Position in column orders
int Width .......................... // Column width
```
*-Special feature is '**Dynamic Settings**' with which Attributes values can be defined at runtime, for all usage types.  

Under the hood library uses most efficient packages in their domain:  
-[Sylvan.Data.Excel](https://github.com/MarkPflug/Sylvan.Data.Excel) for Reading  
-[LargeXlsx](https://github.com/salvois/LargeXlsx) for Writing as it has formatting option and is still pretty quick.  
Library has only those 2 dependecies that themselves are fully self-containd, and as such are quite thin.  
(LargeXsql has transitive dependency on *SharpCompress* which is somewhat bigger ~1 MB).  

While doing research for optimal tool, criteria were to be Open Source, with code on Github and having Nuget.  
Also to be actively maintained, have certain period of development with proven record of usage (Git commits, starts and Nuget downloads).  
Comparison of several packages for the optimal and fastest one:  
[ExcelIO.NetLibs Compare](https://docs.google.com/spreadsheets/d/1rF4QEoDmTLB4cbbVL575276vhnfhyfX-KxGk-rcJAiA/edit?gid=0#gid=0)
