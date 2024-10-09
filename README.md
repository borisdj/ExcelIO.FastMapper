# ExcelIO.FastMapper
**Excel InputOutput Mapper** lib. for Import to and Export from between Poco class and .xlsx file using attribute annotation on model Properties.
In addition it has several useful confing options and Export also supports columns formatting.  
Both Reading and Writing are **very fast** while the package itself is lightweight with minimum depencencies.

Attributes:  
`[ExcelColumn]`

Under the hood libary uses most efficient packages in their domain:  
-[Sylvan.Data.Excel](https://github.com/MarkPflug/Sylvan.Data.Excel){:target="_blank"} for Reading  
-[LargeXlsx](https://github.com/salvois/LargeXlsx) for Writing as it has formatting option but is still pretty quick.  
Library has only those 2 dependecies that themselves are fully self-containd, and as such are quite thin.  
While doing research for optimal tool, other criteria were to be Open Source, with code on Github and having Nuget.  
Then to be actively maintained, have certain period of development with proven record of usage (Git commits, starts and Nuget downloads).  
