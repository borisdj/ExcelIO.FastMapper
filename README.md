# ExcelIO.FastMapper
Excel IO Mapper Import and Export to and from Poco class and .xlsx file using attribute annotation on model Properties and having columns formatting.  
Both Reading and Writing are very fast while the package is lightweight with minimum depencencies.

Attributes:  
[ExcelColumn]

Under the hood libary uses most efficient packages in their domain: [Sylvan.Data.Excel] for Reading, and [LargeXlsx] for Writing as it has formatting option but is still quick.  
IT has only those 2 dependecies that them selfves are fully selfcontaind, and as such are pretty thin.  
While doing research for optimal tool, other criterias were to be Open Source, with code on Github and Nuget package.  
Then to have active maintenance, and certain period of development with proven record of usage (Git commits, starts and Nuget downloads).  
