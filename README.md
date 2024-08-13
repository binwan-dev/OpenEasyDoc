# OpenEasyDoc
OpenXml Wrapper  for excel word ...

## Install  
`dotnet add package OpenEasyDoc`  
  
## Usage  
### Excel  
1. Export data to excel
```
var excelDoc = new EasyExcelDocument("your-excel-path");
var sheet = excelDoc.CreateSheet(1, "sheet-name");
sheet.NewRow().WriteCellData(1).WriteCellData("2").EndRow();
sheet.Close();
excelDoc.Close().Dispose();
```
