# ExcelTemplate

**ExcelTemplate** is a project for exporting Excel with the Excel template, which is based on [NPOI](http://npoi.codeplex.com). All you need is to provide a **DataSet** Object, and to config the file **ExportConfig.xml**, which define the logic where and how to write the Data from the DataSet.
  
All you should provide is as follow:
  
1. **Excel templates**. The application just simply copy cell styles in the excles and fullfil data to avoid lots of lines of codes for apearance properties setting in excel.
  
2. The **DataSet** object. Tell application what data you need to display.
   
3. The **ExportConfig.xml** file. Defne how to disply data in memory after copying the templates of excel. The *Exceltemplae* defines many common components and rules for rendering data.
   
