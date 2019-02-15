The purpose i used node-red is to create a flow can auto fetch data from salesforce and upload them to a BI cloud serve 
named MotionBoard Cloud. 

![Complete Image](/img/flow.png)

MotionBoard is a powerful BI tool can analyze and visualize business data. The data sources can be various of database, excel and csv.

The question is the csv file i get from node-red do not recognize Japanese. About this I have tried some tests. It is related to PC enviroment.   

The Excel file could get clean data. But, when it was uploaded to MotionBoard cloud, the file became unreadable. While it was opened and saved once, it came to be readable magically. 

I mailed the MotionBoard company about the error and got the reply: the difference of XML namespace specify method between the excel files created in excel application and in node-red.

It seems a difficult problem. 

In the robotic process some simple VBA scripts can be useful: Open the origin excel file and saved it regularlly. To avoid error, copy the origin file as a new file is recommended.

```VBA Script
Private Sub Workbook_Open()

NewTime = Now + TimeValue("00:00:10")

Application.OnTime NewTime, "ThisWorkbook.openSave"

End Sub


Sub openSave()  're-edit

Dim dataExcel, Workbook, sheet
Set dataExcel = CreateObject("Excel.Application")

NewTime = Now + TimeValue("00:10:00")

Set Workbook = dataExcel.Workbooks.Open(ThisWorkbook.Path & "\noderedtest.xlsx")

Workbook.Save

Workbook.Close

Call CopyFile 'change path to avert read-in error

Application.OnTime NewTime, "ThisWorkbook.openSave"
  
End Sub

Sub CopyFile()

Dim fs As Object

Dim strFile As String

Dim strNewFile As String

strFile = ThisWorkbook.Path & "\noderedtest.xlsx"

strNewFile = ThisWorkbook.Path & "\MBデータ\noderedtest.xlsx"

Set fs = CreateObject("Scripting.FileSystemObject")

fs.CopyFile strFile, strNewFile


Set fs = Nothing

End Sub

```
