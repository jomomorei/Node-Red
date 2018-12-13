The purpose i used node-red is to create a flow can auto fetch data from salesforce and upload them to a BI cloud serve 
named MotionBoard Cloud. 

![Complete Image](/img/flow.png)

MotionBoard is a powerful BI tool can analyze and visualize business data, similar to Power BI.

The question is the csv file i get from node-red do not support Japanese which i have mentioned in the other post.I have tried times to change the encoding before the data were exported. But it didn't work. Because the export node file fixed the final encoding. Japanese is not in the consideration.

The best solution is to rewrite the output file node. However, it may spend a little more time on it. Maybe in future.

On the other hand, the Excel file seems to have got clean data. However, when i upload it to MotionBoard, the file became unreadable. it made me a little sad. Then i open it and saved it, magically, it came to be readable. 

I mailed the MotionBoard company about the error and got the reason: the difference of XML namespace specify method between the excel files created in excel application and in node-red.

Then I wrote some VBA commands to make the excel file opened and saved regularlly into a new folder at the scheduled time.
```
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
