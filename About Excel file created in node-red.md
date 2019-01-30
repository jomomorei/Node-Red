The purpose i used node-red is to create a flow can auto fetch data from salesforce and upload them to a BI cloud serve 
named MotionBoard Cloud. 

![Complete Image](/img/flow.png)

MotionBoard is a powerful BI tool can analyze and visualize business data, similar to Power BI.

The question is the csv file i get from node-red do not support Japanese which i have mentioned in the other post.I have tried times to change the encoding before the data were exported. But it didn't work. But when run on the another PC, the Japanese was recognized. The most possible reason is the different system setting of PCs.   

The Excel file could get clean data. However, when it was uploaded to MotionBoard cloud, the file became unreadable. While it was opened and saved once, magically, it came to be readable. 

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
