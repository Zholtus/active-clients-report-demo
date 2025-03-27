Option Explicit

Dim man(100, 4) As String, bus_base(100) As String, i As Integer, j As Integer
Dim filename As String
Dim pathname As String
Public run_macro As Boolean
Public run_macro1 As Boolean
Dim str1 As String
Dim str2 As String
Dim str3 As String
Dim str4 As String
Dim q, k As String
Dim con As Variant


Sub RemoveConnectionsR1()
Dim wC As WorkbookConnection

For Each wC In ActiveWorkbook.Connections

        wC.ODBCConnection.Connection = "OLEDB;Provider=MSDAORA.1;Password=REDACTED;User ID=REDACTED;Data Source=REDACTED"
       
Next wC
End Sub

Sub addConnectionsR()
Dim wC As WorkbookConnection
   
    Open "c:\*****\start.txt" For Input As #1
    Dim s As String
    Input #1, s
    Close #1
    
For Each wC In ActiveWorkbook.Connections

        wC.ODBCConnection.Connection = s
       
Next wC
End Sub

Sub check()

If CheckDate = "1" And _
    Dir("c:\*****\start.txt") <> "" Then
        MsgBox ("OK")
On Error GoTo err1
    End If
err1:
    MsgBox ("NOT OK")
End Sub


Function CheckDate() As String

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sConnString As String
    Dim n As Long
    
    ' Считывание данных с start.txt
    Open "c:\*****\start.txt" For Input As #1
    Dim strS As String
    Input #1, strS
    Close #1
    
 
 
    ' Create the connection string.
    sConnString = strS '"Provider=ORAOLEDB.ORACLE;UseSessionFormat=True;Password=REDACTED & strPassword &";User ID="& strUserId &";Data Source=REDACTED;"
    
    ' Create the Connection and Recordset objects.
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    ' Open the connection and execute.
    conn.Open sConnString
    
    Set rs = conn.Execute("SELECT CASE WHEN d.report_date IS NULL THEN 0 ELSE CASE WHEN (select t.holiday from WORK_DAYS t WHERE t.datevalue = trunc(SYSDATE)) = 0 THEN 1 ELSE 0 END END AS RESULT FROM (select MAX(t.datevalue) last_work_day from WORK_DAYS t WHERE t.datevalue < SYSDATE -1 AND t.holiday = 0) t LEFT JOIN (SELECT MAX (d.report_date) report_date FROM active_kb_delta d) d ON d.report_date = t.last_work_day")
    
    ' Check we have data.
    If Not rs.EOF Then
        ' Transfer result.
     CheckDate = rs.Fields("Result").Value
    
        rs.Close
    Else
        MsgBox "Error: No records returned.", vbCritical
    End If
    
    
    If CBool(conn.State And adStateOpen) Then conn.Close
    Set conn = Nothing
    Set rs = Nothing

End Function



Public Sub SalesData()

If Dir("c:\*****\start.txt") <> "" Then
Dim i As Integer
Application.DisplayAlerts = False
On Error GoTo ERR


Call addConnectionsR

ActiveWorkbook.RefreshAll

pathname = "\\******.kz\******\Active Clients"

If Dir(pathname & "\Archive\", vbDirectory) = "" Then
    MkDir pathname & "\Archive\"
End If

' Копирование графика в тело письма
str1 = pathname & "\" & "Chart.jpg"
ActiveWorkbook.Sheets("График v.3").ChartObjects("Диаграмма2").Chart.Export str1

str4 = pathname & "\" & "Chart4.jpg"
ActiveWorkbook.Sheets("Крупный").ChartObjects("Диаграмма 1").Chart.Export str4

str2 = pathname & "\" & "Chart2.jpg"
ActiveWorkbook.Sheets("Средний").ChartObjects("Диаграмма 1").Chart.Export str2

str3 = pathname & "\" & "Chart3.jpg"
ActiveWorkbook.Sheets("ММБ").ChartObjects("Диаграмма 1").Chart.Export str3

ActiveWorkbook.Sheets("Лист1").Select
ActiveWorkbook.Sheets.Copy

Call RemoveConnectionsR1

filename = pathname & "\*****\Active_Clients" & Format(Date, "ddmmyyyy") & ".xlsx"
ActiveWorkbook.SaveAs filename, , Password:="", writeRespassword:=""

filename = pathname & "\Active_Clients" & Format(Date, "ddmmyyyy") & ".xlsx"
ActiveWorkbook.SaveAs filename, , Password:="", writeRespassword:=""
filename = "\\******\*****\Active_Clients" & Format(Date, "ddmmyyyy") & ".xlsx"

ActiveWorkbook.Close

Workbooks.Open filename:="C:\*****\*****\emails.xlsx"
Sheets("active_clients").Select
Range("A1").Select
i = 0
k = ""
q = ""
While Range("A" & i + 2) <> ""
  k = Range("B" & i + 2).Value & ";"
    q = q + k
    i = i + 1
    k = ""
Wend

ActiveWorkbook.Close

Call SendEmail(q)
bbb:
ThisWorkbook.Save

Application.Quit

Exit Sub

ERR:

Call SendEmail2("*****@*****.kz", "Active clients Error", ERR.Number & "<BR><BR>" & vbCrLf & "<BR><BR>" & ERR.Description & "<BR><BR>")

End If
End Sub

Sub ActiveGraffic()
'
' Макрос1 Макрос
'
Dim myR As Long

'найти последнюю строку и плясать от нее
    ActiveWorkbook.Sheets("chart").Select
    myR = Range("A2").End(xlDown).Row + 1


    Cells(myR, 1).FormulaR1C1 = "=Лист1!R2C3"
    Cells(myR, 2).FormulaR1C1 = "=SUM(Лист1!C[13])-RC[1]-RC[2]"
    Cells(myR, 3).FormulaR1C1 = "=SUMIF(Лист1!C[6],""C"",Лист1!C[12])"
    Cells(myR, 4).FormulaR1C1 = "=SUMIF(Лист1!C[5],""M"",Лист1!C[11])+SUMIF(Лист1!C[5],""R"",Лист1!C[11])+SUMIF(Лист1!C[5],""D"",Лист1!C[11])+SUMIF(Лист1!C[5],"""",Лист1!C[11])"
    Cells(myR, 5).FormulaR1C1 = "=RC[-1]+RC[-2]+RC[-3]"
    Range(Cells(myR, 1), Cells(myR, 5)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    ActiveSheet.Range(Cells(myR - 14, 1), Cells(myR, 5)).RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5), _
        Header:=xlYes
    myR = Range("A2").End(xlDown).Row + 1
    
    Range(Cells(myR - 14, 1), Cells(myR, 5)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A2").Select

End Sub



Private Function SendEmail(ByVal email)
', ByRef file1)

Dim oMSG
Dim oConfig
Dim CFields

Set oMSG = CreateObject("CDO.Message")
Set oConfig = CreateObject("CDO.Configuration")
Set CFields = oConfig.Fields
Set oMSG.Configuration = oConfig

CFields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'CFields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.*****.kz"
'CFields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 2
CFields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "post.*****.kz"
CFields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
CFields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "*****@*****.kz"
CFields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "*******"

CFields("urn:schemas:mailheader:content-language") = "windows-1251"

CFields.Update

oMSG.To = email

oMSG.from = "Name of our depertment <*****@*****.kz>"
oMSG.Subject = "Active clients on" & Format(Date, "dd.mm.yyyy")
oMSG.bodypart.Charset = "windows-1251"

oMSG.htmlbody = "<font color = black> <font face = calibri><font size = 3><b>Добрый день!</b><BR>" & _
"<BR>" & _
"текущий отчет <a href=" & filename & ">Active clients</a> has been updated<BR>" & _
"<BR>График <b>Corporate business</b>" & _
"<BR>" & "<img src = Chart.jpg>" & _
"<BR>" & _
"<BR>График по <b>Big business</b>" & _
"<BR>" & "<img src = Chart4.jpg>" & _
"<BR>" & _
"<BR>График по <b>Medium business</b>" & _
"<BR>" & "<img src = Chart2.jpg>" & _
"<BR>" & _
"<BR>График по <b>small business</b>" & _
"<BR>" & "<img src = Chart3.jpg>" & _
"<BR>" & _
"<a href=""\\*****\*****\Active Clients\Archive"">Archive</a> for previous days.<br>" & _
"<BR>" & _
"Folder access is restricted.<br>" & _
"<BR>" & _
"If your colleagues need access to the report, please use the rules<br>" & _
"<a href=""\\*****\*****\Rules_for_using_the_resource.docx"">Rules for using the resource</a> <br>" & _
"<BR>" & _
"<b>Best regards, my team</b>" & "<BR>" & _
"<BR>" & _
"<BR>"

oMSG.addattachment str1
oMSG.addattachment str2
oMSG.addattachment str3
oMSG.addattachment str4

oMSG.send

Set CFields = Nothing
Set oConfig = Nothing
Set oMSG = Nothing

End Function

Function RangeHTML(rng As Range)
Dim fso As Object
Dim ts As Object
Dim TempFile As String
Dim TempWB As Workbook

    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
rng.Copy

Set TempWB = Workbooks.Add(1)
With TempWB.Sheets(1)
    .Cells(1).PasteSpecial Paste:=8
    .Cells(1).PasteSpecial xlPasteValues, , False, False
    .Cells(1).PasteSpecial xlPasteFormats, , False, False
    .Cells(1).Select
    Application.CutCopyMode = False
    On Error GoTo 0
End With

With TempWB.PublishObjects.Add( _
    SourceType:=xlSourceRange, _
    filename:=TempFile, _
    Sheet:=TempWB.Sheets(1).Name, _
    Source:=TempWB.Sheets(1).UsedRange.Address, _
    HtmlType:=xlHtmlStatic)
    .Publish (True)
End With


Set fso = CreateObject("scripting.FileSystemObject")
Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
RangeHTML = ts.ReadAll
ts.Close
RangeHTML = Replace(RangeHTML, "align=center x:publishsoucre=", _
"align-left x:publishsource=")

TempWB.Close savechanges:=False

Kill TempFile

Set ts = Nothing
Set fso = Nothing
Set TempWB = Nothing


End Function


Private Function SendEmail2(ByVal email, ByVal subj, ByVal body)

Dim oMSG
Dim oConfig
Dim CFields

Set oMSG = CreateObject("CDO.Message")
Set oConfig = CreateObject("CDO.Configuration")
Set CFields = oConfig.Fields
Set oMSG.Configuration = oConfig

CFields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
CFields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "post.*****.kz"
'CFields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "***.**.**.***"
CFields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
CFields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "*****@*****.kz"
CFields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "******"
CFields("urn:scemas:mailheader:content-language") = "windows-1251"

CFields.Update

oMSG.bcc = email
oMSG.from = "Центр анализа и развития данных <*****@*****.kz>"
oMSG.Subject = subj
oMSG.bodypart.Charset = "windows-1251"

oMSG.htmlbody = body
oMSG.send

Set CFields = Nothing
Set oConfig = Nothing
Set oMSG = Nothing

End Function

