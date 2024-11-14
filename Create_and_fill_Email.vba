Sub RunButton()

ActiveWorkbook.save

Call CreateEmail_Range

Application.Wait (Now + TimeValue("0:00:01"))

End Sub

Sub CreateEmail_Range()
    Dim rng As Range
    Dim OutApp As Object
    Dim OutMail As Object
    Dim OutAttach As Object
    Dim filePath As String
    Dim StrBody As String
    
    Dim sht As Worksheet
    Dim lastrow As Long
    Dim LastColumn As Long
    Dim StartCell As Range

    Set sht = Worksheets("Copy")
    Set StartCell = Worksheets("Copy").Range("D4")

    Set rng = Nothing
    
    Worksheets("Copy").Activate

'Find Last Row and Column
  lastrow = sht.Cells(sht.Rows.count, StartCell.Column).End(xlUp).Row
  LastColumn = sht.Cells(StartCell.Row, sht.Columns.count).End(xlToLeft).Column

    
    On Error Resume Next
    'Only the visible cells in the selection
        'Set rng = Selection.SpecialCells(xlCellTypeVisible)
    'Fixed cell range
    Set rng = sht.Range(StartCell, sht.Cells(lastrow, LastColumn))

    On Error GoTo 0

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    filePath = ThisWorkbook.Path & "\Carelon Manual Report - " & Format(Date, "mmmm") & ".xlsm"
    On Error Resume Next
    With OutMail
        .To = "Sean.Harrish@CVSHealth.com;Anna.Riccardino@CVSHealth.com;Jerrald.Andreason@CVSHealth.com;Scott.Miller@CVSHealth.com;_IngenioRxPCSSupervisors@CVSHealth.com"
        .CC = "Rohan.Bhardwaj@CVSHealth.com;Joshua.Mier@CVSHealth.com;Devon.Scott@CVSHealth.com;Jolene.Colon@CVSHealth.com;siddharth.ardalkar@cvshealth.com;Demorise.Abron@CVSHealth.com;Kristen.Corradetti@CVSHealth.com;Michael.Erana@CVSHealth.com;Christopher.Esterbrook@CVSHealth.com;Lacy.Long@CVSHealth.com;Jamerson.Lovett@CVSHealth.com;Marissa.Maddy@CVSHealth.com;Amanda.Martinez5@CVSHealth.com;Rush.Miller@CVSHealth.com;Mirelle.Pereda@CVSHealth.com;Sandra.Pritts@CVSHealth.com>;RaShika.Skeen@CVSHealth.com;Michelle.Waltos@CVSHealth.com;Frank.Thomas@cvshealth.com;Douglass.Shotwell@cvshealth.com;Crystal.Boyer@CVSHealth.com"
        .BCC = ""
        .Subject = "Carelon Manual SV LV Report " & Format(Date, "MM/DD/YY") & " @ " & CalculateAdjustedTime & " CST"
        If Dir(ThisWorkbook.Path & "\Carelon Manual Report - " & Format(Date, "mmmm") & ".xlsm") <> "" Then
            Set OutAttach = .Attachments.Add(ThisWorkbook.Path & "\Carelon Manual Report - " & Format(Date, "mmmm") & ".xlsm")
        Else
                MsgBox "File not found", vbExclamation
        End If
        StrBody = "<br/>" & "<br/>" & "<font face=""Calibri"" size=""4"" color=""black"">" & "Thank you," & _
                "<br/>" & _
                "<br/>" & "<font face=""Calibri"" size=""4"" color=""black"">" & "Resource Planning|" & "<font face=""Calibri"" size=""4"" color=""red"">Specialty Pharmacy Operations, CarelonRx " & _
                "<br/>" & "<img src='C:\Users\c104237\Pictures\Camera Roll\cvs logo.png'>"
        .HTMLBody = "<font face=""Calibri"" size=""4"" color=""black"">" & _
        "Current totals and completion percentages for this interval's Carelon SV LV @ " & CalculateAdjustedTime & " CST" & RangetoHTML(rng) & StrBody
        '.Attachments.Add "C:\Users\c104237\Documents\Ingenio\ingenioManualStats_v3.xlsm"
        .Display
        
    End With
    On Error GoTo 0

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
    Set OutAttach = Nothing
Worksheets("Report").Activate
 
 
  
End Sub



Function RangetoHTML(rng As Range)

    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
   End With

    'Publish the sheet to an htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close SaveChanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

Function CalculateAdjustedTime() As String
    Dim currTime As Date
    Dim reportTime As Date
    Dim halfHour As Date
    Dim adjSubject As String
    
    currTime = Now
    reportTime = currTime
    reportTime = DateAdd("h", -1, currTime)
    
    ' Get the nearest half-hour time
    halfHour = TimeValue(Format(reportTime, "hh") & ":" & IIf(Minute(reportTime) < 30, "00", "30"))

    ' Adjust the time if within 10 minutes of the next half-hour
    Dim timeDiff As Long
    If timeDiff >= -10 And timeDiff <= 0 Then
        halfHour = DateAdd("n", 30, halfHour)
    End If
    
    ' Construct the subject based on the adjusted time
    adjustedSubject = Format(halfHour, "h:mm AM/PM")
    
    ' Return the adjusted subject
    CalculateAdjustedTime = adjustedSubject
End Function
