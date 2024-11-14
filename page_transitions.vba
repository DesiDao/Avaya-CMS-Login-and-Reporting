Sub radarShift()
If Module3.showKey("radar") Then
    cleanShift "radar"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Report")
    
    With ws
        .Range("S19").Value = "RADAR"
        .Range("S19:V19").Merge
        .Range("S19:V19").Borders.LineStyle = xlContinuous
        .Range("S19:V19").Borders.Weight = xlMedium
        .Range("S20").Value = "Cache?"
        .Range("S20:T20").Merge
        .Range("S20:T20").Borders.LineStyle = xlContinuous
        .Range("S20:T20").Borders.Weight = xlMedium
        .Range("U20").Value = (Dir("C:\Users\c289894\AppData\Roaming\Avaya\CMS Supervisor R19\Cache", vbDirectory) <> "")
        .Range("U20").Borders.LineStyle = xlContinuous
        .Range("U20").Borders.Weight = xlMedium
        .Range("V20").Value = "-"
        .Range("V20").Borders.LineStyle = xlContinuous
        .Range("V20").Borders.Weight = xlMedium
        .Range("S21:T21").Merge
        .Range("S21:T21").Value = "Process"
        .Range("U21").Value = "PID"
        .Range("V21").Value = "Created"
        
        'Grab current processes
        Module3.Tare
       
    End With
    Set ws = Nothing
End If
End Sub

Function cleanShift(name As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Report")
    Application.ScreenUpdating = False
    
    Select Case name
    Case "cred"
        With ws
            .Range("M20:M21").Value = ""
            .Range("P20:P21").Value = ""
            .Range("M23:M24").Value = ""
            .Range("Q24").Validation.Delete
            .Range("Q24").Value = ""
            
            .Range("M21:N21").Borders(xlEdgeTop).LineStyle = xlLineStyleNone
            .Range("M21:N21").Borders.LineStyle = xlLineStyleNone
            
            .Range("M24:O24").Borders(xlEdgeTop).LineStyle = xlLineStyleNone
            .Range("M24:O24").Borders.LineStyle = xlLineStyleNone
            
            .Range("P21:Q21").Borders(xlEdgeTop).LineStyle = xlLineStyleNone
            .Range("P21:Q21").Borders.LineStyle = xlLineStyleNone

            .Range("Q24").Borders(xlEdgeTop).LineStyle = xlLineStyleNone
            .Range("Q24").Borders.LineStyle = xlLineStyleNone
            
            
            .Range("M20:N20").UnMerge
            .Range("M21:N21").UnMerge
            .Range("P20:Q20").UnMerge
            .Range("P21:Q21").UnMerge
            .Range("M23:N23").UnMerge
            .Range("M24:O24").Validation.Delete
            .Range("M24:O24").UnMerge
        End With
    Case "log"
        With ws
            For Each cell In .Range("M26:N34")
                cell.UnMerge
                cell.ClearContents
            Next cell
        End With
    Case "radar"
        With ws
            .Range("S18:U18").UnMerge
            For Each cell In .Range("S18:V40")
                cell.Borders(xlEdgeTop).LineStyle = xlLineStyleNone
                cell.Borders.LineStyle = xlLineStyleNone
                cell.UnMerge
                cell.ClearContents
                cell.Interior.Color = RGB(255, 255, 255)
            Next cell
        End With
    Case "all"
        cleanShift "cred"
        cleanShift "log"
        cleanShift "radar"
    End Select
    Set ws = Nothing
    Application.ScreenUpdating = True
End Function



Sub credShift()
    
    Dim wsPaste As Worksheet, wsPaste2 As Worksheet, serverArea As Range, inputArea As Range, serverOpt As String, inputOpt As String
    Set wsPaste = ThisWorkbook.Sheets("Paste")
    Set wsPaste2 = ThisWorkbook.Sheets("Paste 2")
    serverOpt = "eaz1acmspp01v.corp.cvscaremark.com,eri1acmspp01hv.corp.cvscaremark.com"
    If Module3.showKey("radar") Then
        inputOpt = "cancel,start,scout, radar"
    Else
        inputOpt = "cancel,start,scout"
    End If
            
            
    Range("M20").Value = "Username"
    Range("M20:N20").Merge
    
    Range("M21:N21").Merge
    With Range("M21:N21").Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    Range("P20").Value = "Password"
    Range("P20:Q20").Merge
    
    Range("P21:Q21").Merge
    With Range("P21:Q21").Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    Range("M23").Value = "Server"
    Range("M23:N23").Merge
    
    With Range("M24:O24")
        .Merge
        With .Validation
            .Delete
            .Add xlValidateList, xlValidAlertStop, xlBetween, serverOpt
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
        With .Borders
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        
    End With
    
    With Range("Q24")
        With .Validation
            .Delete
            .Add xlValidateList, xlValidAlertStop, xlBetween, inputOpt
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
        With .Borders
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    End With
 
End Sub

Function logShift(step As Integer)
    Dim base As Range, ws As Worksheet, temp As String
    Set ws = ThisWorkbook.Sheets("Report")
    temp = "Creating Server…/Logging in…/Opening Catalog…/Pulling TSF…/Adding to Paste…/Pulling wAHT.../Adding to Paste 2…/Opening Outlook…/Creating Message…"
    
    If step < 0 Then Application.ScreenUpdating = False
    If step > 0 Then ws.Range("M25").Offset(step).Value = Split(temp, "/")(step - 1)
    
    radarShift
    
    Application.ScreenUpdating = True
    ActiveSheet.Calculate
    Application.Wait (Now + TimeValue("0:00:01"))
    ActiveSheet.Calculate
    Application.ScreenUpdating = False
    
    Set ws = Nothing

End Function

