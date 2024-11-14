Function ExtractUniqueID(ByVal fullPath As String, ByVal subStr As String) As String
    Dim uniqueID As String
    Dim startPos As Integer
    Dim endPos As Integer
    
    ' Initialize uniqueID in case 'Users' substring is not found or ID is not present
    uniqueID = ""
    
    ' Find the position of "Users" in the string
    startPos = InStr(1, fullPath, subStr, vbTextCompare)

    If startPos > 0 Then ' If "Users" is found in the string
        ' Find the position of the next backslash after "Users"
        endPos = InStr(startPos + Len(subStr), fullPath, "\")

        If endPos > startPos Then
            ' Extract the portion of the string between the backslashes after "Users"
            uniqueID = Mid(fullPath, startPos + Len(subStr), endPos - startPos - Len(subStr))
        End If
    End If

    ' Return the extracted uniqueID
    ExtractUniqueID = uniqueID
End Function

Sub Tare()
    Dim objWMIService As Object
    Dim colProcesses As Object
    Dim objProcess As Object
    Dim result As Variant
    Dim ws As Worksheet
    Dim count As Integer
    Dim datee As String
    
    Set ws = ThisWorkbook.Sheets("Report")
    result = ""
    count = 0
    
    ' Connect to WMI service
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    
    ' Query processes
    Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'acs_ssh.exe' OR Name = 'acsSRV.exe' OR Name = 'acsCNTRL.exe' OR Name = 'ACSScript.exe' OR Name = 'acsApp.exe' OR Name = 'acsRep.exe'")
    
    ' Loop through processes
    For Each objProcess In colProcesses
        count = count + 1
        
        ' Set alternating row color
        If count Mod 2 = 0 Then
            ws.Range("S" & 21 + count & ":V" & 21 + count).Interior.Color = RGB(211, 211, 211)
        End If
        
        ws.Range("S" & 21 + count & ":" & "T" & 21 + count).Merge
        ws.Range("S" & 21 + count).Value = objProcess.name
        ws.Range("U" & 21 + count).Value = objProcess.ProcessId
        ws.Range("V" & 21 + count).Value = displayTDiff(objProcess.creationDate)
    Next objProcess
    
    
    ' Clean up
    Set objWMIService = Nothing
    Set colProcesses = Nothing
End Sub


Function Login(user As String, pass As String, serv As String)
Dim cvsApp As Object, cvsSrv As Object, cvsRep As Object, info As Object, cvsConn As Object, Log As Object ', temp As Object, Rep As Object
Set cvsApp = CreateObject("ACSUP.cvsApplication")
Set cvsSrv = CreateObject("ACSUPSRV.cvsServer")
Set cvsRep = CreateObject("ACSREP.cvsReport")
Set cvsConn = CreateObject("ACSCN.cvsConnection")
cvsConn.bAutoRetry = True

'Clear Current Data
Module1.clearpastesheet
ThisWorkbook.Sheets("Report").Activate
Delta.logShift -1

'Connection
Delta.logShift 1
If cvsApp.CreateServer(user, pass, "", serv, False, "ENU", cvsSrv, cvsConn) Then
    Delta.logShift 2
    If cvsConn.Login(user, pass, serv, "ENU", "", False) Then
    Else
        MsgBox "Login failed", vbInformation, "WOMP WOMP"
    End If
Else
    MsgBox "Server creation failed", vbInformation, "WOMP WOMP"
End If

Delta.logShift 3
'Reports
Dim wsPaste As Worksheet, wsPaste2 As Worksheet
Set wsPaste = ThisWorkbook.Sheets("Paste")
Set wsPaste2 = ThisWorkbook.Sheets("Paste 2")
Delta.logShift 4
'TSF
cvsSrv.Reports.ACD = 1
Set info = cvsSrv.Reports.Reports("Integrated\Designer\Comparison Report With TSF")

If info Is Nothing Then
      If cvsSrv.Interactive Then
          MsgBox "The report Integrated\Designer\Splits/Skills with SL CDS Vers. was not found on ACD 1.", vbCritical Or vbOKOnly, "Avaya CMS Supervisor"
      Else
          Set Log = CreateObject("ACSERR.cvsLog")
          Log.AutoLogWrite "The report Integrated\Designer\Splits/Skills with SL CDS Vers. was not found on ACD 1."
          Set Log = Nothing
      End If
Else

    temp = cvsSrv.Reports.CreateReport(info, Rep)
    If temp Then
        Rep.SetProperty "Splits/Skills", "2342;2376;2377;2378;2379;2380;2355;2356;2359;2441;2442;2443;2445;2446;2449;2291;2300;2367;2368;2393;2394;2395;2396;2397;2398;2399;2400;2401;2402;2403;2404;2405;2406;2407;2408;2409;2410;2411;2412;2413;2414;2415;2101;2311;2312;2221;2230;2371;2372;2181;2267;2268;2351;2352;2161;2170;2271;2273;2341;2121;2130;2256;2261;2263;2321;2322;2141;2150;2331;2332;2201;2204;2207;2361;2362;2276;2381;2110;2426;2429;2431;2433;2434;1326;2392;2258;2278;2294;2297;2298;2313;2318;2328;2336;2251;2252;2104;2107;2108;2124;2127;2128;2144;2147;2148;2164;2167;2168;2187;2188;2208;2224;2227;2228;2251;2252;2357;2384;2387;2388;2444;2447;2448;2315;2316;2317;2319;2323;2324;2325;2327;2329;2333;2334;2337;2343;2313;2314;2318;2326;2328;2336;2253;2314;2272;2318;2282;2313;2336;2257;2326;2283;2326;2262;2310;2320;2330;2340;2350;2360;2366;2370;2390;2266;2281;2392;1326;2097;2155;1822;2344;2416;2151;2152;2153;2154;2156;2157;712;713;714;715;716;717;718;719;720;721;722;2450"
        Rep.ExportData "", 9, 0, True, True, True
        Delta.logShift 5
        wsPaste.Range("A1").PasteSpecial Paste:=xlPasteAll
    End If

End If

Rep.Quit

'If Not cvsSrv.Interactive Then cvsSrv.ActiveTasks.Remove Rep.TaskID
Set info = Nothing
Set temp = Nothing
Set Rep = Nothing

'wAHT
Delta.logShift 6
cvsSrv.Reports.ACD = 1
Set info = cvsSrv.Reports.Reports("Integrated\Designer\Comparison Report >10sec wAHT 2020")

If info Is Nothing Then
      If cvsSrv.Interactive Then
          MsgBox "Report not found...did we rename it?", vbCritical Or vbOKOnly, "Avaya CMS Supervisor"
      Else
          Set Log = CreateObject("ACSERR.cvsLog")
          Log.AutoLogWrite "The report Integrated\Designer\Splits/Skills with SL CDS Vers. was not found on ACD 1."
          Set Log = Nothing
      End If
Else
    temp = cvsSrv.Reports.CreateReport(info, Rep)
    If temp Then
        Rep.SetProperty "Splits/Skills", "2342;2376;2377;2378;2379;2380;2355;2356;2359;2441;2442;2443;2445;2446;2449;2291;2300;2367;2368;2393;2394;2395;2396;2397;2398;2399;2400;2401;2402;2403;2404;2405;2406;2407;2408;2409;2410;2411;2412;2413;2414;2415;2101;2311;2312;2221;2230;2371;2372;2181;2267;2268;2351;2352;2161;2170;2271;2273;2341;2121;2130;2256;2261;2263;2321;2322;2141;2150;2331;2332;2201;2204;2207;2361;2362;2276;2381;2110;2426;2429;2431;2433;2434;1326;2392;2258;2278;2294;2297;2298;2313;2318;2328;2336;2251;2252;2104;2107;2108;2124;2127;2128;2144;2147;2148;2164;2167;2168;2187;2188;2208;2224;2227;2228;2251;2252;2357;2384;2387;2388;2444;2447;2448;2315;2316;2317;2319;2323;2324;2325;2327;2329;2333;2334;2337;2343;2313;2314;2318;2326;2328;2336;2253;2314;2272;2318;2282;2313;2336;2257;2326;2283;2326;2262;2310;2320;2330;2340;2350;2360;2366;2370;2390;2266;2281;2392;1326;2097;2155;1822;2344;2416;2151;2152;2153;2154;2156;2157;712;713;714;715;716;717;718;719;720;721;722;2450"
        Rep.ExportData "", 9, 0, True, True, True
        Delta.logShift 7
        wsPaste2.Range("A1").PasteSpecial Paste:=xlPasteAll
    End If
End If


'Log & Clean
Set info = Nothing
Set temp = Nothing
If Not cvsSrv.Interactive Then
    cvsSrv.ActiveTasks.Remove Rep.TaskID
    cvsApp.Servers.Remove cvsSrv.ServerKey
End If
Set Rep = Nothing

cvsRep.Quit
Set cvsRep = Nothing

cvsConn.Logout
cvsConn.Disconnect
Set cvsConn = Nothing

cvsSrv.Connected = False
Set cvsSrv = Nothing

Set cvsApp = Nothing

Delta.logShift 8
Delta.logShift 9

Delta.cleanShift "cred"
Delta.cleanShift "log"
Delta.cleanShift "radar"
Rep = "%AppData%\Avaya\CMS Supervisor R19\Cache"
On Error Resume Next
If Dir(Rep) <> "" Then Kill Rep
On Error GoTo 0
Set Rep = Nothing
Module2.RunButton

End Function

Sub ImportDataFromCSV()
    Dim wb As Workbook
    Dim wsPaste As Worksheet, wsPaste2 As Worksheet
    Dim csvPathPaste As String, csvPathPaste2 As String
    Dim scriptPath As String
    Dim fileExistsTSF As Boolean, fileExistsAHT As Boolean
    Dim currPath As String
    
    
    ' Set file paths
    currPath = ExtractUniqueID(ThisWorkbook.Path, "Users\")
    currAya = ExtractUniqueID(ThisWorkbook.Path, "Profiles\")
    csvPathPaste = "C:\Users\" & currPath & "\AppData\Roaming\Avaya\CMS Supervisor R19\Profiles\" & currAya & "\Scripts\Reports\TSF.csv"
    csvPathPaste2 = "C:\Users\" & currPath & "\AppData\Roaming\Avaya\CMS Supervisor R19\Profiles\" & currAya & "\Scripts\Reports\AHT.csv"
    scriptPath = """C:\Users\" & currPath & "\AppData\Roaming\Avaya\CMS Supervisor R19\Profiles\" & currAya & "\Scripts\Reports\Reports.acsup"""
    
    'Clear Current Data
    Module1.clearpastesheet

    ' Execute the Avaya CMS script
    On Error Resume Next
    Shell "cmd /c " & scriptPath, vbNormalFocus
    On Error GoTo 0 ' Disable error handling (optional)

    If Err.Number <> 0 Then
        MsgBox "Error executing the script: " & Err.Description, vbExclamation, "Script Execution Error"
    End If

    ' Disable alerts to prevent confirmation prompts
    Application.DisplayAlerts = False

    ' Monitoring file existence
    Do While Not (fileExistsTSF And fileExistsAHT)
        fileExistsTSF = Dir(csvPathPaste) <> ""
        fileExistsAHT = Dir(csvPathPaste2) <> ""
        
        If fileExistsTSF And fileExistsAHT Then
            ' Open TSF.csv and copy data to Paste worksheet
            Workbooks.Open Filename:=csvPathPaste
            Set wb = ActiveWorkbook
            Set wsPaste = ThisWorkbook.Sheets("Paste")
            wb.Sheets(1).UsedRange.Copy wsPaste.Range("A1")
            wb.Close SaveChanges:=False

            ' Open AHT.csv and copy data to Paste 2 worksheet
            Workbooks.Open Filename:=csvPathPaste2
            Set wb = ActiveWorkbook
            Set wsPaste2 = ThisWorkbook.Sheets("Paste 2")
            wb.Sheets(1).UsedRange.Copy wsPaste2.Range("A1")
            wb.Close SaveChanges:=False
            
            ' Enable alerts again
            Application.DisplayAlerts = True
            
            ' Delete CSV files
            Kill csvPathPaste
            Kill csvPathPaste2

            'Call RunButton
            Module2.RunButton
            
            Exit Do
        End If
        
        Application.Wait Now + TimeValue("0:00:05") ' Wait for 1 second before rechecking
    Loop
End Sub

Sub termCache()

    If (Dir("C:\Users\c289894\AppData\Roaming\Avaya\CMS Supervisor R19\Cache", vbDirectory) <> "") Then
        Shell "C:\Users\c289894\Desktop\Avaya Clear v19.bat", vbNormalFocus
    End If
End Sub

Function saveKey(key As String, opt As String)
    'default is False
    SaveSetting "RP", "Options", key, opt
End Function

Function showKey(key As String) As Boolean
    'default is False
    Dim temp As String
    temp = GetSetting("RP", "Options", key, "Failed")
    If temp = "Failed" Then 'If there is no setting an I ask for it, execute and save
        saveKey key, "True"
        temp = "True"
    End If
    showKey = (temp = "True")
End Function

Function pressKey(key As String)
    Dim temp As String
    temp = GetSetting("RP", "Options", key, "Failed")
        If temp = "Failed" Then 'If there is no setting an I ask for it, execute and save
        saveKey key, "True"
        temp = "True"
    End If
    saveKey key, CStr(Not (temp = "True"))
End Function

Sub delKey()
    'default is False
    DeleteSetting "RP", "Options"
End Sub

Sub eee()
    'saveKey "Radar", "False"
    MsgBox showKey("Radar")
    'pressKey "Radar"
    'MsgBox showKey("Radar")
End Sub

Function termProcess(ByVal PID As Long)
    Dim shellCommand As String
    shellCommand = "taskkill /f /pid " & PID
End Function


Function displayTDiff(creationDateStr As String) As String
    Dim creationDate As Date, currentDate As Date
    Dim timeDifference As Double
    Dim timeDifferenceInDays As Long
    Dim timeDifferenceInHours As Long
    Dim timeDifferenceInMinutes As Long
    Dim timeDifferenceInSeconds As Long
    Dim displayText As String

    ' Convert the string representation to a VBA Date data type
    creationDate = DateSerial(Left(creationDateStr, 4), Mid(creationDateStr, 5, 2), Mid(creationDateStr, 7, 2)) + _
                   TimeSerial(Mid(creationDateStr, 9, 2), Mid(creationDateStr, 11, 2), Mid(creationDateStr, 13, 2))

    ' Get the current date and time
    currentDate = Now()

    ' Calculate the difference in time between the current date and time and the creation date
    timeDifference = currentDate - creationDate

    ' Calculate the difference in days, hours, minutes, and seconds then Construct the display text
    timeDifferenceInDays = Int(timeDifference)
    timeDifferenceInHours = Int((timeDifference - timeDifferenceInDays) * 24)
    timeDifferenceInMinutes = Int(((timeDifference - timeDifferenceInDays) * 24 - timeDifferenceInHours) * 60)
    timeDifferenceInSeconds = Int((((timeDifference - timeDifferenceInDays) * 24 - timeDifferenceInHours) * 60 - timeDifferenceInMinutes) * 60)


    If Int(timeDifference) > 0 Then
        displayText = timeDifferenceInDays & " day(s)"
    ElseIf timeDifferenceInHours > 0 Then
        displayText = timeDifferenceInHours & " hour(s)"
    ElseIf timeDifferenceInMinutes > 0 Then
        displayText = timeDifferenceInMinutes & " min(s)"
    Else
        displayText = timeDifferenceInSeconds & " sec(s)"
    End If

    ' Display the result
    displayTDiff = displayText
End Function

