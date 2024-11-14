Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("M20")) Is Nothing Then
        HandleRequest Target.Value
    End If
    If Not Intersect(Target, Me.Range("Q24")) Is Nothing Then
        If Target.Value = "cancel" Then
            Application.EnableEvents = False
            cleanShift "all"
            Application.EnableEvents = True
        ElseIf Target.Value = "start" And Range("M21").Value <> "" And Range("P21").Value <> "" And Range("M24").Value <> "" Then
            Application.EnableEvents = False
            Module3.Login Me.Range("M21").Value, Me.Range("P21").Value, Me.Range("M24").Value
            Application.EnableEvents = True
        ElseIf Target.Value = "scout" Then
            Module3.pressKey "radar"
            Me.Range("Q24").ClearContents
            If Module3.showKey("radar") Then
                Delta.radarShift
            Else
                Delta.cleanShift "radar"
            End If
        ElseIf Target.Value = "radar" Then
            Application.EnableEvents = False
            Delta.radarShift
            Me.Range("Q24").ClearContents
            Application.EnableEvents = True
        End If
    End If
End Sub

Function HandleRequest(code As String)
    Application.EnableEvents = False
        If code = "gg" Then
            Me.Range("M20").ClearContents
            Module3.Login "[USER]", "[PASS]", "[SERVER]"
        ElseIf code = "re" Then
            Me.Range("M20").ClearContents
            Module3.termCache
        ElseIf code = "im" Then
            Me.Range("M20").ClearContents
            ImportDataFromCSV
        ElseIf code = "Desi" Then
            Me.Range("M20").ClearContents
            Delta.credShift
            Delta.radarShift
        ElseIf code = "scout" Then
            Module3.pressKey "radar"
            Me.Range("M20").ClearContents
            If Module3.showKey("radar") Then
                Delta.radarShift
            Else
                Delta.cleanShift "radar"
            End If
        ElseIf code = "radar" Then
            Delta.radarShift
            Me.Range("M20").ClearContents
        ElseIf code = "clear" Then
            Delta.cleanShift "all"
            Me.Range("M20").ClearContents
        End If
    Application.EnableEvents = True
End Function
