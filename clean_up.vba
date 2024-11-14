
Sub ClearSkillsContents()
Worksheets("Skills Holding").Activate
    Range("b4:D200").ClearContents
    Range("b4").Select
End Sub

Sub ClearACWContents()
Worksheets("ACW").Activate
    Range("A3:I200").ClearContents
    Range("B4").Select
End Sub

Sub ClearBreakContents()
Worksheets("Break").Activate
    Range("A3:J200").ClearContents
    Range("B3").Select
End Sub
Sub ClearRRContents()
Worksheets("Restroom").Activate
    Range("A3:J200").ClearContents
    Range("B3").Select
End Sub

Sub clearpastesheet()
Worksheets("Paste 2").Activate
    Range("A1:O999").ClearContents
    Range("A1").Select

Worksheets("Paste").Activate
    Range("A1:L999").ClearContents
    Range("A1").Select

End Sub

Sub clear_Converter()
Worksheets("Min Converter").Select
    Range("C1:AZ100000").ClearContents
End Sub

Sub clear_AUX()
Worksheets("AUX").Select
    Range("A2:Z100000").ClearContents
    Range("D2").Select
End Sub

