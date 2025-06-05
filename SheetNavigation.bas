Attribute VB_Name = "SheetNavigation"
Sub GoToNextSheet()
    On Error Resume Next
    If ActiveSheet.Index < Sheets.Count Then
        Sheets(ActiveSheet.Index + 1).Activate
    End If
End Sub

Sub GoToPreviousSheet()
    On Error Resume Next
    If ActiveSheet.Index > 1 Then
        Sheets(ActiveSheet.Index - 1).Activate
    End If
End Sub
