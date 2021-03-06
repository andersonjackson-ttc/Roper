Attribute VB_Name = "IndScheduleMod"
Option Explicit
' create schedule for individual therapist
Sub createIndSched()
    Dim cell1 As range
    Dim ucStr As String
    Dim cell2 As range
    Dim cell3 As range
    Dim r1 As range
    Dim r2 As range
    Dim r3 As range
    Dim str1 As String
    Dim str2 As String
    Dim i As Integer

  
   ' clear previous data and highlighting
   Call clearIndSched
   
   ' turn off screen updating
   Application.ScreenUpdating = False

   ' look at each row in 3W, 8P & 3P schedules; if initials match initials in row
   ' copy room and schedule row and paste into all schedules


   ' initialize row counter
   i = 3
   ' initialize cell1 as therapist initials cell
   Set cell1 = Sheets("Ind Schedule").range("IndSchedInitials")
   ' make initials upper case if they aren't
   ucStr = UCase(cell1.value)
   
    ' look at 3W Schedule
    For Each r1 In Sheets("3W Schedule").range("SchedGrid3W").Rows
        ' account for possible ADL
        str1 = "ADL " + CStr(ucStr)
        str2 = "ADL" + CStr(ucStr)
        If Application.WorksheetFunction.CountIf(r1, ucStr) > 0 Or Application.WorksheetFunction.CountIf(r1, str1) > 0 Or Application.WorksheetFunction.CountIf(r1, str2) > 0 Then
            ' get room
            cell1.Offset(i, -14).value = r1.Cells(1, 1).value
            ' get schedule for room
            range(r1.Cells(1, 2), r1.Cells(1, 23)).Copy
            range(cell1.Offset(i, -10), cell1.Offset(i, 11)).PasteSpecial xlPasteValues
            ' go to next row in schedule
            i = i + 1
        End If
    Next r1
    ' look at 8P schedule
    For Each r2 In Sheets("8P Schedule").range("SchedGrid8P").Rows
        str1 = "ADL " + CStr(ucStr)
        str2 = "ADL" + CStr(ucStr)
        If Application.WorksheetFunction.CountIf(r2, ucStr) > 0 Or Application.WorksheetFunction.CountIf(r2, str1) > 0 Or Application.WorksheetFunction.CountIf(r2, str2) > 0 Then
            cell1.Offset(i, -14).value = r2.Cells(1, 1).value
            range(r2.Cells(1, 2), r2.Cells(1, 23)).Copy
            range(cell1.Offset(i, -10), cell1.Offset(i, 11)).PasteSpecial xlPasteValues
            ' go to next row in schedule
            i = i + 1
        End If
    Next r2
    
    ' look at 3P schedule
    For Each r3 In Sheets("3P Schedule").range("SchedGrid3P").Rows
        str1 = "ADL " + CStr(ucStr)
        str2 = "ADL" + CStr(ucStr)
        If Application.WorksheetFunction.CountIf(r3, ucStr) > 0 Or Application.WorksheetFunction.CountIf(r3, str1) > 0 Or Application.WorksheetFunction.CountIf(r3, str2) > 0 Then
            cell1.Offset(i, -14).value = r3.Cells(1, 1).value
            range(r3.Cells(1, 2), r3.Cells(1, 23)).Copy
            range(cell1.Offset(i, -10), cell1.Offset(i, 11)).PasteSpecial xlPasteValues
            ' go to next row in schedule
            i = i + 1
        End If
    Next r3
    
    ' get notes
    For Each cell3 In Sheets("All Therapists").range("AllTherapistsInitials")
        
        If cell3.value = UCase(range("IndSchedInitials").value) Then
            Sheets("Ind Schedule").range("IndSchedNoteRef") = cell3.Offset(0, 25).value
        End If
    Next cell3
    
    
    ' apply highlighting
   Call schedCondFormat(Sheets("Ind Schedule").range("IndSchedEvalCell"), Sheets("Ind Schedule").range("IndSchedIntCell"), _
        Sheets("Ind Schedule").range("IndSchedRooms"), Sheets("Ind Schedule").range("IndSchedInitials"), _
        Sheets("Ind Schedule"))

    ' turn off cut copy mode
    Application.CutCopyMode = False
     
   ' deselect row and go back to initials box
   Sheets("Ind Schedule").range("IndSchedInitials").Select
   
   ' turn on screen updating
   Application.ScreenUpdating = True
   
   ' deselect row and go back to initials box
   Sheets("Ind Schedule").range("IndSchedInitials").Select
   
End Sub



