Attribute VB_Name = "TherapistFormMod"
Option Explicit
' read therapist form after it is imported and put rooms into All Therapists sheet
Public Sub getTherapistsRoomsAndNotes(therapistSheet As Worksheet)
    Dim targetSheet As Worksheet
    Dim therSheet As Worksheet
    Dim initialsDict As Object
    Dim roomsDict As Object
    Dim initialsCell As range
    Dim requestedRooms As String
    Dim roomsList As Object
    Dim roomsArray() As String
    Dim rm As Variant
    Dim nextCell As range
    Dim note As range
    Dim atNote As range
    Dim shtNote As String
    Dim atNoteVal As String
    Dim i As Integer
    Dim newCell As range
    
    ' turn off screen updating
    Application.ScreenUpdating = False
    
    ' set sheets
    Set targetSheet = ActiveWorkbook.Sheets("All Therapists")
    Set therSheet = therapistSheet
    
    ' create  dictionary objects
    Set initialsDict = createInitialsDict()
    Set roomsDict = createRoomsDict()
    
    ' loop through rooms sheet to get initials
    ' if initials exist in dictionary (are valid), and rooms exist, get the rooms
    ' and add them to All Therapists
    For Each initialsCell In therSheet.range("$B$2:$B$12")
        ' create arrayList
        Set roomsList = CreateObject("System.Collections.ArrayList")
        If initialsDict.Exists(UCase(initialsCell.value)) Then
           ' get rooms as entered
           requestedRooms = initialsCell.Offset(0, 2).value
           shtNote = initialsCell.Offset(0, 3).value
           ' remove commas and subsitute spaces
           requestedRooms = Replace(requestedRooms, ",", " ")
           requestedRooms = Replace(requestedRooms, "  ", " ")
           roomsArray() = Split(requestedRooms, " ")
           For Each rm In roomsArray
                roomsList.Add (rm)
           Next rm
           
            ' get notes
        Set atNote = targetSheet.range(initialsDict(UCase(initialsCell.value))).Offset(0, 25)
        ' if note is empty, add note; otherwise concatanate note
        If IsEmpty(atNote) Or atNote.value = "" Or atNote.value = "" Then
            atNote.value = shtNote
        Else
            atNoteVal = atNote.value
            atNote.value = atNoteVal + "; " + shtNote
        End If
           
           
           For i = 0 To (roomsList.Count - 1)
                ' avoid skipping a cell because there's a space in the array
                Set nextCell = targetSheet.range(initialsDict(UCase(initialsCell.value))).Offset(0, 4)
                If roomsList.Item(i) <> " " And roomsDict.Exists(UCase(roomsList.Item(i))) Then
                    
                    For Each newCell In range(nextCell, nextCell.Offset(0, 17))
                        If IsEmpty(newCell) Or newCell.value = "" Or newCell.value = " " Then
                            newCell.value = UCase(roomsList.Item(i))
                            Exit For
                        End If
                    Next newCell
                End If
           Next i
        End If

    Next initialsCell
    
    ' highlight duplicates
    Call highlightDuplicateRooms(targetSheet, targetSheet.range("AllTherapistsOTRooms"))
    Call highlightDuplicateRooms(targetSheet, targetSheet.range("AllTherapistsPTRooms"))
    
    ' reset  dictionary objects
    Set initialsDict = Nothing
    Set roomsDict = Nothing
    
    ' turn on screen updating
    Application.ScreenUpdating = True
End Sub


' Copies the Therapist Sheet from the TherapistForm.xlsm Workbook to the SchedulingWorkbook.
' If the sheet already exists, it is deleted and replaced with a new version. The copy is placed AFTER the Ind Schedule sheet.

Public Sub copyTherapistSheet()

Dim fName As String
Dim targetWorkbook As Workbook

' set filename that contains sheet to copy and targetWorkbook (which may also have
' to be changed to a string and set to a file path)
fName = "C:\Users\a\Desktop\Roper-6\spreadsheets\Therapist-Form2.xlsm"
Set targetWorkbook = Workbooks("SchedulingWorkbook.xlsm")

If Not GetWorksheet("Rooms3W") Is Nothing Then
    Application.DisplayAlerts = False
    Worksheets("Rooms3W").Delete
    Application.DisplayAlerts = True
End If

Workbooks.Open Filename:=fName
Sheets("Rooms3W").Copy After:=targetWorkbook.Sheets("Ind Schedule")
Workbooks("Therapist-Form2.xlsm").Close savechanges:=False

End Sub






