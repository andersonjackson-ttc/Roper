Attribute VB_Name = "TherapistSelectorCheckBoxes"
Sub CreateDailyTherapists_Click()
    'Each call checks,clears,sorts, and adds new therapists to a single discipline on a single floor
    Application.ScreenUpdating = False
    Call AddDailyTherapists(range("NamesOT3WAllTherapists"), range("TrueFalseOT3W"), Sheets("All Therapists").range("D6"), range("Sort3WOT"))
    Call AddDailyTherapists(range("NamesOT8PAllTherapists"), range("TrueFalseOT8P"), Sheets("All Therapists").range("D18"), range("Sort8POT"))
    Call AddDailyTherapists(range("NamesOT3PAllTherapists"), range("TrueFalseOT3P"), Sheets("All Therapists").range("D30"), range("Sort3POT"))
    Call AddDailyTherapists(range("NamesPT3WAllTherapists"), range("TrueFalsePT3W"), Sheets("All Therapists").range("D42"), range("Sort3WPT"))
    Call AddDailyTherapists(range("NamesPT8PAllTherapists"), range("TrueFalsePT8P"), Sheets("All Therapists").range("D54"), range("Sort8PPT"))
    Call AddDailyTherapists(range("NamesPT3PAllTherapists"), range("TrueFalsePT3P"), Sheets("All Therapists").range("D66"), range("Sort3PPT"))
    Call AddDailyTherapists(range("NamesSP3WAllTherapists"), range("TrueFalseSP3W"), Sheets("All Therapists").range("D78"), range("Sort3WSP"))
    Call AddDailyTherapists(range("NamesSP8PAllTherapists"), range("TrueFalseSP8P"), Sheets("All Therapists").range("D90"), range("Sort8PSP"))
    Call AddDailyTherapists(range("NamesSP3PAllTherapists"), range("TrueFalseSP3P"), Sheets("All Therapists").range("D102"), range("Sort3PSP"))
    Call AddDailyTherapists(range("NamesREC3WAllTherapists"), range("TrueFalseREC3W"), Sheets("All Therapists").range("D114"), range("Sort3WREC"))
    Call AddDailyTherapists(range("NamesREC8PAllTherapists"), range("TrueFalseREC8P"), Sheets("All Therapists").range("D126"), range("Sort8PREC"))
    Call AddDailyTherapists(range("NameSREC3PAllTherapists"), range("TrueFalseREC3P"), Sheets("All Therapists").range("D138"), range("Sort3PREC"))
    Application.ScreenUpdating = True
End Sub
Private Sub AddDailyTherapists(PasteToRange As range, TrueFalseRange As range, StartCell As range, SortRange As range)
    'callls clear and sort method then puts all checked names into an array, removes the ones that are already present then
    'the remaining names to the discipline/floor area
    Call ClearUnselectedTherapists(PasteToRange, TrueFalseRange, StartCell, SortRange)
    
    Dim Names(0 To 11) As String 'create array
    i = 0
    For Each cel In TrueFalseRange
        If cel.value = True Then
            Names(i) = cel.Parent.Cells(cel.Row, 4).value 'fill array with checked names
            i = i + 1
        End If
    Next cel
    For Each n In PasteToRange
        For j = 0 To UBound(Names)
            If Names(j) = n.value Then 'set name slots in array to empty if they appear in the PasteToRange
                Names(j) = ""
            End If
        Next j
    Next n
    StartCell.Activate 'start at top row of discipline/floor
    For k = 0 To UBound(Names)
        Do While (ActiveCell.value <> "-") 'move down if cell doesn't have a hyphen
            ActiveCell.Offset(rowOffset:=1, columnOffset:=0).Activate
        Loop
        If Names(k) <> "" Then
            ActiveCell.value = Names(k) 'add non-empty array slots to PasteToRange
        End If
    Next k
End Sub
Public Sub ClearUnselectedTherapists(PasteToRange As range, TrueFalseRange As range, StartCell As range, SortRange As range)
    Sheets("All Therapists").Activate
    StartCell.Activate 'start at top row of discipline/floor
    'remove names from PasteToRange if they have a check by them
    For Each cell In TrueFalseRange
        'remove all names that aren't checked
        If cell.value = False Then
            Name = cell.Parent.Cells(cell.Row, 4).value
            If Name <> "-" Then 'skip names are empty hyphen spaces
                For Each cel In PasteToRange 'remove names, rooms, and notes from all therapists if it has been unchecked
                    If Name = cel.value Then
                        cel.value = "-"
                        cel.Offset(0, 1).range("A1:R1").Select
                        Selection.ClearContents
                        cel.Offset(0, 1).range("T1:V1").Select
                        Selection.ClearContents
                        Exit For
                    End If
                Next cel
            End If
        End If
    Next cell
    'sort the remaining rows
    With ActiveWorkbook.Worksheets("All Therapists").Sort
        .SetRange SortRange
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


Public Sub DailyTherapists(control As IRibbonControl)
    Call CreateDailyTherapists_Click
End Sub
