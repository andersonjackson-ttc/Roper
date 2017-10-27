Attribute VB_Name = "SchedulingCallbacks"
Option Explicit

' These callbacks are for use in the custom ui xml file
' xml requires callback function to trigger macro

' callback for create all schedules in scheduling tab
Public Sub allSchedCallBack(control As IRibbonControl)
    Call Create_All_Schedules
End Sub

' callback for create all diddies in scheduling tab
Public Sub allDiddiesCallBack(control As IRibbonControl)
    Call Create_All_Diddies
End Sub

' callback for Import Rooms and Notes in scheduling tab
Public Sub importRooms(control As IRibbonControl)
    Call copyTherapistSheet
    Call clearPtOtColumns
    Call clearAllTherapistsNotesAndRooms
    Call getTherapistsRoomsAndNotes(Sheets("Rooms"))
    ' record last row read when sheet was imported
    Call getLastRow(Sheets("Rooms"), Sheets("All Therapists").range("LastRowCell3W"))
    Call lastTimeCreated(Sheets("All Therapists").range("AllTherapistsTimeCreatedCell"))
    Sheets("All Therapists").Select
End Sub
' callback for Add Rooms and Notes in scheduling tab
Public Sub addRooms(control As IRibbonControl)
    Dim str As String
    Dim lng As Long
    
    str = Sheets("All Therapists").range("LastRowCell3W").value
    lng = CLng(str)
    Call copyTherapistSheet
    ' update last row read after adding new rooms after copying sheet
    Call getLastRow(Sheets("Rooms"), Sheets("All Therapists").range("LastRowCell3W"))
    Call getTherapistsRoomsAndNotes(Sheets("Rooms"), lng)
    Call lastTimeCreated(Sheets("All Therapists").range("AllTherapistsTimeCreatedCell"))
    Sheets("All Therapists").Select
End Sub

' callback for Save 3W in scheduling tab
Public Sub save3W(control As IRibbonControl)
    Call Save3WSchedule
End Sub
' callback for Save 8P in scheduling tab
Public Sub save8P(control As IRibbonControl)
    Call Save8PSchedule
End Sub
' callback for Save 3P in scheduling tab
Public Sub save3P(control As IRibbonControl)
    Call Save3PSchedule
End Sub

