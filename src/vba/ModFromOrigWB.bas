Attribute VB_Name = "ModFromOrigWB"
' procedure from original workbook
Sub SelectRange()
'
' SelectRange Macro
'

'
    range("A1:AA32").Select
End Sub
' Procedure from original workbook
Sub PublishRange(range As range)
'
' PublishRange Macro

'Remove the password protection
    ActiveSheet.Unprotect

'Path = file path where the data will be saved ******(change this to your appropriate file location)******
    Dim Path As String
    Path = "C:\Users\Matt\Desktop\Roper Files\"

'ASName = The current active sheet
    Dim ASName As String
    ASName = ActiveSheet.Name

'FName = FileName
    Dim FName As String

    FName = Path & ASName & ".htm"

    With ActiveWorkbook.PublishObjects.Add(xlSourceRange, _
       FName, ASName, range, _
        xlHtmlStatic, ASName, ASName)
        .Publish (True)
        .AutoRepublish = False
    End With



'Enable password protection again
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

' Save the workbook

    ActiveWorkbook.Save

End Sub
' Procedure From original workbook
Sub LaunchHyperLink()
'
' LaunchHyperLink Macro
'

'
    range("G41").Select
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
End Sub



