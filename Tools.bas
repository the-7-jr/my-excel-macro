Attribute VB_Name = "Tools"
Option Explicit

Dim function_name As String

Sub delete_name_definitions()

    Dim name_definition As Name
    
    On Error Resume Next
    
    Call application_set
    
    function_name = "delete_name_definitions"
    
    For Each name_definition In ActiveWorkbook.Names
        name_definition.Delete
    Next
    
    Call application_reset
    
    On Error GoTo 0
End Sub

Sub unlink()
    
    Dim wb As Workbook
    Dim vntLink As Variant
    Dim i As Integer
    
    On Error Resume Next
    
    Call application_set
    
    function_name = "unlink"
    
    Set wb = ActiveWorkbook
    
    vntLink = wb.LinkSources(xlLinkTypeExcelLinks)
    
    If IsArray(vntLink) Then
        For i = 1 To UBound(vntLink)
            wb.BreakLink vntLink(i), xlLinkTypeExcelLinks
        Next i
    End If
    
    Call application_reset
    
    On Error GoTo 0

End Sub

Sub application_set()
    ' Omit extra applicaiton processing during macro processing
    With Application
        .ScreenUpdating = False             ' Omit drawing
        .Calculate = xlCalculationManual    ' Changed to manual calculation
        .DisplayAlerts = False              ' Omit warning
    End With
End Sub

Sub application_reset()
    ' Reset application settings to normal
    With Application
        .ScreenUpdating = True
        .Calculate = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
End Sub
