Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Public Sub UnhideSheets(control As IRibbonControl)
'Created for Neset Consulting by Dragon Wood (October, 2010)
'Unhiding all the sheets after Macros are enabled

    ActiveWorkbook.Unprotect Password:="NCS117"
    
    Dim ws As Worksheet
        Application.DisplayAlerts = False
    For Each ws In ActiveWorkbook.Worksheets
        ws.Visible = True
    Next ws
        Application.DisplayAlerts = True
    
    ActiveWorkbook.Protect Password:="NCS117"
    
End Sub

Public Sub BoldCell(control As IRibbonControl)
'Created for Neset Consulting Service by Dragon Wood (August, 2011)
'Bolds or Unbolds selcted cell

    Dim b As Range
    Set b = ActiveCell
    
    With ActiveSheet
        .Unprotect Password:="NCS117"
        
        b.Select
    
        If b.Font.Bold = True Then
        b.Font.Bold = False
        Else
        b.Font.Bold = True
        End If
    
        .Protect Password:="NCS117"
    
    End With
End Sub

Public Sub ItalicCell(control As IRibbonControl)
'Created for Neset Consulting Service by Dragon Wood (August, 2011)
'Bolds or Unbolds selcted cell

    Dim i As Range
    Set i = ActiveCell
    
    With ActiveSheet
        .Unprotect Password:="NCS117"
        
        i.Select
    
        If i.Font.Italic = True Then
        i.Font.Italic = False
        Else
        i.Font.Italic = True
        End If
    
        .Protect Password:="NCS117"
    
    End With
End Sub

Sub NCSWebsite(control As IRibbonControl)
'Created for Neset Consulting Service by Dragon Wood (September, 2011)
'Opens the NCS Website

    ActiveWorkbook.FollowHyperlink Address:="http://www.nesetconsulting.com"
    
End Sub