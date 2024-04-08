Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'Created by Dragon Wood (October, 2010)

'Keeps the Instructions sheet protected
    Application.ScreenUpdating = False

    ActiveSheet.Unprotect Password:="NCS117"
    ActiveSheet.Protect Password:="NCS117"
    
    Application.ScreenUpdating = True
End Sub