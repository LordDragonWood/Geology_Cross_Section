Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Private Sub Workbook_Activate()
'Created for Neset Consulting by Dragon Wood (October, 2010)
    
    ActiveWorkbook.Protect Password:="NCS117"
    
End Sub

Private Sub Workbook_Open()
'Created for Neset Consulting by Dragon Wood (October, 2010)

'Activates the Splash Screen
    frmNCSSplash.Show

'Forces the file to open on the Instructions page
     Application.Goto Sheets("Instructions").Range("A1"), True
     
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    'This code will undo PASTE and instead do a PASTE SPECIAL VALUES which will
    'allow you to retain FORMATS in all of the cells in all of the sheets, but will
    'also allow the user to COPY and PASTE data
 
    Dim UndoString As String
    Dim srce As Range
    On Error GoTo err_handler
    UndoString = Application.CommandBars("Standard").Controls("&Undo").List(1)
    
    If VBA.Left(UndoString, 5) <> "Paste" And UndoString <> "Auto Fill" Then
        
        Exit Sub
        
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Undo
            
            
    If UndoString = "Auto Fill" Then
        
        Set srce = Selection
        
        srce.Copy
        
        Target.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                    
        Application.SendKeys "{ESC}"
        Union(Target, srce).Select
       
    Else
    
        Target.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                    
    End If
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
err_handler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
'Created for Neset Consulting by Dragon Wood (October, 2010)

'Unprotects the workbook so final changes can be made before close
    ActiveWorkbook.Unprotect Password:="NCS117"
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    
'Hides all the sheets before closing
    For Each ws In ActiveWorkbook.Worksheets
    If ws.Name <> ("Instructions") Then ws.Visible = xlSheetHidden
    Next ws

'Puts the focus of the workbook on the instructions page for next start up.
    Application.Goto Sheets("Instructions").Range("A1"), True
    
'Saves the workbook
    If Not Me.Saved Then Me.Save
End Sub