Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit
'Created for Neset Consulting by Dragon Wood (August, 2011)
'Formats the Splash screen so there is no close button

Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
 
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 
Private Declare Function DrawMenuBar Lib "User32" (ByVal hwnd As Long) As Long
 
Sub RemoveCaption(objForm As Object)
     
    Dim lStyle          As Long
    Dim hMenu           As Long
    Dim mhWndForm       As Long
     
    If Val(Application.Version) < 9 Then
        mhWndForm = FindWindow("ThunderXFrame", objForm.Caption) 'XL97
    Else
        mhWndForm = FindWindow("ThunderDFrame", objForm.Caption) 'XL2000+
    End If
    
    lStyle = GetWindowLong(mhWndForm, -16)
    lStyle = lStyle And Not &HC00000
    SetWindowLong mhWndForm, -16, lStyle
    DrawMenuBar mhWndForm
     
End Sub
 
Sub ShowForm()
     
    frmNCSSplash.Show False
     
End Sub