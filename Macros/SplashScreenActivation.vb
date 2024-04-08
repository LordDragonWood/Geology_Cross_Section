Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1


Private Sub UserForm_Activate()
'Created for Neset Consulting by Dragon Wood (October, 2010)

    Call RemoveCaption(Me)
    
    Application.OnTime Now + TimeValue("00:00:05"), "KillTheForm"

End Sub