Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Public Sub ExportAsite(control As IRibbonControl)
'Created for Neset Consulting Service by Dragon Wood (August, 2011)
'Exports the data to the asite.txt file and puts it in the same directory as this Workbook
    
    Application.ScreenUpdating = False
    
    With Sheets("Asite")
    .Unprotect Password:="NCS117"
    
    Dim intAsite As Integer
    Dim LR As Long
    Dim i As Long
    Dim strFile As String
    
    strFile = ThisWorkbook.Path & "\" & "asite.txt"
    LR = Range("A" & Rows.Count).End(xlUp).Row
    intAsite = FreeFile

    Open strFile For Output As #intAsite
    
    For i = 1 To LR
        Print #intAsite, .Range("A" & i).Value; vbTab; .Range("B" & i).Value; vbTab; .Range("C" & i).Value; vbTab; .Range("D" & i).Value; vbTab; .Range("E" & i).Value; vbTab; .Range("F" & i).Value; vbTab; .Range("G" & i).Value; vbTab; .Range("H" & i).Value; vbTab; .Range("I" & i).Value; vbTab; .Range("J" & i).Value
    
    Next i
    
    Close #intAsite
    
    .Protect Password:="NCS117"
    
    End With
    
    Application.ScreenUpdating = True

End Sub

Public Sub ExportSurveys(control As IRibbonControl)
'Created for Neset Consulting by Dragon Wood (August, 2011)
' Exports the data to the survey.txt file and puts it in the same directory as this Workbook
    
    Application.ScreenUpdating = False
    
    With Sheets("Surveys")
    .Unprotect Password:="NCS117"
    
    Dim intSurvey As Integer
    Dim LS As Long
    Dim s As Long
    Dim strSfile As String
    
    strSfile = ThisWorkbook.Path & "\" & "surveys.txt"
    LS = Range("C" & Rows.Count).End(xlUp).Row
    intSurvey = FreeFile
    Open strSfile For Output As intSurvey
    
    For s = 11 To LS
        Print #intSurvey, .Range("C" & s).Value; vbTab; .Range("G" & s).Value; vbTab; .Range("D" & s).Value; vbTab; .Range("E" & s).Value; vbTab; .Range("I" & s).Value; vbTab; .Range("K" & s).Value; vbTab; .Range("H" & s).Value; vbTab; .Range("O" & s).Value

    Next s

    Close #intSurvey
    
    .Protect Password:="NCS117"
    
    End With
    Application.ScreenUpdating = True

End Sub

Sub CreatePDFs(control As IRibbonControl)
'Created for Neset Consulting Service by Dragon Wood (October, 2011)
'Saves the reports in PDF Format

    Dim strFile As String
    
    strFile = ThisWorkbook.Path & "\" & "Cross Section Files" & "\"

    If Dir(strFile, vbDirectory) = "" Then
    MkDir (strFile)
    End If
    

 'Save the Cross Section page as a PDF
    Application.Goto Sheets("Cross Section").Range("A1"), True
    With Sheets("Cross Section")
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=strFile & "Cross Section.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    End With

 'Save the Dip Calculator page as a PDF
    Application.Goto Sheets("Dip Calculator").Range("A1"), True
    With Sheets("Dip Calculator")
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=strFile & "Dip Calculator.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    End With

 'Save the Formation Projection page as a PDF
    Application.Goto Sheets("Formation Projection").Range("A1"), True
    With Sheets("Formation Projection")
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=strFile & "Formation Projection.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    End With

 'Save the Surveys page as a PDF
    Application.Goto Sheets("Surveys").Range("A1"), True
    With Sheets("Surveys")
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=strFile & "Surveys.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    End With

        
End Sub