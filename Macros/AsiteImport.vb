Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Sub GetTieIn(control As IRibbonControl)
'Created for Neset Consulting Service by Dragon Wood (September, 2011)
'Collects and inputs the Tie In Survey inforamtion

    Dim strTieInSurveyDepth As String
    Dim strTieInSurveyInc As String
    Dim strTieInSurveyAzi As String
    Dim strTieInSurveyTVD As String
    Dim strTieInSurveyVS As String
    Dim strTieInSurveyNS As String
    Dim strTieInSurveyEW As String

'Make the Surveys page the focus point
 
    Application.Goto Sheets("Surveys").Range("A1"), True
    
 'Check the Survey Tie In fields for content. If there, use the content, if not provide an input box for entering the data.
 
    With Sheets("Surveys")
    '.Unprotect Password:="NCS117"

       .Range("C11").Select
        strTieInSurveyDepth = InputBox("Please Enter The Tie-In Survey Depth.", "Tie-In Survey Depth")
        .Unprotect Password:="NCS117"
        ActiveCell.FormulaR1C1 = strTieInSurveyDepth
    
       .Range("D11").Select
        strTieInSurveyInc = InputBox("Please Enter The Tie-In Survey Inclination (Incl).", "Tie-In Survey Inclination (Incl)")
        .Unprotect Password:="NCS117"
        ActiveCell.FormulaR1C1 = strTieInSurveyInc
    
       .Range("E11").Select
        strTieInSurveyAzi = InputBox("Please Enter The Tie-In Survey Azimuth (Azi).", "Tie-In Survey Azimuth (Azi)")
        .Unprotect Password:="NCS117"
        ActiveCell.FormulaR1C1 = strTieInSurveyAzi
    
       .Range("G11").Select
        strTieInSurveyTVD = InputBox("Please Enter The Tie-In Survey True Vertical Depth (TVD).", "Tie-In Survey True Vertical Depth (TVD)")
        .Unprotect Password:="NCS117"
        ActiveCell.FormulaR1C1 = strTieInSurveyTVD
    
       .Range("H11").Select
        strTieInSurveyVS = InputBox("Please Enter The Tie-In Survey Vertical Section (VS).", "Tie-In Survey Vertical Section (VS)")
        .Unprotect Password:="NCS117"
        ActiveCell.FormulaR1C1 = strTieInSurveyVS
    
       .Range("I11").Select
        strTieInSurveyNS = InputBox("Please Enter The Tie-In Survey Northings (N/S).", "Tie-In Survey Northings (N/S)")
        .Unprotect Password:="NCS117"
        ActiveCell.FormulaR1C1 = strTieInSurveyNS
    
       .Range("K11").Select
        strTieInSurveyEW = InputBox("Please Enter The Tie-In Survey Eastings (E/W).", "Tie-In Survey Eastings (E/W)")
        .Unprotect Password:="NCS117"
        ActiveCell.FormulaR1C1 = strTieInSurveyEW
    

   .Protect Password:="NCS117"
    
    End With
 
 End Sub
 
Sub GetGammaGasScale(control As IRibbonControl)
'Created for Neset Consulting Service by Dragon Wood (September, 2011)
'Collects and inputs the Gamma and Total Gas Scale inforamtion

    Dim strGammaScale As String
    Dim strTotalGasScale As String
    Dim strStartDepth As String
    Dim strEndDepth As String


'Make the Cross Section page the focus point
 
    Application.Goto Sheets("Asite").Range("A2"), True
    
 'Check the Survey Tie In fields for content. If there, use the content, if not provide an input box for entering the data.
 
    With Sheets("Asite")
    '.Unprotect Password:="NCS117"

       .Range("AA2").Select
        strGammaScale = InputBox("Please Enter The Gamma Scale (Default is 100).", "Gamma Scale Range")
        .Unprotect Password:="NCS117"
        ActiveCell.FormulaR1C1 = strGammaScale
    
       .Range("AA3").Select
        strTotalGasScale = InputBox("Please Enter The Total Gas Scale (Default Matches Proposed Hole Depth.", "Total Gas Scale Range")
        .Unprotect Password:="NCS117"
        ActiveCell.FormulaR1C1 = strTotalGasScale
    
       .Range("AA4").Select
        strStartDepth = InputBox("Please Enter The Start Depth For The Gamma/Total Gas Chart.", "Gamma/Total Gas Start Depth")
        .Unprotect Password:="NCS117"
        ActiveCell.FormulaR1C1 = strStartDepth
    
       .Range("AA5").Select
        strEndDepth = InputBox("Please Enter The End Depth For The Gamma/Total Gas Chart.", "Gamma/Total Gas End Depth")
        .Unprotect Password:="NCS117"
        ActiveCell.FormulaR1C1 = strEndDepth
    

   .Protect Password:="NCS117"
    
    End With
 
 End Sub