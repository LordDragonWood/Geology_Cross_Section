Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Public Sub ShowChromatograph(control As IRibbonControl)
'Created for Neset Consulting Service by Dragon Wood (August, 2011)
'Displays the Chromatograph labels on the Asite, Log Inputs & Morning Report Sheets

    With Sheets("Asite")
        .Unprotect Password:="NCS117"
    Application.ScreenUpdating = False

'Renaming the columns in Asite for use of Chromatograph
    If .Range("D1").Value = "Gamma" Then
        .Range("D1").Value = "C1"
    Else
        .Range("D1").Value = "Gamma"
    End If
    If .Range("E1").Value = "TVD" Then
        .Range("E1").Value = "C2"
    Else
        .Range("E1").Value = "TVD"
    End If
    If .Range("F1").Value = "" Then
        .Range("F1").Value = "C3"
    Else
        .Range("F1").Value = ""
    End If
    If .Range("G1").Value = "" Then
        .Range("G1").Value = "iC4"
    Else
        .Range("G1").Value = ""
    End If
    If .Range("H1").Value = "" Then
        .Range("H1").Value = "nC4"
    Else
        .Range("H1").Value = ""
    End If
    If .Range("I1").Value = "Gamma" Then
        .Range("I1").Value = ""
    Else
        .Range("I1").Value = "Gamma"
    End If
    If .Range("J1").Value = "TVD" Then
        .Range("J1").Value = ""
    Else
        .Range("J1").Value = "TVD"
    End If

    Application.ScreenUpdating = True

    .Protect Password:="NCS117"

    End With
    
    With Sheets("Log Inputs")
        .Unprotect Password:="NCS117"
    Application.ScreenUpdating = False
    
    If .Range("I1").Value = "" Then
        .Range("I1").Value = "CHROMATOGRAPH MEASURED DEPTHS"
    Else
        .Range("I1").Value = ""
    End If
    
    If .Range("J2").Value = "" Then
        .Range("J2").Value = "1"
    Else
        .Range("J2").Value = ""
    End If
    If .Range("K2").Locked = False Then
        .Range("K2").Locked = True
        .Range("K2").Borders(xlEdgeBottom).LineStyle = xlNone
    Else
        .Range("K2").Locked = False
        .Range("K2").Borders(xlEdgeBottom).LineStyle = xlContinuous
    End If
   
    If .Range("J3").Value = "" Then
        .Range("J3").Value = "2"
    Else
        .Range("J3").Value = ""
    End If
    If .Range("K3").Locked = False Then
        .Range("K3").Locked = True
        .Range("K3").Borders(xlEdgeBottom).LineStyle = xlNone
    Else
        .Range("K3").Locked = False
        .Range("K3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    End If
    
    If .Range("J4").Value = "" Then
        .Range("J4").Value = "3"
    Else
        .Range("J4").Value = ""
    End If
    If .Range("K4").Locked = False Then
        .Range("K4").Locked = True
        .Range("K4").Borders(xlEdgeBottom).LineStyle = xlNone
    Else
        .Range("K4").Locked = False
        .Range("K4").Borders(xlEdgeBottom).LineStyle = xlContinuous
    End If
    
    If .Range("J5").Value = "" Then
        .Range("J5").Value = "4"
    Else
        .Range("J5").Value = ""
    End If
    If .Range("K5").Locked = False Then
        .Range("K5").Locked = True
        .Range("K5").Borders(xlEdgeBottom).LineStyle = xlNone
    Else
        .Range("K5").Locked = False
        .Range("K5").Borders(xlEdgeBottom).LineStyle = xlContinuous
    End If
    
    If .Range("J6").Value = "" Then
        .Range("J6").Value = "5"
    Else
        .Range("J6").Value = ""
    End If
    If .Range("K6").Locked = False Then
        .Range("K6").Locked = True
        .Range("K6").Borders(xlEdgeBottom).LineStyle = xlNone
    Else
        .Range("K6").Locked = False
        .Range("K6").Borders(xlEdgeBottom).LineStyle = xlContinuous
    End If
    
    If .Range("J13").Value = "" Then
        .Range("J13").Value = "CHROMATOGRAPH 1"
        .Range("J13:K13").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("J13:K13").Borders(xlEdgeTop).Weight = xlMedium
        .Range("J13").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J13").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K13").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K13").Borders(xlEdgeRight).Weight = xlMedium
        .Range("J13:K13").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("J13:K13").Borders(xlEdgeBottom).Weight = xlMedium
    Else
        .Range("J13").Value = ""
        .Range("J13:K13").Borders(xlEdgeTop).LineStyle = xlNone
        .Range("J13").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K13").Borders(xlEdgeRight).LineStyle = xlNone
        .Range("J13:K13").Borders(xlEdgeBottom).LineStyle = xlNone
    End If
    If .Range("J14").Value = "" Then
        .Range("J14").Value = "HW "
        .Range("J14").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J14").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K14").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K14").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J14").Value = ""
        .Range("J14").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K14").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J15").Value = "" Then
        .Range("J15").Value = "C1 "
        .Range("J15").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J15").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K15").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K15").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J15").Value = ""
        .Range("J15").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K15").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J16").Value = "" Then
        .Range("J16").Value = "C2 "
        .Range("J16").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J16").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K16").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K16").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J16").Value = ""
        .Range("J16").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K16").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J17").Value = "" Then
        .Range("J17").Value = "C3 "
        .Range("J17").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J17").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K17").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K17").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J17").Value = ""
        .Range("J17").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K17").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J18").Value = "" Then
        .Range("J18").Value = "iC4 "
        .Range("J18").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J18").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K18").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K18").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J18").Value = ""
        .Range("J18").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K18").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J19").Value = "" Then
        .Range("J19").Value = "nC4 "
        .Range("J19").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J19").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("J19").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("J19").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("K19").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K19").Borders(xlEdgeRight).Weight = xlMedium
        .Range("K19").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("K19").Borders(xlEdgeBottom).Weight = xlMedium
    Else
        .Range("J19").Value = ""
        .Range("J19").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("J19").Borders(xlEdgeBottom).LineStyle = xlNone
        .Range("K19").Borders(xlEdgeRight).LineStyle = xlNone
        .Range("K19").Borders(xlEdgeBottom).LineStyle = xlNone
    End If
    
    If .Range("J21").Value = "" Then
        .Range("J21").Value = "CHROMATOGRAPH 2"
        .Range("J21:K21").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("J21:K21").Borders(xlEdgeTop).Weight = xlMedium
        .Range("J21").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J21").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K21").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K21").Borders(xlEdgeRight).Weight = xlMedium
        .Range("J21:K21").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("J21:K21").Borders(xlEdgeBottom).Weight = xlMedium
    Else
        .Range("J21").Value = ""
        .Range("J21:K21").Borders(xlEdgeTop).LineStyle = xlNone
        .Range("J21").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K21").Borders(xlEdgeRight).LineStyle = xlNone
        .Range("J21:K21").Borders(xlEdgeBottom).LineStyle = xlNone
    End If
    If .Range("J22").Value = "" Then
        .Range("J22").Value = "HW "
        .Range("J22").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J22").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K22").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K22").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J22").Value = ""
        .Range("J22").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K22").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J23").Value = "" Then
        .Range("J23").Value = "C1 "
        .Range("J23").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J23").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K23").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K23").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J23").Value = ""
        .Range("J23").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K23").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J24").Value = "" Then
        .Range("J24").Value = "C2 "
        .Range("J24").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J24").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K24").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K24").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J24").Value = ""
        .Range("J24").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K24").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J25").Value = "" Then
        .Range("J25").Value = "C3 "
        .Range("J25").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J25").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K25").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K25").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J25").Value = ""
        .Range("J25").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K25").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J26").Value = "" Then
        .Range("J26").Value = "iC4 "
        .Range("J26").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J26").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K26").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K26").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J26").Value = ""
        .Range("J26").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K26").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J27").Value = "" Then
        .Range("J27").Value = "nC4 "
        .Range("J27").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J27").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("J27").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("J27").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("K27").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K27").Borders(xlEdgeRight).Weight = xlMedium
        .Range("K27").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("K27").Borders(xlEdgeBottom).Weight = xlMedium
    Else
        .Range("J27").Value = ""
        .Range("J27").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("J27").Borders(xlEdgeBottom).LineStyle = xlNone
        .Range("K27").Borders(xlEdgeRight).LineStyle = xlNone
        .Range("K27").Borders(xlEdgeBottom).LineStyle = xlNone
    End If
    
    If .Range("J29").Value = "" Then
        .Range("J29").Value = "CHROMATOGRAPH 3"
        .Range("J29:K29").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("J29:K29").Borders(xlEdgeTop).Weight = xlMedium
        .Range("J29").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J29").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K29").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K29").Borders(xlEdgeRight).Weight = xlMedium
        .Range("J29:K29").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("J29:K29").Borders(xlEdgeBottom).Weight = xlMedium
    Else
        .Range("J29").Value = ""
        .Range("J29:K29").Borders(xlEdgeTop).LineStyle = xlNone
        .Range("J29").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K29").Borders(xlEdgeRight).LineStyle = xlNone
        .Range("J29:K29").Borders(xlEdgeBottom).LineStyle = xlNone
    End If
    If .Range("J30").Value = "" Then
        .Range("J30").Value = "HW "
        .Range("J30").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J30").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K30").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K30").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J30").Value = ""
        .Range("J30").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K30").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J31").Value = "" Then
        .Range("J31").Value = "C1 "
        .Range("J31").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J31").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K31").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K31").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J31").Value = ""
        .Range("J31").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K31").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J32").Value = "" Then
        .Range("J32").Value = "C2 "
        .Range("J32").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J32").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K32").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K32").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J32").Value = ""
        .Range("J32").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K32").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J33").Value = "" Then
        .Range("J33").Value = "C3 "
        .Range("J33").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J33").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K33").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K33").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J33").Value = ""
        .Range("J33").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K33").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J34").Value = "" Then
        .Range("J34").Value = "iC4 "
        .Range("J34").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J34").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K34").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K34").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J34").Value = ""
        .Range("J34").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K34").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J35").Value = "" Then
        .Range("J35").Value = "nC4 "
        .Range("J35").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J35").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("J35").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("J35").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("K35").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K35").Borders(xlEdgeRight).Weight = xlMedium
        .Range("K35").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("K35").Borders(xlEdgeBottom).Weight = xlMedium
    Else
        .Range("J35").Value = ""
        .Range("J35").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("J35").Borders(xlEdgeBottom).LineStyle = xlNone
        .Range("K35").Borders(xlEdgeRight).LineStyle = xlNone
        .Range("K35").Borders(xlEdgeBottom).LineStyle = xlNone
    End If
    
    If .Range("J37").Value = "" Then
        .Range("J37").Value = "CHROMATOGRAPH 4"
        .Range("J37:K37").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("J37:K37").Borders(xlEdgeTop).Weight = xlMedium
        .Range("J37").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J37").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K37").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K37").Borders(xlEdgeRight).Weight = xlMedium
        .Range("J37:K37").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("J37:K37").Borders(xlEdgeBottom).Weight = xlMedium
    Else
        .Range("J37").Value = ""
        .Range("J37:K37").Borders(xlEdgeTop).LineStyle = xlNone
        .Range("J37").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K37").Borders(xlEdgeRight).LineStyle = xlNone
        .Range("J37:K37").Borders(xlEdgeBottom).LineStyle = xlNone
    End If
    If .Range("J38").Value = "" Then
        .Range("J38").Value = "HW "
        .Range("J38").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J38").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K38").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K38").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J38").Value = ""
        .Range("J38").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K38").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J39").Value = "" Then
        .Range("J39").Value = "C1 "
        .Range("J39").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J39").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K39").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K39").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J39").Value = ""
        .Range("J39").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K39").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J40").Value = "" Then
        .Range("J40").Value = "C2 "
        .Range("J40").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J40").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K40").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K40").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J40").Value = ""
        .Range("J40").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K40").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J41").Value = "" Then
        .Range("J41").Value = "C3 "
        .Range("J41").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J41").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K41").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K41").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J41").Value = ""
        .Range("J41").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K41").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J42").Value = "" Then
        .Range("J42").Value = "iC4 "
        .Range("J42").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J42").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K42").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K42").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J42").Value = ""
        .Range("J42").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K42").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J43").Value = "" Then
        .Range("J43").Value = "nC4 "
        .Range("J43").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J43").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("J43").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("J43").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("K43").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K43").Borders(xlEdgeRight).Weight = xlMedium
        .Range("K43").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("K43").Borders(xlEdgeBottom).Weight = xlMedium
    Else
        .Range("J43").Value = ""
        .Range("J43").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("J43").Borders(xlEdgeBottom).LineStyle = xlNone
        .Range("K43").Borders(xlEdgeRight).LineStyle = xlNone
        .Range("K43").Borders(xlEdgeBottom).LineStyle = xlNone
    End If
    
    If .Range("J45").Value = "" Then
        .Range("J45").Value = "CHROMATOGRAPH 5"
        .Range("J45:K45").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("J45:K45").Borders(xlEdgeTop).Weight = xlMedium
        .Range("J45").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J45").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K45").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K45").Borders(xlEdgeRight).Weight = xlMedium
        .Range("J45:K45").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("J45:K45").Borders(xlEdgeBottom).Weight = xlMedium
    Else
        .Range("J45").Value = ""
        .Range("J45:K45").Borders(xlEdgeTop).LineStyle = xlNone
        .Range("J45").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K45").Borders(xlEdgeRight).LineStyle = xlNone
        .Range("J45:K45").Borders(xlEdgeBottom).LineStyle = xlNone
    End If
    If .Range("J46").Value = "" Then
        .Range("J46").Value = "HW "
        .Range("J46").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J46").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K46").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K46").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J46").Value = ""
        .Range("J46").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K46").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J47").Value = "" Then
        .Range("J47").Value = "C1 "
        .Range("J47").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J47").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K47").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K47").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J47").Value = ""
        .Range("J47").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K47").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J48").Value = "" Then
        .Range("J48").Value = "C2 "
        .Range("J48").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J48").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K48").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K48").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J48").Value = ""
        .Range("J48").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K48").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J49").Value = "" Then
        .Range("J49").Value = "C3 "
        .Range("J49").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J49").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K49").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K49").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J49").Value = ""
        .Range("J49").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K49").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J50").Value = "" Then
        .Range("J50").Value = "iC4 "
        .Range("J50").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J50").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("K50").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K50").Borders(xlEdgeRight).Weight = xlMedium
    Else
        .Range("J50").Value = ""
        .Range("J50").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("K50").Borders(xlEdgeRight).LineStyle = xlNone
    End If
    If .Range("J51").Value = "" Then
        .Range("J51").Value = "nC4 "
        .Range("J51").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J51").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("J51").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("J51").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("K51").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K51").Borders(xlEdgeRight).Weight = xlMedium
        .Range("K51").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("K51").Borders(xlEdgeBottom).Weight = xlMedium
    Else
        .Range("J51").Value = ""
        .Range("J51").Borders(xlEdgeLeft).LineStyle = xlNone
        .Range("J51").Borders(xlEdgeBottom).LineStyle = xlNone
        .Range("K51").Borders(xlEdgeRight).LineStyle = xlNone
        .Range("K51").Borders(xlEdgeBottom).LineStyle = xlNone
    End If

    Application.ScreenUpdating = True

    .Protect Password:="NCS117"

    End With
    

End Sub