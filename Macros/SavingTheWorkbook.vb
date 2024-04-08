Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Sub SaveToFolder(control As IRibbonControl)
'Created for Neset Consulting Service by Dragon Wood (November, 2011)
'Saves the file to a set folder path based on the cell values on the Surveys page

'Declare the Variables for Saving the File
    Dim fileSaveName As String
    Dim wellSaveName As String
    Dim rigSaveName As String
    Dim compSaveName As String

'Declare the Variables for the Directory Path
    Dim fileRootPath As String
    Dim fileSavePath As String
    Dim dirDepth As Long
    Dim nextDir As Long
    Dim tempDir As String
    Dim x As Long

'Declare the Varialbles for the Input Boxes
    Dim compInput As String
    Dim wellInput As String
    Dim rigInput As String

'Declare the Varibles for the Message Boxes
    Dim msgCreate As String
    Dim msgStyle As String
    Dim msgTitle As String

'Declare the Constants for the Shortcut
    Const sctLocation As String = "Desktop"
    Const sctLinkExt As String = ".lnk"

'Declare the Objects for the Shortcut
    Dim oFolder As Object
    Dim oShortcut As Object

'Declare the variables for the Shortcut Path
    Dim sctDeskTopPath As String
    Dim sctShortcut As String
    Dim sctSep As String
    Dim sctName As String
    Dim sctSaveName As String


 
 'Unhide the sheets if still hidden
    
    ActiveWorkbook.Unprotect Password:="NCS117"
    
    Dim ws As Worksheet
        Application.DisplayAlerts = False
    For Each ws In ActiveWorkbook.Worksheets
        ws.Visible = True
    Next ws
        Application.DisplayAlerts = True
    
    ActiveWorkbook.Protect Password:="NCS117"
 
 'Make the Surveys page the focus point
 
    Application.Goto Sheets("Surveys").Range("A1"), True
    
 'Check the Company Name, Well Name, and Rig Name fields for content. If there, use the content, if not provide an input box for entering the data.
 
    With Sheets("Surveys")
    If .Range("D2").Value = "" Then
        .Range("D2").Select
        compInput = InputBox("Please enter the Company Name.", "Company Name")
        ActiveCell.FormulaR1C1 = compInput
    End If
    If .Range("D5").Value = "" Then
        .Range("D5").Select
        wellInput = InputBox("Please fill in the Well Name.", "Well Name")
        ActiveCell.FormulaR1C1 = wellInput
    End If
    If .Range("D6").Value = "" Then
        .Range("D6").Select
        rigInput = InputBox("Please fill in the Rig Name.", "Rig Name")
        ActiveCell.FormulaR1C1 = rigInput
    End If
    
        fileSaveName = CleanFileName(.Range("D5").Value) & " - Cross Section" & ".xlsm"
        wellSaveName = CleanFileName(.Range("D5").Value) & "\"
        rigSaveName = CleanFileName(.Range("D6").Value) & "\"
        compSaveName = CleanFileName(.Range("D2").Value) & "\"
        sctSaveName = CleanFileName(.Range("D5").Value)
    End With
    
 'Set the Root Path
 
    fileRootPath = "C:\Neset\Wells\"
    
 'Set the sub paths
 
    fileSavePath = fileRootPath & compSaveName & rigSaveName & wellSaveName
    
 'Check for directory name, use it if there, create it if not
 
    If Dir(fileSavePath, vbDirectory) = "" Then
        dirDepth = Len(fileSavePath) - Len(Replace(fileSavePath, "\", ""))
        nextDir = InStr(fileSavePath, "\")
        For x = 1 To dirDepth - 1
            nextDir = InStr(nextDir + 1, fileSavePath, "\")
            tempDir = Left(fileSavePath, nextDir)
            If Dir(tempDir, vbDirectory) = "" Then MkDir tempDir
        Next x
    End If
    
'Save the workbook under the new directory and name

    ActiveWorkbook.SaveAs Filename:=fileSavePath & fileSaveName


On Error GoTo ErrHandle
'Determine the path to the Desktop
    Set oFolder = CreateObject("WScript.Shell")
    sctDeskTopPath = oFolder.SpecialFolders(sctLocation)

'Determine the location of the Shortcut
    sctShortcut = Mid(fileSavePath, InStrRev(fileSavePath, "\") + 1)

'Separate the workbook name from the rest of the link
    sctSep = Application.PathSeparator
    sctName = sctSep & sctSaveName
    sctShortcut = sctDeskTopPath & sctName & sctLinkExt
    
'Create the Shortcut
    Set oShortcut = oFolder.CreateShortCut(sctShortcut)

'Establish the link to the File
    With oShortcut
        .TargetPath = fileSavePath
         .Save
    End With
    
'Clear Memory
    Set oFolder = Nothing
    Set oShortcut = Nothing

'Display the Success Message
    msgCreate = "The file can be found here: " & fileSavePath & vbNewLine
    msgCreate = msgCreate & "There is a shortcut to the folder on your desktop."
    msgStyle = 0
    msgTitle = "The file was created successfully"
    MsgBox msgCreate, msgStyle, msgTitle
     
    Exit Sub
     
'Display the Error Message
ErrHandle:
    msgCreate = "Please close and reopen the original file."
    msgStyle = 48
    msgTitle = "The file was not created"
    MsgBox msgCreate, msgStyle, msgTitle
    
End Sub

Function CleanFileName(sFileName As String, Optional ReplaceInvalidwith As String = "") As String
    'Removes invalid filename characters
    
    Const InvalidChars As String = "%~:\/?*<>|"""
    Dim ThisChar As Long
    CleanFileName = sFileName
    For ThisChar = 1 To Len(InvalidChars)
        CleanFileName = Replace(CleanFileName, Mid(InvalidChars, ThisChar, 1), ReplaceInvalidwith)
    Next
End Function

Public Sub SaveCopyCrossSection(control As IRibbonControl)
'Created for Neset Consulting Service by Dragon Wood (August, 2011)
'Increments the reference cell for the new Workbook Name
'Creates a copy of this workbook in the same directory as the original
'Provides the option of changing the code to specify a specific directory

'Creates a reference point in the workbook for incrementing

    With Sheets("Instructions")
        .Unprotect Password:="NCS117"
    
    Dim strLogNumber As String
    strLogNumber = .Range("AG5").Value
 
    If Len(strLogNumber) = 0 Then
        strLogNumber = " (Copy)1"
    Else
        strLogNumber = " (Copy)" & CLng(Mid(strLogNumber, 7, Len(strLogNumber))) + 1
    End If
 
        .Range("AG5").Value = strLogNumber
        
        .Protect Password:="NCS117"
    
    End With

'Creates a copy of the log and renames it according to how many times you've copied the file

    With ActiveWorkbook
        .Unprotect Password:="NCS117"
    
    Dim strFilename As Variant
    Dim CurrentName As Variant
    Dim Sheet_Name As String
    Dim Current_Path As String
 
    CurrentName = Split(ThisWorkbook.Name, ".")
    Sheet_Name = "Instructions"
    Current_Path = ThisWorkbook.Path
 
    Set strFilename = Sheets(Sheet_Name).Range("AG5")


    stPos = InStr(1, CurrentName(0), " (Copy)", vbBinaryCompare)

    If stPos > 1 Then
        CurrentName(0) = Mid(CurrentName(0), 1, stPos - 1)
    Else
    End If
 
        .SaveAs Filename:=Current_Path & "\" & CurrentName(0) & strFilename.Value & ".xlsm"
        
        .Protect Password:="NCS117"
    
    End With

End Sub