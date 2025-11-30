REM ***** BASIC *****
REM State Machine Generator for IEC 61131-3 Structured Text
REM This module converts state machine definitions in spreadsheet to ST Function Block code

Option Explicit

Sub GenerateStateMachine()
    ' Main entry point - generates complete ST code
    Dim sCode As String
    Dim oSheet As Object
    Dim oDoc As Object
    
    oDoc = ThisComponent
    
    ' Generate all sections
    sCode = GenerateFunctionBlockHeader(oDoc)
    sCode = sCode & GenerateVarDeclarations(oDoc)
    sCode = sCode & GenerateStateMachineLogic(oDoc)
    sCode = sCode & GenerateFunctionBlockFooter()
    
    ' Display or save the result
    ExportToFile(sCode, oDoc)
    
End Sub

Function GenerateFunctionBlockHeader(oDoc As Object) As String
    ' Generate FUNCTION_BLOCK declaration and metadata comments
    Dim oSheet As Object
    Dim sName As String
    Dim sAuthor As String
    Dim sVersion As String
    Dim sCode As String
    
    oSheet = oDoc.Sheets.getByName("Config")
    sName = GetCellValue(oSheet, 1, 0)      ' B1 - Program Name
    sAuthor = GetCellValue(oSheet, 1, 1)    ' B2 - Author
    sVersion = GetCellValue(oSheet, 1, 2)   ' B3 - Version
    
    sCode = "(* ======================================== *)" & Chr(10)
    sCode = sCode & "(* State Machine: " & sName & " *)" & Chr(10)
    If Len(sAuthor) > 0 Then
        sCode = sCode & "(* Author: " & sAuthor & " *)" & Chr(10)
    End If
    sCode = sCode & "(* Version: " & sVersion & " *)" & Chr(10)
    sCode = sCode & "(* Generated: " & Format(Now(), "YYYY-MM-DD HH:MM:SS") & " *)" & Chr(10)
    sCode = sCode & "(* ======================================== *)" & Chr(10)
    sCode = sCode & Chr(10)
    sCode = sCode & "FUNCTION_BLOCK " & sName & Chr(10)
    
    GenerateFunctionBlockHeader = sCode
End Function

Function GenerateVarDeclarations(oDoc As Object) As String
    ' Generate VAR_INPUT, VAR_OUTPUT, and VAR sections
    Dim sCode As String
    
    sCode = Chr(10)
    sCode = sCode & GenerateInputVars(oDoc)
    sCode = sCode & Chr(10)
    sCode = sCode & GenerateOutputVars(oDoc)
    sCode = sCode & Chr(10)
    sCode = sCode & GenerateInternalVars(oDoc)
    
    GenerateVarDeclarations = sCode
End Function

Function GenerateInputVars(oDoc As Object) As String
    ' Generate VAR_INPUT section
    Dim oSheet As Object
    Dim sCode As String
    Dim iRow As Integer
    Dim sVarName As String
    Dim sDescription As String
    
    oSheet = oDoc.Sheets.getByName("Inputs")
    sCode = "VAR_INPUT" & Chr(10)
    
    ' Start from row 2 (row index 1) - skip header
    iRow = 1
    Do While True
        sVarName = GetCellValue(oSheet, 1, iRow)  ' Column B
        If Len(sVarName) = 0 Then Exit Do
        
        sDescription = GetCellValue(oSheet, 2, iRow)  ' Column C
        
        If Len(sDescription) > 0 Then
            sCode = sCode & "    " & sVarName & " : BOOL;  (* " & sDescription & " *)" & Chr(10)
        Else
            sCode = sCode & "    " & sVarName & " : BOOL;" & Chr(10)
        End If
        
        iRow = iRow + 1
    Loop
    
    sCode = sCode & "END_VAR" & Chr(10)
    
    GenerateInputVars = sCode
End Function

Function GenerateOutputVars(oDoc As Object) As String
    ' Generate VAR_OUTPUT section
    Dim oSheet As Object
    Dim sCode As String
    Dim iRow As Integer
    Dim sVarName As String
    Dim sDescription As String
    
    oSheet = oDoc.Sheets.getByName("Outputs")
    sCode = "VAR_OUTPUT" & Chr(10)
    
    ' Start from row 2 (row index 1) - skip header
    iRow = 1
    Do While True
        sVarName = GetCellValue(oSheet, 1, iRow)  ' Column B
        If Len(sVarName) = 0 Then Exit Do
        
        sDescription = GetCellValue(oSheet, 2, iRow)  ' Column C
        
        If Len(sDescription) > 0 Then
            sCode = sCode & "    " & sVarName & " : BOOL;  (* " & sDescription & " *)" & Chr(10)
        Else
            sCode = sCode & "    " & sVarName & " : BOOL;" & Chr(10)
        End If
        
        iRow = iRow + 1
    Loop
    
    sCode = sCode & "END_VAR" & Chr(10)
    
    GenerateOutputVars = sCode
End Function

Function GenerateInternalVars(oDoc As Object) As String
    ' Generate VAR section with state enumeration and state variable
    Dim oSheet As Object
    Dim sCode As String
    Dim sStateVarName As String
    Dim sInitialState As String
    
    oSheet = oDoc.Sheets.getByName("Config")
    sStateVarName = GetCellValue(oSheet, 1, 4)  ' B5 - State Variable Name
    sInitialState = GetCellValue(oSheet, 1, 5)  ' B6 - Initial State
    
    sCode = "VAR" & Chr(10)
    sCode = sCode & "    " & sStateVarName & " : " & GenerateStateEnum(oDoc) & " := " & sInitialState & ";" & Chr(10)
    sCode = sCode & "END_VAR" & Chr(10)
    
    GenerateInternalVars = sCode
End Function

Function GenerateStateEnum(oDoc As Object) As String
    ' Generate state enumeration type inline
    Dim oSheet As Object
    Dim sCode As String
    Dim iRow As Integer
    Dim sStateName As String
    Dim bFirst As Boolean
    
    oSheet = oDoc.Sheets.getByName("States")
    sCode = "("
    bFirst = True
    
    ' Start from row 2 (row index 1) - skip header
    iRow = 1
    Do While True
        sStateName = GetCellValue(oSheet, 1, iRow)  ' Column B
        If Len(sStateName) = 0 Then Exit Do
        
        If Not bFirst Then
            sCode = sCode & ", "
        End If
        sCode = sCode & sStateName
        bFirst = False
        
        iRow = iRow + 1
    Loop
    
    sCode = sCode & ")"
    
    GenerateStateEnum = sCode
End Function

Function GenerateStateMachineLogic(oDoc As Object) As String
    ' Generate main state machine CASE statement
    Dim oSheet As Object
    Dim sCode As String
    Dim sStateVarName As String
    
    oSheet = oDoc.Sheets.getByName("Config")
    sStateVarName = GetCellValue(oSheet, 1, 4)  ' B5 - State Variable Name
    
    sCode = Chr(10)
    sCode = sCode & "(* State Machine Logic *)" & Chr(10)
    sCode = sCode & "CASE " & sStateVarName & " OF" & Chr(10)
    sCode = sCode & GenerateStateCases(oDoc)
    sCode = sCode & "END_CASE;" & Chr(10)
    
    GenerateStateMachineLogic = sCode
End Function

Function GenerateStateCases(oDoc As Object) As String
    ' Generate each state case with actions and transitions
    Dim oStatesSheet As Object
    Dim sCode As String
    Dim iRow As Integer
    Dim sStateName As String
    
    oStatesSheet = oDoc.Sheets.getByName("States")
    sCode = ""
    
    ' Start from row 2 (row index 1) - skip header
    iRow = 1
    Do While True
        sStateName = GetCellValue(oStatesSheet, 1, iRow)  ' Column B
        If Len(sStateName) = 0 Then Exit Do
        
        sCode = sCode & Chr(10)
        sCode = sCode & "    " & sStateName & ":" & Chr(10)
        sCode = sCode & GenerateStateActions(oDoc, sStateName)
        sCode = sCode & GenerateStateTransitions(oDoc, sStateName)
        
        iRow = iRow + 1
    Loop
    
    GenerateStateCases = sCode
End Function

Function GenerateStateActions(oDoc As Object, sStateName As String) As String
    ' Generate output assignments for a specific state
    Dim oActionsSheet As Object
    Dim oOutputsSheet As Object
    Dim sCode As String
    Dim iStateRow As Integer
    Dim iOutputCol As Integer
    Dim sOutputName As String
    Dim iValue As Integer
    
    oActionsSheet = oDoc.Sheets.getByName("StateActions")
    oOutputsSheet = oDoc.Sheets.getByName("Outputs")
    
    ' Find the state row in StateActions sheet
    iStateRow = FindStateRow(oActionsSheet, sStateName)
    If iStateRow = -1 Then
        GenerateStateActions = ""
        Exit Function
    End If
    
    sCode = "        (* Actions *)" & Chr(10)
    
    ' Iterate through output columns (columns B to Q, indices 1 to 16)
    For iOutputCol = 1 To 16
        sOutputName = GetCellValue(oOutputsSheet, 1, iOutputCol)  ' Get output name from Outputs sheet
        If Len(sOutputName) = 0 Then Exit For
        
        iValue = Val(GetCellValue(oActionsSheet, iOutputCol, iStateRow))
        
        If iValue = 1 Then
            sCode = sCode & "        " & sOutputName & " := TRUE;" & Chr(10)
        Else
            sCode = sCode & "        " & sOutputName & " := FALSE;" & Chr(10)
        End If
    Next iOutputCol
    
    sCode = sCode & Chr(10)
    
    GenerateStateActions = sCode
End Function

Function FindStateRow(oSheet As Object, sStateName As String) As Integer
    ' Find row index for a given state name in StateActions sheet
    Dim iRow As Integer
    Dim sCellValue As String
    
    iRow = 1  ' Start from row 2 (index 1) - skip header
    Do While True
        sCellValue = GetCellValue(oSheet, 0, iRow)  ' Column A
        If Len(sCellValue) = 0 Then
            FindStateRow = -1
            Exit Function
        End If
        
        If sCellValue = sStateName Then
            FindStateRow = iRow
            Exit Function
        End If
        
        iRow = iRow + 1
    Loop
End Function

Function GenerateStateTransitions(oDoc As Object, sFromState As String) As String
    ' Generate IF-ELSIF chain for state transitions
    Dim oSheet As Object
    Dim sCode As String
    Dim iRow As Integer
    Dim sFrom As String
    Dim sTo As String
    Dim sCondition As String
    Dim sComment As String
    Dim bFirst As Boolean
    Dim oConfigSheet As Object
    Dim sStateVarName As String
    
    oSheet = oDoc.Sheets.getByName("Transitions")
    oConfigSheet = oDoc.Sheets.getByName("Config")
    sStateVarName = GetCellValue(oConfigSheet, 1, 4)  ' B5
    
    sCode = "        (* Transitions *)" & Chr(10)
    bFirst = True
    
    ' Start from row 2 (row index 1) - skip header
    iRow = 1
    Do While True
        sFrom = GetCellValue(oSheet, 0, iRow)  ' Column A
        If Len(sFrom) = 0 Then Exit Do
        
        If sFrom = sFromState Then
            sTo = GetCellValue(oSheet, 1, iRow)       ' Column B
            sCondition = GetCellValue(oSheet, 2, iRow) ' Column C
            sComment = GetCellValue(oSheet, 3, iRow)   ' Column D
            
            If bFirst Then
                sCode = sCode & "        IF "
                bFirst = False
            Else
                sCode = sCode & "        ELSIF "
            End If
            
            sCode = sCode & sCondition & " THEN"
            If Len(sComment) > 0 Then
                sCode = sCode & "  (* " & sComment & " *)"
            End If
            sCode = sCode & Chr(10)
            sCode = sCode & "            " & sStateVarName & " := " & sTo & ";" & Chr(10)
        End If
        
        iRow = iRow + 1
    Loop
    
    If Not bFirst Then
        sCode = sCode & "        END_IF;" & Chr(10)
    Else
        sCode = sCode & "        (* No transitions *)" & Chr(10)
    End If
    
    GenerateStateTransitions = sCode
End Function

Function GenerateFunctionBlockFooter() As String
    ' Generate FUNCTION_BLOCK closing
    GenerateFunctionBlockFooter = Chr(10) & "END_FUNCTION_BLOCK" & Chr(10)
End Function

Function GetCellValue(oSheet As Object, iCol As Integer, iRow As Integer) As String
    ' Get cell value as string (0-indexed)
    Dim oCell As Object
    Dim sValue As String
    
    On Error Resume Next
    oCell = oSheet.getCellByPosition(iCol, iRow)
    sValue = Trim(oCell.getString())
    
    GetCellValue = sValue
End Function

Sub ExportToFile(sCode As String, oDoc As Object)
    ' Export generated code to a file
    Dim oSheet As Object
    Dim sFileName As String
    Dim sFilePath As String
    Dim oFilePicker As Object
    Dim oSimpleFileAccess As Object
    Dim oTextOutputStream As Object
    Dim oOutputStream As Object
    
    oSheet = oDoc.Sheets.getByName("Config")
    sFileName = GetCellValue(oSheet, 1, 1) & ".st"
    
    ' Create file picker dialog
    oFilePicker = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
    oFilePicker.Initialize(Array(com.sun.star.ui.dialogs.TemplateDescription.FILESAVE_SIMPLE))
    oFilePicker.setDefaultName(sFileName)
    oFilePicker.appendFilter("Structured Text", "*.st")
    oFilePicker.appendFilter("All Files", "*.*")
    oFilePicker.setCurrentFilter("Structured Text")
    
    If oFilePicker.Execute() = com.sun.star.ui.dialogs.ExecutableDialogResults.OK Then
        sFilePath = oFilePicker.Files(0)
        
        ' Write to file using UNO services (cross-platform compatible)
        On Error GoTo ErrorHandler
        oSimpleFileAccess = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
        
        ' Delete file if it exists
        If oSimpleFileAccess.exists(sFilePath) Then
            oSimpleFileAccess.kill(sFilePath)
        End If
        
        ' Create output stream
        oOutputStream = oSimpleFileAccess.openFileWrite(sFilePath)
        oTextOutputStream = CreateUnoService("com.sun.star.io.TextOutputStream")
        oTextOutputStream.setOutputStream(oOutputStream)
        oTextOutputStream.setEncoding("UTF-8")
        
        ' Write the code
        oTextOutputStream.writeString(sCode)
        
        ' Close streams
        oTextOutputStream.closeOutput()
        oOutputStream.closeOutput()
        
        MsgBox "State machine code generated successfully!" & Chr(10) & Chr(10) & _
               "File: " & ConvertFromUrl(sFilePath), 64, "Success"
        Exit Sub
        
ErrorHandler:
        MsgBox "Error writing file: " & Error$ & Chr(10) & Chr(10) & _
               "Path: " & sFilePath, 16, "Error"
    End If
End Sub


Sub PreviewStateMachine()
    ' Preview generated code in a scrollable dialog
    Dim sCode As String
    Dim oDoc As Object
    Dim oDialogModel As Object
    Dim oDialog As Object
    Dim oTextModel As Object
    Dim oButtonModel As Object
    Dim oSaveButtonModel As Object
    
    oDoc = ThisComponent
    
    ' Generate code
    sCode = GenerateFunctionBlockHeader(oDoc)
    sCode = sCode & GenerateVarDeclarations(oDoc)
    sCode = sCode & GenerateStateMachineLogic(oDoc)
    sCode = sCode & GenerateFunctionBlockFooter()
    
    ' Create dialog
    oDialogModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    oDialogModel.Width = 300
    oDialogModel.Height = 250
    oDialogModel.Title = "State Machine Preview"
    
    ' Create scrollable text field
    oTextModel = oDialogModel.createInstance("com.sun.star.awt.UnoControlEditModel")
    oTextModel.PositionX = 5
    oTextModel.PositionY = 5
    oTextModel.Width = 290
    oTextModel.Height = 215
    oTextModel.MultiLine = True
    oTextModel.VScroll = True
    oTextModel.HScroll = True
    oTextModel.ReadOnly = True
    oTextModel.Text = sCode
    oTextModel.FontName = "Courier New"
    oTextModel.FontHeight = 10
    oDialogModel.insertByName("TextField", oTextModel)
    
    ' Create Save button
    oSaveButtonModel = oDialogModel.createInstance("com.sun.star.awt.UnoControlButtonModel")
    oSaveButtonModel.PositionX = 80
    oSaveButtonModel.PositionY = 225
    oSaveButtonModel.Width = 60
    oSaveButtonModel.Height = 15
    oSaveButtonModel.Label = "Save to File"
    oSaveButtonModel.PushButtonType = 0  ' Standard button
    oDialogModel.insertByName("SaveButton", oSaveButtonModel)
    
    ' Create Close button
    oButtonModel = oDialogModel.createInstance("com.sun.star.awt.UnoControlButtonModel")
    oButtonModel.PositionX = 160
    oButtonModel.PositionY = 225
    oButtonModel.Width = 50
    oButtonModel.Height = 15
    oButtonModel.Label = "Close"
    oButtonModel.PushButtonType = 2  ' OK button
    oDialogModel.insertByName("CloseButton", oButtonModel)
    
    ' Create and execute dialog
    oDialog = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    oDialog.setModel(oDialogModel)
    oDialog.setVisible(True)
    
    ' Add action listener for Save button
    Dim oSaveButton As Object
    oSaveButton = oDialog.getControl("SaveButton")
    Dim oActionListener As Object
    oActionListener = CreateUnoListener("SaveButton_", "com.sun.star.awt.XActionListener")
    GlobalCodeToSave = sCode  ' Store code globally for save action
    GlobalDoc = oDoc
    oSaveButton.addActionListener(oActionListener)
    
    oDialog.execute()
    oDialog.dispose()
    
End Sub

' Global variables for save action
Global GlobalCodeToSave As String
Global GlobalDoc As Object

Sub SaveButton_actionPerformed(oEvent As Object)
    ' Save button action handler
    ExportToFile(GlobalCodeToSave, GlobalDoc)
End Sub

Sub SaveButton_disposing(oEvent As Object)
    ' Disposing handler (required for listener)
End Sub
