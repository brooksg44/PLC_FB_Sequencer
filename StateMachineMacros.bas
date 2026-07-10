REM ***** BASIC *****
REM State Machine Generator for IEC 61131-3 Structured Text
REM Converts the state machine definition in this spreadsheet into an ST Function Block.
REM
REM Entry points:
REM   PreviewStateMachine  - validate, then preview the generated code in a dialog
REM   GenerateStateMachine - validate, then export the generated code to a .st file
REM
REM Config sheet options (label in column A, value in column B; rows may be in any order):
REM   Program Name          - Function Block name (required)
REM   Author                - metadata comment (optional)
REM   Version               - metadata comment (optional)
REM   State Variable Name   - internal state variable name (required)
REM   Initial State         - starting state, must exist in the States sheet (required)
REM   State Enum Type Name  - if set, states are declared as a named TYPE before the
REM                           Function Block (portable to CODESYS/TwinCAT); if blank,
REM                           an inline enumeration is used
REM   State Output Variable - if set, an INT output with this name reports the current
REM                           state number (row order in the States sheet) for
REM                           HMI/diagnostics

Option Explicit

' Globals used by the preview dialog's Save button listener
Global GlobalCodeToSave As String
Global GlobalDoc As Object

' ================= Entry points =================

Sub GenerateStateMachine()
    ' Validate the spreadsheet, generate ST code, and export to a .st file
    Dim sCode As String

    sCode = BuildValidatedCode(ThisComponent)
    If Len(sCode) = 0 Then Exit Sub

    ExportToFile(sCode, ThisComponent)
End Sub

Sub PreviewStateMachine()
    ' Validate the spreadsheet and preview the generated code in a scrollable dialog
    Dim sCode As String
    Dim oDialogModel As Object
    Dim oDialog As Object
    Dim oTextModel As Object
    Dim oButtonModel As Object
    Dim oSaveButtonModel As Object

    sCode = BuildValidatedCode(ThisComponent)
    If Len(sCode) = 0 Then Exit Sub

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
    Dim oActionListener As Object
    oSaveButton = oDialog.getControl("SaveButton")
    oActionListener = CreateUnoListener("SaveButton_", "com.sun.star.awt.XActionListener")
    GlobalCodeToSave = sCode  ' Store code globally for save action
    GlobalDoc = ThisComponent
    oSaveButton.addActionListener(oActionListener)

    oDialog.execute()
    oDialog.dispose()
End Sub

Sub SaveButton_actionPerformed(oEvent As Object)
    ' Save button action handler
    ExportToFile(GlobalCodeToSave, GlobalDoc)
End Sub

Sub SaveButton_disposing(oEvent As Object)
    ' Disposing handler (required for listener)
End Sub

' ================= Code assembly =================

Function BuildValidatedCode(oDoc As Object) As String
    ' Validate the spreadsheet; on success return the generated code.
    ' Returns "" if errors were found or the user cancelled after warnings.
    Dim sErrors As String
    Dim sWarnings As String

    Call ValidateStateMachine(oDoc, sErrors, sWarnings)

    If Len(sErrors) > 0 Then
        MsgBox "Cannot generate code. Fix these problems first:" & NL() & NL() & sErrors, 16, "Validation Errors"
        BuildValidatedCode = ""
        Exit Function
    End If

    If Len(sWarnings) > 0 Then
        If MsgBox("Warnings:" & NL() & NL() & sWarnings & NL() & NL() & "Generate code anyway?", 4 + 48, "Validation Warnings") <> 6 Then
            BuildValidatedCode = ""
            Exit Function
        End If
    End If

    BuildValidatedCode = BuildStateMachineCode(oDoc)
End Function

Function BuildStateMachineCode(oDoc As Object) As String
    ' Assemble the complete Function Block source
    Dim sCode As String

    sCode = GenerateFunctionBlockHeader(oDoc)
    sCode = sCode & GenerateEnumTypeBlock(oDoc)
    sCode = sCode & "FUNCTION_BLOCK " & GetConfigValue(oDoc, "Program Name") & NL()
    sCode = sCode & GenerateVarDeclarations(oDoc)
    sCode = sCode & GenerateStateMachineLogic(oDoc)
    sCode = sCode & NL() & "END_FUNCTION_BLOCK" & NL()

    BuildStateMachineCode = sCode
End Function

Function GenerateFunctionBlockHeader(oDoc As Object) As String
    ' Generate metadata comment banner
    Dim sName As String
    Dim sAuthor As String
    Dim sVersion As String
    Dim sCode As String

    sName = GetConfigValue(oDoc, "Program Name")
    sAuthor = GetConfigValue(oDoc, "Author")
    sVersion = GetConfigValue(oDoc, "Version")

    sCode = "(* ======================================== *)" & NL()
    sCode = sCode & "(* State Machine: " & sName & " *)" & NL()
    If Len(sAuthor) > 0 Then
        sCode = sCode & "(* Author: " & sAuthor & " *)" & NL()
    End If
    If Len(sVersion) > 0 Then
        sCode = sCode & "(* Version: " & sVersion & " *)" & NL()
    End If
    sCode = sCode & "(* Generated: " & Format(Now(), "YYYY-MM-DD HH:MM:SS") & " *)" & NL()
    sCode = sCode & "(* ======================================== *)" & NL()
    sCode = sCode & NL()

    GenerateFunctionBlockHeader = sCode
End Function

Function GenerateEnumTypeBlock(oDoc As Object) As String
    ' Generate a named TYPE declaration for the states, if configured.
    ' Returns "" when 'State Enum Type Name' is blank (inline enum is used instead).
    Dim oSheet As Object
    Dim sTypeName As String
    Dim sCode As String
    Dim iRow As Integer
    Dim sStateName As String
    Dim sDescription As String
    Dim sNextState As String

    sTypeName = GetConfigValue(oDoc, "State Enum Type Name")
    If Len(sTypeName) = 0 Then
        GenerateEnumTypeBlock = ""
        Exit Function
    End If

    oSheet = oDoc.Sheets.getByName("States")
    sCode = "TYPE " & sTypeName & " :" & NL()
    sCode = sCode & "(" & NL()

    iRow = 1
    Do While True
        sStateName = GetCellValue(oSheet, 1, iRow)  ' Column B
        If Len(sStateName) = 0 Then Exit Do

        sDescription = GetCellValue(oSheet, 2, iRow)  ' Column C
        sNextState = GetCellValue(oSheet, 1, iRow + 1)

        sCode = sCode & "    " & sStateName
        If Len(sNextState) > 0 Then
            sCode = sCode & ","
        End If
        If Len(sDescription) > 0 Then
            sCode = sCode & "  (* " & sDescription & " *)"
        End If
        sCode = sCode & NL()

        iRow = iRow + 1
    Loop

    sCode = sCode & ");" & NL()
    sCode = sCode & "END_TYPE" & NL()
    sCode = sCode & NL()

    GenerateEnumTypeBlock = sCode
End Function

Function GenerateVarDeclarations(oDoc As Object) As String
    ' Generate VAR_INPUT, VAR_OUTPUT, and VAR sections
    Dim sCode As String
    Dim sStateOut As String
    Dim sStateOutDecl As String

    sStateOut = GetConfigValue(oDoc, "State Output Variable")
    sStateOutDecl = ""
    If Len(sStateOut) > 0 Then
        sStateOutDecl = "    " & sStateOut & " : INT;  (* Current state number (row order in States sheet) *)" & NL()
    End If

    sCode = NL()
    sCode = sCode & GenerateIOVars(oDoc, "Inputs", "VAR_INPUT", "")
    sCode = sCode & NL()
    sCode = sCode & GenerateIOVars(oDoc, "Outputs", "VAR_OUTPUT", sStateOutDecl)
    sCode = sCode & NL()
    sCode = sCode & GenerateInternalVars(oDoc)

    GenerateVarDeclarations = sCode
End Function

Function GenerateIOVars(oDoc As Object, sSheetName As String, sSection As String, sExtraDecl As String) As String
    ' Generate a VAR_INPUT or VAR_OUTPUT section from an I/O sheet
    Dim oSheet As Object
    Dim sCode As String
    Dim iRow As Integer
    Dim sVarName As String
    Dim sDescription As String

    oSheet = oDoc.Sheets.getByName(sSheetName)
    sCode = sSection & NL()

    ' Start from row 2 (row index 1) - skip header
    iRow = 1
    Do While True
        sVarName = GetCellValue(oSheet, 1, iRow)  ' Column B
        If Len(sVarName) = 0 Then Exit Do

        sDescription = GetCellValue(oSheet, 2, iRow)  ' Column C

        If Len(sDescription) > 0 Then
            sCode = sCode & "    " & sVarName & " : BOOL;  (* " & sDescription & " *)" & NL()
        Else
            sCode = sCode & "    " & sVarName & " : BOOL;" & NL()
        End If

        iRow = iRow + 1
    Loop

    sCode = sCode & sExtraDecl
    sCode = sCode & "END_VAR" & NL()

    GenerateIOVars = sCode
End Function

Function GenerateInternalVars(oDoc As Object) As String
    ' Generate VAR section with the state variable
    Dim sCode As String
    Dim sStateVarName As String
    Dim sInitialState As String
    Dim sEnumType As String
    Dim sTypeText As String

    sStateVarName = GetConfigValue(oDoc, "State Variable Name")
    sInitialState = GetConfigValue(oDoc, "Initial State")
    sEnumType = GetConfigValue(oDoc, "State Enum Type Name")

    If Len(sEnumType) > 0 Then
        sTypeText = sEnumType
    Else
        sTypeText = GenerateStateEnum(oDoc)
    End If

    sCode = "VAR" & NL()
    sCode = sCode & "    " & sStateVarName & " : " & sTypeText & " := " & sInitialState & ";" & NL()
    sCode = sCode & "END_VAR" & NL()

    GenerateInternalVars = sCode
End Function

Function GenerateStateEnum(oDoc As Object) As String
    ' Generate the inline state enumeration, e.g. (IDLE, RUNNING, ERROR)
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
    ' Generate the main state machine CASE statement
    Dim sCode As String
    Dim sStateVarName As String
    Dim sInitialState As String

    sStateVarName = GetConfigValue(oDoc, "State Variable Name")
    sInitialState = GetConfigValue(oDoc, "Initial State")

    sCode = NL()
    sCode = sCode & "(* State Machine Logic *)" & NL()
    sCode = sCode & "CASE " & sStateVarName & " OF" & NL()
    sCode = sCode & GenerateStateCases(oDoc)
    sCode = sCode & NL()
    sCode = sCode & "    ELSE" & NL()
    sCode = sCode & "        (* Unknown state: recover to initial state *)" & NL()
    sCode = sCode & "        " & sStateVarName & " := " & sInitialState & ";" & NL()
    sCode = sCode & "END_CASE;" & NL()

    GenerateStateMachineLogic = sCode
End Function

Function GenerateStateCases(oDoc As Object) As String
    ' Generate each state case with actions and transitions
    Dim oStatesSheet As Object
    Dim sCode As String
    Dim iRow As Integer
    Dim iStateIndex As Integer
    Dim sStateName As String

    oStatesSheet = oDoc.Sheets.getByName("States")
    sCode = ""

    ' Start from row 2 (row index 1) - skip header
    iRow = 1
    iStateIndex = 1
    Do While True
        sStateName = GetCellValue(oStatesSheet, 1, iRow)  ' Column B
        If Len(sStateName) = 0 Then Exit Do

        sCode = sCode & NL()
        sCode = sCode & "    " & sStateName & ":" & NL()
        sCode = sCode & GenerateStateActions(oDoc, sStateName, iStateIndex)
        sCode = sCode & GenerateStateTransitions(oDoc, sStateName)

        iRow = iRow + 1
        iStateIndex = iStateIndex + 1
    Loop

    GenerateStateCases = sCode
End Function

Function GenerateStateActions(oDoc As Object, sStateName As String, iStateIndex As Integer) As String
    ' Generate output assignments for a specific state.
    ' Output names are read from the StateActions header row, so column order
    ' in StateActions is authoritative and there is no fixed output limit.
    Dim oActionsSheet As Object
    Dim sCode As String
    Dim iStateRow As Integer
    Dim iCol As Integer
    Dim sOutputName As String
    Dim sStateOut As String

    oActionsSheet = oDoc.Sheets.getByName("StateActions")
    sStateOut = GetConfigValue(oDoc, "State Output Variable")

    sCode = "        (* Actions *)" & NL()

    If Len(sStateOut) > 0 Then
        sCode = sCode & "        " & sStateOut & " := " & iStateIndex & ";" & NL()
    End If

    iStateRow = FindRowByValue(oActionsSheet, 0, sStateName)
    If iStateRow = -1 Then
        sCode = sCode & "        (* No StateActions row for this state: outputs hold their previous values *)" & NL()
    Else
        ' Iterate StateActions header columns until the first empty header cell
        iCol = 1
        Do While True
            sOutputName = GetCellValue(oActionsSheet, iCol, 0)  ' Header row
            If Len(sOutputName) = 0 Then Exit Do

            If Val(GetCellValue(oActionsSheet, iCol, iStateRow)) = 1 Then
                sCode = sCode & "        " & sOutputName & " := TRUE;" & NL()
            Else
                sCode = sCode & "        " & sOutputName & " := FALSE;" & NL()
            End If

            iCol = iCol + 1
        Loop
    End If

    sCode = sCode & NL()

    GenerateStateActions = sCode
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
    Dim sStateVarName As String

    oSheet = oDoc.Sheets.getByName("Transitions")
    sStateVarName = GetConfigValue(oDoc, "State Variable Name")

    sCode = "        (* Transitions *)" & NL()
    bFirst = True

    ' Start from row 2 (row index 1) - skip header
    iRow = 1
    Do While True
        sFrom = GetCellValue(oSheet, 0, iRow)  ' Column A
        If Len(sFrom) = 0 Then Exit Do

        If SameText(sFrom, sFromState) Then
            sTo = GetCellValue(oSheet, 1, iRow)        ' Column B
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
            sCode = sCode & NL()
            sCode = sCode & "            " & sStateVarName & " := " & sTo & ";" & NL()
        End If

        iRow = iRow + 1
    Loop

    If Not bFirst Then
        sCode = sCode & "        END_IF;" & NL()
    Else
        sCode = sCode & "        (* No transitions *)" & NL()
    End If

    GenerateStateTransitions = sCode
End Function

' ================= Validation =================

Sub ValidateStateMachine(oDoc As Object, ByRef sErrors As String, ByRef sWarnings As String)
    ' Cross-check the whole spreadsheet before generating code.
    ' Errors block generation; warnings ask for confirmation.
    Dim oSheet As Object
    Dim oTransSheet As Object
    Dim aStates(499) As String
    Dim nStates As Integer
    Dim aIO(999) As String
    Dim nIO As Integer
    Dim aOutputs(499) As String
    Dim nOutputs As Integer
    Dim i As Integer
    Dim iRow As Integer
    Dim iCol As Integer
    Dim sName As String
    Dim bFound As Boolean
    Dim sFrom As String
    Dim sTo As String
    Dim sCondition As String

    sErrors = ""
    sWarnings = ""

    ' --- Required sheets ---
    Dim aSheetNames(5) As String
    aSheetNames(0) = "Config"
    aSheetNames(1) = "Inputs"
    aSheetNames(2) = "Outputs"
    aSheetNames(3) = "States"
    aSheetNames(4) = "Transitions"
    aSheetNames(5) = "StateActions"
    For i = 0 To 5
        If Not oDoc.Sheets.hasByName(aSheetNames(i)) Then
            Call AddMsg(sErrors, "Missing sheet: " & aSheetNames(i))
        End If
    Next i
    If Len(sErrors) > 0 Then Exit Sub  ' cannot check anything else

    ' --- Config values ---
    Dim sProgName As String
    Dim sStateVar As String
    Dim sInitState As String
    Dim sEnumType As String
    Dim sStateOut As String

    sProgName = GetConfigValue(oDoc, "Program Name")
    sStateVar = GetConfigValue(oDoc, "State Variable Name")
    sInitState = GetConfigValue(oDoc, "Initial State")
    sEnumType = GetConfigValue(oDoc, "State Enum Type Name")
    sStateOut = GetConfigValue(oDoc, "State Output Variable")

    If Len(sProgName) = 0 Then
        Call AddMsg(sErrors, "Config: 'Program Name' is empty")
    ElseIf Not IsValidIdentifier(sProgName) Then
        Call AddMsg(sErrors, "Config: Program Name '" & sProgName & "' is not a valid IEC identifier")
    End If

    If Len(sStateVar) = 0 Then
        Call AddMsg(sErrors, "Config: 'State Variable Name' is empty")
    ElseIf Not IsValidIdentifier(sStateVar) Then
        Call AddMsg(sErrors, "Config: State Variable Name '" & sStateVar & "' is not a valid IEC identifier")
    End If

    If Len(sEnumType) > 0 And Not IsValidIdentifier(sEnumType) Then
        Call AddMsg(sErrors, "Config: State Enum Type Name '" & sEnumType & "' is not a valid IEC identifier")
    End If

    If Len(sStateOut) > 0 And Not IsValidIdentifier(sStateOut) Then
        Call AddMsg(sErrors, "Config: State Output Variable '" & sStateOut & "' is not a valid IEC identifier")
    End If

    ' --- States ---
    oSheet = oDoc.Sheets.getByName("States")
    nStates = 0
    iRow = 1
    Do While True
        sName = GetCellValue(oSheet, 1, iRow)  ' Column B
        If Len(sName) = 0 Then Exit Do

        If Not IsValidIdentifier(sName) Then
            Call AddMsg(sErrors, "States: '" & sName & "' is not a valid IEC identifier")
        End If
        If InList(aStates, nStates, sName) Then
            Call AddMsg(sErrors, "States: duplicate state '" & sName & "'")
        End If
        aStates(nStates) = sName
        nStates = nStates + 1

        iRow = iRow + 1
    Loop
    If nStates = 0 Then
        Call AddMsg(sErrors, "States: no states defined")
    End If

    If Len(sInitState) = 0 Then
        Call AddMsg(sErrors, "Config: 'Initial State' is empty")
    ElseIf nStates > 0 And Not InList(aStates, nStates, sInitState) Then
        Call AddMsg(sErrors, "Config: Initial State '" & sInitState & "' is not defined in the States sheet")
    End If

    ' --- Inputs / Outputs ---
    nIO = 0
    nOutputs = 0

    oSheet = oDoc.Sheets.getByName("Inputs")
    iRow = 1
    Do While True
        sName = GetCellValue(oSheet, 1, iRow)  ' Column B
        If Len(sName) = 0 Then Exit Do

        If Not IsValidIdentifier(sName) Then
            Call AddMsg(sErrors, "Inputs: '" & sName & "' is not a valid IEC identifier")
        End If
        If InList(aIO, nIO, sName) Then
            Call AddMsg(sErrors, "Inputs: duplicate variable name '" & sName & "'")
        End If
        aIO(nIO) = sName
        nIO = nIO + 1

        iRow = iRow + 1
    Loop

    oSheet = oDoc.Sheets.getByName("Outputs")
    iRow = 1
    Do While True
        sName = GetCellValue(oSheet, 1, iRow)  ' Column B
        If Len(sName) = 0 Then Exit Do

        If Not IsValidIdentifier(sName) Then
            Call AddMsg(sErrors, "Outputs: '" & sName & "' is not a valid IEC identifier")
        End If
        If InList(aIO, nIO, sName) Then
            Call AddMsg(sErrors, "Outputs: duplicate variable name '" & sName & "' (also check Inputs)")
        End If
        aIO(nIO) = sName
        nIO = nIO + 1
        aOutputs(nOutputs) = sName
        nOutputs = nOutputs + 1

        iRow = iRow + 1
    Loop

    If Len(sStateVar) > 0 And InList(aIO, nIO, sStateVar) Then
        Call AddMsg(sErrors, "Config: State Variable Name '" & sStateVar & "' collides with an input/output name")
    End If
    If Len(sStateOut) > 0 And InList(aIO, nIO, sStateOut) Then
        Call AddMsg(sErrors, "Config: State Output Variable '" & sStateOut & "' collides with an input/output name")
    End If

    ' --- Transitions ---
    oTransSheet = oDoc.Sheets.getByName("Transitions")
    iRow = 1
    Do While True
        sFrom = GetCellValue(oTransSheet, 0, iRow)  ' Column A
        If Len(sFrom) = 0 Then Exit Do

        sTo = GetCellValue(oTransSheet, 1, iRow)        ' Column B
        sCondition = GetCellValue(oTransSheet, 2, iRow) ' Column C

        If Not InList(aStates, nStates, sFrom) Then
            Call AddMsg(sErrors, "Transitions row " & (iRow + 1) & ": From state '" & sFrom & "' is not defined in the States sheet")
        End If
        If Len(sTo) = 0 Then
            Call AddMsg(sErrors, "Transitions row " & (iRow + 1) & ": To state is empty")
        ElseIf Not InList(aStates, nStates, sTo) Then
            Call AddMsg(sErrors, "Transitions row " & (iRow + 1) & ": To state '" & sTo & "' is not defined in the States sheet")
        End If
        If Len(sCondition) = 0 Then
            Call AddMsg(sErrors, "Transitions row " & (iRow + 1) & " ('" & sFrom & "' -> '" & sTo & "'): condition is empty")
        End If

        iRow = iRow + 1
    Loop

    ' --- StateActions ---
    oSheet = oDoc.Sheets.getByName("StateActions")

    ' Header columns must be declared outputs
    iCol = 1
    Do While True
        sName = GetCellValue(oSheet, iCol, 0)  ' Header row
        If Len(sName) = 0 Then Exit Do

        If Not InList(aOutputs, nOutputs, sName) Then
            Call AddMsg(sErrors, "StateActions: header column '" & sName & "' is not defined in the Outputs sheet")
        End If

        iCol = iCol + 1
    Loop

    ' Declared outputs that never appear in the header are never assigned
    For i = 0 To nOutputs - 1
        If Not InActionsHeader(oSheet, aOutputs(i)) Then
            Call AddMsg(sWarnings, "StateActions: output '" & aOutputs(i) & "' has no column, so it is never assigned")
        End If
    Next i

    ' Every state should have a StateActions row
    For i = 0 To nStates - 1
        If FindRowByValue(oSheet, 0, aStates(i)) = -1 Then
            Call AddMsg(sWarnings, "StateActions: no row for state '" & aStates(i) & "': its outputs will hold their previous values")
        End If
    Next i

    ' Rows that do not match any state are ignored
    iRow = 1
    Do While True
        sName = GetCellValue(oSheet, 0, iRow)  ' Column A
        If Len(sName) = 0 Then Exit Do

        If Not InList(aStates, nStates, sName) Then
            Call AddMsg(sWarnings, "StateActions: row '" & sName & "' does not match any state (ignored)")
        End If

        iRow = iRow + 1
    Loop

    ' --- Unreachable states ---
    For i = 0 To nStates - 1
        If Not SameText(aStates(i), sInitState) Then
            bFound = False
            iRow = 1
            Do While True
                sName = GetCellValue(oTransSheet, 0, iRow)  ' Column A (row terminator)
                If Len(sName) = 0 Then Exit Do
                If SameText(GetCellValue(oTransSheet, 1, iRow), aStates(i)) Then
                    bFound = True
                    Exit Do
                End If
                iRow = iRow + 1
            Loop
            If Not bFound Then
                Call AddMsg(sWarnings, "State '" & aStates(i) & "' is unreachable: no transition leads to it")
            End If
        End If
    Next i
End Sub

' ================= Helpers =================

Function NL() As String
    ' Line ending for generated code (CRLF for Windows PLC IDE compatibility)
    NL = Chr(13) & Chr(10)
End Function

Sub AddMsg(ByRef sList As String, sMsg As String)
    ' Append a bullet line to an error/warning list
    If Len(sList) > 0 Then
        sList = sList & NL()
    End If
    sList = sList & "- " & sMsg
End Sub

Function SameText(sA As String, sB As String) As Boolean
    ' Case-insensitive comparison (ST identifiers are case-insensitive)
    SameText = (UCase(Trim(sA)) = UCase(Trim(sB)))
End Function

Function IsValidIdentifier(sName As String) As Boolean
    ' Check IEC 61131-3 identifier rules: letter or underscore first,
    ' then letters, digits, or underscores
    Const FIRST_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ_"
    Const OTHER_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ_0123456789"
    Dim i As Integer

    IsValidIdentifier = False
    If Len(sName) = 0 Then Exit Function
    If InStr(FIRST_CHARS, UCase(Mid(sName, 1, 1))) = 0 Then Exit Function
    For i = 2 To Len(sName)
        If InStr(OTHER_CHARS, UCase(Mid(sName, i, 1))) = 0 Then Exit Function
    Next i
    IsValidIdentifier = True
End Function

Function InList(aList() As String, nCount As Integer, sValue As String) As Boolean
    ' Case-insensitive membership test in the first nCount entries
    Dim i As Integer

    InList = False
    For i = 0 To nCount - 1
        If SameText(aList(i), sValue) Then
            InList = True
            Exit Function
        End If
    Next i
End Function

Function GetConfigValue(oDoc As Object, sLabel As String) As String
    ' Look up a Config value by its label in column A.
    ' Robust to rows being moved, inserted, or a header row being added.
    Dim oSheet As Object
    Dim iRow As Integer

    GetConfigValue = ""
    If Not oDoc.Sheets.hasByName("Config") Then Exit Function
    oSheet = oDoc.Sheets.getByName("Config")

    For iRow = 0 To 49
        If SameText(GetCellValue(oSheet, 0, iRow), sLabel) Then
            GetConfigValue = GetCellValue(oSheet, 1, iRow)
            Exit Function
        End If
    Next iRow
End Function

Function FindRowByValue(oSheet As Object, iCol As Integer, sValue As String) As Integer
    ' Find the row index of a value in a column, scanning from row 2 until
    ' the first empty cell. Returns -1 if not found.
    Dim iRow As Integer
    Dim sCell As String

    iRow = 1
    Do While True
        sCell = GetCellValue(oSheet, iCol, iRow)
        If Len(sCell) = 0 Then
            FindRowByValue = -1
            Exit Function
        End If
        If SameText(sCell, sValue) Then
            FindRowByValue = iRow
            Exit Function
        End If
        iRow = iRow + 1
    Loop
End Function

Function InActionsHeader(oSheet As Object, sName As String) As Boolean
    ' Check whether a name appears in the StateActions header row (from column B)
    Dim iCol As Integer
    Dim sCell As String

    InActionsHeader = False
    iCol = 1
    Do While True
        sCell = GetCellValue(oSheet, iCol, 0)
        If Len(sCell) = 0 Then Exit Function
        If SameText(sCell, sName) Then
            InActionsHeader = True
            Exit Function
        End If
        iCol = iCol + 1
    Loop
End Function

Function GetCellValue(oSheet As Object, iCol As Integer, iRow As Integer) As String
    ' Get cell value as trimmed string (0-indexed)
    Dim oCell As Object
    Dim sValue As String

    On Error Resume Next
    oCell = oSheet.getCellByPosition(iCol, iRow)
    sValue = Trim(oCell.getString())

    GetCellValue = sValue
End Function

' ================= File export =================

Sub WriteStringToUrl(sUrl As String, sText As String)
    ' Write a string to a file URL as UTF-8 (cross-platform UNO services)
    Dim oSimpleFileAccess As Object
    Dim oOutputStream As Object
    Dim oTextOutputStream As Object

    oSimpleFileAccess = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")

    ' Delete file if it exists
    If oSimpleFileAccess.exists(sUrl) Then
        oSimpleFileAccess.kill(sUrl)
    End If

    oOutputStream = oSimpleFileAccess.openFileWrite(sUrl)
    oTextOutputStream = CreateUnoService("com.sun.star.io.TextOutputStream")
    oTextOutputStream.setOutputStream(oOutputStream)
    oTextOutputStream.setEncoding("UTF-8")
    oTextOutputStream.writeString(sText)
    oTextOutputStream.closeOutput()
    oOutputStream.closeOutput()
End Sub

Sub ExportToFile(sCode As String, oDoc As Object)
    ' Export generated code to a file chosen by the user
    Dim sFileName As String
    Dim sFilePath As String
    Dim oFilePicker As Object

    sFileName = GetConfigValue(oDoc, "Program Name") & ".st"

    ' Create file picker dialog
    oFilePicker = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
    oFilePicker.Initialize(Array(com.sun.star.ui.dialogs.TemplateDescription.FILESAVE_SIMPLE))
    oFilePicker.setDefaultName(sFileName)
    oFilePicker.appendFilter("Structured Text", "*.st")
    oFilePicker.appendFilter("All Files", "*.*")
    oFilePicker.setCurrentFilter("Structured Text")

    If oFilePicker.Execute() = com.sun.star.ui.dialogs.ExecutableDialogResults.OK Then
        sFilePath = oFilePicker.Files(0)

        On Error GoTo ErrorHandler
        WriteStringToUrl(sFilePath, sCode)

        MsgBox "State machine code generated successfully!" & Chr(10) & Chr(10) & _
               "File: " & ConvertFromUrl(sFilePath), 64, "Success"
        Exit Sub

ErrorHandler:
        MsgBox "Error writing file: " & Error$ & Chr(10) & Chr(10) & _
               "Path: " & sFilePath, 16, "Error"
    End If
End Sub
