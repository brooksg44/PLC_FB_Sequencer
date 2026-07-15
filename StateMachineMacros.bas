REM ***** BASIC *****
REM State Machine Generator for IEC 61131-3 Structured Text
REM Converts the state machine definition in this spreadsheet into an ST Function
REM Block plus a wrapper PROGRAM that instantiates it, runs the external timers,
REM and maps physical I/O.
REM
REM Entry points:
REM   PreviewStateMachine  - validate, then preview both generated files in a dialog
REM   GenerateStateMachine - validate, then export both generated files (.st)
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
REM   PRG Name              - wrapper PROGRAM name (optional); defaults to PRG_<base>,
REM                           where <base> is the Program Name without a leading 'FB_'
REM                           and trailing 'Seq'
REM   FB Instance Name      - FB instance variable in the wrapper PROGRAM (optional);
REM                           defaults to G_<base>
REM
REM Wrapper PROGRAM data (all optional, backward compatible):
REM   Inputs/Outputs column D 'Physical Address' - e.g. %IX0.0 for inputs, %QX0.0 for
REM       outputs. Inputs with an address become physical vars (name AT %IX..); outputs
REM       with an address get a qx-prefixed physical var refreshed from the FB output.
REM       If no address is given anywhere, addresses are auto-assigned sequentially.
REM   Timers sheet - one row per external TON, columns:
REM       A: Timer Name  B: IN Signal  C: Preset Time  D: Done Input  E: Comment
REM       F: IN Signal Address (optional)
REM       The timer runs in the wrapper (tX(IN := <IN Signal>, PT := <Preset Time>))
REM       and its .Q output feeds the FB input named in 'Done Input'. An IN Signal
REM       naming an FB output (e.g. a step flag) is qualified with the instance name.
REM       If the IN Signal is a physical sensor that is not an FB input/output, give
REM       its address in column F and the wrapper declares it (name AT %IX.. : BOOL).

Option Explicit

' Globals used by the preview dialog's Save button listener
Global GlobalCodeToSave As String
Global GlobalPrgToSave As String
Global GlobalDoc As Object

' ================= Entry points =================

Sub GenerateStateMachine()
    ' Validate the spreadsheet, generate the FB and wrapper PRG, and export both
    Dim sFbCode As String
    Dim sPrgCode As String

    sFbCode = BuildValidatedCode(ThisComponent)
    If Len(sFbCode) = 0 Then Exit Sub

    sPrgCode = BuildProgramCode(ThisComponent)

    ExportBoth(sFbCode, sPrgCode, ThisComponent)
End Sub

Sub PreviewStateMachine()
    ' Validate the spreadsheet and preview both generated files in a scrollable dialog
    Dim sFbCode As String
    Dim sPrgCode As String
    Dim sPreview As String
    Dim oDialogModel As Object
    Dim oDialog As Object
    Dim oTextModel As Object
    Dim oButtonModel As Object
    Dim oSaveButtonModel As Object

    sFbCode = BuildValidatedCode(ThisComponent)
    If Len(sFbCode) = 0 Then Exit Sub

    sPrgCode = BuildProgramCode(ThisComponent)

    sPreview = "(* ========== File 1: " & GetConfigValue(ThisComponent, "Program Name") & ".st ========== *)" & NL() & NL()
    sPreview = sPreview & sFbCode & NL()
    sPreview = sPreview & "(* ========== File 2: " & GetPrgName(ThisComponent) & ".st ========== *)" & NL() & NL()
    sPreview = sPreview & sPrgCode

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
    oTextModel.Text = sPreview
    oTextModel.FontName = "Courier New"
    oTextModel.FontHeight = 10
    oDialogModel.insertByName("TextField", oTextModel)

    ' Create Save button
    oSaveButtonModel = oDialogModel.createInstance("com.sun.star.awt.UnoControlButtonModel")
    oSaveButtonModel.PositionX = 80
    oSaveButtonModel.PositionY = 225
    oSaveButtonModel.Width = 60
    oSaveButtonModel.Height = 15
    oSaveButtonModel.Label = "Save to Files"
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
    GlobalCodeToSave = sFbCode  ' Store code globally for save action
    GlobalPrgToSave = sPrgCode
    GlobalDoc = ThisComponent
    oSaveButton.addActionListener(oActionListener)

    oDialog.execute()
    oDialog.dispose()
End Sub

Sub SaveButton_actionPerformed(oEvent As Object)
    ' Save button action handler
    ExportBoth(GlobalCodeToSave, GlobalPrgToSave, GlobalDoc)
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

' ================= Wrapper PROGRAM generation =================

Function DeriveBaseName(oDoc As Object) As String
    ' Machine base name derived from the Program Name by stripping a leading
    ' 'FB_' and a trailing 'Seq' (e.g. FB_CrimpDeviceSeq -> CrimpDevice)
    Dim sName As String
    Dim sBase As String

    sName = GetConfigValue(oDoc, "Program Name")
    sBase = sName
    If UCase(Left(sBase, 3)) = "FB_" Then
        sBase = Mid(sBase, 4)
    End If
    If Len(sBase) > 3 And UCase(Right(sBase, 3)) = "SEQ" Then
        sBase = Left(sBase, Len(sBase) - 3)
    End If
    If Len(sBase) = 0 Then
        sBase = sName
    End If

    DeriveBaseName = sBase
End Function

Function GetPrgName(oDoc As Object) As String
    ' Wrapper PROGRAM name: Config 'PRG Name', or PRG_<base> when blank
    Dim sName As String

    sName = GetConfigValue(oDoc, "PRG Name")
    If Len(sName) = 0 Then
        sName = "PRG_" & DeriveBaseName(oDoc)
    End If

    GetPrgName = sName
End Function

Function GetInstanceName(oDoc As Object) As String
    ' FB instance variable name: Config 'FB Instance Name', or G_<base> when blank
    Dim sName As String

    sName = GetConfigValue(oDoc, "FB Instance Name")
    If Len(sName) = 0 Then
        sName = "G_" & DeriveBaseName(oDoc)
    End If

    GetInstanceName = sName
End Function

Function HasTimers(oDoc As Object) As Boolean
    ' Whether the optional Timers sheet exists and has at least one row
    Dim oSheet As Object

    HasTimers = False
    If Not oDoc.Sheets.hasByName("Timers") Then Exit Function
    oSheet = oDoc.Sheets.getByName("Timers")
    HasTimers = (Len(GetCellValue(oSheet, 0, 1)) > 0)
End Function

Function FindTimerByDoneInput(oDoc As Object, sInputName As String) As String
    ' Return the name of the timer whose 'Done Input' column matches an FB input,
    ' or "" when the input is not driven by a timer
    Dim oSheet As Object
    Dim iRow As Integer
    Dim sTimerName As String

    FindTimerByDoneInput = ""
    If Not oDoc.Sheets.hasByName("Timers") Then Exit Function
    oSheet = oDoc.Sheets.getByName("Timers")

    iRow = 1
    Do While True
        sTimerName = GetCellValue(oSheet, 0, iRow)  ' Column A
        If Len(sTimerName) = 0 Then Exit Function
        If SameText(GetCellValue(oSheet, 3, iRow), sInputName) Then  ' Column D
            FindTimerByDoneInput = sTimerName
            Exit Function
        End If
        iRow = iRow + 1
    Loop
End Function

Function IsTimerInSignal(oDoc As Object, sName As String) As Boolean
    ' Whether a variable is used as the IN signal of a timer (e.g. a step flag)
    Dim oSheet As Object
    Dim iRow As Integer

    IsTimerInSignal = False
    If Not oDoc.Sheets.hasByName("Timers") Then Exit Function
    oSheet = oDoc.Sheets.getByName("Timers")

    iRow = 1
    Do While True
        If Len(GetCellValue(oSheet, 0, iRow)) = 0 Then Exit Function
        If SameText(GetCellValue(oSheet, 1, iRow), sName) Then  ' Column B
            IsTimerInSignal = True
            Exit Function
        End If
        iRow = iRow + 1
    Loop
End Function

Function IsOutputName(oDoc As Object, sName As String) As Boolean
    ' Whether a name is declared in the Outputs sheet
    Dim oSheet As Object
    Dim iRow As Integer
    Dim sVarName As String

    IsOutputName = False
    oSheet = oDoc.Sheets.getByName("Outputs")

    iRow = 1
    Do While True
        sVarName = GetCellValue(oSheet, 1, iRow)  ' Column B
        If Len(sVarName) = 0 Then Exit Function
        If SameText(sVarName, sName) Then
            IsOutputName = True
            Exit Function
        End If
        iRow = iRow + 1
    Loop
End Function

Function SheetHasAnyAddress(oSheet As Object) As Boolean
    ' Whether any row of an I/O sheet has a Physical Address (column D)
    Dim iRow As Integer

    SheetHasAnyAddress = False
    iRow = 1
    Do While True
        If Len(GetCellValue(oSheet, 1, iRow)) = 0 Then Exit Function  ' Column B terminates
        If Len(GetCellValue(oSheet, 3, iRow)) > 0 Then  ' Column D
            SheetHasAnyAddress = True
            Exit Function
        End If
        iRow = iRow + 1
    Loop
End Function

Function UseAutoAddresses(oDoc As Object) As Boolean
    ' Auto-assign sequential addresses only when no Physical Address is given anywhere
    UseAutoAddresses = Not (SheetHasAnyAddress(oDoc.Sheets.getByName("Inputs")) _
        Or SheetHasAnyAddress(oDoc.Sheets.getByName("Outputs")))
End Function

Function AutoAddress(sPrefix As String, iIndex As Integer) As String
    ' Sequential bit address: index 0 -> <prefix>0.0, index 8 -> <prefix>1.0
    AutoAddress = sPrefix & (iIndex \ 8) & "." & (iIndex Mod 8)
End Function

Function BuildProgramCode(oDoc As Object) As String
    ' Assemble the wrapper PROGRAM that instantiates the Function Block,
    ' runs the external timers, and maps the physical I/O
    Dim sCode As String
    Dim sFbName As String
    Dim sDocName As String

    sFbName = GetConfigValue(oDoc, "Program Name")
    sDocName = ""
    On Error Resume Next
    sDocName = oDoc.getTitle()
    On Error Goto 0

    sCode = "(* ======================================== *)" & NL()
    sCode = sCode & "(* Wrapper program for " & sFbName & " *)" & NL()
    If HasTimers(oDoc) Then
        sCode = sCode & "(* Runs the timers externally and maps physical I/O. *)" & NL()
    Else
        sCode = sCode & "(* Maps physical I/O to the state machine. *)" & NL()
    End If
    If Len(sDocName) > 0 Then
        sCode = sCode & "(* " & sFbName & " is generated from " & sDocName & " *)" & NL()
    End If
    sCode = sCode & "(* Generated: " & Format(Now(), "YYYY-MM-DD HH:MM:SS") & " *)" & NL()
    sCode = sCode & "(* ======================================== *)" & NL()
    sCode = sCode & NL()
    sCode = sCode & "PROGRAM " & GetPrgName(oDoc) & NL()
    sCode = sCode & NL()
    sCode = sCode & GenerateProgramVars(oDoc)
    sCode = sCode & NL()
    sCode = sCode & GenerateTimerCalls(oDoc)
    sCode = sCode & GenerateFBCall(oDoc)
    sCode = sCode & GenerateOutputRefresh(oDoc)
    sCode = sCode & "END_PROGRAM" & NL()

    BuildProgramCode = sCode
End Function

Function GenerateProgramVars(oDoc As Object) As String
    ' Generate the wrapper VAR section: physical inputs, physical output image,
    ' the FB instance, and the external timers
    Dim oSheet As Object
    Dim sCode As String
    Dim sBlock As String
    Dim iRow As Integer
    Dim iAutoIdx As Integer
    Dim sName As String
    Dim sDescription As String
    Dim sAddr As String
    Dim sComment As String
    Dim bAuto As Boolean

    bAuto = UseAutoAddresses(oDoc)
    sCode = "VAR" & NL()

    ' Physical inputs (timer-driven inputs are fed from TON.Q instead)
    oSheet = oDoc.Sheets.getByName("Inputs")
    iAutoIdx = 0
    iRow = 1
    Do While True
        sName = GetCellValue(oSheet, 1, iRow)  ' Column B
        If Len(sName) = 0 Then Exit Do

        If Len(FindTimerByDoneInput(oDoc, sName)) = 0 Then
            sDescription = GetCellValue(oSheet, 2, iRow)  ' Column C
            sAddr = GetCellValue(oSheet, 3, iRow)         ' Column D
            If Len(sAddr) = 0 And bAuto Then
                sAddr = AutoAddress("%IX", iAutoIdx)
            End If

            If Len(sAddr) > 0 Then
                sCode = sCode & "    " & sName & " AT " & sAddr & " : BOOL;"
            Else
                sCode = sCode & "    " & sName & " : BOOL;  (* no physical address assigned *)"
            End If
            If Len(sDescription) > 0 Then
                sCode = sCode & "  (* " & sDescription & " *)"
            End If
            sCode = sCode & NL()
            iAutoIdx = iAutoIdx + 1
        End If

        iRow = iRow + 1
    Loop

    ' Physical vars for timer IN signals that are not FB inputs/outputs
    ' (Timers sheet column F 'IN Signal Address')
    If oDoc.Sheets.hasByName("Timers") Then
        oSheet = oDoc.Sheets.getByName("Timers")
        iRow = 1
        Do While True
            sName = GetCellValue(oSheet, 0, iRow)  ' Column A
            If Len(sName) = 0 Then Exit Do

            sAddr = GetCellValue(oSheet, 5, iRow)  ' Column F
            If Len(sAddr) > 0 Then
                sCode = sCode & "    " & GetCellValue(oSheet, 1, iRow) & " AT " & sAddr & " : BOOL;  (* IN signal for " & sName & " *)" & NL()
            End If

            iRow = iRow + 1
        Loop
    End If

    ' Physical output image (qx-prefixed vars for outputs with an address)
    oSheet = oDoc.Sheets.getByName("Outputs")
    sBlock = ""
    iAutoIdx = 0
    iRow = 1
    Do While True
        sName = GetCellValue(oSheet, 1, iRow)  ' Column B
        If Len(sName) = 0 Then Exit Do

        sDescription = GetCellValue(oSheet, 2, iRow)  ' Column C
        sAddr = GetCellValue(oSheet, 3, iRow)         ' Column D
        If Len(sAddr) = 0 And bAuto And Not IsTimerInSignal(oDoc, sName) Then
            sAddr = AutoAddress("%QX", iAutoIdx)
        End If

        If Len(sAddr) > 0 Then
            sBlock = sBlock & "    qx" & sName & " AT " & sAddr & " : BOOL;"
            If Len(sDescription) > 0 Then
                sBlock = sBlock & "  (* " & sDescription & " *)"
            End If
            sBlock = sBlock & NL()
            iAutoIdx = iAutoIdx + 1
        End If

        iRow = iRow + 1
    Loop
    If Len(sBlock) > 0 Then
        sCode = sCode & NL() & sBlock
    End If

    ' FB instance and timer instances
    sCode = sCode & NL()
    sCode = sCode & "    " & GetInstanceName(oDoc) & " : " & GetConfigValue(oDoc, "Program Name") & ";" & NL()

    If oDoc.Sheets.hasByName("Timers") Then
        oSheet = oDoc.Sheets.getByName("Timers")
        iRow = 1
        Do While True
            sName = GetCellValue(oSheet, 0, iRow)  ' Column A
            If Len(sName) = 0 Then Exit Do

            sComment = GetCellValue(oSheet, 4, iRow)  ' Column E
            sCode = sCode & "    " & sName & " : TON;"
            If Len(sComment) > 0 Then
                sCode = sCode & "  (* " & sComment & " *)"
            End If
            sCode = sCode & NL()

            iRow = iRow + 1
        Loop
    End If

    sCode = sCode & "END_VAR" & NL()

    GenerateProgramVars = sCode
End Function

Function GenerateTimerCalls(oDoc As Object) As String
    ' Generate the external timer calls; their done flags feed the FB inputs
    Dim oSheet As Object
    Dim sCode As String
    Dim iRow As Integer
    Dim sTimerName As String
    Dim sInSignal As String
    Dim sPreset As String

    If Not HasTimers(oDoc) Then
        GenerateTimerCalls = ""
        Exit Function
    End If

    oSheet = oDoc.Sheets.getByName("Timers")
    sCode = "(* Timers are driven by the step flag outputs of the state machine *)" & NL()

    iRow = 1
    Do While True
        sTimerName = GetCellValue(oSheet, 0, iRow)  ' Column A
        If Len(sTimerName) = 0 Then Exit Do

        sInSignal = GetCellValue(oSheet, 1, iRow)  ' Column B
        sPreset = GetCellValue(oSheet, 2, iRow)    ' Column C

        ' FB output names (step flags) are qualified with the instance name
        If IsOutputName(oDoc, sInSignal) Then
            sInSignal = GetInstanceName(oDoc) & "." & sInSignal
        End If

        sCode = sCode & sTimerName & "(IN := " & sInSignal & ", PT := " & sPreset & ");" & NL()

        iRow = iRow + 1
    Loop

    sCode = sCode & NL()

    GenerateTimerCalls = sCode
End Function

Function GenerateFBCall(oDoc As Object) As String
    ' Generate the Function Block call with all inputs mapped; inputs listed as a
    ' timer 'Done Input' are wired from the timer's .Q output
    Dim oSheet As Object
    Dim sCode As String
    Dim iRow As Integer
    Dim sName As String
    Dim sTimerName As String
    Dim bFirst As Boolean

    oSheet = oDoc.Sheets.getByName("Inputs")
    sCode = GetInstanceName(oDoc) & "("
    bFirst = True

    iRow = 1
    Do While True
        sName = GetCellValue(oSheet, 1, iRow)  ' Column B
        If Len(sName) = 0 Then Exit Do

        If bFirst Then
            sCode = sCode & NL()
            bFirst = False
        Else
            sCode = sCode & "," & NL()
        End If

        sTimerName = FindTimerByDoneInput(oDoc, sName)
        If Len(sTimerName) > 0 Then
            sCode = sCode & "    " & sName & " := " & sTimerName & ".Q"
        Else
            sCode = sCode & "    " & sName & " := " & sName
        End If

        iRow = iRow + 1
    Loop

    sCode = sCode & ");" & NL()
    sCode = sCode & NL()

    GenerateFBCall = sCode
End Function

Function GenerateOutputRefresh(oDoc As Object) As String
    ' Generate the physical output image refresh for the mapped outputs
    Dim oSheet As Object
    Dim sCode As String
    Dim sBody As String
    Dim iRow As Integer
    Dim sName As String
    Dim sAddr As String
    Dim bAuto As Boolean
    Dim bMapped As Boolean

    bAuto = UseAutoAddresses(oDoc)
    oSheet = oDoc.Sheets.getByName("Outputs")
    sBody = ""

    iRow = 1
    Do While True
        sName = GetCellValue(oSheet, 1, iRow)  ' Column B
        If Len(sName) = 0 Then Exit Do

        sAddr = GetCellValue(oSheet, 3, iRow)  ' Column D
        bMapped = (Len(sAddr) > 0)
        If Not bMapped And bAuto And Not IsTimerInSignal(oDoc, sName) Then
            bMapped = True
        End If

        If bMapped Then
            sBody = sBody & "qx" & sName & " := " & GetInstanceName(oDoc) & "." & sName & ";" & NL()
        End If

        iRow = iRow + 1
    Loop

    If Len(sBody) = 0 Then
        GenerateOutputRefresh = ""
        Exit Function
    End If

    sCode = "(* Refresh the physical output image *)" & NL()
    sCode = sCode & sBody
    sCode = sCode & NL()

    GenerateOutputRefresh = sCode
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
    Dim nInputs As Integer
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
    Dim sPrgName As String
    Dim sInstName As String

    sProgName = GetConfigValue(oDoc, "Program Name")
    sStateVar = GetConfigValue(oDoc, "State Variable Name")
    sInitState = GetConfigValue(oDoc, "Initial State")
    sEnumType = GetConfigValue(oDoc, "State Enum Type Name")
    sStateOut = GetConfigValue(oDoc, "State Output Variable")
    sPrgName = GetConfigValue(oDoc, "PRG Name")
    sInstName = GetConfigValue(oDoc, "FB Instance Name")

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

    If Len(sPrgName) > 0 And Not IsValidIdentifier(sPrgName) Then
        Call AddMsg(sErrors, "Config: PRG Name '" & sPrgName & "' is not a valid IEC identifier")
    End If

    If Len(sInstName) > 0 And Not IsValidIdentifier(sInstName) Then
        Call AddMsg(sErrors, "Config: FB Instance Name '" & sInstName & "' is not a valid IEC identifier")
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
    nInputs = nIO

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
    If InList(aIO, nIO, GetInstanceName(oDoc)) Then
        Call AddMsg(sErrors, "Config: FB Instance Name '" & GetInstanceName(oDoc) & "' collides with an input/output name")
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

    ' --- Timers (optional sheet, used by the wrapper PROGRAM) ---
    Dim aTimers(199) As String
    Dim nTimers As Integer
    Dim aDone(199) As String
    Dim nDone As Integer
    Dim sTimerName As String
    Dim sInSignal As String
    Dim sPreset As String
    Dim sDone As String
    Dim sInAddr As String
    Dim aInDecl(199) As String
    Dim nInDecl As Integer

    nTimers = 0
    nDone = 0
    nInDecl = 0
    If oDoc.Sheets.hasByName("Timers") Then
        oSheet = oDoc.Sheets.getByName("Timers")
        iRow = 1
        Do While True
            sTimerName = GetCellValue(oSheet, 0, iRow)  ' Column A
            If Len(sTimerName) = 0 Then Exit Do

            sInSignal = GetCellValue(oSheet, 1, iRow)  ' Column B
            sPreset = GetCellValue(oSheet, 2, iRow)    ' Column C
            sDone = GetCellValue(oSheet, 3, iRow)      ' Column D
            sInAddr = GetCellValue(oSheet, 5, iRow)    ' Column F

            If Not IsValidIdentifier(sTimerName) Then
                Call AddMsg(sErrors, "Timers row " & (iRow + 1) & ": Timer Name '" & sTimerName & "' is not a valid IEC identifier")
            End If
            If InList(aTimers, nTimers, sTimerName) Then
                Call AddMsg(sErrors, "Timers row " & (iRow + 1) & ": duplicate timer '" & sTimerName & "'")
            End If
            If InList(aIO, nIO, sTimerName) Then
                Call AddMsg(sErrors, "Timers row " & (iRow + 1) & ": Timer Name '" & sTimerName & "' collides with an input/output name")
            End If
            aTimers(nTimers) = sTimerName
            nTimers = nTimers + 1

            If Len(sPreset) = 0 Then
                Call AddMsg(sErrors, "Timers row " & (iRow + 1) & " ('" & sTimerName & "'): Preset Time is empty")
            ElseIf UCase(Left(sPreset, 2)) <> "T#" Then
                Call AddMsg(sWarnings, "Timers row " & (iRow + 1) & " ('" & sTimerName & "'): Preset Time '" & sPreset & "' does not look like a time literal (e.g. T#2S)")
            End If

            If Len(sDone) = 0 Then
                Call AddMsg(sErrors, "Timers row " & (iRow + 1) & " ('" & sTimerName & "'): Done Input is empty")
            Else
                If Not InList(aIO, nInputs, sDone) Then
                    Call AddMsg(sErrors, "Timers row " & (iRow + 1) & " ('" & sTimerName & "'): Done Input '" & sDone & "' is not defined in the Inputs sheet")
                End If
                If InList(aDone, nDone, sDone) Then
                    Call AddMsg(sErrors, "Timers row " & (iRow + 1) & ": input '" & sDone & "' is driven by more than one timer")
                End If
                aDone(nDone) = sDone
                nDone = nDone + 1
            End If

            If Len(sInSignal) = 0 Then
                Call AddMsg(sErrors, "Timers row " & (iRow + 1) & " ('" & sTimerName & "'): IN Signal is empty")
            ElseIf InList(aIO, nIO, sInSignal) Then
                If Len(sInAddr) > 0 Then
                    Call AddMsg(sErrors, "Timers row " & (iRow + 1) & " ('" & sTimerName & "'): IN Signal '" & sInSignal & "' is already a declared input/output; remove the IN Signal Address")
                End If
            ElseIf Len(sInAddr) > 0 Then
                If Not IsValidIdentifier(sInSignal) Then
                    Call AddMsg(sErrors, "Timers row " & (iRow + 1) & " ('" & sTimerName & "'): IN Signal Address requires a plain variable name as IN Signal, not an expression")
                ElseIf InList(aInDecl, nInDecl, sInSignal) Then
                    Call AddMsg(sErrors, "Timers row " & (iRow + 1) & " ('" & sTimerName & "'): IN Signal '" & sInSignal & "' is declared by more than one timer row")
                Else
                    aInDecl(nInDecl) = sInSignal
                    nInDecl = nInDecl + 1
                End If
            Else
                Call AddMsg(sWarnings, "Timers row " & (iRow + 1) & " ('" & sTimerName & "'): IN Signal '" & sInSignal & "' is not a declared input/output; it is passed through as-is")
            End If

            iRow = iRow + 1
        Loop
    End If

    ' --- Physical addresses (optional column D on Inputs/Outputs) ---
    Dim aAddr(999) As String
    Dim nAddr As Integer
    Dim sAddr As String

    nAddr = 0
    oSheet = oDoc.Sheets.getByName("Inputs")
    iRow = 1
    Do While True
        sName = GetCellValue(oSheet, 1, iRow)  ' Column B
        If Len(sName) = 0 Then Exit Do

        sAddr = GetCellValue(oSheet, 3, iRow)  ' Column D
        If Len(sAddr) > 0 Then
            If UCase(Left(sAddr, 2)) <> "%I" Then
                Call AddMsg(sWarnings, "Inputs row " & (iRow + 1) & " ('" & sName & "'): Physical Address '" & sAddr & "' does not start with %I")
            End If
            If InList(aAddr, nAddr, sAddr) Then
                Call AddMsg(sErrors, "Inputs row " & (iRow + 1) & " ('" & sName & "'): duplicate Physical Address '" & sAddr & "'")
            End If
            aAddr(nAddr) = sAddr
            nAddr = nAddr + 1
            If Len(FindTimerByDoneInput(oDoc, sName)) > 0 Then
                Call AddMsg(sWarnings, "Inputs row " & (iRow + 1) & " ('" & sName & "'): this input is driven by a timer, so its Physical Address is ignored")
            End If
        End If

        iRow = iRow + 1
    Loop

    oSheet = oDoc.Sheets.getByName("Outputs")
    iRow = 1
    Do While True
        sName = GetCellValue(oSheet, 1, iRow)  ' Column B
        If Len(sName) = 0 Then Exit Do

        sAddr = GetCellValue(oSheet, 3, iRow)  ' Column D
        If Len(sAddr) > 0 Then
            If UCase(Left(sAddr, 2)) <> "%Q" Then
                Call AddMsg(sWarnings, "Outputs row " & (iRow + 1) & " ('" & sName & "'): Physical Address '" & sAddr & "' does not start with %Q")
            End If
            If InList(aAddr, nAddr, sAddr) Then
                Call AddMsg(sErrors, "Outputs row " & (iRow + 1) & " ('" & sName & "'): duplicate Physical Address '" & sAddr & "'")
            End If
            aAddr(nAddr) = sAddr
            nAddr = nAddr + 1
        End If

        iRow = iRow + 1
    Loop

    If oDoc.Sheets.hasByName("Timers") Then
        oSheet = oDoc.Sheets.getByName("Timers")
        iRow = 1
        Do While True
            sName = GetCellValue(oSheet, 0, iRow)  ' Column A
            If Len(sName) = 0 Then Exit Do

            sAddr = GetCellValue(oSheet, 5, iRow)  ' Column F
            If Len(sAddr) > 0 Then
                If UCase(Left(sAddr, 2)) <> "%I" Then
                    Call AddMsg(sWarnings, "Timers row " & (iRow + 1) & " ('" & sName & "'): IN Signal Address '" & sAddr & "' does not start with %I")
                End If
                If InList(aAddr, nAddr, sAddr) Then
                    Call AddMsg(sErrors, "Timers row " & (iRow + 1) & " ('" & sName & "'): duplicate Physical Address '" & sAddr & "'")
                End If
                aAddr(nAddr) = sAddr
                nAddr = nAddr + 1
            End If

            iRow = iRow + 1
        Loop
    End If

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

Function ParentFolderUrl(sUrl As String) As String
    ' Folder part of a file URL (everything before the last '/')
    Dim i As Integer

    For i = Len(sUrl) To 1 Step -1
        If Mid(sUrl, i, 1) = "/" Then
            ParentFolderUrl = Left(sUrl, i - 1)
            Exit Function
        End If
    Next i
    ParentFolderUrl = sUrl
End Function

Sub ExportBoth(sFbCode As String, sPrgCode As String, oDoc As Object)
    ' Export the Function Block to a file chosen by the user, then write the
    ' wrapper PROGRAM as <PRG Name>.st into the same folder
    Dim sFileName As String
    Dim sFilePath As String
    Dim sPrgPath As String
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
        sPrgPath = ParentFolderUrl(sFilePath) & "/" & GetPrgName(oDoc) & ".st"

        On Error GoTo ErrorHandler
        WriteStringToUrl(sFilePath, sFbCode)
        WriteStringToUrl(sPrgPath, sPrgCode)

        MsgBox "State machine code generated successfully!" & Chr(10) & Chr(10) & _
               "Function Block: " & ConvertFromUrl(sFilePath) & Chr(10) & _
               "Wrapper PROGRAM: " & ConvertFromUrl(sPrgPath), 64, "Success"
        Exit Sub

ErrorHandler:
        MsgBox "Error writing file: " & Error$ & Chr(10) & Chr(10) & _
               "Paths: " & sFilePath & Chr(10) & sPrgPath, 16, "Error"
    End If
End Sub
