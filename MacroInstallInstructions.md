# State Machine Generator - Macro Installation and Usage

## Overview
These macros convert a state machine definition in `ST_StateMachine_Generator.fods` into IEC 61131-3 compliant Structured Text (ST) code. Two files are generated:

1. The state machine **Function Block** (e.g. `FB_CrimpDeviceSeq.st`)
2. A wrapper **PROGRAM** (e.g. `PRG_CrimpDevice.st`) that instantiates the Function Block, runs the external TON timers, and maps the physical I/O

## Installation Instructions

### Method 1: Import the Module (Recommended)

1. Open `ST_StateMachine_Generator.fods` in LibreOffice Calc
2. Go to **Tools → Macros → Edit Macros** (or press Alt+F11)
3. In the LibreOffice Basic IDE, select your document in the left tree (under "My Macros" or the document name)
4. Right-click on the document name and select **Insert → Module**
5. Delete the default content in the new module
6. Open `StateMachineMacros.bas` in a text editor, copy all content
7. Paste the content into the LibreOffice Basic module editor
8. Save (Ctrl+S or File → Save)
9. Close the Basic IDE

### Method 2: Direct Text Import

1. Open `ST_StateMachine_Generator.fods` in LibreOffice Calc
2. Go to **Tools → Macros → Organize Macros → LibreOffice Basic**
3. Select your document and click **Edit**
4. In the Basic IDE, go to **File → Import Basic...**
5. Browse and select `StateMachineMacros.bas`
6. Click **OK** to import

## Usage

### Running the Generator

Once the macros are installed, you can generate ST code:

1. **Configure your state machine** in the spreadsheet:
   - **Config** sheet: Set program name, author, version, state variable name, initial state. Values are found by the label in column A, so the rows can be reordered freely. Two optional settings:
     - **State Enum Type Name**: if set, the states are declared as a named `TYPE` before the Function Block (required by most IDEs such as CODESYS and TwinCAT); if blank, an inline enumeration is used
     - **State Output Variable**: if set, an INT output with this name reports the current state number for HMI/diagnostics; if blank, it is omitted
     - **PRG Name**: wrapper PROGRAM name; if blank, defaults to `PRG_<base>` where `<base>` is the Program Name without a leading `FB_` and trailing `Seq`
     - **FB Instance Name**: FB instance variable in the wrapper PROGRAM; if blank, defaults to `G_<base>`
   - **Inputs** sheet: Define input variables (BOOL type assumed). Optional column D **Physical Address** (e.g. `%IX0.0`) is used by the wrapper PROGRAM; inputs fed by a timer need no address. If no address is given anywhere, addresses are auto-assigned sequentially.
   - **Outputs** sheet: Define output variables (BOOL type assumed). Optional column D **Physical Address** (e.g. `%QX0.0`): outputs with an address get a `qx`-prefixed physical variable in the wrapper, refreshed from the FB output each cycle. Leave step-flag outputs unaddressed.
   - **States** sheet: Define all states in your state machine
   - **Transitions** sheet: Define state transitions with conditions
   - **StateActions** sheet: Define output values (0 or 1) for each state; the column headers name the outputs and must match the Outputs sheet
   - **Timers** sheet (optional): one row per external TON run by the wrapper PROGRAM, columns: **Timer Name** (e.g. `tX2`), **IN Signal** (an FB output such as step flag `X2` is qualified with the instance name; an input name is used directly), **Preset Time** (e.g. `T#2S`), **Done Input** (the FB input fed from the timer's `.Q`), **Comment**, and **IN Signal Address** (optional; if the IN Signal is a physical sensor that is not an FB input/output, give its `%IX` address and the wrapper declares it)

2. **Generate the code**:
   - Go to **Tools → Macros → Run Macro**
   - Navigate to your document → StateMachineMacros (or the module name you used)
   - Select `GenerateStateMachine` from the list
   - Click **Run**

3. **Fix any validation problems**:
   - The spreadsheet is validated before any code is generated
   - **Errors** (missing/duplicate/invalid state and variable names, transitions to undefined states, empty conditions, unknown StateActions columns, etc.) block generation and are listed with their sheet and row
   - **Warnings** (a state with no StateActions row, an output with no StateActions column, unreachable states, etc.) are listed and ask whether to continue

4. **Save the generated files**:
   - A file picker dialog will appear (default name is the Program Name)
   - Choose the location and name for the Function Block .st file
   - Click **Save**
   - The wrapper PROGRAM is written as `<PRG Name>.st` into the same folder (an existing file with that name is overwritten)
   - You'll see a success message with both file paths

### Preview Code (Optional)

To preview the generated code before saving:
- Run the `PreviewStateMachine` macro instead
- Both files are shown in a scrollable dialog with a **Save to Files** button

## Generated Code Structure

The macro generates IEC 61131-3 compliant ST code with the following structure:

```
(* Header with metadata *)

TYPE <StateEnumTypeName> :        (only when State Enum Type Name is set)
(
    STATE1,  (* description *)
    STATE2
);
END_TYPE

FUNCTION_BLOCK <ProgramName>

VAR_INPUT
    <input variables as BOOL>
END_VAR

VAR_OUTPUT
    <output variables as BOOL>
    <stateOutput> : INT;           (only when State Output Variable is set)
END_VAR

VAR
    <stateVariable> : <StateEnumTypeName or inline enum> := <InitialState>;
END_VAR

(* State Machine Logic *)
CASE <stateVariable> OF
    STATE1:
        (* Actions *)
        <stateOutput> := 1;
        OUT_01 := TRUE;
        OUT_02 := FALSE;

        (* Transitions *)
        IF <condition> THEN
            <stateVariable> := STATE2;
        END_IF;

    STATE2:
        ...

    ELSE
        (* Unknown state: recover to initial state *)
        <stateVariable> := <InitialState>;
END_CASE;

END_FUNCTION_BLOCK
```

The wrapper PROGRAM has this structure (see `PRG_CrimpDevice.st` for a real example):

```
(* Header with metadata *)

PROGRAM <PRGName>

VAR
    <input> AT %IX0.0 : BOOL;      (physical inputs; timer-done inputs omitted)
    qx<output> AT %QX0.0 : BOOL;   (physical outputs, for outputs with an address)

    <instance> : <ProgramName>;    (the Function Block instance)
    <timer> : TON;                 (one per Timers row)
END_VAR

(* Timers are driven by the step flag outputs of the state machine *)
<timer>(IN := <instance>.<step flag>, PT := T#2S);

<instance>(
    <input> := <input>,
    <doneInput> := <timer>.Q);

(* Refresh the physical output image *)
qx<output> := <instance>.<output>;

END_PROGRAM
```

## Example Output

See `examples/FBSeq.st` for the exact output generated from the sample data shipped in the spreadsheet. `CrimpDevice_StateMachine.fods` is a complete worked example with timers and physical addresses; it regenerates `FB_CrimpDeviceSeq.st` and the shipped `PRG_CrimpDevice.st`.

## Features

✓ IEC 61131-3 compliant ST code generation
✓ Validation of the whole definition before generation
✓ Automatic VAR_INPUT and VAR_OUTPUT declarations
✓ Named TYPE or inline state enumeration
✓ CASE-based state machine structure with defensive ELSE branch
✓ Condition-based state transitions with comments
✓ State-specific output actions (matched by StateActions column headers)
✓ Optional current-state INT output for HMI/diagnostics
✓ Wrapper PROGRAM generation with FB instantiation, external timers, and physical I/O mapping
✓ Optional per-variable physical addresses with automatic sequential fallback
✓ Header comments with metadata and timestamp
✓ File export with .st extension (CRLF line endings)

## Customization

You can modify the generated code format by editing these functions in `StateMachineMacros.bas`:

- `GenerateFunctionBlockHeader()` - Header format and comments
- `GenerateEnumTypeBlock()` - Named TYPE declaration format
- `GenerateStateActions()` - Output assignment format
- `GenerateStateTransitions()` - Transition logic format
- `BuildProgramCode()` / `GenerateProgramVars()` / `GenerateTimerCalls()` / `GenerateFBCall()` / `GenerateOutputRefresh()` - Wrapper PROGRAM format
- `ValidateStateMachine()` - Validation rules

## Notes

- All inputs and outputs are assumed to be BOOL type (standard for PLC I/O)
- State transitions are evaluated in the order they appear in the Transitions sheet
- The first matching transition condition will be executed
- State and variable name matching is case-insensitive (as in ST itself)
- Transition conditions are passed through as-is; their ST syntax is not checked
- If you need additional variable types (INT, REAL, etc.), you'll need to extend the macro
- The generated code follows standard IEC 61131-3 ST syntax and should work with most PLC programming tools

## Troubleshooting

**Problem:** Macro doesn't appear in the list
- **Solution:** Make sure the macro is saved in the document (not "My Macros")
- Try closing and reopening the document

**Problem:** Validation errors when running the macro
- **Solution:** Each message names the sheet and row of the problem; fix them in the spreadsheet and rerun
- Check that all required sheets exist: Config, Inputs, Outputs, States, Transitions, StateActions
- Make sure all state names in Transitions and StateActions exist in the States sheet

**Problem:** Generated code has errors in the PLC IDE
- **Solution:** Check transition conditions use valid ST syntax (conditions are not syntax-checked)
- If your IDE rejects the inline enumeration, set **State Enum Type Name** in the Config sheet to generate a named TYPE instead

## Support

For issues or feature requests, modify the macro code as needed for your specific requirements.
