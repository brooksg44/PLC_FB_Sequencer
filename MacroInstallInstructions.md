# State Machine Generator - Macro Installation and Usage

## Overview
These macros convert a state machine definition in `ST_StateMachine_Generator.ods` into IEC 61131-3 compliant Structured Text (ST) Function Block code.

## Installation Instructions

### Method 1: Import the Module (Recommended)

1. Open `ST_StateMachine_Generator.ods` in LibreOffice Calc
2. Go to **Tools → Macros → Edit Macros** (or press Alt+F11)
3. In the LibreOffice Basic IDE, select your document in the left tree (under "My Macros" or the document name)
4. Right-click on the document name and select **Insert → Module**
5. Delete the default content in the new module
6. Open `StateMachineMacros.bas` in a text editor, copy all content
7. Paste the content into the LibreOffice Basic module editor
8. Save (Ctrl+S or File → Save)
9. Close the Basic IDE

### Method 2: Direct Text Import

1. Open `ST_StateMachine_Generator.ods` in LibreOffice Calc
2. Go to **Tools → Macros → Organize Macros → LibreOffice Basic**
3. Select your document and click **Edit**
4. In the Basic IDE, go to **File → Import Basic...**
5. Browse and select `StateMachineMacros.bas`
6. Click **OK** to import

## Usage

### Running the Generator

Once the macros are installed, you can generate ST code:

1. **Configure your state machine** in the spreadsheet:
   - **Config** sheet: Set program name, author, version, state variable name, initial state
   - **Inputs** sheet: Define input variables (BOOL type assumed)
   - **Outputs** sheet: Define output variables (BOOL type assumed)
   - **States** sheet: Define all states in your state machine
   - **Transitions** sheet: Define state transitions with conditions
   - **StateActions** sheet: Define output values (0 or 1) for each state

2. **Generate the code**:
   - Go to **Tools → Macros → Run Macro**
   - Navigate to your document → StateMachineMacros (or the module name you used)
   - Select `GenerateStateMachine` from the list
   - Click **Run**
   
3. **Save the generated file**:
   - A file picker dialog will appear
   - Choose the location and name for your .st file
   - Click **Save**
   - You'll see a success message with the file path

### Preview Code (Optional)

To preview the generated code before saving:
- Run the `PreviewStateMachine` macro instead
- Note: Preview is limited to 2000 characters in the message box

## Generated Code Structure

The macro generates IEC 61131-3 compliant ST code with the following structure:

```
(* Header with metadata *)

FUNCTION_BLOCK <ProgramName>

VAR_INPUT
    <input variables as BOOL>
END_VAR

VAR_OUTPUT
    <output variables as BOOL>
END_VAR

VAR
    <stateVariable> : (STATE1, STATE2, ...) := <InitialState>;
END_VAR

(* State Machine Logic *)
CASE <stateVariable> OF
    STATE1:
        (* Actions *)
        OUT_01 := TRUE;
        OUT_02 := FALSE;
        
        (* Transitions *)
        IF <condition> THEN
            <stateVariable> := STATE2;
        END_IF;
        
    STATE2:
        ...
END_CASE;

END_FUNCTION_BLOCK
```

## Example Output

Based on the example data in your spreadsheet, the generated code would look like:

```
(* ======================================== *)
(* State Machine: MyStateMachine *)
(* Version: 1.0 *)
(* Generated: 2025-11-28 20:27:00 *)
(* ======================================== *)

FUNCTION_BLOCK MyStateMachine

VAR_INPUT
    IN_01 : BOOL;
    IN_02 : BOOL;
    IN_03 : BOOL;
    IN_04 : BOOL;
    IN_05 : BOOL;
    IN_06 : BOOL;
    IN_07 : BOOL;
    IN_08 : BOOL;
    IN_09 : BOOL;
    IN_10 : BOOL;
    IN_11 : BOOL;
    IN_12 : BOOL;
    IN_13 : BOOL;
    IN_14 : BOOL;
    IN_15 : BOOL;
    IN_16 : BOOL;
END_VAR

VAR_OUTPUT
    OUT_01 : BOOL;
    OUT_02 : BOOL;
    OUT_03 : BOOL;
    OUT_04 : BOOL;
    OUT_05 : BOOL;
    OUT_06 : BOOL;
    OUT_07 : BOOL;
    OUT_08 : BOOL;
    OUT_09 : BOOL;
    OUT_10 : BOOL;
    OUT_11 : BOOL;
    OUT_12 : BOOL;
    OUT_13 : BOOL;
    OUT_14 : BOOL;
    OUT_15 : BOOL;
    OUT_16 : BOOL;
END_VAR

VAR
    currentState : (IDLE, STATE_2, STATE_3) := IDLE;
END_VAR

(* State Machine Logic *)
CASE currentState OF

    IDLE:
        (* Actions *)
        OUT_01 := FALSE;
        OUT_02 := FALSE;
        OUT_03 := FALSE;
        OUT_04 := FALSE;
        OUT_05 := FALSE;
        OUT_06 := FALSE;
        OUT_07 := FALSE;
        OUT_08 := FALSE;
        OUT_09 := FALSE;
        OUT_10 := FALSE;
        OUT_11 := FALSE;
        OUT_12 := FALSE;
        OUT_13 := FALSE;
        OUT_14 := FALSE;
        OUT_15 := FALSE;
        OUT_16 := FALSE;

        (* Transitions *)
        IF IN_01 AND IN_02 THEN  (* Example transition *)
            currentState := STATE_2;
        END_IF;

    STATE_2:
        (* Actions *)
        OUT_01 := TRUE;
        OUT_02 := FALSE;
        OUT_03 := FALSE;
        OUT_04 := FALSE;
        OUT_05 := FALSE;
        OUT_06 := FALSE;
        OUT_07 := FALSE;
        OUT_08 := FALSE;
        OUT_09 := FALSE;
        OUT_10 := FALSE;
        OUT_11 := FALSE;
        OUT_12 := FALSE;
        OUT_13 := FALSE;
        OUT_14 := FALSE;
        OUT_15 := FALSE;
        OUT_16 := FALSE;

        (* Transitions *)
        IF IN_03 THEN
            currentState := STATE_3;
        END_IF;

    STATE_3:
        (* Actions *)
        OUT_01 := TRUE;
        OUT_02 := TRUE;
        OUT_03 := FALSE;
        OUT_04 := FALSE;
        OUT_05 := FALSE;
        OUT_06 := FALSE;
        OUT_07 := FALSE;
        OUT_08 := FALSE;
        OUT_09 := FALSE;
        OUT_10 := FALSE;
        OUT_11 := FALSE;
        OUT_12 := FALSE;
        OUT_13 := FALSE;
        OUT_14 := FALSE;
        OUT_15 := FALSE;
        OUT_16 := FALSE;

        (* Transitions *)
        IF NOT IN_01 THEN
            currentState := IDLE;
        END_IF;

END_CASE;

END_FUNCTION_BLOCK
```

## Features

✓ IEC 61131-3 compliant ST code generation
✓ Automatic VAR_INPUT and VAR_OUTPUT declarations
✓ Inline state enumeration type
✓ CASE-based state machine structure
✓ Condition-based state transitions with comments
✓ State-specific output actions
✓ Header comments with metadata and timestamp
✓ File export with .st extension

## Customization

You can modify the generated code format by editing these functions in `StateMachineMacros.bas`:

- `GenerateFunctionBlockHeader()` - Header format and comments
- `GenerateStateActions()` - Output assignment format
- `GenerateStateTransitions()` - Transition logic format
- `GenerateFunctionBlockFooter()` - Footer format

## Notes

- All inputs and outputs are assumed to be BOOL type (standard for PLC I/O)
- State transitions are evaluated in the order they appear in the Transitions sheet
- The first matching transition condition will be executed
- If you need additional variable types (INT, REAL, etc.), you'll need to extend the macro
- The generated code follows standard IEC 61131-3 ST syntax and should work with most PLC programming tools

## Troubleshooting

**Problem:** Macro doesn't appear in the list
- **Solution:** Make sure the macro is saved in the document (not "My Macros")
- Try closing and reopening the document

**Problem:** Error when running macro
- **Solution:** Check that all required sheets exist: Config, Inputs, Outputs, States, Transitions, StateActions
- Verify that the Config sheet has values in the expected cells

**Problem:** Generated code has errors
- **Solution:** Check transition conditions use valid ST syntax
- Verify state names don't contain spaces or special characters
- Make sure all state names in Transitions exist in States sheet

## Support

For issues or feature requests, modify the macro code as needed for your specific requirements.
