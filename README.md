# State Machine Generator for IEC 61131-3

Generate IEC 61131-3 compliant Structured Text (ST) Function Blocks from spreadsheet-based state machine definitions.

## Overview

This tool converts state machine specifications defined in a LibreOffice Calc spreadsheet into production-ready Structured Text code for PLCs. Define your states, transitions, and I/O actions in an intuitive spreadsheet format, then generate standards-compliant Function Block code with a single click.

## Features

- ✓ **IEC 61131-3 Compliant** - Generates standard ST Function Block code
- ✓ **Spreadsheet-Based Design** - Visual, tabular state machine definition
- ✓ **Complete I/O Management** - Automatic VAR_INPUT and VAR_OUTPUT declarations
- ✓ **State Enumeration** - Type-safe state definitions
- ✓ **Conditional Transitions** - Support for complex boolean expressions
- ✓ **State Actions** - Define outputs for each state
- ✓ **Preview & Export** - Scrollable preview with direct file export
- ✓ **Metadata Support** - Author, version, and timestamp in generated code

## Quick Start

1. Open `ST_StateMachine_Generator.ods` in LibreOffice Calc
2. Define your state machine in the spreadsheet tabs
3. Install the macros (see Installation below)
4. Run **Tools → Macros → Run Macro → PreviewStateMachine**
5. Review and save your generated ST code

## Files

- **ST_StateMachine_Generator.ods** - State machine definition spreadsheet
- **StateMachineMacros.bas** - LibreOffice Basic macro code
- **MacroInstallInstructions.md** - Detailed installation and usage guide
- **README.md** - This file

## Installation

### Quick Install

1. Open `ST_StateMachine_Generator.ods`
2. Press **Alt+F11** to open the Basic IDE
3. Right-click on your document name → **Insert → Module**
4. Copy the contents of `StateMachineMacros.bas` and paste into the new module
5. Save (Ctrl+S) and close the IDE

See `MacroInstallInstructions.md` for detailed installation instructions and alternative methods.

## Usage

### 1. Define Your State Machine

The spreadsheet has six tabs for defining your state machine:

#### Config Tab
- **Program Name** - Function Block name
- **Author** - Your name (optional)
- **Version** - Version number
- **State Variable Name** - Internal state variable (e.g., `currentState`)
- **Initial State** - Starting state

#### Inputs Tab
- Define boolean input variables (IN_01, IN_02, etc.)
- Add descriptions for documentation

#### Outputs Tab
- Define boolean output variables (OUT_01, OUT_02, etc.)
- Add descriptions for documentation

#### States Tab
- List all states in your state machine (IDLE, RUNNING, ERROR, etc.)
- Add descriptions for each state

#### Transitions Tab
- **From State** - Source state
- **To State** - Destination state
- **Condition** - Boolean expression (e.g., `IN_01 AND IN_02`)
- **Comment** - Optional transition description

#### StateActions Tab
- Matrix of states (rows) and outputs (columns)
- Enter `1` for TRUE, `0` for FALSE
- Defines output values for each state

### 2. Generate Code

Two macros are available:

**PreviewStateMachine** (Recommended)
- Opens scrollable dialog with full code preview
- "Save to File" button to export
- "Close" button to cancel

**GenerateStateMachine**
- Directly opens save dialog
- Immediately exports to .st file

### 3. Use in Your PLC Project

Import the generated `.st` file into your PLC programming environment (CODESYS, TwinCAT, etc.) and instantiate the Function Block in your program.

## Example

Based on the sample data in the spreadsheet:

```
FUNCTION_BLOCK MyStateMachine

VAR_INPUT
    IN_01 : BOOL;
    IN_02 : BOOL;
    IN_03 : BOOL;
END_VAR

VAR_OUTPUT
    OUT_01 : BOOL;
    OUT_02 : BOOL;
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

        (* Transitions *)
        IF IN_01 AND IN_02 THEN  (* Example transition *)
            currentState := STATE_2;
        END_IF;

    STATE_2:
        (* Actions *)
        OUT_01 := TRUE;
        OUT_02 := FALSE;

        (* Transitions *)
        IF IN_03 THEN
            currentState := STATE_3;
        END_IF;

    STATE_3:
        (* Actions *)
        OUT_01 := TRUE;
        OUT_02 := TRUE;

        (* Transitions *)
        IF NOT IN_01 THEN
            currentState := IDLE;
        END_IF;

END_CASE;

END_FUNCTION_BLOCK
```

## Customization

Edit `StateMachineMacros.bas` to customize:

- Code formatting and style
- Comment styles
- Variable naming conventions
- Additional variable types (INT, REAL, TIME, etc.)
- Entry/exit actions for states
- Default output behaviors

## Limitations

- Currently supports BOOL types only for I/O
- Maximum 16 inputs and 16 outputs (easily extended)
- Single state variable per Function Block
- Transitions evaluated in order (first match wins)

## Troubleshooting

**Macro doesn't run**
- Ensure macros are enabled in LibreOffice (Tools → Options → Security → Macro Security)
- Verify macro is saved in the document, not "My Macros"

**Wrong values in generated code**
- Check Config sheet has values in correct rows
- Verify state names match exactly between States and Transitions sheets
- Ensure StateActions has all states listed

**File save errors**
- Check write permissions for target directory
- Avoid special characters in filenames
- On macOS, select a standard directory (Documents, Desktop, etc.)

## Requirements

- LibreOffice Calc 6.0 or later
- macOS, Windows, or Linux
- PLC programming environment supporting IEC 61131-3 ST

## License

Free to use and modify for personal and commercial projects.

## Contributing

Suggestions and improvements welcome! Modify the macro code to suit your specific needs.

## Support

For detailed usage instructions, see `MacroInstallInstructions.md`.

## Version History

- **1.0** (2025-11-29) - Initial release
  - Basic state machine generation
  - BOOL I/O support
  - Scrollable preview dialog
  - File export functionality
