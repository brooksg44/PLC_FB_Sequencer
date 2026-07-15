# State Machine Generator for IEC 61131-3

Generate IEC 61131-3 compliant Structured Text (ST) Function Blocks from spreadsheet-based state machine definitions.

## Overview

This tool converts state machine specifications defined in a LibreOffice Calc spreadsheet into production-ready Structured Text code for PLCs. Define your states, transitions, and I/O actions in an intuitive spreadsheet format, then generate standards-compliant code with a single click. Two files are generated:

1. **`FB_<name>.st`** - the state machine Function Block
2. **`PRG_<name>.st`** - a wrapper PROGRAM that instantiates the Function Block, runs the external TON timers, and maps the physical I/O (`%IX`/`%QX` addresses)

## Features

- ✓ **IEC 61131-3 Compliant** - Generates standard ST Function Block code
- ✓ **Spreadsheet-Based Design** - Visual, tabular state machine definition
- ✓ **Validation Before Generation** - Cross-checks states, transitions, and actions; errors block generation with a clear list of problems
- ✓ **Complete I/O Management** - Automatic VAR_INPUT and VAR_OUTPUT declarations
- ✓ **State Enumeration** - Named `TYPE` declaration (portable) or inline enum
- ✓ **Conditional Transitions** - Support for complex boolean expressions
- ✓ **State Actions** - Define outputs for each state
- ✓ **Defensive CASE ELSE** - Unknown state values recover to the initial state
- ✓ **State Number Output** - Optional INT output for HMI/diagnostics
- ✓ **Preview & Export** - Scrollable preview with direct file export
- ✓ **Metadata Support** - Author, version, and timestamp in generated code
- ✓ **Wrapper PROGRAM Generation** - PRG_* code that calls the Function Block, with physical I/O mapping and external timers
- ✓ **Timers Sheet** - Declare TONs driven by step flags; their done flags feed back into the sequencer inputs
- ✓ **Physical Addressing** - Per-variable `%IX`/`%QX` addresses, or automatic sequential assignment

## Quick Start

1. Open `ST_StateMachine_Generator.fods` in LibreOffice Calc
2. Define your state machine in the spreadsheet tabs
3. Install the macros (see Installation below)
4. Run **Tools → Macros → Run Macro → PreviewStateMachine**
5. Review and save your generated ST code

## Files

- **ST_StateMachine_Generator.fods** - State machine definition spreadsheet (flat-XML OpenDocument, diff-friendly in git)
- **StateMachineMacros.bas** - LibreOffice Basic macro code
- **MacroInstallInstructions.md** - Detailed installation and usage guide
- **CrimpDevice_StateMachine.fods** - Complete worked example with timers and physical addresses; regenerates `FB_CrimpDeviceSeq.st` and the shipped `PRG_CrimpDevice.st`
- **examples/FBSeq.st** - Output generated from the sample data in the spreadsheet
- **README.md** - This file

## Installation

### Quick Install

1. Open `ST_StateMachine_Generator.fods`
2. Press **Alt+F11** to open the Basic IDE
3. Right-click on your document name → **Insert → Module**
4. Copy the contents of `StateMachineMacros.bas` and paste into the new module
5. Save (Ctrl+S) and close the IDE

See `MacroInstallInstructions.md` for detailed installation instructions and alternative methods.

## Usage

### 1. Define Your State Machine

The spreadsheet has six tabs for defining your state machine:

#### Config Tab

Values are looked up by the label in column A, so rows can be reordered or added to freely.

- **Program Name** - Function Block name (required)
- **Author** - Your name (optional)
- **Version** - Version number (optional)
- **State Variable Name** - Internal state variable (e.g., `currentState`) (required)
- **Initial State** - Starting state; must exist in the States tab (required)
- **State Enum Type Name** - If set (e.g., `E_FBSeq_State`), states are declared as a named `TYPE` before the Function Block, which most IDEs (CODESYS, TwinCAT) require. Leave blank to use an inline enumeration instead.
- **State Output Variable** - If set (e.g., `stateNum`), an INT output with this name reports the current state number (row order in the States tab) for HMI/diagnostics. Leave blank to omit.
- **PRG Name** - Wrapper PROGRAM name (optional). Defaults to `PRG_<base>`, where `<base>` is the Program Name without a leading `FB_` and trailing `Seq` (e.g., `FB_CrimpDeviceSeq` → `PRG_CrimpDevice`).
- **FB Instance Name** - Name of the FB instance variable in the wrapper PROGRAM (optional). Defaults to `G_<base>`.

#### Inputs Tab
- Define boolean input variables (IN_01, IN_02, etc.)
- Add descriptions for documentation
- **Physical Address** (column D, optional) - e.g., `%IX0.0`; the wrapper PROGRAM declares the input `AT` this address. Inputs listed as a timer's *Done Input* need no address (they are fed from the timer's `.Q`). If no address is given anywhere, addresses are auto-assigned sequentially.

#### Outputs Tab
- Define boolean output variables (OUT_01, OUT_02, etc.)
- Add descriptions for documentation
- **Physical Address** (column D, optional) - e.g., `%QX0.0`; the wrapper PROGRAM declares a `qx`-prefixed physical output `AT` this address and refreshes it from the FB output each cycle. Leave step-flag outputs (X1, X2, ...) without an address so they stay internal.

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
- Column headers name the outputs, so columns can be in any order (they must match the Outputs tab)
- Enter `1` for TRUE, `0` for FALSE
- Defines output values for each state

#### Timers Tab (optional)

One row per external TON run by the wrapper PROGRAM:

- **Timer Name** - TON instance name (e.g., `tX2`)
- **IN Signal** - What starts the timer; an FB output name (e.g., step flag `X2`) is qualified with the instance name (`G_CrimpDevice.X2`), an input name is used directly
- **Preset Time** - Time literal (e.g., `T#2S`)
- **Done Input** - The FB input fed from the timer's `.Q` output (e.g., `T1CrimpDone`)
- **Comment** - Optional, appears next to the TON declaration
- **IN Signal Address** - Optional; if the IN Signal is a physical sensor that is not an FB input/output (e.g., a level switch only the timer reads), give its `%IX` address here and the wrapper declares it

Delete the sheet or leave it empty if the state machine needs no external timers.

### 2. Generate Code

Both macros first validate the spreadsheet. Errors (undefined states, empty conditions, duplicate or invalid names, etc.) block generation and are listed in a dialog; warnings (e.g., a state with no StateActions row, unreachable states) ask for confirmation.

**PreviewStateMachine** (Recommended)
- Opens scrollable dialog previewing both generated files
- "Save to Files" button to export
- "Close" button to cancel

**GenerateStateMachine**
- Directly opens save dialog (default filename is the Program Name)
- Exports the Function Block to the chosen file and writes the wrapper `<PRG Name>.st` into the same folder

### 3. Use in Your PLC Project

Import both generated `.st` files into your PLC programming environment (CODESYS, TwinCAT, etc.) and add the PRG_* program to a cyclic task. The wrapper already instantiates the Function Block, runs the timers, and refreshes the physical outputs.

## Example

See `examples/FBSeq.st` for the complete output generated from the sample data in the spreadsheet. The structure:

```
(* Header with metadata and timestamp *)

TYPE E_FBSeq_State :
(
    IDLE,  (* Initial idle state *)
    STATE_2,
    STATE_3
);
END_TYPE

FUNCTION_BLOCK FBSeq

VAR_INPUT
    IN_01 : BOOL;
    ...
END_VAR

VAR_OUTPUT
    OUT_01 : BOOL;
    ...
    stateNum : INT;  (* Current state number (row order in States sheet) *)
END_VAR

VAR
    currentState : E_FBSeq_State := IDLE;
END_VAR

(* State Machine Logic *)
CASE currentState OF

    IDLE:
        (* Actions *)
        stateNum := 1;
        OUT_01 := FALSE;
        ...

        (* Transitions *)
        IF IN_01 AND IN_02 THEN  (* Example transition *)
            currentState := STATE_2;
        END_IF;

    ...

    ELSE
        (* Unknown state: recover to initial state *)
        currentState := IDLE;
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

## Limitations

- Currently supports BOOL types only for I/O (plus the optional INT state output)
- Single state variable per Function Block
- Transitions evaluated in order (first match wins)
- Transition conditions are passed through as-is; their ST syntax is not checked

## Troubleshooting

**Macro doesn't run**
- Ensure macros are enabled in LibreOffice (Tools → Options → Security → Macro Security)
- Verify macro is saved in the document, not "My Macros"

**Validation errors when generating**
- The error dialog lists each problem with its sheet and row; fix them in the spreadsheet and rerun
- State names must match (case-insensitively) between the States, Transitions, and StateActions sheets
- Every transition needs a non-empty condition

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

- **1.2** (2026-07-15)
  - Wrapper PROGRAM (PRG_*) generation alongside the Function Block: FB instantiation, external TON timers, and physical I/O mapping
  - New optional Timers sheet (Timer Name, IN Signal, Preset Time, Done Input, Comment, IN Signal Address)
  - New optional Physical Address column (D) on the Inputs and Outputs sheets, with automatic sequential assignment when omitted
  - New optional Config entries: PRG Name and FB Instance Name (sensible defaults derived from the Program Name)
  - Preview dialog shows both files; export writes both files in one step
  - Validation extended to timers, done inputs, and duplicate/misprefixed physical addresses
- **1.1** (2026-07-10)
  - Validation pass before generation (errors block, warnings confirm)
  - Config values looked up by label instead of fixed cell positions
  - Optional named `TYPE` enum declaration (Config: State Enum Type Name)
  - Optional INT state number output for HMI/diagnostics (Config: State Output Variable)
  - Defensive `ELSE` branch in the CASE statement (recovers to initial state)
  - StateActions columns matched by header name; 16-output limit removed
  - Fixed default export filename (was Author, now Program Name)
  - CRLF line endings in generated code
  - Spreadsheet renamed to `.fods` to match its actual flat-XML format
- **1.0** (2025-11-29) - Initial release
  - Basic state machine generation
  - BOOL I/O support
  - Scrollable preview dialog
  - File export functionality
