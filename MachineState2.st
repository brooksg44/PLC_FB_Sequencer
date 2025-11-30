(* ======================================== *)
(* State Machine:  *)
(* Author: 1.0 *)
(* Version:  *)
(* Generated: 2025-11-29 16:41:58 *)
(* ======================================== *)

FUNCTION_BLOCK 

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
        (* No actions *)

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
        (* No actions *)

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
        (* No actions *)

        (* Transitions *)
        IF NOT IN_01 THEN
            currentState := IDLE;
        END_IF;
END_CASE;

END_FUNCTION_BLOCK
