(* ======================================== *)
(* Wrapper program for FB_CrimpDeviceSeq *)
(* Runs the timers externally and maps physical I/O. *)
(* FB_CrimpDeviceSeq is generated from CrimpDevice_StateMachine.fods *)
(* ======================================== *)

PROGRAM PRG_CrimpDevice

VAR
    S1Start AT %IX0.0 : BOOL;   (* Start button *)
    S2Confirm AT %IX0.1 : BOOL; (* Unload confirmation button *)
    qxA1Crimp AT %QX0.0 : BOOL; (* Crimp actuator *)
    qxP1Unload AT %QX0.1 : BOOL; (* Unload actuator *)

    G_CrimpDevice : FB_CrimpDeviceSeq;
    tX2 : TON;  (* Crimp time *)
    tX3 : TON;  (* Pause time *)
END_VAR

(* Timers are driven by the step flag outputs of the state machine *)
tX2(IN := G_CrimpDevice.X2, PT := T#2S);
tX3(IN := G_CrimpDevice.X3, PT := T#5S);

G_CrimpDevice(
    S1Start := S1Start,
    S2Confirm := S2Confirm,
    T1CrimpDone := tX2.Q,
    T2PauseDone := tX3.Q);

(* Refresh the physical output image *)
qxA1Crimp := G_CrimpDevice.A1Crimp;
qxP1Unload := G_CrimpDevice.P1Unload;

END_PROGRAM
