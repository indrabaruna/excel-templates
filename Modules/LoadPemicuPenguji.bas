Sub LoadPemicuPenguji()
'=================unlock pemicupenguji======================
Call LockEWS.unlockDetailPemicu
Call LockEWS.unlockDetailPenguji
Call LockEWS.unlockPemicu
Call LockEWS.unlockPenguji

'=================Progress_Bar======================
Call Progress_Bar.Progress(5)
Application.Wait (Now + TimeValue("0:00:01"))

Call Account_B1_loadPemicuAndPenguji
Call Account_B2_loadPemicuAndPenguji
Call Account_B3_loadPemicuAndPenguji
Call Account_B4_loadPemicuAndPenguji
Call Account_B5_loadPemicuAndPenguji

Call Progress_Bar.Progress(25)
Application.Wait (Now + TimeValue("0:00:01"))

Call Account_C1_loadPemicuAndPenguji
Call Account_C2_loadPemicuAndPenguji
Call Account_C3_loadPemicuAndPenguji
Call Account_C4_loadPemicuAndPenguji
Call Account_C5_loadPemicuAndPenguji
Call Account_C6_loadPemicuAndPenguji
Call Account_C7_loadPemicuAndPenguji
Call Account_C8_loadPemicuAndPenguji
Call Account_C9_loadPemicuAndPenguji
Call Account_C10_loadPemicuAndPenguji

Call Progress_Bar.Progress(50)
Application.Wait (Now + TimeValue("0:00:01"))

Call Account_C11_loadPemicuAndPenguji
Call Account_C12_loadPemicuAndPenguji
Call Account_C13_loadPemicuAndPenguji
Call Account_C14_loadPemicuAndPenguji
Call Account_D_loadPemicuAndPenguji
Call Account_E1_loadPemicuAndPenguji
Call Account_E2_loadPemicuAndPenguji
Call Account_E3_loadPemicuAndPenguji
Call Account_E4_loadPemicuAndPenguji
Call Account_E5_loadPemicuAndPenguji

Call Progress_Bar.Progress(75)
Application.Wait (Now + TimeValue("0:00:01"))

Call Account_E6_loadPemicuAndPenguji
Call Account_E7_loadPemicuAndPenguji
Call Account_E8_loadPemicuAndPenguji
Call Account_E9_loadPemicuAndPenguji
Call Account_E10_loadPemicuAndPenguji
Call Account_E11_loadPemicuAndPenguji
Call Account_E12_loadPemicuAndPenguji
Call Account_E13_loadPemicuAndPenguji
Call Account_F1_loadPemicuAndPenguji
Call Account_F2_loadPemicuAndPenguji
Call Account_F3_loadPemicuAndPenguji
Call Account_F4_loadPemicuAndPenguji
Call Account_G1_loadPemicuAndPenguji
Call Account_G2_loadPemicuAndPenguji

Call Progress_Bar.Progress(90)
Application.Wait (Now + TimeValue("0:00:01"))

Call Account_G3_loadPemicuAndPenguji
Call Account_G4_loadPemicuAndPenguji
Call Account_H1_loadPemicuAndPenguji
Call Account_I1_loadPemicuAndPenguji
Call Account_J1_loadPemicuAndPenguji
Call Account_K1_loadPemicuAndPenguji
Call Account_L1_loadPemicuAndPenguji
Call Account_M1_loadPemicuAndPenguji
Call Account_M2_loadPemicuAndPenguji
Call Account_M3_loadPemicuAndPenguji

Call Progress_Bar.Progress(99)
Application.Wait (Now + TimeValue("0:00:01"))

Call Account_P1_loadPemicuAndPenguji
Call Account_Q1_loadPemicuAndPenguji

Call ProgressBar.Progress(100, False)

' ================= LockPemicuPenguji ======================

Call LockEWS.lockDetailPemicu
Call LockEWS.lockDetailPenguji
Call LockEWS.lockPemicu
Call LockEWS.lockPenguji

Application.Wait (Now + TimeValue("0:00:5"))
Unload Progress_Bar


End Sub
