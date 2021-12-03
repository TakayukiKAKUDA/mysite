Attribute VB_Name = "EditModule"
Option Explicit

Sub DBInsertAll()
    Dim i As Long, iMax As Long
    
    EditSheet.Activate
    iMax = MaxCount(1)
    
    Call DBConnect("E")
    On Error GoTo Err_Handler
    StrSQL = "SELECT * FROM ‰p’PŒêDATABASE;"
    adoRs.Open StrSQL, adoCn, adOpenDynamic, adLockOptimistic
    
    adoCn.BeginTrans
    For i = 6 To iMax
        adoRs.AddNew
        adoRs.Fields("‰p’PŒê") = Cells(i, 2).Value
        adoRs.Fields("•iŽŒ") = Cells(i, 3).Value
        adoRs.Fields("“ú–{Œê–ó") = Cells(i, 4).Value
        adoRs.Fields("‹æŠÔ") = Cells(i, 5).Value
        adoRs.Update
    Next i
    adoCn.CommitTrans
    
    Call DBCutOff
    MsgBox "COMPLETE!", vbInformation
    Exit Sub
Err_Handler:
    adoCn.RollbackTrans
    Call DBCutOff
    MsgBox Error$, vbCritical
End Sub
