Attribute VB_Name = "DBIModule"
Option Explicit

Sub CastR_I(RN As Long)
    Cells(RN, 8).Value = adoRs!AID
    Cells(RN, 9).Value = adoRs!ID
    Cells(RN, 10).Value = adoRs!種類
    Cells(RN, 11).Value = adoRs!情報
End Sub

Sub DBIInsert(X As Long)
    EditSheet.Activate
    Call DBConnect("E")
    On Error GoTo Err_Handler
    StrSQL = "SELECT * FROM 追加情報;"
    adoRs.Open StrSQL, adoCn, adOpenDynamic, adLockOptimistic
    adoCn.BeginTrans
    adoRs.AddNew
    adoRs.Fields("ID") = X
    adoRs.Update
    adoCn.CommitTrans
    Call CastR_I(2)
    Call CastR_I(3)
    Call DBCutOff
    Judge = True
    MsgBox "新規レコードを作成。", vbInformation
    Exit Sub
Err_Handler:
    adoCn.RollbackTrans
    Call DBCutOff
    Judge = False
    MsgBox Error$
End Sub

Sub DBIUpdate(X As Long)
    Dim XStrR As Variant, XStrW As Variant, XStrF As Variant
    Dim i As Long, iMax As Long, iStart As Long, iEnd As Long
    
    EditSheet.Activate
    iStart = 8
    iEnd = 11
    
    Call DBConnect("E")
    On Error GoTo Err_Handler
    adoCn.BeginTrans
    StrSQL = "SELECT * FROM 追加情報 WHERE AID = " & X & ";"
    adoRs.Open StrSQL, adoCn, adOpenKeyset, adLockOptimistic
    
    If adoRs.BOF = True And adoRs.EOF = True Then
        Call DBCutOff
        Judge = False
        MsgBox "対象データが見つかりません。", vbCritical
        Exit Sub
    End If
    
    Call CastR_I(2)
    
    XStrR = Range(Cells(2, iStart), Cells(2, iEnd))
    XStrW = Range(Cells(3, iStart), Cells(3, iEnd))
    XStrF = Range(Cells(4, iStart), Cells(4, iEnd))
    
    For i = LBound(XStrF, 2) To UBound(XStrF, 2)
        If IsEmpty(XStrW(1, i)) = False Then XStrW(1, i) = StrConv(XStrW(1, i), vbNarrow)
        If StrComp(XStrW(1, i), XStrR(1, i)) = 0 Then
            XStrF(1, i) = False
        Else
            XStrF(1, i) = True
        End If
    Next i
    
    Range(Cells(2, iStart), Cells(2, iEnd)) = XStrR
    Range(Cells(3, iStart), Cells(3, iEnd)) = XStrW
    Range(Cells(4, iStart), Cells(4, iEnd)) = XStrF
    
    If XStrF(1, 1) = True Then adoRs.Fields("AID") = Val(XStrW(1, 1))
    If XStrF(1, 2) = True Then adoRs.Fields("ID") = Val(XStrW(1, 2))
    If XStrF(1, 3) = True Then adoRs.Fields("種類") = CStr(XStrW(1, 3))
    If XStrF(1, 4) = True Then adoRs.Fields("情報") = CStr(XStrW(1, 4))
    adoRs.Update
    adoCn.CommitTrans
    
    Call DBCutOff
    Judge = True
    MsgBox "更新完了。", vbInformation
    Exit Sub
Err_Handler:
    adoCn.RollbackTrans
    Call DBCutOff
    Judge = False
    MsgBox Error$
End Sub

Sub DBISelectSingle(X As Long)
    Dim i As Long, iMax As Long, iStart As Long, iEnd As Long
    EditSheet.Activate
    iStart = 8
    iEnd = 11
    Range(Cells(2, iStart), Cells(3, iEnd)).ClearContents
    Call DBConnect("E")
    On Error GoTo Err_Handler
    StrSQL = "SELECT * FROM 追加情報 WHERE AID = " & X & ";"
    adoRs.Open StrSQL, adoCn, adOpenForwardOnly, adLockReadOnly
    
    If adoRs.BOF = True And adoRs.EOF = True Then
        Call DBCutOff
        Judge = False
        MsgBox "対象データが見つかりませんでした。", vbCritical, "Error"
        Exit Sub
    End If
    
    Call CastR_I(2)
    Call CastR_I(3)
    Call DBCutOff
    Judge = True
    Exit Sub
Err_Handler:
    Call DBCutOff
    Judge = False
    MsgBox Error$, vbCritical
    Debug.Print Error$
    Debug.Print StrSQL
End Sub

Sub DBISelect(X As Long)
    Dim i As Long, iMax As Long, iStart As Long, iEnd As Long
    EditSheet.Activate
    iStart = 8
    iEnd = 11
    iMax = MaxCount(iStart)
    If iMax >= 6 Then Range(Cells(6, iStart), Cells(iMax, iEnd)).ClearContents
    
    Call DBConnect("E")
    On Error GoTo Err_Handler
    StrSQL = "SELECT * FROM 追加情報 WHERE ID = " & X & " ORDER BY AID ASC;"
    adoRs.Open StrSQL, adoCn, adOpenForwardOnly, adLockReadOnly
    
    If adoRs.BOF = True And adoRs.EOF = True Then
        Call DBCutOff
        Judge = False
        Exit Sub
    End If
    
    i = 6
    Do Until adoRs.EOF
        Call CastR_I(i)
        i = i + 1
        adoRs.MoveNext
    Loop
    
    Call DBCutOff
    Judge = True
    Exit Sub
Err_Handler:
    Call DBCutOff
    Judge = False
    MsgBox Error$, vbCritical
    Debug.Print Error$
    Debug.Print StrSQL
End Sub
