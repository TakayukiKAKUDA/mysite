Attribute VB_Name = "DBEModule"
Option Explicit

Sub CastR_E(RN As Long)
    Cells(RN, 1).Value = adoRs!ID
    Cells(RN, 2).Value = adoRs!�p�P��
    Cells(RN, 3).Value = adoRs!�i��
    Cells(RN, 4).Value = adoRs!���{���
    Cells(RN, 5).Value = adoRs!���
    Cells(RN, 6).Value = adoRs!����
End Sub

Sub DBEInsert()
    EditSheet.Activate
    Call DBConnect("E")
    On Error GoTo Err_Handler
    adoCn.BeginTrans
    StrSQL = "SELECT * FROM �p�P��DATABASE;"
    adoRs.Open StrSQL, adoCn, adOpenDynamic, adLockOptimistic
    adoRs.AddNew
    adoRs.Update
    adoCn.CommitTrans
    Call CastR_E(2)
    Call CastR_E(3)
    Call DBCutOff
    Judge = True
    MsgBox "�V�K���R�[�h���쐬�B", vbInformation
    Exit Sub
Err_Handler:
    adoCn.RollbackTrans
    Call DBCutOff
    Judge = False
    MsgBox Error$
End Sub

Sub DBEUpdate(X As Long)
    Dim XStrR As Variant, XStrW As Variant, XStrF As Variant
    Dim i As Long, iMax As Long, iStart As Long, iEnd As Long
    
    EditSheet.Activate
    iStart = 1
    iEnd = 6
    
    Call DBConnect("E")
    On Error GoTo Err_Handler
    adoCn.BeginTrans
    StrSQL = "SELECT * FROM �p�P��DATABASE WHERE ID = " & X & ";"
    adoRs.Open StrSQL, adoCn, adOpenKeyset, adLockOptimistic
    
    If adoRs.BOF = True And adoRs.EOF = True Then
        Call DBCutOff
        Judge = False
        MsgBox "�Ώۃf�[�^��������܂���B", vbCritical
        Exit Sub
    End If
    
    Call CastR_E(2)
    
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
    
    If XStrF(1, 1) = True Then adoRs.Fields("ID") = Val(XStrW(1, 1))
    If XStrF(1, 2) = True Then adoRs.Fields("�p�P��") = CStr(XStrW(1, 2))
    If XStrF(1, 3) = True Then adoRs.Fields("�i��") = CStr(XStrW(1, 3))
    If XStrF(1, 4) = True Then adoRs.Fields("���{���") = CStr(XStrW(1, 4))
    If XStrF(1, 5) = True Then adoRs.Fields("���") = CStr(XStrW(1, 5))
    If XStrF(1, 6) = True Then adoRs.Fields("����") = CStr(XStrW(1, 6))
    adoRs.Update
    adoCn.CommitTrans
    
    Call DBCutOff
    Judge = True
    MsgBox "�X�V�����B", vbInformation
    Exit Sub
Err_Handler:
    adoCn.RollbackTrans
    Call DBCutOff
    Judge = False
    MsgBox Error$
End Sub

Sub DBESelectSingle(X As Long)
    Dim i As Long, iMax As Long, iStart As Long, iEnd As Long
    EditSheet.Activate
    iStart = 1
    iEnd = 6
    Range(Cells(2, iStart), Cells(3, iEnd)).ClearContents
    Call DBConnect("E")
    On Error GoTo Err_Handler
    StrSQL = "SELECT * FROM �p�P��DATABASE WHERE ID = " & X & ";"
    adoRs.Open StrSQL, adoCn, adOpenForwardOnly, adLockReadOnly
    
    If adoRs.BOF = True And adoRs.EOF = True Then
        Call DBCutOff
        Judge = False
        MsgBox "�Ώۃf�[�^��������܂���ł����B", vbCritical, "Error"
        Exit Sub
    End If
    
    Call CastR_E(2)
    Call CastR_E(3)
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

Sub DBESelectSearch()
    Dim i As Long, iMax As Long, iStart As Long, iEnd As Long
    Dim Term(4) As String, TermX As String
    EditSheet.Activate
    iStart = 1
    iEnd = 6
    
    iMax = MaxCount(1)
    If iMax >= 6 Then Range(Cells(6, iStart), Cells(iMax, iEnd)).ClearContents
    
    For i = LBound(Term) To UBound(Term)
        Term(i) = ""
    Next i
    
    With EditForm
        If .TextBox5.Text <> "" Then Term(0) = "ID = " & Val(.TextBox5.Text) & " "
        If .TextBox6.Text <> "" Then Term(1) = "�p�P�� LIKE '" & CStr(.TextBox6.Text) & "' "
        If .ComboBox3.Value <> "" Then Term(2) = "�i�� LIKE '" & CStr(.ComboBox3.Value) & "' "
        If .TextBox7.Text <> "" Then Term(3) = "���{��� LIKE '" & CStr(.TextBox7.Text) & "' "
        If .ComboBox4.Value <> "" Then Term(4) = "��� LIKE '" & CStr(.ComboBox4.Value) & "' "
    End With
    
    TermX = "WHERE"
    For i = LBound(Term) To UBound(Term)
        If Term(i) <> "" Then TermX = TermX & " " & Term(i) & "AND"
    Next i
    TermX = StrCut(TermX, 3, "R")
    If StrComp(TermX, "WH") = 0 Then TermX = ""
    
    On Error GoTo Err_Handler
    Call DBConnect("E")
    StrSQL = "SELECT * FROM �p�P��DATABASE " & TermX & "ORDER BY ID DESC;"
    adoRs.Open StrSQL, adoCn, adOpenKeyset, adLockReadOnly
    
    If adoRs.BOF = True And adoRs.EOF = True Then
        Call DBCutOff
        Judge = False
        MsgBox "�Ώۃf�[�^��������܂���ł����B", vbCritical, "ERROR"
        Exit Sub
    End If
    
    i = 6
    adoRs.MoveFirst
    Do Until adoRs.EOF
        Call CastR_E(i)
        i = i + 1
        adoRs.MoveNext
    Loop
    
    Call DBCutOff
    Judge = True
    Exit Sub
Err_Handler:
    Call DBCutOff
    Judge = False
    MsgBox Error$
    Debug.Print Error$
    Debug.Print StrSQL
End Sub
