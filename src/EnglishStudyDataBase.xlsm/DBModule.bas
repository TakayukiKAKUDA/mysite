Attribute VB_Name = "DBModule"
Option Explicit

Public adoCn As ADODB.Connection
Public adoRs As ADODB.Recordset
Public StrSQL As String

Sub DBConnect(PStr As String)
    Dim DBPath As String
    Select Case PStr
        Case "E"
            'DBPath = "C:\Users\spect\Documents\WorkSpace\英単語帳VT.accdb"
            DBPath = "C:\Users\spect\Documents\英単語帳.accdb"
        Case Else
            MsgBox "パス指定文字に不正があります。", vbCritical
            Exit Sub
    End Select
    Set adoCn = New ADODB.Connection
    Set adoRs = New ADODB.Recordset
    adoCn.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & DBPath & ";"
    adoCn.Open
End Sub

Sub DBCutOff()
    If Not adoRs Is Nothing Then adoRs.Close
    Set adoRs = Nothing
    If Not adoCn Is Nothing Then adoCn.Close
    Set adoCn = Nothing
End Sub
