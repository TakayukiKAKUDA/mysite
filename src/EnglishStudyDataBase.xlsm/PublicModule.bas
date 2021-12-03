Attribute VB_Name = "PublicModule"
Option Explicit

Public Judge As Boolean
Public StrOut As String

Public Function MaxCount(N As Long) As Long
    MaxCount = Cells(ActiveSheet.Rows.Count, N).End(xlUp).Row
End Function

Public Function TerminalCount(N As Long) As Long
    TerminalCount = Cells(N, ActiveSheet.Columns.Count).End(xlToLeft).Column
End Function

Public Function StrCut(StrI, N As Long, D As String) As String
    Dim L As Long
    If TypeName(StrI) <> "String" Then Exit Function
    L = Len(StrI)
    If L < N Then Exit Function
    Select Case D
        Case "R"
            StrCut = Left(StrI, L - N)
        Case "L"
            StrCut = Right(StrI, L - N)
        Case Else
            Exit Function
    End Select
End Function

Sub SheetSetter()
    Cells.Font.Size = 10
    Cells.Font.Name = "ƒƒCƒŠƒI"
    Cells.HorizontalAlignment = xlCenter
    Columns.AutoFit
End Sub

Sub Boot_UF()
    EditForm.Show
End Sub

Sub SheetSetter2()
    Cells.Borders.Color = vbWhite
End Sub
