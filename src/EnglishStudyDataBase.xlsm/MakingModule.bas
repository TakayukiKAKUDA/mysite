Attribute VB_Name = "MakingModule"
Option Explicit

Sub ChainCellValues()
    Dim InCB As Long, InCE As Long, OutC As Long
    Dim i As Long, iMax As Long, j As Long
    InCB = 1
    InCE = 3
    OutC = InCE + 1
    iMax = MaxCount(InCB)
    Range(Cells(2, OutC), Cells(iMax, OutC)).ClearContents
    For i = 2 To iMax
        For j = InCB To InCE
            Cells(i, OutC).Value = Cells(i, OutC).Value & Cells(i, j).Value
        Next j
    Next i
    Columns.AutoFit
End Sub
