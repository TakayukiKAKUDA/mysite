VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EditForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6730
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   16090
   OleObjectBlob   =   "EditForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "EditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CastTE()
    TextBox1.Text = Cells(3, 1).Value
    TextBox2.Text = Cells(3, 2).Value
    ComboBox1.Value = Cells(3, 3).Value
    TextBox3.Text = Cells(3, 4).Value
    ComboBox2.Value = Cells(3, 5).Value
    TextBox4.Text = Cells(3, 6).Value
End Sub

Private Sub CastTA()
    TextBox8.Text = Cells(3, 8).Value
    TextBox9.Text = Cells(3, 9).Value
    ComboBox5.Value = Cells(3, 10).Value
    TextBox10.Text = Cells(3, 11).Value
End Sub

Private Sub CastVE()
    Cells(3, 1).Value = TextBox1.Text
    Cells(3, 2).Value = TextBox2.Text
    Cells(3, 3).Value = ComboBox1.Value
    Cells(3, 4).Value = TextBox3.Text
    Cells(3, 5).Value = ComboBox2.Value
    Cells(3, 6).Value = TextBox4.Text
End Sub

Private Sub CastVA()
    Cells(3, 8).Value = TextBox8.Text
    Cells(3, 9).Value = TextBox9.Text
    Cells(3, 10).Value = ComboBox5.Value
    Cells(3, 11).Value = TextBox10.Text
End Sub

Sub CastListE()
    Dim i As Long, iMax As Long
    ListView1.ListItems.Clear
    iMax = MaxCount(1)
    If iMax <= 5 Then Exit Sub
    For i = 6 To iMax
        With ListView1.ListItems.Add
            .Text = Cells(i, 1).Value
            .SubItems(1) = Cells(i, 2).Value
            .SubItems(2) = Cells(i, 3).Value
            .SubItems(3) = Cells(i, 4).Value
        End With
    Next i
End Sub

Sub CastListA()
    Dim i As Long, iMax As Long
    ListView2.ListItems.Clear
    iMax = MaxCount(8)
    If iMax <= 5 Then Exit Sub
    For i = 6 To iMax
        With ListView2.ListItems.Add
            .Text = Cells(i, 8).Value
            .SubItems(1) = Cells(i, 10).Value
            .SubItems(2) = Cells(i, 11).Value
        End With
    Next i
End Sub

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    Call DBESelectSearch
    Call CastListE
End Sub

Private Sub CommandButton3_Click()
    TextBox5.Text = ""
    TextBox6.Text = ""
    ComboBox3.Value = ""
    TextBox7.Text = ""
    ComboBox4.Value = ""
End Sub

Private Sub CommandButton4_Click()
    Call DBEInsert
    Call CastTE
    Call DBESelectSearch
    Call CastListE
    TextBox2.SetFocus
End Sub

Private Sub CommandButton5_Click()
    If TextBox1.Text = "" Then Exit Sub
    Call CastVE
    Call DBEUpdate(Val(TextBox1.Text))
    If Judge = False Then Exit Sub
    Call CastTE
    Call DBESelectSearch
    Call CastListE
End Sub

Private Sub CommandButton6_Click()
    If TextBox1.Text = "" Then Exit Sub
    Call DBIInsert(Val(TextBox1.Text))
    Call DBISelect(Val(TextBox1.Text))
    Call CastListA
End Sub

Private Sub CommandButton7_Click()
    If TextBox8.Text = "" Then Exit Sub
    Call CastVA
    Call DBIUpdate(Val(TextBox8.Text))
    If Judge = False Then Exit Sub
    Call CastTA
    Call DBISelect(Val(TextBox1.Text))
    Call CastListA
End Sub

Private Sub ListView1_DblClick()
    On Error GoTo Err_Handler
    Call DBESelectSingle(Val(ListView1.SelectedItem))
    If Judge = False Then Exit Sub
    Call CastTE
    Call DBISelect(Val(TextBox1.Text))
    Call CastListA
    Exit Sub
Err_Handler:
    MsgBox Err.Description, vbCritical, "Error: #" & Err.Number
End Sub

Private Sub ListView2_DblClick()
    On Error GoTo Err_Handler
    Call DBISelectSingle(Val(ListView2.SelectedItem))
    If Judge = False Then Exit Sub
    Call CastTA
    Exit Sub
Err_Handler:
    MsgBox Err.Description, vbCritical, "Error: #" & Err.Number
End Sub

Private Sub UserForm_Initialize()
    Dim XCtrl As Variant
    Dim MC As Long, SC As Long, AC As Long, FC As Long
    
    MC = RGB(230, 230, 230)
    SC = RGB(197, 168, 128)
    AC = RGB(83, 46, 28)
    FC = RGB(15, 15, 15)
    
    Me.BackColor = MC
    Me.Caption = Me.Name
    
    For Each XCtrl In Controls
        With XCtrl
            .Font.Name = "メイリオ"
            .Font.Size = 10
            Select Case TypeName(XCtrl)
                Case "Frame"
                    .BackColor = MC
                    .ForeColor = FC
                    .SpecialEffect = 6
                Case "Label"
                    .BackColor = SC
                    .ForeColor = FC
                    .TextAlign = fmTextAlignCenter
                    .SpecialEffect = 6
                Case "TextBox"
                    .BackColor = MC
                    .ForeColor = FC
                    .TextAlign = fmTextAlignCenter
                    .SpecialEffect = 6
                Case "ComboBox"
                    .BackColor = MC
                    .ForeColor = FC
                    .TextAlign = fmTextAlignCenter
                    .SpecialEffect = 6
                Case "CommandButton"
                    .Font.Size = 12
                    .BackColor = AC
                    .ForeColor = SC
                Case Else
            End Select
        End With
    Next XCtrl
    
    With ListView1
        .Font.Name = "メイリオ"
        .Font.Size = 10
        .ForeColor = FC
        .BackColor = MC
        .View = lvwReport
        .LabelEdit = lvwManual
        .HideSelection = False
        .AllowColumnReorder = True
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Add 1, "S_ID", "ID", 54
        .ColumnHeaders.Add 2, "S_English", "英単語", 78
        .ColumnHeaders.Add 3, "S_POS", "品詞", 54
        .ColumnHeaders.Add 4, "S_Japanese", "日本語訳", 78
    End With
    
    With ListView2
        .Font.Name = "メイリオ"
        .Font.Size = 10
        .ForeColor = FC
        .BackColor = MC
        .View = lvwReport
        .LabelEdit = lvwManual
        .HideSelection = False
        .AllowColumnReorder = True
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Add 1, "S_AID", "AID"
        .ColumnHeaders.Add 2, "S_TYPE", "種類"
        .ColumnHeaders.Add 3, "S_INFO", "情報"
    End With
    
    With ComboBox1
        .AddItem ""
        .AddItem "名詞"
        .AddItem "代名詞"
        .AddItem "形容詞"
        .AddItem "副詞"
        .AddItem "動詞"
        .AddItem "前置詞"
        .AddItem "接続詞"
        .AddItem "間投詞"
    End With
    
    With ComboBox2
        .AddItem "TOEIC 金フレ CHAPTER01"
        .AddItem "TOEIC 金フレ CHAPTER02"
        .AddItem "TOEIC 金フレ CHAPTER03"
        .AddItem "TOEIC 金フレ CHAPTER04"
    End With
    
    ComboBox3.List = ComboBox1.List
    ComboBox4.List = ComboBox2.List
    
    With ComboBox5
        .AddItem "文章"
        .AddItem "後方配置"
        .AddItem "前方配置"
        .AddItem "類義語"
        .AddItem "対義語"
    End With
    
    Frame1.Caption = "LIST"
    Frame2.Caption = "編集欄"
    Frame3.Caption = "検索条件"
    Frame4.Caption = "追加情報"
    
    Label1.Caption = "ID"
    Label2.Caption = "英単語"
    Label3.Caption = "品詞"
    Label4.Caption = "日本語訳"
    Label5.Caption = "区間"
    Label6.Caption = "メモ"
    
    Label7.Caption = "ID"
    Label8.Caption = "英単語"
    Label9.Caption = "品詞"
    Label10.Caption = "日本語訳"
    Label11.Caption = "区間"
    
    Label12.Caption = "AID"
    Label13.Caption = "ID"
    Label14.Caption = "種類"
    Label15.Caption = "情報"
    
    CommandButton1.Caption = "CLOSE"
    CommandButton2.Caption = "SEARCH"
    CommandButton3.Caption = "CLEAR"
    
    CommandButton4.Caption = "NEW"
    CommandButton5.Caption = "SAVE"
    CommandButton6.Caption = "NEW"
    CommandButton7.Caption = "SAVE"
    
    TextBox1.Locked = True
    TextBox8.Locked = True
    TextBox9.Locked = True
    TextBox4.TextAlign = fmTextAlignLeft
    TextBox4.MultiLine = True
    TextBox10.TextAlign = fmTextAlignLeft
    TextBox10.MultiLine = True
    EditSheet.Activate
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    HomeSheet.Activate
End Sub
