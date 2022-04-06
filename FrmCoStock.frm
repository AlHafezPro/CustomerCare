VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCoStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ⁄—Ì› „” Êœ⁄ ÃœÌœ"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   8640
   Begin VB.CheckBox ChkIsHall 
      Alignment       =   1  'Right Justify
      Caption         =   "„⁄„·"
      Height          =   345
      Left            =   120
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1110
      Width           =   1275
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   3570
      Top             =   2850
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   4875
      Left            =   60
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Width           =   8565
      _cx             =   15108
      _cy             =   8599
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.TextBox TxtStrName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3270
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1110
      Width           =   4365
   End
   Begin VB.TextBox TxtStrNo 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1110
      Width           =   945
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3180
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   38
      ImageHeight     =   38
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   38
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoStock.frx":6209F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   1217
      ButtonWidth     =   1191
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   33
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   38
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   37
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "≈”„ «·„” Êœ⁄"
      Height          =   195
      Index           =   1
      Left            =   6630
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   780
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "—ﬁ„ «·„” Êœ⁄"
      Height          =   195
      Index           =   0
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   810
      Width           =   915
   End
End
Attribute VB_Name = "FrmCoStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ColStrid = 1
Const ColStrNo = 2
Const ColStrName = 3
Const ColAccNo = 4
Const ColIsHall = 5

Dim Flag As Boolean

Sub SaveRec()
On Error GoTo ErrorHandler
Dim Id As Integer, RsMax As New ADODB.Recordset
With flexGrid
'If TxtStrNo.Tag <> "" Then
''Update
'    Sqltext = "Update NameStr Set StrName ='" & TxtStrName.Text & "',EmpNo=" & EmpNo & " Where Id=" & TxtStrNo.Tag
'    de.con.Execute (Sqltext)
'    Vrow = .Row
'Else
'Insert
    If Trim(TxtStrNo.Text) <> "" And Trim(TxtStrName.Text) <> "" Then
        sqlText = "Insert Into NameStr(StrNo , StrName, AccNo , isHall , EmpNo) Values(" & TxtStrNo.Text & ",'" & TxtStrName.Text & "',''," & ChkIsHall.Value & "," & empNo & ")"
        de.con.Execute (sqlText)
        sqlText = "Select Max(Id)MaxId From NameStr"
        Set RsMax = de.con.Execute(sqlText)
        TxtStrNo.Tag = RsMax!maxId
        .AddItem ""
        Vrow = .Rows - 1
    End If
'End If
    .TextMatrix(.Rows - 1, ColStrid) = TxtStrNo.Tag
    .TextMatrix(.Rows - 1, ColStrNo) = TxtStrNo.Text
    .TextMatrix(.Rows - 1, ColStrName) = TxtStrName.Text
    .TextMatrix(.Rows - 1, ColAccNo) = ""
    .TextMatrix(.Rows - 1, ColIsHall) = ChkIsHall.Value
    FillFormating
    .Col = ColStrid
    .Sort = flexSortNumericDescending
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Sub PrintData()
With cr1
    .Connect = ConnectName("")
    .ReportFileName = App.Path + "\Reports\RepCoStore.rpt"
    .DiscardSavedData = True
    .WindowState = crptMaximized
    .Action = 1
End With
End Sub
Sub NewRec()
    TxtStrNo.Tag = ""
    TxtStrNo.Text = ""
    TxtStrName.Text = ""
    ChkIsHall.Value = False
    TxtStrNo.SetFocus
End Sub
Function DeleteRow(Grid As VSFlexGrid, Vrow As Integer, Col As Integer, Table As String, Id As String) As Boolean
On Error GoTo ErrorHandler
With Grid
    sqlText = "Delete From " & Table & " Where " & Id & " = " & .TextMatrix(Vrow, Col)
    de.con.Execute (sqlText)
End With
DeleteRow = True
Exit Function
ErrorHandler:
DeleteRow = False
MsgBox (Err.Description)
End Function
Sub FillGrid()
    Dim rs As New ADODB.Recordset
    sqlText = "Select Id , StrNo , StrName , AccNo , isHall From NameStr Order By StrNo"
    Set rs = de.con.Execute(sqlText)
    Set flexGrid.DataSource = rs
End Sub
Sub init()
    Flag = False
    FillGrid
    Me.top = 0
    Me.left = 0
    FillFormating
    With flexGrid
        .Cols = 6
        .Editable = flexEDKbdMouse
    End With
    Flag = True
End Sub
Sub FillFormating()
    fs = "|>" + "Id"
    fs = fs + "|>" + "—ﬁ„ «·„” Êœ⁄"
    fs = fs + "|>" + "≈”„ «·„” Êœ⁄"
    fs = fs + "|>" + "—ﬁ„ «·Õ”«»"
    fs = fs + "|>" + ""
    With flexGrid
        .FormatString = fs
        .ColWidth(ColStrid) = 0
        SetColWidths ColStrNo, flexGrid
        SetColWidths ColStrName, flexGrid
        SetColWidths ColAccNo, flexGrid
        
        .ColDataType(ColIsHall) = flexDTBoolean
        .ColWidth(ColIsHall) = 300
    End With
End Sub
Sub SetColWidths(ByVal ColNo As Integer, flexGrid As VSFlexGrid)
    Dim i, J, s, w
    With flexGrid
            s = 0
            For i = 0 To .Rows - 1
                w = TextWidth(.TextMatrix(i, ColNo))
                If w > s Then s = w
            Next i
            .ColWidth(ColNo) = s + 300
    End With
End Sub


Private Sub ChkIsHall_Click()
If ChkIsHall.Value Then
    ChkIsHall.Caption = "’«·Â"
Else
    ChkIsHall.Caption = "„⁄„·"
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ActiveControl.name <> "FlexGrid" Then
        SendKeys "{Tab}"
        SendKeys "{Home}+{End}"
    End If
End If
End Sub

Private Sub Form_Load()
init
End Sub
Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With flexGrid
        sqlText = "Update NameStr Set StrName = '" & .TextMatrix(Row, ColStrName) & "',IsHall=" & Val(.TextMatrix(Row, ColIsHall)) & ",AccNo='" & .TextMatrix(Row, ColAccNo) & "', EmpNo=" & empNo & " Where Id=" & .TextMatrix(Row, ColStrid)
        de.con.Execute (sqlText)
End With
End Sub

Private Sub FlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
If Col = ColStrid Or Col = ColStrNo Then cancel = True
End Sub

Private Sub FlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim FirstRow As Integer, LastRow As Integer, Vrow As Integer
With flexGrid
If KeyCode = vbKeyDelete Then
    If MsgBox("Â· √‰  „ √ﬂœ „‰ ⁄„·Ì… «·Õ–›", vbYesNo + vbDefaultButton2, "Õ–› «·”Ã·«  «·„Õœœ…") = vbYes Then
        If .Row >= .RowSel Then
            FirstRow = .Row
            LastRow = .RowSel
        Else
            FirstRow = .RowSel
            LastRow = .Row
        End If
        For i = FirstRow To LastRow Step -1
            Vrow = i
            If DeleteRow(flexGrid, Vrow, ColStrid, "NameStr", "Id") Then
                .RemoveItem Vrow
            End If
        Next
    End If
End If
End With
End Sub

Private Sub FlexGrid_RowColChange()
If Flag Then
    With flexGrid
        TxtStrNo.Tag = .TextMatrix(.Row, ColStrid)
        TxtStrNo.Text = .TextMatrix(.Row, ColStrNo)
        TxtStrName.Text = .TextMatrix(.Row, ColStrName)
        ChkIsHall.Value = IIf(Val(.TextMatrix(.Row, ColIsHall)) = 0, 0, 1)
    End With
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        NewRec
    Case 3
        SaveRec
    Case 5
        PrintData
    Case 7
        Unload Me
End Select
End Sub

Private Sub TxtStrName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SaveRec
    End If
End Sub
