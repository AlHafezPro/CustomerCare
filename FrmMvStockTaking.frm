VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMvStockTaking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ã—œ…  «·„Ê«œ "
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   11910
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5325
      Left            =   60
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   5775
      _cx             =   10186
      _cy             =   9393
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
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
      ScrollTrack     =   0   'False
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
      ExplorerBar     =   0
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
   Begin VB.TextBox TxtStkNo 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4140
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   1725
   End
   Begin MSDataListLib.DataCombo ComboStr 
      Height          =   315
      Left            =   6840
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   690
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   90
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
            Picture         =   "FrmMvStockTaking.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvStockTaking.frx":6209F
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
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   1217
      ButtonWidth     =   1191
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   15
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   5385
      Left            =   60
      TabIndex        =   2
      Top             =   1080
      Width           =   11805
      _cx             =   20823
      _cy             =   9499
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ „Œ“‰Ì"
      Height          =   285
      Left            =   5940
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.Label LStkName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   750
      Width           =   4005
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·„” Êœ⁄"
      Height          =   285
      Left            =   11250
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   630
   End
End
Attribute VB_Name = "FrmMvStockTaking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ColStrid = 1
Const ColStrNo = 2
Const ColStrName = 3
Const ColStkId = 4
Const ColStkNo = 5
Const ColStkName = 6
Const ColStkIn = 7
Const ColStkOut = 8
Const ColBalance = 9

Dim Flag As Boolean
Const ColStkId_1 = 1
Const ColStkno_1 = 2
Const ColStkname_1 = 3


Dim Stmov As MvStockType
Dim PrevBalance As Double
Sub ReturnOldData(ByVal vRow As Integer)
With flexGrid
    .TextMatrix(vRow, ColBalance) = PrevBalance
End With
End Sub

Function SelectRow(Stkid As Double) As Boolean
With flexGrid
If Stkid = 0 Then
    SelectRow = True
    Exit Function
Else
    For i = 1 To .Rows - 1
        If .TextMatrix(i, ColStkId) = Stkid Then
            .Row = i
            .TopRow = i
          '  .Row = i
            SelectRow = True
            Exit Function
        End If
    Next
End If
SelectRow = False
End With
End Function
Function MaxRec() As Double
    Dim RsMax As New ADODB.Recordset
    sqltext = "Select isnull(Max(ByanId),0) as MaxByanId From Stmov"
    Set RsMax = de.con.Execute(sqltext)
    If RsMax!MaxByanId = 0 Then
        MaxRec = 1
    Else
        MaxRec = RsMax!MaxByanId + 1
    End If
End Function
Sub FillRec(ByVal CurrentQty As Double, ByVal PrevQty As Double, ByVal Strid As Integer, ByVal Stkid As Double)
With Stmov
    .ByanId = MaxRec()
    .MovDate = Format(Date, "mm/dd/yyyy")
    .Qty = IIf(PrevQty > CurrentQty, PrevQty - CurrentQty, CurrentQty - PrevQty)
    .QtyType = IIf(PrevQty > CurrentQty, 1, 0)
    .Stkid = Stkid
    .Strid = Strid
    .DocType = 2 'Ã—œ
End With
End Sub
Sub FillActiveControl(List As VSFlexGrid)
    With List
        If ActiveControl.Text <> "" Then
            If Not ActiveControl.DataChanged Then Exit Sub
            Flag = False
            ActiveControl.Text = IIf(.Visible = False, "", .TextMatrix(.Row, ColStkno_1))
            LStkName.Caption = IIf(.Visible = False, "", .TextMatrix(.Row, ColStkname_1))
            Flag = True
            ActiveControl.Tag = IIf(.Visible = False, "", .TextMatrix(.Row, ColStkId_1))
        Else
            ActiveControl.Text = ""
            ActiveControl.Tag = ""
            LStkName.Caption = ""
        End If
        List.Visible = False
        ActiveControl.DataChanged = False
    End With
End Sub

Sub FillFormatVSFlex(flexGrid As VSFlexGrid)

    Fs = "|ID"
    Fs = Fs + "|<" + "«·—ﬁ„ «·„Œ“‰Ì"
    Fs = Fs + "|<" + "«·≈”„"
    With flexGrid
        .Visible = False
        .FormatString = Fs
            .ColWidth(ColStkId_1) = 0
            SetColWidths ColStkno_1, flexGrid
            SetColWidths ColStkname_1, flexGrid
            .Visible = True
    End With

End Sub


Sub MoveCursor(KeyCode As Integer)
On Error Resume Next
With Grid
    If KeyCode = vbKeyDown Then
        .Row = .Row + 1
    ElseIf KeyCode = vbKeyUp Then
        .Row = .Row - 1
    End If
If Not .RowIsVisible(.Row) Then
    .TopRow = .Row
End If
.Col = 0
.ColSel = .Cols - 1
End With
End Sub

Sub FillList(sqltext As String, Field1 As String, Field2 As String, List As VSFlexGrid)
    Set Rs = de.con.Execute(sqltext)
    If Rs.RecordCount > 0 Then
        Set List.DataSource = Rs
        FillFormatVSFlex List
        List.Row = 1
        List.Col = 1
        List.ColSel = List.Cols - 1
        List.Visible = True
        TxtStkNo.SetFocus
    Else
        List.Text = ""
        List.Visible = False
        TxtStkNo.SetFocus
    End If
End Sub
Function UpdateStmov(CurrentQty As Double, PrevQty As Double, Strid As Integer, Stkid As Double) As Boolean
On Error GoTo errorhandler
With Stmov
    FillRec CurrentQty, PrevQty, Strid, Stkid
    sqltext = "Insert into Stmov(ByanId , StkId  ,  StrId , Movdate , DocType , Qty , QtyType,EmpNo)Values("
    sqltext = sqltext & .ByanId & "," & .Stkid & "," & .Strid & ",'" & .MovDate & "'," & .DocType & "," & .Qty & "," & .QtyType & "," & EmpNo & ")"
    de.con.Execute (sqltext)
End With
UpdateStmov = True
Exit Function
errorhandler:
UpdateStmov = False
End Function
Sub SearchRec(Strid As Integer)
Dim RsBalance As New ADODB.Recordset
    sqltext = "Select StrId , StrNo , StrName  , StkId  , StkNo , StkName , 0 ,0 , Qty   From BalanceQry Where Strid=" & Strid
    Set RsBalance = de.con.Execute(sqltext)
    Set flexGrid.DataSource = RsBalance
    FillFormatString
End Sub

Sub Init()
    Dim RsSTr As New ADODB.Recordset
    sqltext = "Select Id , StrNo , StrName From NameStr Order by StrNo"
    Set RsSTr = de.con.Execute(sqltext)
    Set ComboStr.RowSource = RsSTr
    ComboStr.ListField = "StrName"
    ComboStr.BoundColumn = "Id"
    If RsSTr.RecordCount > 0 Then
        RsSTr.MoveFirst
        RsSTr.MoveFirst
        ComboStr.BoundText = RsSTr!ID
    End If
    FillFormatString
    Top = 0
    Left = 0
    flexGrid.Editable = flexEDKbdMouse
    flexGrid.Rows = 1
End Sub
Sub PrintData()

End Sub

Sub FillFormatString()
    Fs = "|>" + "ÚStrId"
    Fs = Fs + "|>" + "«·„” Êœ⁄"
    Fs = Fs + "|>" + "«·≈”„"
    Fs = Fs + "|>" + "StkId"
    Fs = Fs + "|>" + "«·—ﬁ„ «·„Œ“‰Ì"
    Fs = Fs + "|>" + "«·≈”„"
    Fs = Fs + "|>" + "«·œ«Œ·"
    Fs = Fs + "|>" + "«·Œ«—Ã"
    Fs = Fs + "|>" + "«·—’Ìœ"
    With flexGrid
        .Cols = 11
        .FormatString = Fs
        .ColWidth(ColStkId) = 0
        .ColWidth(ColStrid) = 0
        .ColWidth(ColStrNo) = 0
        .ColWidth(ColStrName) = 0
        SetColWidths ColStkIn, flexGrid
        SetColWidths ColStkOut, flexGrid
        SetColWidths ColStkNo, flexGrid
        SetColWidths ColStkName, flexGrid
        SetColWidths ColBalance, flexGrid
   End With
End Sub

Sub SetColWidths(ColNo As Integer, flexGrid As VSFlexGrid)

    With flexGrid
         .AutoSize (ColNo)
    End With
End Sub



Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo errorhandler
With flexGrid
    If .TextMatrix(Row, ColBalance) < 0 Then
        MsgBox "«·—’Ìœ ·«Ì”„Õ ", vbExclamation, " ‰»ÌÂ"
        .TextMatrix(Row, ColBalance) = PrevBalance
    Else
        If PrevBalance <> CDbl(.TextMatrix(Row, ColBalance)) Then
            If Not UpdateStmov(.TextMatrix(Row, ColBalance), PrevBalance, ComboStr.BoundText, .TextMatrix(Row, ColStkId)) Then
                 MsgBox Err.Description & Chr(13) & "Œÿ√ ›Ì «· ⁄œÌ·", vbExclamation, "Œÿ√"
            End If
        End If
    End If
End With
Exit Sub
errorhandler:
MsgBox Err.Description
ReturnOldData Row
End Sub

Private Sub FlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo errorhandler
    With flexGrid
        If Col <> ColBalance Then
            Cancel = False
        Else
            PrevBalance = .TextMatrix(Row, ColBalance)
        End If
    End With
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    If ActiveControl.Name = "FlexGrid" Then
        SendKeys "{home}+{end}"
    End If
End If
End Sub

Private Sub Form_Load()
    Init
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        SearchRec ComboStr.BoundText
    Case 3
       Unload Me
End Select
End Sub

Private Sub TxtStkNo_Change()
If Flag Then
    Dim sqltext As String
    If Trim(TxtStkNo.Text) = "" Then
        TxtStkNo.Tag = ""
        Grid.Visible = False
        Exit Sub
    End If
        sqltext = "Select top 15 Id , StkNo , StkName From CoStock  where StkNo like " & LikeExpression(TxtStkNo.Text) & " Or Stkname Like " & LikeExpression(TxtStkNo.Text)
    FillList sqltext, "Id", "StkNo", Grid
End If
End Sub

Private Sub TxtStkNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = False
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode
    Flag = True
End Sub

Private Sub txtStkNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            FillActiveControl Grid
            If Not SelectRow(Val(TxtStkNo.Tag)) Then
                MsgBox "«·„«œ… €Ì— „ÊÃÊœ…", vbQuestion, " ‰»ÌÂ"
                TxtStkNo.SetFocus
                SendKeys "{home}+{end}"
            Else
'                FlexGrid.Col = FlexGrid.Cols - 1
                flexGrid.Col = ColBalance
            End If
            'SendKeys "{home}+{end}"
    End If
End Sub

