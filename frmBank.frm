VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBank 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÊÑãíÒ ÇáÈÇÑßÒÏ"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   11265
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   2895
      Left            =   5370
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   1995
      _cx             =   3519
      _cy             =   5106
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
   Begin VB.TextBox TxtStkName 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6060
      TabIndex        =   0
      Top             =   1050
      Width           =   5235
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   390
      Top             =   1470
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3210
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
            Picture         =   "frmBank.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBank.frx":6209F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   1217
      ButtonWidth     =   1191
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   37
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "1"
                  Text            =   "ÊÞÑíÑ ááÈäß"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "2"
                  Text            =   "ÊÞÑíÑ ãÕÇÑíÝ ÇáÊæØíä"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   34
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   4530
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   0
         Width           =   6825
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   4425
      Left            =   60
      TabIndex        =   5
      Top             =   1470
      Width           =   11235
      _cx             =   19817
      _cy             =   7805
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
      ScrollTips      =   -1  'True
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
   Begin VB.Label LBalance 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÑÕíÏ ÇáãÇÏå"
      Height          =   195
      Index           =   0
      Left            =   2220
      TabIndex        =   11
      Top             =   810
      Width           =   780
   End
   Begin VB.Label LStkName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   3060
      TabIndex        =   1
      Top             =   1080
      Width           =   2955
   End
   Begin Threed.SSCheck ChkManual 
      Height          =   345
      Left            =   60
      TabIndex        =   4
      Top             =   1080
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   609
      _Version        =   131074
      ForeColor       =   8388608
      Caption         =   "íÏæí"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÅÓã ÇáãÇÏå"
      Height          =   195
      Index           =   4
      Left            =   5280
      TabIndex        =   9
      Top             =   810
      Width           =   705
   End
   Begin Threed.SSCheck chkBarcode 
      Height          =   345
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   609
      _Version        =   131074
      ForeColor       =   8388608
      Caption         =   "ÈÇÑßæÏ"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÑÞã ÇáãÇÏå"
      Height          =   195
      Index           =   1
      Left            =   10620
      TabIndex        =   8
      Top             =   810
      Width           =   675
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OK As Boolean, Flag As Boolean, Pos As Integer
'Dim BankRec As BankEmpType

Const ColStkId = 1
Const ColStkNo = 2
Const ColStkName = 3
Const ColBalance = 4
Const ColPrice = 5
Const ColIsBarcode = 6
Const ColIsManual = 7


Const ColStkId_1 = 1
Const ColStkno_1 = 2
Const ColStkname_1 = 3
Const ColIsBarcode_1 = 4
Const ColIsManual_1 = 5

Dim AllowAdd As Boolean


Sub MoveCursor(KeyCode As Integer, FlexGrid As VSFlexGrid)
On Error Resume Next
If Not FlexGrid.Visible Then Exit Sub
With FlexGrid
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

Sub FillFormating(ByVal i As Integer)
If i = 1 Then
    Fs = "|>" + "StkId"
    Fs = Fs + "|>" + "ÑÞã ÇáãÇÏÉ"
    Fs = Fs + "|>" + "ÅÓã ÇáãÇÏå"
    Fs = Fs + "|>" + "ÑÕíÏ ÇáãÓÊæÏÚ"
    Fs = Fs + "|>" + "ÓÚÑ ÇáãÇÏå"
    Fs = Fs + "|>" + "ÈÇÑßæÏ"
    Fs = Fs + "|>" + "íÏæí"
    With FlexGrid
        .FormatString = Fs
        .Cols = 8
        .ColWidth(ColStkId) = 0
        SetColWidths ColStkNo, FlexGrid
        SetColWidths ColStkName, FlexGrid
        SetColWidths ColBalance, FlexGrid
        SetColWidths ColPrice, FlexGrid
        SetColWidths ColIsBarcode, FlexGrid
        SetColWidths ColIsManual, FlexGrid
        .ColDataType(ColIsBarcode) = flexDTBoolean
        .ColDataType(ColIsManual) = flexDTBoolean
    End With
ElseIf i = 2 Then
    Fs = "|>" + "stkid"
    Fs = Fs + "|>" + "ÑÞã ÇáãÇÏå"
    Fs = Fs + "|>" + "ÅÓã ÇáãÇÏå"
    Fs = Fs + "|>" + "ÇáÈÇÑßæÏ"
    Fs = Fs + "|>" + "íÏæí"
    With Grid
        .FormatString = Fs
        .Cols = 6
        .ColWidth(ColStkId_1) = 0
        SetColWidths ColStkno_1, Grid
        SetColWidths ColStkname_1, Grid
        SetColWidths ColIsBarcode_1, Grid
        SetColWidths ColIsManual_1, Grid
        .ColDataType(ColIsBarcode_1) = flexDTBoolean
        .ColDataType(ColIsManual_1) = flexDTBoolean
    End With
End If
End Sub

Sub SetColWidths(ByVal ColNo As Integer, FlexGrid As VSFlexGrid)
    With FlexGrid
        .AutoSize (ColNo)
    End With
End Sub
Sub ChangeCursor(ByVal X As Integer)
If X = 1 Then
    With TxtStkName
       Grid.Top = .Top + .Height
       Grid.Left = .Left
       Grid.Width = .Width
    End With
End If
End Sub

Sub init()
Dim Rs As New ADODB.Recordset
    Top = 0
    Left = 0
    FlexGrid.Editable = flexEDKbdMouse
    OK = True
    sqltext = "Select c1.id , c1.StkNo , StkName , s1.FnlQnt , CliPrice , isbarcode , ismanual  from CoStock c1 left outer join Stkinf s1  on c1.Id = s1.Stkid"
    Set Rs = de.con.Execute(sqltext)
    Set FlexGrid.DataSource = Rs
    FillFormating 1
End Sub

Private Sub chkBarcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ChkManual.SetFocus
End If
End Sub
Function UpdateItemNo(Stkid As Double, barcode As Integer, manual As Integer) As Boolean
On Error GoTo errorhandler
sqltext = "update costock set isbarcode=" & barcode & ",IsManual=" & manual & " where id=" & Stkid
de.con.Execute (sqltext)
UpdateItemNo = True
Exit Function
errorhandler:
MsgBox Err.Description
UpdateItemNo = False
End Function
Function GetRow(Stkid As Double) As Integer
With FlexGrid
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, ColStkId)) = Stkid Then
            GetRow = i
            Exit Function
        End If
    Next
End With
End Function
Function MovToItemNo(Stkid As Double) As Integer
On Error GoTo errorhandler
Dim vRow As Integer
vRow = GetRow(Stkid)
If vRow = 0 Then Exit Function
With FlexGrid
    .TextMatrix(vRow, ColIsBarcode) = Val(chkBarcode.Value)
    .TextMatrix(vRow, ColIsManual) = Val(ChkManual.Value)
    If Not .RowIsVisible(vRow) Then
        .TopRow = vRow
    End If
End With

Exit Function
errorhandler:
MsgBox Err.Description
MoveToRec = 1
End Function


Private Sub ChkManual_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If UpdateItemNo(TxtStkName.Tag, Val(chkBarcode.Value), Val(ChkManual.Value)) Then
        MovToItemNo (TxtStkName.Tag)
         TxtStkName.Tag = 0
        OK = False
        TxtStkName.Text = ""
        LStkName.Caption = ""
        LBalance.Caption = ""
        chkBarcode.Value = ssCBUnchecked
        ChkManual.Value = ssCBUnchecked
        OK = True
        TxtStkName.SelStart = 0
        TxtStkName.SelLength = Len(TxtStkName.Text)
        TxtStkName.SetFocus
    End If
End If
End Sub

Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo errorhandler
With FlexGrid
sqltext = "update costock set isbarcode =" & Val(.TextMatrix(Row, ColIsBarcode)) & ",ismanual=" & Val(.TextMatrix(Row, ColIsManual)) & " where id=" & .TextMatrix(Row, ColStkId)
de.con.Execute (sqltext)
End With
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

Private Sub FlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col <> ColIsBarcode And Col <> ColIsManual Then
Cancel = True
End If
End Sub

Private Sub Form_Load()
init

End Sub
Function GetStkName(ByVal Stkid As Integer) As String
On Error GoTo errorhandler
Dim Rs As New ADODB.Recordset
sqltext = "Select StkName From CoStock Where id = " & Stkid
Set Rs = de.con.Execute(sqltext)
GetStkName = Rs!StkName
Exit Function
errorhandler:
GetStkName = ""
End Function

Private Sub Grid_RowColChange()
If Flag Then
    OK = False
    With Grid
       Select Case Pos
        Case 1
            TxtStkName.Tag = .TextMatrix(.Row, ColStkId_1)
            TxtStkName.Text = .TextMatrix(.Row, ColStkno_1)
            LStkName.Caption = .TextMatrix(.Row, ColStkname_1)
            chkBarcode.Value = Val(.TextMatrix(.Row, ColIsBarcode_1))
            ChkManual.Value = Val(.TextMatrix(.Row, ColIsManual_1))
            LBalance.Caption = GetBalance(.TextMatrix(.Row, ColStkno_1))
       End Select
    End With
    OK = True
End If
End Sub
Function ItemsCount() As Integer
On Error GoTo errorhandler
Dim Rs As New ADODB.Recordset
sqltext = "Select Count(*) CountRec FROM dbo.costock Where cliprice is not null"
Set Rs = de.con.Execute(sqltext)
EmployeesCount = Rs!CountRec
Exit Function
errorhandler:
EmployeesCount = 0
End Function
Sub PrintData(ByVal i As Integer)
On Error GoTo errorhandler
With cr1
    .Connect = ODBCString
    Select Case i
        Case 1
            .Formulas(0) = ""
            .Formulas(1) = ""
            .Formulas(2) = ""
            .Formulas(3) = ""
            .Formulas(4) = ""
            .ReportFileName = App.Path + "\Reports\RepBank1.rpt"
        Case 2
            .Formulas(0) = "Value-date='" & TxtValueDate.Text & "'"
            .Formulas(1) = ""
            .Formulas(2) = ""
            .Formulas(3) = "FeesBank=" & Val(TxtFees.Text)
            .Formulas(4) = "Managment='" & TxtManagment.Text & "'"
            .ReportFileName = App.Path + "\Reports\RepBank2.rpt"
        Case 3
            .Formulas(0) = ""
            .Formulas(1) = "BankFees=" & TxtFees.Text
            .Formulas(2) = "FeesTotal=" & Val(TxtFees.Text) * EmployeesCount
            .Formulas(3) = ""
            .Formulas(4) = ""
            .ReportFileName = App.Path + "\Reports\RepBank3.rpt"
    End Select
    .SQLQuery = "SELECT    year, Month, FullName, PaidSalWithProdBank, BankAccNo FROM dbo.PaidSalaries Where Bank=1"
    .DiscardSavedData = True
    .WindowState = crptMaximized
    .Action = 1
End With
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub
Private Sub ExportToExcel()

Dim objXL As Excel.Application
Dim objWB As Excel.Workbook
Dim objWS As Excel.Worksheet
Dim r As Long
Dim c As Long
Set objXL = New Excel.Application
Set objWB = objXL.Workbooks.Add
Set objWS = objWB.Worksheets(1)
'ProgressBar1.Visible = True
'ProgressBar1.Min = 0
'ProgressBar1.Max = FlexGrid.Rows
With objWS
    For r = 0 To FlexGrid.Rows - 1
'    ProgressBar1.Value = r
    For c = 0 To FlexGrid.Cols - 1
         If c = 8 Then
            'Selection.NumberFormat = "General"
         End If
        .Cells(r + 1, c + 1) = FlexGrid.TextMatrix(r, c)
Next
Next
.Cells.Columns.AutoFit
End With
'ProgressBar1.Visible = False
objXL.Visible = True
Set objWS = Nothing
Set objWB = Nothing
Set objXL = Nothing

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        'PrintData 1
    Case 3
        
        'ExportToExcel
    Case 5
        Unload Me
End Select
End Sub

Function Found(stkno As String) As Boolean
Dim Rs As New ADODB.Recordset
sqltext = "Select Count(*) CountRec  From costock Where stkno='" & stkno & "'"
Set Rs = de.con.Execute(sqltext)
If Rs!CountRec > 0 Then
    Found = True
Else
    Found = False
End If
End Function

Function FillVariables() As Boolean
On Error GoTo errorhandler
    If Val(TxtStkName.Tag) = 0 Then
        FillVariables = False
        Exit Function
    End If
FillVariables = True
Exit Function
errorhandler:
FillVariables = False
MsgBox Err.Description
End Function

Function FillStructure() As Boolean
On Error GoTo errorhandler
If FillVariables Then
    With BankRec
        .EmpNo = Val(TxtStkName.Tag)
        .EnglishName = TxtEnglishName.Text
        .Bank = Abs(Val(ChkBank.Value))
        .AccNo = TxtAccNo.Text
    End With
End If
FillStructure = True
Exit Function
errorhandler:
FillStructure = False
MsgBox Err.Description
End Function
Function SaveRec() As Boolean
On Error GoTo errorhandler
If FillStructure Then
    With BankRec
'        If Not Found(.Modno) Then
'            Sqltext = "insert into CoModelDiscount(ModNo , Discount)Values(" & .Modno & "," & .Amount & ")"
'            AllowAdd = True
'        Else
            sqltext = "Update Employee  SEt  Bank=" & .Bank & ",BankAccNo='" & .AccNo & "',EnglishName='" & .EnglishName & "'  Where Empno=" & .EmpNo
            AllowAdd = False
            de.con.Execute (sqltext)
'        End If
    End With
End If
SaveRec = True
Exit Function
errorhandler:
SaveRec = False
MsgBox Err.Description
End Function
Function GetEmployeeType(ByVal EmpNo As Integer) As String
Dim Rs As New ADODB.Recordset
sqltext = "Select Case When Type= 2 Then 'ãÓÊÞíá' Else '' end Type From EmpFullName Where EmpNo =" & EmpNo
Set Rs = de.con.Execute(sqltext)
GetEmployeeType = Rs!Type
End Function
Sub insertintoGrid()
Dim vRow As Integer
With FlexGrid
    If AllowAdd Then
        .AddItem ""
        vRow = .Rows - 1
    Else
        vRow = GetRow(BankRec.EmpNo)
        If vRow = 0 Then vRow = .Rows - 1
    End If
    .TextMatrix(vRow, ColEmpNo) = BankRec.EmpNo
    .TextMatrix(vRow, Colfullname) = TxtStkName.Text
    .TextMatrix(vRow, ColEnglishName) = TxtEnglishName.Text
    .TextMatrix(vRow, ColAccNo) = TxtAccNo.Text
    .TextMatrix(vRow, ColBankChk) = BankRec.Bank
    .TextMatrix(vRow, ColEmpType) = GetEmployeeType(BankRec.EmpNo)
    If Not .RowIsVisible(vRow) Then
        .TopRow = vRow
    End If
End With
End Sub
Function GetCount(ByVal i As Integer) As Integer
Dim Rs As New ADODB.Recordset
sqltext = "Select Count(*) CountRec From paidsalaries Where isnull(Bank,0)=" & i
Set Rs = de.con.Execute(sqltext)
GetCount = Rs!CountRec
End Function
Private Sub ChkBank_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If SaveRec Then
        insertintoGrid
        TxtStkName.Tag = 0
        OK = False
        TxtStkName.Text = ""
        TxtAccNo.Text = ""
        ChkBank.Value = ssCBUnchecked
        OK = True
        LBank.Caption = GetCount(1)
        LUnBank.Caption = GetCount(0)
        TxtStkName.SelStart = 0
        TxtStkName.SelLength = Len(TxtStkName.Text)
        TxtStkName.SetFocus
    End If
End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Tag
    Case 1
        PrintData 2
    Case 2
        PrintData 3
End Select
End Sub

Private Sub TxtAccNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ChkBank.SetFocus
End If
End Sub

Private Sub TxtStkName_Change()
On Error GoTo errorhandler
Dim RsSearch As New ADODB.Recordset
If TxtStkName.Text = "" Then
    TxtStkName.Tag = 0
    Grid.Visible = False
    Exit Sub
End If
If OK Then
    Flag = flase
    sqltext = "Select top 30 c1.id , StkNo , stkname  , isbarcode , ismanual From costock c1  Where c1.stkname Like" & LikeExpression(TxtStkName.Text) & " Or stkno Like" & LikeExpression(TxtStkName.Text)
    Set RsSearch = de.con.Execute(sqltext)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        Grid.Row = 0
        FillFormating 2
        ChangeCursor 1
        Grid.Visible = True
    Else
        TxtStkName.Tag = 0
        Grid.Visible = False
    End If
    Flag = True
End If
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub


Private Sub TxtStkName_GotFocus()
TxtStkName.Tag = 0
OK = False
TxtStkName.Text = ""
LStkName.Caption = ""
LBalance.Caption = ""
chkBarcode.Value = ssCBUnchecked
ChkManual.Value = ssCBUnchecked
OK = True

Pos = 1
End Sub

Private Sub TxtStkName_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True

End Sub


Function GetEnglishName(ByVal EmpNo As Integer) As String
    Dim Rs As New ADODB.Recordset
    sqltext = "Select isnull(EnglishName,'') EnglishName From Employee  Where EmpNo = " & EmpNo
    Set Rs = de.con.Execute(sqltext)
    GetEnglishName = Rs!EnglishName
End Function

Private Sub TxtStkName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible And Val(TxtStkName.Tag) = 0 Then
        TxtStkName.Tag = Grid.TextMatrix(1, ColStkId_1)
        OK = False
        TxtStkName.Text = Grid.TextMatrix(1, ColStkno_1)
        OK = True
        LStkName.Caption = Grid.TextMatrix(1, ColStkname_1)
        LBalance.Caption = GetBalance(Grid.TextMatrix(1, ColStkno_1))
        chkBarcode.Value = Val(Grid.TextMatrix(1, ColIsBarcode_1))
        ChkManual.Value = Val(Grid.TextMatrix(1, ColIsManual_1))
        
    End If
    Grid.Visible = False
    chkBarcode.SetFocus
    SendKeys "{Home}+{End}"
End If
End Sub

Private Sub TxtEnglishName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{home}+{end}"
    TxtAccNo.SetFocus
End If
End Sub
Private Sub TxtManagment_Change()
On Error GoTo errorhandler
Dim RsSearch As New ADODB.Recordset
If TxtManagment.Text = "" Then
    TxtManagment.Tag = 0
    Grid.Visible = False
    Exit Sub
End If
If OK Then
    Flag = flase
    sqltext = "Select EmpNo , fULLnAME  From EmpFullName e1 Where FullName Like" & LikeExpression(TxtManagment.Text) & " Or Convert(varchar(255),EmpNo) Like" & LikeExpression(TxtManagment.Text)
    Set RsSearch = de.con.Execute(sqltext)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        Grid.Row = 0
        FillFormating 2
        ChangeCursor 2
        Grid.Visible = True
    Else
        TxtManagment.Tag = 0
        Grid.Visible = False
    End If
    Flag = True
End If
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub


Private Sub TxtManagment_GotFocus()
Pos = 2
End Sub

Private Sub TxtManagment_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True

End Sub



Private Sub TxtManagment_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible And Val(TxtManagment.Tag) = 0 Then
        TxtManagment.Tag = Grid.TextMatrix(1, ColEmpNo_1)
        OK = False
        TxtManagment.Text = Grid.TextMatrix(1, ColFullName_1)
        OK = True
    End If
    Grid.Visible = False
    TxtManagment.SetFocus
    SendKeys "{Home}+{End}"
End If
End Sub

