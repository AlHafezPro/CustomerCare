VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmViewMaintCallOrders 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����� �������  �������"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7890
   ScaleMode       =   0  'User
   ScaleWidth      =   17308.2
   Begin Crystal.CrystalReport cr1 
      Left            =   4860
      Top             =   3750
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   615
      Left            =   2310
      TabIndex        =   6
      Top             =   720
      Width           =   14235
      _ExtentX        =   25109
      _ExtentY        =   1085
      _Version        =   131074
      ForeColor       =   255
      Begin VB.TextBox TxtSearch 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   90
         Width           =   12975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   13830
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   90
         Width           =   285
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   555
      Left            =   60
      TabIndex        =   2
      Top             =   7260
      Width           =   16485
      _ExtentX        =   29078
      _ExtentY        =   979
      _Version        =   131074
      Begin VB.Label LSelectedClaims 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   12870
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   150
         Width           =   2265
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������� �������"
         Height          =   195
         Index           =   1
         Left            =   15360
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��� �������"
         Height          =   195
         Index           =   0
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   150
         Width           =   780
      End
      Begin VB.Label LCount 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   150
         Width           =   2265
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   5835
      Left            =   60
      TabIndex        =   1
      Top             =   1380
      Width           =   16485
      _cx             =   29078
      _cy             =   10292
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3570
      Top             =   1200
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
            Picture         =   "FrmViewMaintCallOrders.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmViewMaintCallOrders.frx":6209F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   16575
      _ExtentX        =   29236
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
            Object.ToolTipText     =   "����� ������� ������� �������"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   1620
         TabIndex        =   15
         Top             =   60
         Width           =   12405
         _ExtentX        =   21881
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   13500
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   30
         Width           =   3015
         Begin MSMask.MaskEdBox TxtDate 
            Height          =   375
            Left            =   570
            TabIndex        =   12
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "����� �����"
            Height          =   195
            Index           =   2
            Left            =   2190
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   30
            Width           =   780
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   9300
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   90
         Width           =   2325
      End
   End
   Begin Threed.SSCheck ChkAllOrders 
      Height          =   315
      Left            =   90
      TabIndex        =   14
      Top             =   810
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   131074
      ForeColor       =   16711680
      Caption         =   "������� �������"
   End
   Begin VB.Menu mnu 
      Caption         =   "select"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect 
         Caption         =   "�����"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "�����"
      End
   End
End
Attribute VB_Name = "FrmViewMaintCallOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ColChk = 1
Const ColMaintCallNo = 2
Const ColCallNo = 3
Const ColCallDate = 4
Const ColCallTime = 5
Const ColClientName = 6
Const ColClientZoneNo = 7
Const ColClientZoneName = 8
Const ColClientAddress = 9
Const ColClientPhoneNBr = 10
Const ColClientMobilPhoneNbr = 11
Const colProdFamNo = 12
Const ColProdFamillyName = 13
Const ColCallDEscription = 14
Const ColCallNotes = 15
Const ColCallDefineName = 16
Const ColClientEntryEmpNo = 17
Const ColClientEntryName = 18
Const ColClaimsRepeatForTheClient = 19
 
 Dim maintDataService_ As New MaintDataService


Sub FillFormating(ByVal i As Integer, FlexGrid As VSFlexGrid)
If i = 1 Then
   
    fs = "|>" + ""
    fs = fs + "|>" + "��� ������"
    fs = fs + "|>" + "��� �����"
    fs = fs + "|>" + "����� �����"
    fs = fs + "|>" + "��� �����"
    fs = fs + "|>" + "��� ������"
    fs = fs + "|>" + "ZoneNo"
    fs = fs + "|>" + "�������"
    fs = fs + "|>" + "����� ������"
    fs = fs + "|>" + "��� ������"
    fs = fs + "|>" + "��� ��������"
    fs = fs + "|>" + "ProdFamNo"
    fs = fs + "|>" + "�������"
    fs = fs + "|>" + "������"
    fs = fs + "|>" + "�������"
    fs = fs + "|>" + "�������"
    fs = fs + "|>" + "CallReceiverEmpNo"
    fs = fs + "|>" + "���� ������"
    fs = fs + "|>" + "����� ������"

     With FlexGrid
        .FormatString = fs
        .Cols = 20
        
         .ColWidth(ColChk) = 300
        SetColWidths ColMaintCallNo, FlexGrid
        SetColWidths ColCallNo, FlexGrid
        .ColWidth(ColMaintState) = 0
        SetColWidths ColCallDate, FlexGrid
        SetColWidths ColCallTime, FlexGrid
        .ColWidth(colProdFamNo) = 0
        SetColWidths ColProdFamillyName, FlexGrid
        SetColWidths ColCallDEscription, FlexGrid
        SetColWidths ColClientName, FlexGrid
        .ColWidth(ColClientZoneNo) = 0
        SetColWidths ColClientZoneName, FlexGrid
        SetColWidths ColClientAddress, FlexGrid
        SetColWidths ColClientPhoneNBr, FlexGrid
        SetColWidths ColClientMobilPhoneNbr, FlexGrid
        SetColWidths ColCallNotes, FlexGrid
        SetColWidths ColCallDefineName, FlexGrid
        .ColWidth(ColClientEntryEmpNo) = 0
        SetColWidths ColClientEntryName, FlexGrid
        .ColWidth(ColClaimsRepeatForTheClient) = 0
        .ColDataType(ColChk) = flexDTBoolean
    End With
End If
End Sub

'Sub ShowToolBars()
'For i = 1 To Toolbar1.Buttons.Count - 1
'    If Val(Toolbar1.Buttons(i).Tag) <> 0 Then
'        Toolbar1.Buttons(i).Visible = Gettag(empNo, Toolbar1.Buttons(i).Tag)
'    End If
'Next
'End Sub

Sub init()
top = 0
left = 0
TxtDate.Text = Format(Date, "dd/mm/yyyy")
FlexGrid.Rows = 1
FillFormating 1, FlexGrid

FillGrid TxtDate.Text, ChkAllOrders.Value
FillMaintCallsStatistics

FlexGrid.Editable = flexEDKbdMouse


'ShowToolBars


End Sub

Sub FillMaintCallsStatistics()
On Error GoTo ErrorHandler

LCount.Caption = FlexGrid.Rows - 1
LSelectedClaims.Caption = GetSelectedClaimsCount

Exit Sub
ErrorHandler:
MsgBox Err.Description

End Sub

Sub FillGrid(vdate As String, allOrders As Boolean, Optional searchFormula As String = "")
On Error GoTo ErrorHandler
Dim rsMaintCallOrdersInfo As New ADODB.Recordset
Set rsMaintCallOrdersInfo = maintDataService_.GetMaintCallOrdersInfo(vdate, allOrders, searchFormula)
Set FlexGrid.DataSource = rsMaintCallOrdersInfo
FillFormating 1, FlexGrid
ColorRepeatedClaims
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Sub ColorRepeatedClaims()
On Error GoTo ErrorHandler
Dim Vrow As Integer
    With FlexGrid
        For i = 1 To .Rows - 1
            If .TextMatrix(i, ColClaimsRepeatForTheClient) > 1 Then
                Vrow = i
                 ColorRow Vrow, &HFFFFC0, FlexGrid
            End If
        Next
    End With
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Private Sub ChkAllOrders_Click(Value As Integer)
If ChkAllOrders.Value Then
    ChkAllOrders.Caption = "������� �������"
    Toolbar1.Buttons(1).Enabled = False
Else
    ChkAllOrders.Caption = "������� �������"
    Toolbar1.Buttons(1).Enabled = True

End If
FillGrid TxtDate.Text, ChkAllOrders.Value, TxtSearch.Text
FillMaintCallsStatistics

End Sub

Function GetSelectedClaimsCount() As Integer
On Error GoTo ErrorHandler
Dim selectedClaims As Integer
With FlexGrid
    For i = 1 To .Rows - 1
        If .TextMatrix(i, ColChk) Then
            selectedClaims = selectedClaims + 1
        End If
    Next i
End With
If selectedClaims < 0 Then
    GetSelectedClaimsCount = 0
Else
    GetSelectedClaimsCount = selectedClaims
End If
Exit Function
ErrorHandler:
MsgBox Err.Description
GetSelectedClaimsCount = 0
End Function
Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrorHandler
With FlexGrid
        maintDataService_.ChangeIsCheckMaintCall .TextMatrix(Row, ColCallNo), .TextMatrix(Row, ColChk)
End With
LSelectedClaims.Caption = GetSelectedClaimsCount
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub


Private Sub FlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)


If Col <> ColChk Then cancel = True
If ChkAllOrders.Value Then cancel = True
End Sub

Private Sub FlexGrid_DblClick()
On Error GoTo ErrorHandler
If FlexGrid.TextMatrix(FlexGrid.Row, ColMaintCallNo) = "" Then
    Exit Sub
End If
If Gettag(empNo, 39) Then
    Dim FrmMaintCall As New FrmMaintCallNew
    With FlexGrid
        idCallNo = .TextMatrix(.Row, ColMaintCallNo)
        LoadForm = True
        FrmMaintCall.Show
    End With
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub FlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim FirstRow As Integer, LastRow As Integer, Vrow As Integer
With FlexGrid
If KeyCode = vbKeyDelete Then
    If MsgBox("�� ��� ����� �� ����� �����", vbYesNo + vbDefaultButton2, "��� ������� �������") = vbYes Then
        If .Row >= .RowSel Then
            FirstRow = .Row
            LastRow = .RowSel
        Else
            FirstRow = .RowSel
            LastRow = .Row
        End If
        For i = FirstRow To LastRow Step -1
            Vrow = i
            If .Rows = 1 Then
                'UpdateRecords i
            Else
                  'RemoveFromMvStock .TextMatrix(i, Colid)
                  
                  If DeleteRow(FlexGrid, Vrow, ColCallNo, "MaintCallOrder", "CallNo") Then
                    .RemoveItem Vrow
                    
                End If
            End If
        Next
        FillMaintCallsStatistics
        .Col = ColCallNo
        .SetFocus
    End If
End If
End With
End Sub

Private Sub FlexGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button And vbRightButton Then
            PopupMenu mnu
    End If
End Sub


Private Sub Form_Load()
init
End Sub

Function ExecProcedure() As Boolean
On Error GoTo ErrorHandler
    sqlText = "sp_GetClientInfo"
    de.con.Execute (sqlText)
    ExecProcedure = True
Exit Function
ErrorHandler:
ExecProcedure = False
MsgBox Err.Description
End Function

Function GetSelectedIds(Vcol As Integer) As String
On Error GoTo ErrorHandler
Dim maintCallOrdersString As String
If FlexGrid.Rows = 1 Then
    GetSelectedIds = ""
    Exit Function
End If
maintCallOrdersString = ""
ProgressBar1.Min = 0
ProgressBar1.Max = FlexGrid.Rows - 1
ProgressBar1.Visible = True
ProgressBar1.Value = 0
With FlexGrid
    For i = 1 To .Rows - 1
        ProgressBar1.Value = ProgressBar1.Value + 1
        If .TextMatrix(i, ColChk) Then
            If .TextMatrix(i, Vcol) <> "" Then
                maintCallOrdersString = maintCallOrdersString & "," & .TextMatrix(i, Vcol)
            End If
        End If
    Next
End With
ProgressBar1.Visible = False
GetSelectedIds = Mid(maintCallOrdersString, 2)
Exit Function
ErrorHandler:
GetSelectedIds = ""
MsgBox Err.Description
End Function
Sub TransferToMaintCallAndAndPrint()
On Error GoTo ErrorHandler
Dim unPrintedMaintCallOrders As String
If MsgBox("�� ��� ����� �� ����� ������� ������� ������� ��� ��� ������� �", vbQuestion + vbYesNo + vbDefaultButton2, "����� ������� ������� �������") = vbYes Then
    Screen.MousePointer = vbHourglass
    unPrintedMaintCallOrders = GetSelectedIds(ColCallNo)
    If unPrintedMaintCallOrders = "" Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    TransferOrdersToMaintCall unPrintedMaintCallOrders
    FillGrid TxtDate.Text, ChkAllOrders.Value, TxtSearch.Text
    FillMaintCallsStatistics
    Screen.MousePointer = vbDefault
End If
Exit Sub
ErrorHandler:
Screen.MousePointer = vbDefault
MsgBox Err.Description
End Sub

Sub updateFlexGrid(EnumMaintCallStateRec As EnumMaintCallState)
If unPrintedMaintCalls Then Exit Sub
    With FlexGrid
        .Visible = False
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, ColMantStatePrinted)) = 0 And EnumMaintCallStateRec = UnderwayAndPrinted Then
                .TextMatrix(i, UnderwayAndPrinted) = 1
        
            Else
                .TextMatrix(i, UnderwayAndPrinted) = 0
              End If
        Next
        .Visible = True
End With
End Sub
Sub TransferOrdersToMaintCall(unPrintedMaintCallOrders As String)
On Error GoTo ErrorHandler
    maintDataService_.TransferOrdersToMaintCall unPrintedMaintCallOrders
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub


Sub UpdateMaintCallState(unPrintedMaintCalls As String, EnumMaintCallStateRec As EnumMaintCallState)
On Error GoTo ErrorHandler
    maintDataService_.UpdateMaintCallState unPrintedMaintCalls, EnumMaintCallStateRec
    FillMaintCallsStatistics
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub mnuCancel_Click()
UpdateRow FlexGrid, False
FillMaintCallsStatistics
End Sub

Private Sub mnuSelect_Click()
UpdateRow FlexGrid, True
FillMaintCallsStatistics
End Sub
Sub UpdateRow(FlexGrid As VSFlexGrid, isChk As Boolean)
On Error GoTo ErrorHandler
With FlexGrid
        If .Row >= .RowSel Then
            FirstRow = .Row
            LastRow = .RowSel
        Else
            FirstRow = .RowSel
            LastRow = .Row
        End If
        For i = FirstRow To LastRow Step -1
                .TextMatrix(i, ColChk) = isChk
                maintDataService_.ChangeIsCheckMaintCall .TextMatrix(i, ColCallNo), .TextMatrix(i, ColChk)
        Next i
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index

    Case 1
        TransferToMaintCallAndAndPrint
    Case 3
        Unload Me
End Select
End Sub

Private Sub TxtDate_Change()
FillGrid TxtDate.Text, ChkAllOrders.Value, TxtSearch.Text
End Sub

'Sub NumberingTheSelectedOrders()
'On Error GoTo errorhandler
'    If ChkAllClaims.Value Then
'        MsgBox "������� ����� ������ ����� �������", vbInformation, "������� �����"
'        Exit Sub
'    End If
'    If MsgBox("���� ����� ������� �������" + Chr(13) + "�� ��� ����� �", vbQuestion + vbYesNo + vbDefaultButton2, "����� ������� �������") = vbYes Then
'        Screen.MousePointer = vbHourglass
'        Dim selectedClaimsWithoutNumber As String
'        selectedClaimsWithoutNumber = GetSelectedIds(ColCallNo)
'        If maintDataService_.NumberingTheSelectedOrders(selectedClaimsWithoutNumber) Then
'            TxtSearch_Change
'        End If
'        Screen.MousePointer = vbDefault
'    Else
'        Exit Sub
'    End If
'
'With FlexGrid
'
'End With
'
'Exit Sub
'errorhandler:
'Screen.MousePointer = vbDefault
'MsgBox Err.Description
'End Sub
Private Sub TxtSearch_Change()
FillGrid TxtDate.Text, ChkAllOrders.Value, TxtSearch.Text
FillMaintCallsStatistics
End Sub

Private Sub TxtSearch_GotFocus()
ChangeToArabic

End Sub
