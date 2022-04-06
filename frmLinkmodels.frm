VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLinkmodels 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÑÈØ ÇáãæÏíáÇÊ"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   15660
   Begin VB.TextBox TxtMaintFamNo 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1050
      Width           =   3465
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   2895
      Left            =   3900
      TabIndex        =   6
      Top             =   3030
      Visible         =   0   'False
      Width           =   2895
      _cx             =   5106
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
   Begin VB.TextBox txtSalesModel 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   5130
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1050
      Width           =   3465
   End
   Begin VB.TextBox txtmaintmodel 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   12120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1050
      Width           =   3465
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   300
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   38
      ImageHeight     =   38
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":5568
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":83C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":AB6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":D466
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":F98B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":12143
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":14B57
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":174A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":1A21D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":1CA6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":1F913
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":2266D
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":2500D
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":27F6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":2A997
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":2D351
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":2FC82
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":32588
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":34FF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":37FA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":3A8CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":3D1FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":3F93F
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":422AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":44B36
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":46D45
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":496A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":4BECC
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":4E980
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":513BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":53ED2
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":56E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":59C38
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":5F744
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":62382
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinkmodels.frx":64F31
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
      Width           =   15660
      _ExtentX        =   27623
      _ExtentY        =   1217
      ButtonWidth     =   1191
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   5535
      Left            =   30
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1470
      Width           =   15585
      _cx             =   27490
      _cy             =   9763
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
   Begin VB.Label LSalesFamNo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1770
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1050
      Width           =   3345
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÇáÚÇÆáå"
      Height          =   195
      Left            =   4635
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   750
      Width           =   420
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÊÚÏíá ÇáÚÇÆáå"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11160
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   750
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ãæÏÈá ÇáãÈíÚÇÊ"
      Height          =   195
      Left            =   7590
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   750
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ãæÏíá ÎÏãå ÇáãÓÊåáß"
      Height          =   195
      Left            =   14115
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   750
      Width           =   1440
   End
End
Attribute VB_Name = "frmLinkmodels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const colMaintModNo = 1
Const colMaintModName = 2
Const colMaintModelDescription = 3
Const ColMaintFamNo = 4
Const ColMaintProdFamName = 5

Const ColSalesModNo = 6
Const ColSalesModName = 7
Const ColSalesModelDescription = 8
Const ColsalesFamNo = 9
Const ColSalesProdFamName = 10
Const colCountRec = 11


Const ColNo = 1
Const ColName = 2



Dim OK, Flag As Boolean, Pos As Integer

Function GetRow(ID As Integer) As Integer
With FlexGrid
    For i = 1 To .Rows - 1
        If .TextMatrix(i, ColNo) = ID Then
            GetRow = i
            Exit Function
        End If
    Next
End With
End Function
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

Sub ChangeCursor(ByVal X As Integer)
If X = 1 Then
    With txtmaintmodel
       Grid.top = .top + .Height
       Grid.left = .left
       Grid.Width = .Width
    End With
ElseIf X = 2 Then
    With txtSalesModel
       Grid.top = .top + .Height
       Grid.left = .left
       Grid.Width = .Width
    End With
ElseIf X = 3 Then
    With TxtMaintFamNo
        Grid.top = .top + .Height
        Grid.left = .left
        Grid.Width = .Width
    End With

End If
End Sub

Sub FillFormating(FlexGrid As VSFlexGrid, Index As Integer)
Select Case Index
    Case 1
        FS = "|>" + "modno"
        FS = FS + "|>" + "ãæÏíá ÎÏãÉ ÇáãÓÊåáß"
        FS = FS + "|>" + "ÇáÔÑÍ"
        FS = FS + "|>" + "MaintFamNo"
        FS = FS + "|>" + "ÚÇÆáå ÎÏãÉ ÇáãÓÊåáß"
        
        FS = FS + "|>" + "SalesModNo"
        FS = FS + "|>" + "ÅÓã ãæÏíá ÇáãÈíÚÇÊ"
        FS = FS + "|>" + "ÇáÔÑÍ"
        FS = FS + "|>" + "SalesFamNo"
        FS = FS + "|>" + "ÚÇÆáå ÈÑäÇãÌ ÇáãÈíÚÇÊ"
        FS = FS + "|>" + "ÚÏÏ ãÑÇÊ ÇáÇÓÊÎÏÇã"
        
        With FlexGrid
            .FormatString = FS
            .ColWidth(colMaintModNo) = 0
            SetColWidths colMaintModName, FlexGrid
            SetColWidths colMaintModelDescription, FlexGrid
            .ColWidth(ColMaintFamNo) = 0
            SetColWidths ColMaintProdFamName, FlexGrid
            
            .ColWidth(ColSalesModNo) = 0
            SetColWidths ColSalesModName, FlexGrid
            SetColWidths ColSalesModelDescription, FlexGrid
            .ColWidth(ColsalesFamNo) = 0
            SetColWidths ColSalesProdFamName, FlexGrid
            SetColWidths colCountRec, FlexGrid
        End With
    Case 2
        FS = "|>" + "Id"
        FS = FS + "|>" + "ÇáÔÑÍ"
        With FlexGrid
            .FormatString = FS
            .ColWidth(ColNo) = 0
            SetColWidths ColName, FlexGrid
        End With
End Select
End Sub

Sub SetColWidths(ByVal ColNo As Integer, FlexGrid As VSFlexGrid)
    With FlexGrid
        .AutoSize (ColNo)
    End With
End Sub


Private Sub FlexGrid_RowColChange()
    If Flag Then
    OK = False
        With FlexGrid
            txtmaintmodel.Tag = .TextMatrix(.Row, colMaintModNo)
            txtmaintmodel.Text = .TextMatrix(.Row, colMaintModName)
            TxtMaintFamNo.Text = .TextMatrix(.Row, ColMaintProdFamName)
            TxtMaintFamNo.Tag = .TextMatrix(.Row, ColMaintFamNo)
            txtSalesModel.Tag = .TextMatrix(.Row, ColSalesModNo)
            txtSalesModel.Text = .TextMatrix(.Row, ColSalesModName)
            LSalesFamNo.Caption = .TextMatrix(.Row, ColSalesProdFamName)
            LSalesFamNo.Tag = .TextMatrix(.Row, ColsalesFamNo)
            
        End With
    OK = True
    End If
End Sub

Private Sub Form_Load()
Init
End Sub
Sub FillGrid()
    Dim Rs As New ADODB.Recordset
    sqltext = "select modno,  MaintSymbol, MaintName, MaintFamNo, MaintProdFamName,SalesModNo, SalesSymbol, SalesName ,SalesFamNo, SalesProdFamName, CountRec from maintmodelsQry"
    Set Rs = de.con.Execute(sqltext)
    Set FlexGrid.DataSource = Rs
    FillFormating FlexGrid, 1
End Sub

Sub Init()
OK = True
Flag = False
    FillGrid
    Me.top = 0
    Me.left = 0
    With FlexGrid
'        .Cols = 3
        .Editable = flexEDKbdMouse
    End With
Flag = True
End Sub

Private Sub Grid_RowColChange()
If Flag Then
    OK = False
    With Grid
       Select Case Pos
        Case 1
            txtmaintmodel.Tag = .TextMatrix(.Row, ColNo)
            txtmaintmodel.Text = .TextMatrix(.Row, ColName)
        Case 2
            txtSalesModel.Tag = .TextMatrix(.Row, ColNo)
            txtSalesModel.Text = .TextMatrix(.Row, ColName)
        Case 3
            TxtMaintFamNo.Tag = .TextMatrix(.Row, ColNo)
            TxtMaintFamNo.Text = .TextMatrix(.Row, ColName)
       End Select
    End With
    OK = True
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    
    Case 1
        Unload Me
End Select
End Sub

Private Sub TxtMaintFamNo_Change()
On Error GoTo errorhandler
Dim RsSearch As New ADODB.Recordset
If TxtMaintFamNo.Text = "" Then
    TxtMaintFamNo.Tag = 0
    Grid.Visible = False
    Exit Sub
End If
If OK Then
    Flag = False
    sqltext = "Select top 10 Prodfamno , prodFamNameA  from adhamproductFamily Where prodFamNameA Like" & LikeExpression(TxtMaintFamNo.Text)
    Set RsSearch = de.con.Execute(sqltext)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        Grid.Row = 0
        FillFormating Grid, 2
        ChangeCursor 3
        Grid.Visible = True
    Else
        TxtMaintFamNo.Tag = 0
        Grid.Visible = False
    End If
    Flag = True
End If
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

Private Sub TxtMaintFamNo_GotFocus()
Pos = 3
End Sub

Private Sub TxtMaintFamNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True
End Sub
Sub UpdateFamily(ModNo As Integer, FamNo As Integer)
On Error GoTo errorhandler
sqltext = "Update AdhamModels set FamNo = " & FamNo & " Where ModNo=" & ModNo
de.con.Execute (sqltext)
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub
Private Sub TxtMaintFamNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible Then
        OK = False
        TxtMaintFamNo.Tag = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
        TxtMaintFamNo.Text = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColName)
        OK = True
    ElseIf Grid.Visible = False And TxtMaintFamNo.Text <> "" And Val(TxtMaintFamNo.Tag) <> 0 Then
        txtSalesModel.SetFocus
        txtSalesModel.SelStart = 0
        txtSalesModel.SelLength = Len(txtSalesModel.Text)
        Exit Sub
    Else
        OK = False
        TxtMaintFamNo.Tag = 0
        TxtMaintFamNo.Text = ""
        OK = True
    End If
    Grid.Visible = False
    If Val(txtmaintmodel.Tag) <> 0 And Val(TxtMaintFamNo.Tag) <> 0 Then
        UpdateFamily Val(txtmaintmodel.Tag), Val(TxtMaintFamNo.Tag)
    End If
    txtSalesModel.SetFocus
    txtSalesModel.SelStart = 0
    txtSalesModel.SelLength = Len(txtSalesModel.Text)
End If
End Sub

Private Sub txtmaintmodel_Change()
On Error GoTo errorhandler
Dim RsSearch As New ADODB.Recordset
If txtmaintmodel.Text = "" Then
    txtmaintmodel.Tag = 0
    Grid.Visible = False
    Exit Sub
End If
If OK Then
    Flag = False
    sqltext = "Select top 10 ModNo , symbol  from AdhamModels  Where Symbol Like" & LikeExpression(txtmaintmodel.Text) & " or Name Like" & LikeExpression(txtmaintmodel.Text)
    Set RsSearch = de.con.Execute(sqltext)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        Grid.Row = 0
        FillFormating Grid, 2
        ChangeCursor 1
        Grid.Visible = True
    Else
        txtmaintmodel.Tag = 0
        Grid.Visible = False
    End If
    Flag = True
End If
Exit Sub
errorhandler:
MsgBox Err.Description

End Sub


Private Sub txtmaintmodel_GotFocus()
Pos = 1
End Sub

Private Sub txtmaintmodel_KeyDown(KeyCode As Integer, Shift As Integer)
 Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True

End Sub
Friend Function GetFamNo(ModNo As Integer, Index As Integer) As PordFAmilyType
On Error GoTo errorhandler
Dim ProdFamRec As PordFAmilyType

Dim Rs As New ADODB.Recordset
If Index = 1 Then
    sqltext = "select ProdFamNo , ProdFamNameA as ProdFamName From AdhamModels a1 inner join AdhamProductFAmily a2 on a1.FamNo = a2.ProdFamNo where a1.ModNo = " & ModNo
ElseIf Index = 2 Then
    sqltext = "select ProdFamNo , ProdFamName From Hafez2020.dbo.Models a1 inner join Hafez2020.dbo.ProductFAmily a2 on a1.FamNo = a2.ProdFamNo where a1.ModNo = " & ModNo
End If
Set Rs = de.con.Execute(sqltext)

ProdFamRec.ProdFamNo = Rs!ProdFamNo
ProdFamRec.ProdFAmName = Rs!ProdFAmName
GetFamNo = ProdFamRec
Exit Function
errorhandler:
MsgBox Err.Description
End Function
Private Sub txtmaintmodel_KeyPress(KeyAscii As Integer)
Dim vRow  As Integer
Dim PordFAmilyRec As PordFAmilyType, SalesFamilyRec As PordFAmilyType
If KeyAscii = 13 Then
    If Grid.Visible Then
        OK = False
        txtmaintmodel.Tag = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
        txtmaintmodel.Text = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColName)
        PordFAmilyRec = GetFamNo(txtmaintmodel.Tag, 1)
        TxtMaintFamNo.Tag = PordFAmilyRec.ProdFamNo
        TxtMaintFamNo.Text = PordFAmilyRec.ProdFAmName
        vRow = GetRow(Val(txtmaintmodel.Tag))
        With FlexGrid
            .TopRow = vRow
        End With
        OK = True
    ElseIf Grid.Visible = False And txtmaintmodel.Text <> "" And Val(txtmaintmodel.Tag) <> 0 Then
        TxtMaintFamNo.SelStart = 0
        TxtMaintFamNo.SelLength = Len(TxtMaintFamNo.Text)
        TxtMaintFamNo.SetFocus
        Exit Sub
    Else
        OK = False
        txtmaintmodel.Tag = 0
        txtmaintmodel.Text = ""
        TxtMaintFamNo.Tag = 0
        TxtMaintFamNo.Text = ""
        OK = True
    End If
Grid.Visible = False
TxtMaintFamNo.SelStart = 0
TxtMaintFamNo.SelLength = Len(TxtMaintFamNo.Text)
TxtMaintFamNo.SetFocus
End If
End Sub

Private Sub txtSalesModel_Change()
On Error GoTo errorhandler
Dim RsSearch As New ADODB.Recordset
If txtSalesModel.Text = "" Then
    txtSalesModel.Tag = 0
    Grid.Visible = False
    Exit Sub
End If
If OK Then
    Flag = False
    sqltext = "Select top 10 ModNo , symbol  from hafez2020.dbo.Models Where symbol Like" & LikeExpression(txtSalesModel.Text)
    Set RsSearch = de.con.Execute(sqltext)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        Grid.Row = 0
        FillFormating Grid, 2
        ChangeCursor 2
        Grid.Visible = True
    Else
        txtSalesModel.Tag = 0
        Grid.Visible = False
    End If
    Flag = True
End If
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

Private Sub txtSalesModel_GotFocus()
Pos = 2
End Sub

Private Sub txtSalesModel_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True
End Sub
Sub UpdateModel(MaintModNo As Integer, SalesModNo)
On Error GoTo errorhandler
    sqltext = "Update adhammodels Set SalesModNo=" & SalesModNo & " Where ModNo=" & MaintModNo
    de.con.Execute (sqltext)
    vRow = GetRow(MaintModNo)
    With FlexGrid
        .TextMatrix(vRow, ColSalesModNo) = SalesModNo
'        .TextMatrix(Vrow, ColSalesModName) = GetModelName(SalesModNo)
    End With
    FillGrid
    With FlexGrid
        .TopRow = vRow
    End With
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub
Private Sub txtSalesModel_KeyPress(KeyAscii As Integer)
Dim PordFAmilyRec As PordFAmilyType
If KeyAscii = 13 Then
    If Grid.Visible Then
        OK = False
        txtSalesModel.Tag = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
        txtSalesModel.Text = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColName)
        PordFAmilyRec = GetFamNo(txtSalesModel.Tag, 2)
        LSalesFamNo.Caption = PordFAmilyRec.ProdFAmName
        LSalesFamNo.Tag = PordFAmilyRec.ProdFamNo
        OK = True
    ElseIf Grid.Visible = False And txtSalesModel.Text <> "" And Val(txtSalesModel.Tag) <> 0 Then
        txtmaintmodel.SetFocus
        txtmaintmodel.SelStart = 0
        txtmaintmodel.SelLength = Len(txtmaintmodel.Text)
        Exit Sub
    Else
        OK = False
        txtSalesModel.Tag = 0
        txtSalesModel.Text = ""
        LSalesFamNo.Caption = ""
        LSalesFamNo.Tag = 0
        OK = True
    End If
    Grid.Visible = False
    If Val(txtmaintmodel.Tag) <> 0 And Val(txtSalesModel.Tag) <> 0 Then
        UpdateModel Val(txtmaintmodel.Tag), Val(txtSalesModel.Tag)
    End If
    txtmaintmodel.SetFocus
    txtmaintmodel.SelStart = 0
    txtmaintmodel.SelLength = Len(txtmaintmodel.Text)
End If
End Sub
