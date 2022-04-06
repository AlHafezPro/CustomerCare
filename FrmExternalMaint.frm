VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmExternalMaint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈Õ’«∆Ì… ⁄‰«ÊÌ‰ «·’Ì«‰… «·Œ«—ÃÌ…"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4575
   ScaleWidth      =   9285
   Begin Crystal.CrystalReport cr1 
      Left            =   1710
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   2895
      Left            =   30
      TabIndex        =   2
      Top             =   1620
      Width           =   9225
      _cx             =   16272
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
   Begin MSMask.MaskEdBox TxtTillDate 
      Height          =   435
      Left            =   6900
      TabIndex        =   1
      Top             =   1110
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   767
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtFromDate 
      Height          =   435
      Left            =   8100
      TabIndex        =   0
      Top             =   1110
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   767
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   150
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
            Picture         =   "FrmExternalMaint.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExternalMaint.frx":6209F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9285
      _ExtentX        =   16378
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
            ImageIndex      =   15
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "1"
                  Text            =   "Õ”» «·Ê—‘…"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "2"
                  Text            =   "Õ”» «·⁄«∆·…"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   37
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "3"
                  Text            =   "Õ”» «·Ê—‘…"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "4"
                  Text            =   "Õ”» «·⁄«∆·…"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   9300
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   90
         Width           =   2325
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "≈·Ï  «—ÌŒ"
      Height          =   195
      Index           =   0
      Left            =   7440
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰  «—ÌŒ"
      Height          =   195
      Left            =   8670
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   600
   End
End
Attribute VB_Name = "FrmExternalMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Colteamno = 1
Const ColTeamName = 2
Const colCash = 3
Const ColCredit = 4
Const ColGarantee = 5
Const ColCount = 6

Const colProdFamNo = 1
Const ColProdFamName = 2
Const ColCash_1 = 3
Const ColCredit_1 = 4
Const ColGarantee_1 = 5
Const ColCount_1 = 6

Sub FillFormating(ByVal i As Integer)
If i = 1 Then
    Fs = "|>" + "TeamNo"
    Fs = Fs + "|>" + "≈”„ «·Ê—‘…"
    Fs = Fs + "|>" + "‰ﬁœÌ"
    Fs = Fs + "|>" + "–„…"
    Fs = Fs + "|>" + "÷„‰ «·ﬂ›«·…"
    Fs = Fs + "|>" + "«·⁄œœ"
    With FlexGrid
        .FormatString = Fs
        .Cols = 7
        .ColWidth(Colteamno) = 0
        SetColWidths ColTeamName, FlexGrid
        SetColWidths colCash, FlexGrid
        SetColWidths ColCredit, FlexGrid
        SetColWidths ColGarantee, FlexGrid
        SetColWidths ColCount, FlexGrid
    End With
ElseIf i = 2 Then
    Fs = "|>" + "ProdFAmNo"
    Fs = Fs + "|>" + "≈”„ «·⁄«∆·…"
    Fs = Fs + "|>" + "‰ﬁœÌ"
    Fs = Fs + "|>" + "–„…"
    Fs = Fs + "|>" + "÷„‰ «·ﬂ›«·…"
    Fs = Fs + "|>" + "«·⁄œœ"
    With FlexGrid
        .FormatString = Fs
        .Cols = 7
        .ColWidth(colProdFamNo) = 0
        SetColWidths ColProdFamName, FlexGrid
        SetColWidths ColCash_1, FlexGrid
        SetColWidths ColCredit_1, FlexGrid
        SetColWidths ColGarantee_1, FlexGrid
        SetColWidths ColCount_1, FlexGrid
    End With
End If
End Sub

Sub SetColWidths(ByVal ColNo As Integer, FlexGrid As VSFlexGrid)
    With FlexGrid
        .AutoSize (ColNo)
    End With
End Sub

Function ExecProcedure(Vindex As Integer) As Boolean
    On Error GoTo errorhandler
    
    Dim rs As New ADODB.Recordset
    
Select Case Vindex
    Case 1 'Õ”» «·Ê—‘…
        Sqltext = "Exec sp_ExternalMaint '" & ConvertControlDate(TxtFromDate.Text) & "','" & ConvertControlDate(TxtTillDate.Text) & "'"
        de.con.Execute (Sqltext)
        

    Case 2 'Õ”» «·⁄«∆·…
        Sqltext = "Exec sp_ExternalMaintByFAmily '" & ConvertControlDate(TxtFromDate.Text) & "','" & ConvertControlDate(TxtTillDate.Text) & "'"
        de.con.Execute (Sqltext)

End Select
ExecProcedure = True
    Exit Function
errorhandler:
   MsgBox Err.Description
ExecProcedure = False
End Function

Sub printRep(Vindex As Integer)
On Error GoTo errorhandler
With cr1
    .Connect = ConnectName("")
    .Formulas(0) = "FromDate='" & TxtFromDate.Text & "'"
    .Formulas(1) = "TillDate='" & TxtTillDate.Text & "'"
   
    Select Case Vindex
        Case 1
            
            .ReportFileName = App.Path + "\Reports\REpExternalMaint.rpt"
        Case 2
            .ReportFileName = App.Path + "\Reports\REpExternalMaintByFAmily.rpt"
    End Select
    .DiscardSavedData = True
    .WindowState = crptMaximized
    .Action = 1
End With
Exit Sub
errorhandler:
   MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Top = 0
    Left = 0
    FlexGrid.Rows = 1
    FillFormating 1
    TxtFromDate.Text = "01/01/" & LTrim(RTrim(Str(Year(Date))))
    TxtTillDate.Text = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim rs As New ADODB.Recordset
Select Case Button.Index
    Case 1
        
         If ExecProcedure(1) Then
                Sqltext = "Select teamno , teamname  , Cash0 , Cash1 , Cash2 , countRec From t_ExternalMaint"
                Set rs = de.con.Execute(Sqltext)
                Set FlexGrid.DataSource = rs
                FillFormating 1
         End If
    Case 3
        If ExecProcedure(1) Then
            printRep (1)
        End If
    Case 5
        Unload Me
End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Tag
    Case 1
    If ExecProcedure(1) Then
            Sqltext = "Select teamno , teamname  , Cash0 , Cash1 , Cash2 , countRec From t_ExternalMaint"
            Set rs = de.con.Execute(Sqltext)
            Set FlexGrid.DataSource = rs
            FillFormating 1
    End If
      
    Case 2
          If ExecProcedure(2) Then
            Sqltext = "Select Prodfamno , prodfamnamea , Cash0 , Cash1 , Cash2 , countRec From t_ExternalMaintByFAmily"
            Set rs = de.con.Execute(Sqltext)
            Set FlexGrid.DataSource = rs
            FillFormating 2
        End If
    Case 3
        If ExecProcedure(1) Then
            printRep 1
        End If
    Case 4
        If ExecProcedure(2) Then
            printRep 2
        End If
End Select
End Sub
