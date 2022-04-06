VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmmntMoveSummary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Ê“Ì⁄ «·«⁄„«·"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   16575
   Begin Crystal.CrystalReport cr1 
      Left            =   5100
      Top             =   3300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6615
      Left            =   0
      TabIndex        =   4
      Top             =   1890
      Visible         =   0   'False
      Width           =   3345
      _cx             =   5900
      _cy             =   11668
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
   Begin Threed.SSFrame SSFrame2 
      Height          =   585
      Left            =   60
      TabIndex        =   14
      Top             =   1260
      Width           =   16485
      _ExtentX        =   29078
      _ExtentY        =   1032
      _Version        =   131074
      Begin VB.CommandButton CmdLinkCalls 
         Caption         =   "—»ÿ «·‘ﬂ«ÊÌ «·„⁄·„Â"
         Height          =   345
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   120
         Width           =   2265
      End
      Begin VB.TextBox TxtAssistantEmpNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2370
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   90
         Width           =   3765
      End
      Begin VB.TextBox TxtTeamName 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   11730
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   60
         Width           =   3765
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„—«›ﬁ"
         Height          =   195
         Index           =   0
         Left            =   6180
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«”„ «·Ê—‘Â"
         Height          =   195
         Left            =   15630
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   120
         Width           =   765
      End
   End
   Begin VB.TextBox TxtSearchFormula 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   780
      Width           =   13185
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   13500
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   60
      Width           =   3015
      Begin MSMask.MaskEdBox TxtDate 
         Height          =   375
         Left            =   570
         TabIndex        =   10
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
         Caption         =   " «—ÌÕ «·‘ﬂÊÌ"
         Height          =   195
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   30
         Width           =   990
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   645
      Left            =   0
      TabIndex        =   6
      Top             =   8520
      Width           =   16485
      _ExtentX        =   29078
      _ExtentY        =   1138
      _Version        =   131074
      Begin VB.Label LCount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   14100
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄œœ «·‘ﬂ«ÊÌ"
         Height          =   345
         Left            =   15420
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   120
         Width           =   1005
      End
   End
   Begin VB.TextBox TxtSearch 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   930
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3390
      Visible         =   0   'False
      Width           =   6765
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
            Picture         =   "FrmmntMoveSummary.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmmntMoveSummary.frx":6209F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   2
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
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   14
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   " ÕœÌÀ «·‘ﬂ«ÊÏ"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "2"
                  Text            =   "ÿ»«⁄Â"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   37
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "1"
                  Text            =   "ÿ»«⁄Â «· ›«’Ì·"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "2"
                  Text            =   "«·«Õ’«∆ÌÂ"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   4530
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   0
         Width           =   6825
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   6615
      Left            =   60
      TabIndex        =   1
      Top             =   1890
      Width           =   16485
      _cx             =   29078
      _cy             =   11668
      Appearance      =   1
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "»ÕÀ"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   16260
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   810
      Width           =   285
   End
   Begin Threed.SSCheck ChkAllClaims 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   131074
      ForeColor       =   16711680
      Caption         =   "«·‘ﬂ«ÊÏ «· Ì ·„  Ê“⁄"
   End
   Begin VB.Menu mnu 
      Caption         =   "select"
      Visible         =   0   'False
      Begin VB.Menu mnuselect 
         Caption         =   " ÕœÌœ"
      End
      Begin VB.Menu mnucancel 
         Caption         =   "«·€«¡"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "ÿ»«⁄Â  ›«’Ì· «·Ê“‘Â"
      End
   End
End
Attribute VB_Name = "FrmmntMoveSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ok As Boolean, Flag As Boolean, Pos As Integer
'Dim BankRec As BankEmpType

Const ColChk = 1
Const ColCallNo = 2
Const ColClientName = 3
Const ColClientPhoneNBr = 4


Const ColCallDate = 5

Const colTeamNo = 6
Const ColTeamName = 7
Const ColAssistantEmpNo = 8
Const ColAssistantFullName = 9


Const ColNo = 1
Const ColName = 2
Dim maintDataService_ As New MaintDataService
Dim oldVdate As String




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

Sub FillFormating(ByVal i As Integer, FlexGrid As VSFlexGrid)
If i = 1 Then
   
    fs = "|>" + "«·—ﬁ„"
    fs = fs + "|>" + "«·≈”„"

    With FlexGrid
        .FormatString = fs
        .Cols = 3

        SetColWidths ColNo, FlexGrid
        SetColWidths ColName, FlexGrid

    End With
ElseIf i = 2 Then
    fs = "|>" + "chk"
    fs = fs + "|>" + "—ﬁ„ «·‘ﬂÊÏ"
    fs = fs + "|>" + "≈”„ «·“»Ê‰"
    fs = fs + "|>" + "—ﬁ„ «·Â« ›"
    fs = fs + "|>" + " «—ÌŒ «·‘ﬂÊÏ"
'    Fs = Fs + "|>" + "«· ÊﬁÌ "
    fs = fs + "|>" + "TeamNo."
    fs = fs + "|>" + "≈”„ «·Ê—‘Â"
    fs = fs + "|>" + "—ﬁ„ «·„—«›ﬁ"
    fs = fs + "|>" + "≈”„ «·„—«›ﬁ"
    With FlexGrid
        .FormatString = fs
        .Cols = 10
        .ColWidth(ColChk) = 500
        .ColDataType(ColChk) = flexDTBoolean
        SetColWidths ColCallNo, FlexGrid
        SetColWidths ColClientName, FlexGrid
        SetColWidths ColClientPhoneNBr, FlexGrid
        
        SetColWidths ColCallDate, FlexGrid
        'SetColWidths ColCallDate, FlexGrid
        .ColWidth(colTeamNo) = 0
        SetColWidths ColTeamName, FlexGrid '= 6000 ', FlexGrid
        SetColWidths ColAssistantEmpNo, FlexGrid
        SetColWidths ColAssistantFullName, FlexGrid '= 6000 ', FlexGrid
    End With
End If
End Sub

'Sub SetColWidths(ByVal ColNo As Integer, FlexGrid As VSFlexGrid)
'    With FlexGrid
'        .AutoSize (ColNo)
'    End With
'End Sub

Sub ChangeCursor(sender As Control, Optional top As Variant = 0, Optional left As Variant = 0)

    With sender
       Grid.top = top + .top + .Height
       Grid.left = left + .left
       Grid.Width = .Width
    End With

End Sub

Sub FillGrid(vdate As String, searchFormula As String, allClaims As Boolean)
On Error GoTo ErrorHandler
    sqlText = "select  CallChk , CallNo, adhamname ,  adhamphon , Convert(varchar(10),CallDateTime,103) CallDateTime ,  TeamNo, TeamName, AssistantEmpNo, AssistantFullName from DistributeCliamsQry "
    sqlText = sqlText & " where  CallNo <> -1"
    sqlText = sqlText & " and  convert(varchar(10),CallDateTime,101) ='" & IIf(IsDate(vdate), ConvertControlDate(vdate), ConvertControlDate(Format(Date, "dd/mm/yyyy"))) & "'"
    
    If Not allClaims Then
        sqlText = sqlText & " and (teamNo is null or teamno=0)"
    End If
    
    If searchFormula <> "" Then
        sqlText = sqlText & " and (convert(varchar(10),CallNo) like " & LikeExpression(searchFormula) & " or convert(varchar(10),LeaderEmpNo)like" & LikeExpression(searchFormula) & " or LeaderFullName like " & LikeExpression(searchFormula) & " or convert(varchar(10),AssistantEmpNo) like " & LikeExpression(searchFormula) & " or convert(varchar(10),AssistantFullName) like" & LikeExpression(searchFormula) & "or TeamName like" & LikeExpression(searchFormula) & "or adhamname like " & LikeExpression(searchFormula) & ")"
    End If
    sqlText = sqlText & " order by CallNo"
    
    Set rs = de.con.Execute(sqlText)
    
    Set FlexGrid.DataSource = rs
    LCount.Caption = FlexGrid.Rows - 1
    FillFormating 2, FlexGrid
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Sub init()
    top = 0
    left = 0
    TxtDate.Text = Format(Date, "dd/mm/yyyy")
    Ok = True
    
    If GetCallsWithoutLink Then
        FillGrid TxtDate.Text, "", False
    End If
    Grid.Rows = 1
    FlexGrid.Editable = flexEDKbdMouse
    'Grid.SelectionMode = flexSelectionListBox
End Sub

Private Sub ChkAllClaims_Click(Value As Integer)
If Abs(ChkAllClaims.Value) = ssCBChecked Then
    ChkAllClaims.Caption = "ﬂ«›Â «·‘ﬂ«ÊÌ"
Else
    ChkAllClaims.Caption = "«·‘ﬂ«ÊÏ «· Ì ·„  Ê“⁄"
End If
FillGrid TxtDate.Text, TxtSearchFormula.Text, ChkAllClaims.Value
End Sub

Private Sub CmdLinkCalls_Click()
On Error GoTo ErrorHandler

If Val(TxtTeamName.Tag) = 0 Then Exit Sub

sqlText = "Update MntMoveSummary set TeamNo=" & TxtTeamName.Tag

If Val(TxtAssistantEmpNo.Tag) <> 0 Then

    sqlText = sqlText & ",attendantNo=" & TxtAssistantEmpNo.Tag
Else
    sqlText = sqlText & ",attendantNo=Null"
End If

sqlText = sqlText & " where Callchk=1 and date='" & ConvertControlDate(TxtDate.Text) & "'"
sqlText = sqlText & "; Update MntMoveSummary Set Callchk=0  where Callchk=1 and TeamNo=" & TxtTeamName.Tag & " and date='" & ConvertControlDate(TxtDate.Text) & "'"
de.con.Execute (sqlText)

FillGrid TxtDate.Text, "", False
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrorHandler
Dim form_ As Form
Set form_ = Screen.ActiveForm
With FlexGrid
    If Col = ColChk Then
        sqlText = "Update MntMoveSummary set CallChk=" & .TextMatrix(Row, ColChk) & " where callno=" & .TextMatrix(Row, ColCallNo)
        de.con.Execute (sqlText)
    ElseIf Col = ColCallDate And Not form_ Is Nothing Then
        
        If form_.Name = Me.Name And MsgBox("Â· «‰  „ √ﬂœ „‰ «· €œÌ·", vbQuestion + vbYesNo + vbDefaultButton2, " ⁄œÌ·  «—ÌŒ «· Ê“Ì⁄ Ê «·‘ﬂÊÏ") = vbYes Then
            UpdateDate .TextMatrix(Row, ColCallDate), .TextMatrix(Row, ColCallNo)
        Else
            .TextMatrix(Row, ColCallDate) = oldVdate
        End If
    End If
End With
Exit Sub
ErrorHandler:
FlexGrid.TextMatrix(Row, ColCallDate) = oldVdate
MsgBox Err.Description
End Sub

Private Sub FlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)

If Col = ColCallDate Or Col = ColChk Or Col = ColClientPhoneNBr Then
    oldVdate = FlexGrid.TextMatrix(FlexGrid.Row, ColCallDate)
    Exit Sub
End If
cancel = True

End Sub

Sub UpdateDate(vdate As String, CallNo As Long)
On Error GoTo ErrorHandler
    maintDataService_.UpdateCallDateAndMntMoveSummaryDate vdate, CallNo
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub FlexGrid_DblClick()
On Error GoTo ErrorHandler

If Gettag(empNo, 39) Then
    Dim FrmMaintCallNew As New FrmMaintCallNew
    With FlexGrid
        idCallNo = .TextMatrix(.Row, ColCallNo)
        LoadForm = True
        FrmMaintCallNew.Show
    End With
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub flexGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    With FlexGrid
    Ok = False
    Select Case .Col
        Case ColTeamName
                ShowEditor FlexGrid, ColTeamName, colTeamNo
        Case ColAssistantFullName
            ShowEditor FlexGrid, ColAssistantFullName, ColAssistantEmpNo
    End Select

        End With
End If
End Sub
Sub ShowEditor(FlexGrid As VSFlexGrid, colText As Integer, ColId As Integer)
With FlexGrid
    Ok = False
    If .Rows > 23 Then
        TxtSearch.Move .left + .CellLeft + 220, .top + .CellTop, .CellWidth, .CellHeight
    Else
        TxtSearch.Move .left + .CellLeft, .top + .CellTop, .CellWidth, .CellHeight
    End If
    TxtSearch.Tag = .TextMatrix(.Row, ColId)
     TxtSearch.Text = .TextMatrix(.Row, colText)
    TxtSearch.Visible = True
    TxtSearch.SelStart = 0
    TxtSearch.SelLength = Len(TxtSearch.Text)
    TxtSearch.SetFocus
    Ok = True
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



Private Sub Grid_RowColChange()
If Grid.Row = 0 Then Exit Sub
If Flag Then
    Ok = False
    With Grid
    
            ActiveControl.Tag = .TextMatrix(.Row, ColNo)
            ActiveControl.Text = .TextMatrix(.Row, ColName)
            
    End With
    Ok = True
End If
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
                sqlText = "Update MntMoveSummary set CallChk=" & IIf(isChk, 1, 0) & " where callno=" & .TextMatrix(i, ColCallNo)
                de.con.Execute (sqlText)
        Next i
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub mnuCancel_Click()
UpdateRow FlexGrid, False
End Sub

Private Sub mnuPrint_Click()

With FlexGrid
PrintRep .TextMatrix(.Row, ColCallDate), 1, .TextMatrix(.Row, colTeamNo)

End With
End Sub

Private Sub mnuSelect_Click()

UpdateRow FlexGrid, True


End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
    If GetCallsWithoutLink Then
        FillGrid TxtDate.Text, "", ChkAllClaims.Value
    End If
    Case 2
    
        PrintRep IIf(IsDate(TxtDate.Text), TxtDate.Text, Format(Now, "dd/mm/yyyy")), 1
    Case 4
        Unload Me
End Select
End Sub
Sub PrintRep(vdate As String, Vindex As Integer, Optional TeamNo As Variant)
On Error GoTo ErrorHandler

With cr1
    .Connect = ConnectName("")
    If Vindex = 1 Then
        .ReportFileName = App.Path + "\Reports\RepmntMoveSummary.rpt"
            .SQLQuery = "select CallNo, CallDateTime, ProdFamName, adhamphon, adhamname, ZoneName,"
            .SQLQuery = .SQLQuery & " CallReceiver, TeamNo, TeamName From dbo.MaintCallsDetailsQry"
            .SQLQuery = .SQLQuery & " Where convert(varchar(10),calldatetime,101) = '" & ConvertControlDate(vdate) & "'"
        If Not IsMissing(TeamNo) Then
            .SQLQuery = .SQLQuery & " and isnull(TeamNo,0) = " & Val(TeamNo)
        End If
        .SQLQuery = .SQLQuery & " order by teamno , callno"
    ElseIf Vindex = 2 Then
    .ReportFileName = App.Path + "\Reports\RepDistributeTeams.rpt"
    
    .SQLQuery = "select TeamNo , teamname ,  attendantNo  , AssistantName ,  date  , countRec  from MaintTeamDistributeStatisticsQry where TeamNo is not null and date='" & ConvertControlDate(vdate) & "' order by countrec desc"
    End If
    .DiscardSavedData = True
    .WindowState = crptMaximized
    
    .Action = 1
End With


Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Function GetCallsWithoutLink() As Boolean
 On Error GoTo ErrorHandler

sqlText = "insert into MntMoveSummary(CallNo , Date)" _
            & " select  m2.CallNo ," _
            & " convert(varchar(10), CallDatetime, 101)" _
            & " from   maintcall m2 left outer join MntMoveSummary m3 on " _
            & " m2.callNo = m3.callNo " _
            & " where m3.callno is null"
de.con.Execute (sqlText)
GetCallsWithoutLink = True
  
Exit Function
ErrorHandler:
MsgBox Err.Description
End Function



Function SaveRec(CallNo As Double, TeamNo As Integer, AssistantEmpNo As Integer)
On Error GoTo ErrorHandler
    sqlText = "Sp_distributeClaims " & CallNo & "," & TeamNo & "," & AssistantEmpNo
    de.con.Execute (sqlText)
SaveRec = True
Exit Function
ErrorHandler:
SaveRec = False
MsgBox Err.Description
End Function
Function GetRow(TeamNo As Integer) As Integer
On Error GoTo ErrorHandler

With FlexGrid
    For i = 1 To .Rows - 1
        If (.TextMatrix(i, colTeamNo) = TeamNo) Then
            GetRow = i
            Exit Function
        End If
    Next i
GetRow = -1
End With
Exit Function
ErrorHandler:
GetRow = -1
MsgBox Err.Description
End Function
'Sub AddToGrid(leaderEmpNo As Integer)
'On Error GoTo errorhandler
'Dim Vrow As Integer
'    sqltext = "select m1.TeamNo , m1.EmpNo  , e1.FirstName + ' ' + e1.LastName as LeaderFullName , assistantempno , e2.FirstName + ' ' + e2.LastName as AssistantFullName  from MaintTeam m1 inner join employee e1 on m1.Empno = e1.EmpNo left outer join employee e2 on m1.assistantempno = e2.empno Where m1.EmpNo= " & leaderEmpNo
'    Set Rs = de.con.Execute(sqltext)
'
'
'    With FlexGrid
'        Vrow = GetRow(Rs!teamNo)
'        If Vrow = -1 Then
'            .AddItem ""
'            Vrow = .Rows - 1
'        End If
'
'        .TextMatrix(Vrow, Colteamno) = Rs!teamNo
'        .TextMatrix(Vrow, ColLeaderEmpNo) = Rs!EmpNo
'        .TextMatrix(Vrow, ColTeamName) = Rs!LeaderFullName
'        .TextMatrix(Vrow, ColAssistantEmpNo) = IIf(IsNull(Rs!assistantEmpNo), "Null", Rs!assistantEmpNo)
'        .TextMatrix(Vrow, ColAssistantFullName) = IIf(IsNull(Rs!AssistantFullName), "", Rs!AssistantFullName)
'        FillFormating 2, FlexGrid
'        If Not .RowIsVisible(Vrow) Then
'            .TopRow = Vrow
'
'        End If
'        .Row = Vrow
'        .Col = 0
'        .ColSel = .Cols - 1
'    End With
'
'Exit Sub
'errorhandler:
'MsgBox Err.Description
'End Sub



Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Tag
    Case 1
        PrintRep IIf(IsDate(TxtDate.Text), TxtDate.Text, Format(Now, "dd/mm/yyyy")), 1
    Case 2
        PrintRep IIf(IsDate(TxtDate.Text), TxtDate.Text, Format(Now, "dd/mm/yyyy")), 2
End Select
End Sub

Private Sub TxtAssistantEmpNo_Change()
Search TxtAssistantEmpNo, "Employee", "FirstName + ' ' + LastName", "EmpNo", True, SSFrame2.top, SSFrame2.left
End Sub

Private Sub TxtAssistantEmpNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid

End Sub

Private Sub TxtAssistantEmpNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible Then
        Ok = False
        TxtAssistantEmpNo.Tag = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColNo), Grid.TextMatrix(Grid.Row, ColNo))
        TxtAssistantEmpNo.Text = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColName), Grid.TextMatrix(Grid.Row, ColName))
        Ok = True
        Grid.Visible = False
    End If
    CmdLinkCalls.SetFocus
    SendKeys "{Home}+{End}"
End If
End Sub

Private Sub TxtDate_Change()
FillGrid TxtDate.Text, TxtSearchFormula.Text, ChkAllClaims.Value
End Sub

Private Sub TxtSearch_Change()
With FlexGrid
Select Case .Col
    Case ColTeamName
        Search TxtSearch, "MaintTeam", "TeamName", "TeamNo"
    Case ColAssistantFullName
         Search TxtSearch, "Employee", "FirstName + ' ' + LastName", "EmpNo"
End Select
End With
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
End Sub

Sub Search(sender As Control, tableName As String, listField As String, dataMember As String, Optional isChangeCursor = False, Optional top As Variant, Optional left As Variant)
On Error GoTo ErrorHandler
Dim RsSearch As New ADODB.Recordset


If Ok Then
    sender.Tag = 0
    If sender.Text = "" Then
        Grid.Visible = False
        Exit Sub
    End If
    Flag = False
    sqlText = "Select top 30 " & dataMember & "," & listField & " From " & tableName & " Where "
    sqlText = sqlText & dataMember & " Like " & LikeExpression(sender.Text) & " Or " & listField & " Like " & LikeExpression(sender.Text)
    Set RsSearch = de.con.Execute(sqlText)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        FillFormating 1, Grid
        If isChangeCursor Then ChangeCursor sender, top, left
        Grid.Visible = True
    Else
        sender.Tag = 0
        Grid.Visible = False
    End If
    Flag = True
End If

         
   Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub


Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible Then
            With FlexGrid
                If .Col = ColTeamName Then
                    .TextMatrix(.Row, colTeamNo) = IIf(Val(TxtSearch.Tag) = 0, Grid.TextMatrix(1, ColNo), TxtSearch.Tag)
                    .TextMatrix(.Row, .Col) = IIf(Val(TxtSearch.Tag) = 0, Grid.TextMatrix(1, ColName), TxtSearch.Text)
                ElseIf .Col = ColAssistantFullName Or .Col = ColAssistantEmpNo Then
                    .TextMatrix(.Row, ColAssistantEmpNo) = IIf(Val(TxtSearch.Tag) = 0, Grid.TextMatrix(1, ColNo), TxtSearch.Tag)
                    .TextMatrix(.Row, .Col) = IIf(Val(TxtSearch.Tag) = 0, Grid.TextMatrix(1, ColName), TxtSearch.Text)
                End If
            SetColWidths .Col, FlexGrid
            End With
        TxtSearch.Visible = False
        Grid.Visible = False
   Else
        If Val(TxtSearch.Tag) = 0 Then
            ClearRow FlexGrid
        End If

         TxtSearch.Visible = False
    End If
With FlexGrid
    SaveRec .TextMatrix(.Row, ColCallNo), Val(.TextMatrix(.Row, colTeamNo)), Val(.TextMatrix(.Row, ColAssistantEmpNo))
End With

FlexGrid.SetFocus
End If
End Sub

Sub ClearRow(FlexGrid As VSFlexGrid)
With FlexGrid
    Select Case FlexGrid.Col
        Case ColAssistantFullName, ColAssistantEmpNo
                
                 .TextMatrix(.Row, ColAssistantEmpNo) = ""
                 .TextMatrix(.Row, ColAssistantFullName) = ""

                 
        Case ColTeamName
           
            .TextMatrix(.Row, colTeamNo) = ""
            .TextMatrix(.Row, .Col) = ""
            .TextMatrix(.Row, ColAssistantEmpNo) = ""
            .TextMatrix(.Row, ColAssistantFullName) = ""
                 
    End Select
End With
End Sub
Function GetAssistantEmpFullName(AssistantEmpNo As Integer)
On Error GoTo ErrorHandler
Dim rs As Recordset
sqlText = "Select firstName + ' ' + LastName as AssistantFullName from Employee  Where Empno=" & AssistantEmpNo
Set rs = de.con.Execute(sqlText)
If rs.RecordCount > 0 Then
    GetAssistantEmpFullName = rs!AssistantFullName
    Exit Function
End If
GetAssistantEmpFullName = ""
Exit Function
ErrorHandler:
GetAssistantEmpFullName = ""
MsgBox Err.Description
End Function


Private Sub TxtSearchFormula_Change()
FillGrid TxtDate.Text, TxtSearchFormula.Text, ChkAllClaims.Value
End Sub

Private Sub TxtTeamName_Change()
Search TxtTeamName, "MaintTeam", "TeamName", "TeamNo", True, SSFrame2.top, SSFrame2.left

End Sub

Private Sub TxtTeamName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid

End Sub

Private Sub TxtTeamName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible Then
        Ok = False
        TxtTeamName.Tag = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColNo), Grid.TextMatrix(Grid.Row, ColNo))
        TxtTeamName.Text = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColName), Grid.TextMatrix(Grid.Row, ColName))
        Ok = True
        Grid.Visible = False
    End If
    TxtAssistantEmpNo.SetFocus
    SendKeys "{Home}+{End}"
End If
End Sub

