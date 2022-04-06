VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmTransferProvinces 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»ÕÀ"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   13770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   13770
   Begin VB.CommandButton CmdSearch 
      Caption         =   "»ÕÀ"
      Height          =   345
      Left            =   8610
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1050
      Width           =   1185
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   495
      Left            =   60
      TabIndex        =   10
      Top             =   6750
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   873
      _Version        =   131074
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„Ã„Ê⁄ «·ﬂ·Ì"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   5
         Left            =   12600
         TabIndex        =   22
         Top             =   60
         Width           =   975
      End
      Begin VB.Label LSumTotal 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   10020
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   60
         Width           =   2565
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„Ã„Ê⁄ «·„Õœœ"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   8
         Left            =   5280
         TabIndex        =   20
         Top             =   30
         Width           =   1005
      End
      Begin VB.Label LSumSelected 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2670
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   30
         Width           =   2565
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·⁄œœ «·ﬂ·Ì"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   10
         Left            =   8820
         TabIndex        =   18
         Top             =   60
         Width           =   765
      End
      Begin VB.Label LCountTotal 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   90
         Width           =   675
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·⁄œœ «·„Õœœ"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   11
         Left            =   1770
         TabIndex        =   16
         Top             =   60
         Width           =   795
      End
      Begin VB.Label LCountSelected 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   930
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   30
         Width           =   675
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexAccount 
      Height          =   5265
      Left            =   90
      TabIndex        =   4
      Top             =   1470
      Width           =   13635
      _cx             =   24051
      _cy             =   9287
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   405
      Left            =   60
      TabIndex        =   3
      Top             =   1050
      Visible         =   0   'False
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
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
            Picture         =   "FrmTransferProvinces.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferProvinces.frx":6209F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   1680
      Top             =   6780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   13770
      _ExtentX        =   24289
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
            Object.ToolTipText     =   "«” ⁄—«÷ «·„·›« "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "‰—ÕÌ· «·»Ì«‰«  ≈·Ï „Õ«”»Â «·’«·« "
            ImageIndex      =   14
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin Threed.SSFrame SSFrame2 
         Height          =   555
         Left            =   6600
         TabIndex        =   12
         Top             =   60
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   979
         _Version        =   131074
         Begin MSDataListLib.DataCombo ComboCompany 
            Height          =   360
            Left            =   90
            TabIndex        =   13
            Top             =   120
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSMask.MaskEdBox txttransferDate 
            Height          =   345
            Left            =   4830
            TabIndex        =   24
            Top             =   120
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   " «—ÌŒ «· —ÕÌ·"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   6045
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   90
            Width           =   915
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Õœœ «·‘—ﬂÂ"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   3270
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   90
            Width           =   780
         End
      End
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
   Begin MSDataListLib.DataCombo ComboPayment 
      Height          =   360
      Left            =   9810
      TabIndex        =   2
      Top             =   1050
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox TxtFromDate 
      Height          =   345
      Left            =   12600
      TabIndex        =   0
      Top             =   1080
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtTillDate 
      Height          =   345
      Left            =   11460
      TabIndex        =   1
      Top             =   1080
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ"
      Height          =   195
      Index           =   2
      Left            =   13260
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ"
      Height          =   195
      Index           =   3
      Left            =   12240
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   10590
      TabIndex        =   7
      Top             =   810
      Width           =   795
   End
   Begin VB.Menu mnuRight 
      Caption         =   "rightmouse"
      Visible         =   0   'False
      Begin VB.Menu mnuselect 
         Caption         =   " ÕœÌœ"
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "≈·€«¡"
      End
   End
End
Attribute VB_Name = "FrmTransferProvinces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ColCallNo = 1
Const ColRepNo = 2
Const ColPaymentTypeNo = 3
Const ColPaymentTypeName = 4
Const colRepDate = 5
Const ColDebAccNoTemp = 6
Const ColCreAccNoTemp = 7
Const ColStkNo = 8
Const ColStkName = 9

Const ColQty = 10
Const ColPrice = 11
Const ColTTlPrice = 12
Const colDescription = 13




Sub FillFormating(ByVal i As Integer, FlexGrid As VSFlexGrid)

If i = 1 Then
    fs = "|>" + "—ﬁ„ «·‘ﬂÊÏ"
    fs = fs + "|>" + "—ﬁ„ «·«’·«Õ"
    fs = fs + "|>" + "—ﬁ„ ÿ—Ìﬁ… «·œ›⁄"
    fs = fs + "|>" + "ÿ—Ìﬁ… «·œ›⁄"
    fs = fs + "|>" + " «—ÌÃ «·«’·«Õ"
    fs = fs + "|>" + "Õ”«» «·„œÌ‰"
    fs = fs + "|>" + "Õ”«» «·œ«∆‰"
    fs = fs + "|>" + "«·—ﬁ„ «·„Œ“‰Ì"
    fs = fs + "|>" + "«·‘—Õ"
    fs = fs + "|>" + "«·ﬂ„Ì…"
    fs = fs + "|>" + "«·”⁄—"
    fs = fs + "|>" + "«·≈Ã„«·Ì"
    fs = fs + "|>" + "«·‘—Õ"
    
    With FlexGrid
        .FormatString = fs
        .Cols = 14
        SetColWidths ColCallNo, FlexGrid
        SetColWidths ColRepNo, FlexGrid
        .ColWidth(ColPaymentTypeNo) = 0
        SetColWidths ColPaymentTypeName, FlexGrid
        SetColWidths colRepDate, FlexGrid
        SetColWidths ColDebAccNoTemp, FlexGrid
        SetColWidths ColCreAccNoTemp, FlexGrid
        SetColWidths ColStkNo, FlexGrid
        SetColWidths ColStkName, FlexGrid
        SetColWidths ColQty, FlexGrid
        SetColWidths ColPrice, FlexGrid
        SetColWidths ColTTlPrice, FlexGrid
        SetColWidths colDescription, FlexGrid
        

        
End With

End If
End Sub

Sub SetColWidths(ByVal ColNo As Integer, FlexGrid As VSFlexGrid)
    With FlexGrid
        .AutoSize (ColNo)
    End With
End Sub
Function GetDebAccNo(stkno As String, PaymentTYpeNo As Integer, CompNo As Integer) As String
On Error GoTo ErrorHandler
GetDebAccNo = ""

Select Case PaymentTYpeNo
    Case 0:
        GetDebAccNo = "1812/00001/01/" & Right("00" + Trim(Str(CompNo)), 2)
    Case 2:
        GetDebAccNo = "2614/00002/01/" & Right("00" + Trim(Str(CompNo)), 2)
End Select

Exit Function
ErrorHandler:
GetDebAccNo = ""
End Function

Function GetCreAccNo(stkno As String, PaymentTYpeNo As Integer, CompNo As Integer) As String
On Error GoTo ErrorHandler
GetCreAccNo = ""

Select Case PaymentTYpeNo
    Case 0:
        GetCreAccNo = "4601/" + GetAccountNo(stkno) + "/01/" & Right("00" + Trim(Str(CompNo)), 2)
    Case 2:
        GetCreAccNo = "4704/" + GetAccountNo(stkno) + "/01/" & Right("00" + Trim(Str(CompNo)), 2)
End Select

Exit Function
ErrorHandler:
GetCreAccNo = ""
End Function

Sub FillGrid(rs As ADODB.Recordset, CompNo As Integer)
On Error GoTo ErrorHandler
If rs.RecordCount = 0 Then Exit Sub
FlexAccount.Rows = 1
rs.MoveFirst
While Not rs.EOF
    With FlexAccount
        .AddItem ""
        .TextMatrix(.Rows - 1, ColCallNo) = rs!CallNo
        .TextMatrix(.Rows - 1, ColRepNo) = rs!repNo
        .TextMatrix(.Rows - 1, ColPaymentTypeNo) = rs!No
        .TextMatrix(.Rows - 1, ColPaymentTypeName) = rs!name
        .TextMatrix(.Rows - 1, colRepDate) = rs!RepDate
        .TextMatrix(.Rows - 1, ColDebAccNoTemp) = GetDebAccNo(rs!PieceNo, rs!No, CompNo)
        .TextMatrix(.Rows - 1, ColCreAccNoTemp) = GetCreAccNo(rs!PieceNo, rs!No, CompNo)
        .TextMatrix(.Rows - 1, ColStkNo) = rs!PieceNo
        .TextMatrix(.Rows - 1, ColStkName) = GetStkName(rs!PieceNo)
        .TextMatrix(.Rows - 1, ColQty) = rs!Qty
        .TextMatrix(.Rows - 1, ColPrice) = rs!Price
        .TextMatrix(.Rows - 1, ColTTlPrice) = IIf(IsNull(rs!Qty), 0, rs!Qty) * IIf(IsNull(rs!Price), 0, rs!Price)
        .TextMatrix(.Rows - 1, colDescription) = "—ﬁ„ «·«’·«Õ " + Str(rs!repNo) + " «· «»⁄ ··‘ﬂÊÏ" + Str(rs!CallNo)
    End With
    
    rs.MoveNext
Wend


Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Function GetAccountNo(stkno As String) As String
On Error GoTo ErrorHandler
GetAccountNo = ""
Dim rs As New ADODB.Recordset
sqlText = "Select AccNo From CoStock Where StkNo='" & stkno & "'"
Set rs = de.con.Execute(sqlText)

If rs.RecordCount > 0 Then
    GetAccountNo = rs!AccNo & ""
End If
Exit Function
ErrorHandler:
GetAccountNo = ""
Exit Function
End Function

Function GetAccNo(stkno As String) As String
On Error GoTo ErrorHandler
GetAccNo = ""
Dim rs As New ADODB.Recordset
sqlText = "Select AccNo From CoStock Where StkNo='" & stkno & "'"
Set rs = de.con.Execute(sqlText)

If rs.RecordCount > 0 Then
    GetAccNo = rs!AccNo & ""
End If
Exit Function
ErrorHandler:
GetAccNo = ""
Exit Function
End Function

Function GetStkName(stkno As String) As String
On Error GoTo ErrorHandler
GetStkName = ""
Dim rs As New ADODB.Recordset
sqlText = "Select StkName From CoStock Where StkNo='" & stkno & "'"
Set rs = de.con.Execute(sqlText)

If rs.RecordCount > 0 Then
    GetStkName = rs!StkName & ""
End If
Exit Function
ErrorHandler:
GetStkName = ""
Exit Function
End Function
Sub FillCombos()
    

    
    If PaymentEmpStr = "" Then Exit Sub
    Dim rsPayment As New ADODB.Recordset
    sqlText = "Select No , Name  From PayMethod Where No in (" & PaymentEmpStr & ")"
    Set rsPayment = de.con.Execute(sqlText)
    Set ComboPayment.RowSource = rsPayment
    ComboPayment.listField = "Name"
    ComboPayment.BoundColumn = "No"
    ComboPayment.BoundText = 0
    
    
    Dim RsCompany As New ADODB.Recordset
    
    sqlText = "Select CompNo , Name From dbo.AccCompany"
    Set RsCompany = de.con.Execute(sqlText)
    Set ComboCompany.RowSource = RsCompany
        ComboCompany.listField = "Name"
    ComboCompany.BoundColumn = "CompNo"
    ComboCompany.BoundText = 1

End Sub


Sub GetData()
On Error GoTo ErrorHandler
Dim rs As New Recordset
If ComboCompany.BoundText = "" Then
MsgBox "·„ Ì „  ÕœÌœ «·‘—ﬂÂ «Ê «·’«·Â", vbExclamation, "attention"
FlexAccount.Rows = 1

Exit Sub
End If
CD.Filter = "*.txt"
CD.ShowOpen

rs.Open CD.FileName, , , , adCmdFile
'Set FlexAccount.DataSource = Rs
FillGrid rs, ComboCompany.BoundText
FillFormating 1, FlexAccount

Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Sub ColorRow(Row As Integer, Color As Long)
With FlexAccount
    For i = 1 To .Cols - 1
        .Col = i
        .Row = Row
        .CellBackColor = Color
    Next
End With
End Sub

Sub init()
top = 0
left = 0
FillCombos

FlexAccount.Rows = 1
FlexAccount.Cols = 13
FillFormating 1, FlexAccount
End Sub

Private Sub CmdSearch_Click()
Dim statisticsType As statisticsType
SearchData TxtFromDate.Text, TxtTillDate.Text, IIf(ComboPayment.Text = "", -1, ComboPayment.BoundText)
FillStatisticsLabels
End Sub
Friend Function Fillstatistics() As statisticsType
On Error GoTo ErrorHandler

Dim statisticsRec   As statisticsType

statisticsRec.SumSelectedValues = 0
statisticsRec.CountSelectedValues = 0
statisticsRec.SumUnSelectedValues = 0
statisticsRec.CountunSelectedValues = 0


With FlexAccount
    For i = 1 To .Rows - 1
         If .RowData(i) Then
            statisticsRec.SumSelectedValues = statisticsRec.SumSelectedValues + .TextMatrix(i, ColTTlPrice)
            statisticsRec.CountSelectedValues = statisticsRec.CountSelectedValues + 1
         Else
            statisticsRec.SumUnSelectedValues = statisticsRec.SumUnSelectedValues + .TextMatrix(i, ColTTlPrice)
            statisticsRec.CountunSelectedValues = statisticsRec.CountunSelectedValues + 1
         End If
    Next
End With
Fillstatistics = statisticsRec
Exit Function
ErrorHandler:
Fillstatistics = statisticsRec
MsgBox Err.Description
End Function
Sub SearchData(FromDate As String, TillDate As String, PaymrntType As Integer)
On Error GoTo ErrorHandler

If Not IsDate(FromDate) Then
    FromDate = "01/01/1900"
End If

If Not IsDate(TillDate) Then
    TillDate = "12/31/" & Trim(Str(Year(Now)))
End If
Dim i As Integer
With FlexAccount

For i = 1 To .Rows - 1


    If FromDate <= .TextMatrix(i, colRepDate) And TillDate >= .TextMatrix(i, colRepDate) And _
    (.TextMatrix(i, ColPaymentTypeNo) = PaymrntType Or PaymrntType = -1) Then
        
        .RowData(i) = 1
        ColorRow i, &HFFFFC0
    Else
         .RowData(i) = 0
        ColorRow i, vbWhite
    End If
Next
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description




End Sub

Private Sub FlexAccount_DblClick()
With FlexAccount
   If .Rows = 1 Then Exit Sub
    .Visible = False
    .Col = ColFixIt
    .RowData(.Row) = .RowData(.Row) Xor 1
    If .RowData(.Row) And 1 Then
        .RowData(.Row) = 1
         'Set .CellPicture = LoadPicture(App.Path + "\Chkon95.bmp")
        ColorRow .Row, &HFFFFC0
    Else
        .RowData(.Row) = 0
         'Set .CellPicture = LoadPicture(App.Path + "\Chkoff95.bmp")
        ColorRow .Row, vbWhite
    End If
    .Visible = True
End With
FillStatisticsLabels
End Sub

Private Sub FlexAccount_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu mnuRight
End If
End Sub

Private Sub Form_Load()
init
End Sub
Sub FillStatisticsLabels()

 Dim statisticsRec   As statisticsType
        statisticsRec = Fillstatistics
        LSumSelected.Caption = statisticsRec.SumSelectedValues
        LCountSelected.Caption = statisticsRec.CountSelectedValues
        LSumTotal.Caption = statisticsRec.SumSelectedValues + statisticsRec.SumUnSelectedValues
        LCountTotal.Caption = statisticsRec.CountSelectedValues + statisticsRec.CountunSelectedValues

End Sub

Private Sub mnuCancel_Click()
With FlexAccount
    .Visible = False
    If .Row > .RowSel Then
        Stp = -1
    Else
        Stp = 1
    End If
    For i = .Row To .RowSel Step Stp
        .Row = i
        .Col = ColFixIt
         .RowData(.Row) = 0
         .TextMatrix(i, ColFixIt) = 0
         'Set .CellPicture = LoadPicture(App.Path + "\Chkoff95.bmp")
        ColorRow .Row, vbWhite
    Next
    .Visible = True
    FillStatisticsLabels
End With
End Sub

Private Sub mnuSelect_Click()
With FlexAccount
    .Visible = False
    If .Row > .RowSel Then
        Stp = -1
    Else
        Stp = 1
    End If
    For i = .Row To .RowSel Step Stp
        .Row = i
        .Col = ColFixIt
         .RowData(.Row) = 1
        .TextMatrix(i, ColFixIt) = 1

'         Set .CellPicture = LoadPicture(App.Path + "\Chkon95.bmp")
        ColorRow .Row, &HFFFFC0
    Next
    .Visible = True
    FillStatisticsLabels
End With
End Sub

Function transferToAccregTemp(transferDate As String) As Boolean
On Error GoTo ErrorHandler
transferToAccregTemp = False
Dim maxi As Double
ProgressBar1.Visible = True
Screen.MousePointer = vbHourglass
SQL = "Select Max(AccRegNoTemp)  As Maxi ," _
            & "Max(AccRegCoSer) as MaxCo ," _
            & "Max(AccRegEntry) as MaxEntry" _
      & " From " & systemConfigration.HafezHallsDatabaseDestination & ".Dbo.AccRegTemp "
Set rs = de.con.Execute(SQL)
maxi = IIf(IsNull(rs!maxi), 0, rs!maxi)
MaxEntry = IIf(IsNull(rs!MaxEntry), 0, rs!MaxEntry)
de.con.BeginTrans



With FlexAccount
ProgressBar1.Min = 1
    ProgressBar1.Max = .Rows
    Flag = True
    For i = 1 To .Rows - 1
        .Row = i
        If .RowData(.Row) = 1 Then
            ProgressBar1.Value = i
            SQL = "select max(AccRegCoSer) as MaxCo from " & systemConfigration.HafezHallsDatabaseDestination & ".Dbo.AccRegTemp where"
            SQL = SQL & " Month(AccRegDate)= Month(" & "'" & ConvertControlDate(transferDate) & "'" & ") and "
            SQL = SQL & " year(AccRegDate)= year( " & "'" & ConvertControlDate(transferDate) & "'" & ") and "
            SQL = SQL & "Right(DebAccNoTemp,2)=" & "'" & Right("0" + .TextMatrix(.Row, ColDebAccNoTemp), 2) & "'"
'            MsgBox Sql
            Set Rs1 = de.con.Execute(SQL)
            MaxCo = IIf(IsNull(Rs1!MaxCo), 0, Rs1!MaxCo)
            MaxCo = MaxCo + 1
            maxi = maxi + 1
            MaxEntry = MaxEntry + 1
            
            sqlText = "Insert into " & systemConfigration.HafezHallsDatabaseDestination & ".Dbo.AccRegTemp (AccRegNoTemp , DebAccNoTemp , CreAccNoTemp , "
            sqlText = sqlText & "DescriptionTemp , AccRegDate , Amount , Internal, FixIt , EmpNo) Values ("
            sqlText = sqlText & maxi & ",'" & .TextMatrix(.Row, ColDebAccNoTemp) & "','" & .TextMatrix(.Row, ColCreAccNoTemp) & "','"
            sqlText = sqlText & .TextMatrix(.Row, colDescription) & "','" & ConvertControlDate(transferDate) & "',"
            sqlText = sqlText & .TextMatrix(.Row, ColTTlPrice) & ",35,0," & empNo & ")"

            de.con.Execute (sqlText)
            

        End If
    Next
    de.con.CommitTrans
    ProgressBar1.Value = .Rows - 1
    End With
ProgressBar1.Visible = False
transferToAccregTemp = True
Exit Function
ErrorHandler:
transferToAccregTemp = False
de.con.RollbackTrans
MsgBox Err.Description
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
   
        GetData
        FillStatisticsLabels
Case 2
    If Not IsDate(txttransferDate.Text) Then
        MsgBox " «—ÌÕ «· —ÕÌ· €Ì— ’ÕÌÕ", vbCritical, " «—ÌŒ «· —ÕÌ·"
        Exit Sub
    End If
    If transferToAccregTemp(txttransferDate.Text) Then
        MsgBox " „  —ÕÌ· «·«’·«Õ« ", vbInformation, " —ÕÌ· «·»Ì‰« "
    End If
        
    Case 4
        
        Unload Me
End Select
End Sub

Private Sub TxtFromDate_Change()
TxtTillDate.Text = TxtFromDate.Text
End Sub
