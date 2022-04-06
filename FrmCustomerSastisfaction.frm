VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCustomerSastisfaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÑÖì ÇáÒÈæä"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   9480
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   5085
      Left            =   30
      TabIndex        =   7
      Top             =   750
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   8969
      _Version        =   131074
      ForeColor       =   128
      Alignment       =   1
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ãáÇÍÙÇÊ"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   2520
         Width           =   570
      End
      Begin VB.Label LDescription 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   2730
         Width           =   2055
      End
      Begin VB.Label lTeamName 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáæÑÔÉ"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   1635
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   3870
         Width           =   465
      End
      Begin VB.Label LRepPrice 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1230
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   4680
         Width           =   885
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáÊßáÝÉ"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   1635
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   4410
         Width           =   450
      End
      Begin VB.Label LCount 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   4680
         Width           =   885
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáÊßÑÇÑ"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   435
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   4440
         Width           =   450
      End
      Begin VB.Label LCallDescription 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1410
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáÔßæì"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   1515
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1200
         Width           =   570
      End
      Begin VB.Label LRepDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   900
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÊÇÑíÎ ÇáÅÕáÇÍ"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   1170
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   660
         Width           =   915
      End
      Begin VB.Label LCallDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   300
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÊÇÑíÎ ÇáÔßæì"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   1095
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   60
         Width           =   990
      End
   End
   Begin VB.CheckBox PhoneChk 
      Alignment       =   1  'Right Justify
      Caption         =   "ÇáåÇÊÝ ÕÍíÍ"
      Height          =   315
      Left            =   2250
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   750
      Width           =   2295
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   2085
      Left            =   2400
      TabIndex        =   3
      Top             =   1350
      Visible         =   0   'False
      Width           =   2385
      _cx             =   4207
      _cy             =   3678
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4530
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   3525
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   4725
      Left            =   2220
      TabIndex        =   0
      Top             =   1080
      Width           =   7215
      _cx             =   12726
      _cy             =   8334
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   810
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
            Picture         =   "FrmCustomerSastisfaction.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":4E7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":7777
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":A126
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":C64B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":EE03
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":11817
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":14169
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":16EDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":1972C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":1C5D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":1F32D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":21CCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":24C2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":27A8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":2A4B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":2CE6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":2F79F
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":323DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":34CE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":3774B
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":3A6FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":3D025
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":3F95A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":42509
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":44C49
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":475B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":49E40
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":4C04F
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":4E9AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":511D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":53C8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":566C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":591DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":5C176
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":5EF42
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerSastisfaction.frx":61BBC
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
      Width           =   9480
      _ExtentX        =   16722
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
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   15
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   2490
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   -30
         Width           =   6915
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "ÑÞã ÇáåÇÊÝ"
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   1
            Left            =   2100
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   180
            Width           =   825
         End
         Begin VB.Label LPhone 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   150
            Width           =   2055
         End
         Begin VB.Label LCustomerName 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3870
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   150
            Width           =   2055
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "ÅÓã ÇáÒÈæä"
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   0
            Left            =   6030
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   150
            Width           =   825
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÅÓã ÇáãÌíÈ Çæ ÕÝÊå"
      Height          =   195
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   750
      Width           =   1335
   End
End
Attribute VB_Name = "FrmCustomerSastisfaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ColQuestionId = 1
Const ColQuestionName = 2
Const ColChoicesId = 3
Const ColChoicesName = 4
Const ColClarification = 5

Dim FillCombosOk As Boolean


'Const ColQuestionName = 2

Sub FillFormating(ByVal i As Integer)
If i = 1 Then
    Fs = "|>" + "ÑÞã ÇáÓÄÇá"
    Fs = Fs + "|>" + "ÇáÓÜÜÜÜÜÜÄÇá"
    Fs = Fs + "|>" + "ÑÞã ÇáÎíÇÑÇÊ"
    Fs = Fs + "|>" + "ÇáÎíÇÑÇÊ"
    Fs = Fs + "|>" + "ÇáÈíÜÜÜÜÜÜÜÜÇä"
    With FlexGrid
        .FormatString = Fs
        .Cols = 6
        .ColWidth(ColQuestionId) = 0
        SetColWidths ColQuestionName, FlexGrid
'        SetColWidths ColChoicesId, FlexGrid
        .ColWidth(ColChoicesId) = 0
        SetColWidths ColChoicesName, FlexGrid
        SetColWidths ColClarification, FlexGrid
'        SetColWidths colStkName, flexGrid
'        SetColWidths colQty, flexGrid
'        SetColWidths colPrice, flexGrid
'        .ColWidth(colPriceTypeId) = 0
'        SetColWidths colPriceTypeName, flexGrid
'        .ColWidth(colPaymentTYpeId) = 0
'        SetColWidths colPaymentTYpeName, flexGrid
    End With
End If
End Sub

Sub SetColWidths(ByVal ColNo As Integer, FlexGrid As VSFlexGrid)
    With FlexGrid
        .AutoSize (ColNo)
    End With
End Sub

Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
FillFormating 1
End Sub

Private Sub FlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
If Col = ColQuestionName Then cancel = True
End Sub

Private Sub FlexGrid_RowColChange()
If FillCombosOk Then
    FillCombos
End If
End Sub

Private Sub PhoneChk_Click()
If PhoneChk.Value Then
    PhoneChk.Caption = "ÇáåÇÊÝ ÎØÃ"
    FlexGrid.Visible = False
Else
    PhoneChk.Caption = "ÇáåÇÊÝ ÕÍíÍ"
    FlexGrid.Visible = True
End If

End Sub
Sub FillLabels()
    With CustomerInformationRec
        LcustomerName.Caption = .AdhamName
        LPhone.Caption = .AdhamPhon
        LCallDate.Caption = .CallDateTime
        LRepDate.Caption = .RepDate
        LCallDescription.Caption = .CallDEscription
        LDescription.Caption = .Description
        LRepPrice.Caption = .RepPrice
        Lcount.Caption = .CountRec
    End With
End Sub
Sub BuildCombo(QuestionId As Integer, Col1 As Integer, col2 As Integer)
Dim RsClass  As New ADODB.Recordset
sqlText = "select  Ser , Response from CustomerCareReplay.dbo.CoQuestionResponce where QuestionId = " & QuestionId
Set RsClass = de.con.Execute(sqlText)

If RsClass.RecordCount > 0 Then
    With FlexGrid
        Lst = .BuildComboList(RsClass, "Response", "Response", vbYellow)
        .ColComboList(Col1) = Lst
        Lst = .BuildComboList(RsClass, "Response", "Ser", vbYellow)
        .ColComboList(col2) = Lst
    End With
Else
    With FlexGrid
        .Rows = 1
    End With
End If

End Sub

            
            



Sub FillCombos()
    With FlexGrid
        BuildCombo .TextMatrix(.Row, ColQuestionId), ColChoicesName, ColChoicesId
    End With
End Sub
Sub FillInitGrid()
Dim rsInit As New ADODB.Recordset
sqlText = "select QuestionId , QuestionName  from CustomerCareReplay.dbo.CoQuestion"
Set rsInit = de.con.Execute(sqlText)
Set FlexGrid.DataSource = rsInit
FillFormating 1
'FillCombos
End Sub
Sub init()
    FillCombosOk = False
    FlexGrid.Editable = flexEDKbdMouse
    FillLabels
    FillInitGrid
    FillCombosOk = True
End Sub
Private Sub Form_Load()
init
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 3
        Unload Me
End Select
End Sub
