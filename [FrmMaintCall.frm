VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmMaintCall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Œœ„… «·„” Â·ﬂ"
   ClientHeight    =   7155
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   11205
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   1365
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1050
      Visible         =   0   'False
      Width           =   11055
      _cx             =   19500
      _cy             =   2408
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
   Begin VB.CommandButton CmdSave 
      Caption         =   "Õ›Ÿ «·‘ﬂÊÏ"
      Height          =   315
      Left            =   1950
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   6450
      Width           =   1545
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "ÿ»«⁄…"
      Height          =   315
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   6450
      Width           =   1395
   End
   Begin VB.TextBox TxtCallVia 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   5910
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   6030
      Width           =   4095
   End
   Begin VB.TextBox txtCallNotes 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   6030
      Width           =   3465
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   2055
      Left            =   90
      TabIndex        =   48
      Top             =   2430
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   3625
      _Version        =   131074
      Begin VB.TextBox TxtHomePhone 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   7740
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   450
         Width           =   1995
      End
      Begin VB.TextBox TxtZoneName 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   7740
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   1995
      End
      Begin VB.TextBox TxtCustomerName 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   90
         Width           =   9705
      End
      Begin VB.TextBox TxtNotes 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1230
         Width           =   9675
      End
      Begin VB.TextBox TxtEmail 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   7740
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1620
         Width           =   1995
      End
      Begin VB.TextBox TxtMobilePhone 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3300
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   3585
      End
      Begin VB.TextBox TxtWorkPhone 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   510
         Width           =   2145
      End
      Begin VB.TextBox TxtAddress 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   870
         Width           =   6795
      End
      Begin Crystal.CrystalReport cr1 
         Left            =   4710
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "≈”„ «·“»Ê‰"
         Height          =   195
         Index           =   3
         Left            =   10230
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   90
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·Â« › «·À«» "
         Height          =   195
         Index           =   4
         Left            =   9810
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   510
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·„Ê»«Ì·"
         Height          =   195
         Index           =   5
         Left            =   6900
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   510
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„‰ÿﬁ…"
         Height          =   195
         Index           =   7
         Left            =   10485
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   870
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·⁄‰Ê«‰"
         Height          =   195
         Index           =   8
         Left            =   7020
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   870
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„·«ÕŸ« "
         Height          =   195
         Index           =   9
         Left            =   10365
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   1230
         Width           =   570
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·»—Ìœ «·≈·ﬂ —Ê‰Ì"
         Height          =   195
         Index           =   19
         Left            =   9840
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1650
         Width           =   1125
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ „ﬂ«‰ «·⁄„·"
         Height          =   195
         Index           =   6
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   540
         Width           =   1035
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid1 
      Height          =   1815
      Left            =   1920
      TabIndex        =   47
      Top             =   2550
      Visible         =   0   'False
      Width           =   2895
      _cx             =   5106
      _cy             =   3201
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
   Begin Threed.SSFrame SSFrame1 
      Height          =   435
      Left            =   30
      TabIndex        =   38
      Top             =   0
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   767
      _Version        =   131074
      Begin VB.Label LCallNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   90
         Width           =   1245
      End
      Begin VB.Label LCallDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6450
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   90
         Width           =   1245
      End
      Begin VB.Label LCalltime 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4170
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   90
         Width           =   1245
      End
      Begin VB.Label LCallrecieverName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   90
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·‘ﬂÊÏ"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   0
         Left            =   10170
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   90
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «·‘ﬂÊÏ"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   7740
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   90
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Êﬁ  «·‘ﬂÊÏ"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   2
         Left            =   5490
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   90
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„” ·„ «·‘ﬂÊÏ"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   3
         Left            =   3090
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   90
         Width           =   1005
      End
   End
   Begin VB.TextBox TxtREpNo 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   8310
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   690
      Width           =   885
   End
   Begin VB.TextBox TxtREpName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4410
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   690
      Width           =   3855
   End
   Begin VB.TextBox TxtCustomer 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   690
      Width           =   4335
   End
   Begin MSDataListLib.DataCombo ComboProductFamily 
      Height          =   315
      Left            =   9180
      TabIndex        =   0
      Top             =   690
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VSFlex8Ctl.VSFlexGrid GRidMaintInfo 
      Height          =   1245
      Left            =   30
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4500
      Width           =   11085
      _cx             =   19553
      _cy             =   2196
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
   Begin MSDataListLib.DataCombo ComboPaymentTYpe 
      Height          =   360
      Left            =   4170
      TabIndex        =   15
      Top             =   6030
      Width           =   1725
      _ExtentX        =   3043
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
   Begin Threed.SSFrame SSFrame2 
      Height          =   375
      Left            =   30
      TabIndex        =   63
      Top             =   6750
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   661
      _Version        =   131074
      Begin VB.Label LDailyCount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8340
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   60
         Width           =   1245
      End
      Begin VB.Label LyearlyCount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   60
         Width           =   1245
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "⁄œœ «·‘ﬂ«ÊÏ «·”‰ÊÌ…"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   6630
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   60
         Width           =   1485
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "⁄œœ «·‘ﬂ«ÊÏ «·ÌÊ„Ì…"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   9660
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   60
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000080&
      Height          =   1365
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1050
      Width           =   11085
      Begin VB.Label LVia 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   990
         Width           =   1665
      End
      Begin VB.Label LNotes 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   990
         Width           =   3195
      End
      Begin VB.Label LAddress 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5010
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   990
         Width           =   6015
      End
      Begin VB.Label LZoneName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   420
         Width           =   2745
      End
      Begin VB.Label LWorkPhone 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2790
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label LMobilePhone 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label LHomePhone 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6210
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label LcustomerName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7950
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   420
         Width           =   3075
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "»„⁄—›… "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   18
         Left            =   1170
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   750
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "„·«ÕŸ« "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   17
         Left            =   4020
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   750
         Width           =   930
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "«·⁄‰Ê«‰"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   16
         Left            =   10410
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   750
         Width           =   585
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„‰ÿﬁ…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   15
         Left            =   2010
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   150
         Width           =   675
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ „ﬂ«‰ «·⁄„·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   3210
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   150
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·„Ê»«Ì·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   13
         Left            =   5190
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·Â« › «·À«» "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   150
         Width           =   1410
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "≈”„ «·“»Ê‰"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   10140
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   150
         Width           =   885
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "⁄œœ „—«  «·“Ì«—…"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10035
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   5760
      Width           =   1110
   End
   Begin VB.Label Lcount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9150
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "»„⁄—›… "
      Height          =   195
      Index           =   10
      Left            =   10620
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   6090
      Width           =   510
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„·«ÕŸ« "
      Height          =   195
      Index           =   21
      Left            =   3570
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   6060
      Width           =   570
   End
   Begin VB.Label LRepeator 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   5760
      Width           =   3405
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·⁄ÿ·"
      Height          =   195
      Index           =   1
      Left            =   8790
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   480
      Width           =   390
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·⁄«∆·…"
      Height          =   195
      Index           =   0
      Left            =   10695
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   450
      Width           =   420
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·“»Ê‰"
      Height          =   195
      Index           =   2
      Left            =   3945
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   450
      Width           =   420
   End
End
Attribute VB_Name = "FrmMaintCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Flag As Boolean
Dim gridok As Boolean
Dim ComboOk As Boolean
Dim OkRepNo As Boolean

Const ColCustomerId = 1
Const ColCustomerName = 2
Const ColCustomerHomephone = 3
Const ColCustomerMobilePhone = 4
Const ColCustomerWorkPhone = 5
Const ColCustomerAddress = 6

Const ColIsPrinted = 1
Const ColCliNo = 2
Const ColCallNO = 3
Const colTeamNo = 4
Const ColTeamName = 5
Const ColCallDate = 6
Const ColCallTime = 8
Const colRepDate = 9
Const ColRepTimeBegin = 0
Const ColREpTimeEnd = 10
Const colRepPrice = 11
Const colSymbol = 12
Const ColProdPurchaseDate = 13
Const colNotes = 14
Const ColClass = 15


Const Colinfocallno = 1
Const ColinfocustomerNo = 2
Const Colinfocustomername = 3
Const Colinfoacustomerhomephone = 4
Const Colinfocalldate = 5
Const Colinfocalltime = 6
Const Colinfocustomeraddress = 7
Const ColinfoCallReceiver = 8
Const ColinfoProductFamilyNo = 9
Const ColinfoNote = 10


Const ColViaNo = 1
Const ColViaName = 2


Const ColZoneNo = 1
Const ColZoneName = 2


Dim MaintCallRec As MntCallRecTYpe
Dim CustomerRec As CustomerRecType


Sub MoveCursor(KeyCode As Integer, Grid As VSFlexGrid)
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

Function GetMaintCall(vdate As String) As Double
Dim rscalls As New ADODB.Recordset
sqlText = "select count(*) as CountRec From MaintCall where calldatetime >='" & ConvertControlDate(vdate) & "'"
Set rscalls = de.con.Execute(sqlText)
GetMaintCall = rscalls!CountRec
End Function

Sub FillFormating(FlexGrid As VSFlexGrid, i As Integer)
If i = 1 Then
    Fs = "|>" & "—ﬁ„ «·“»Ê‰"
    Fs = Fs + "|>" & "≈”„ «·“»Ê‰"
    Fs = Fs + "|>" & "—ﬁ„ «·„‰“·"
    Fs = Fs + "|>" & "—ﬁ„ «·„Ê»«Ì·"
    Fs = Fs + "|>" & "—ﬁ„ «·⁄„·"
    Fs = Fs + "|>" & "«·⁄‰Ê«‰"
    With FlexGrid
        .Visible = False
        .Cols = 7
        .FormatString = Fs
        SetColWidths ColCustomerId, FlexGrid
'        .ColWidth(ColCustomerId) = 0
        SetColWidths ColCustomerName, FlexGrid
        SetColWidths ColCustomerHomephone, FlexGrid
        SetColWidths ColCustomerMobilePhone, FlexGrid
        SetColWidths ColCustomerWorkPhone, FlexGrid
        SetColWidths ColCustomerAddress, FlexGrid
        .Visible = True
    End With
ElseIf i = 2 Then
    
    Fs = "|>" & "Chk"
    Fs = Fs + "|>" & "—ﬁ„ «·“»Ê‰"
    Fs = Fs + "|>" & "—ﬁ„ «·‘ﬂÊÏ"
    Fs = Fs + "|>" & "—ﬁ„ «·›—Ìﬁ"
    Fs = Fs + "|>" & "›—Ìﬁ «·’Ì«‰…"
    Fs = Fs + "|>" & " «—ÌŒ «·‘ﬂÊÏ"
    Fs = Fs + "|>" & "”«⁄… «·‘ﬂÊÏ"
    
    Fs = Fs + "|>" & " «—ÌŒ «·≈’·«Õ"
    Fs = Fs + "|>" & " ÊﬁÌﬁ «·»œ«Ì…"
    Fs = Fs + "|>" & " ÊﬁÌÀ «·‰Â«Ì…"
    Fs = Fs + "|>" & "«·ﬂ·›… "
    Fs = Fs + "|>" & "«·„‰ Ã"
    
    
    Fs = Fs + "|>" & " «—ÌŒ «·‘—«¡"
    Fs = Fs + "|>" & "„·«ÕŸ« "
    Fs = Fs + "|>" & "«· ’‰Ì›"
    
    With FlexGrid
        .Visible = False
        .Cols = 16
        .FormatString = Fs
        
        .ColWidth(ColIsPrinted) = 300
        .ColWidth(ColCliNo) = 0
        .ColWidth(colTeamNo) = 0
        .ColWidth(ColClass) = 0
        
        SetColWidths ColCallNO, FlexGrid
        SetColWidths ColTeamName, FlexGrid
        SetColWidths ColCallDate, FlexGrid
        SetColWidths ColCallTime, FlexGrid
        SetColWidths colRepDate, FlexGrid
        SetColWidths ColRepTimeBegin, FlexGrid
        SetColWidths ColREpTimeEnd, FlexGrid
        SetColWidths colRepPrice, FlexGrid
        SetColWidths colSymbol, FlexGrid
        SetColWidths ColProdPurchaseDate, FlexGrid
        SetColWidths colNotes, FlexGrid
        .ColDataType(ColIsPrinted) = flexDTBoolean
        .Visible = True
    End With
ElseIf i = 3 Then
    Fs = "|>" & "—ﬁ„ «·‘ﬂÊÏ"
    Fs = Fs + "|>" & "—ﬁ„ «·“»Ê‰"
    Fs = Fs + "|>" & "≈”„ «·“»Ê‰"
    Fs = Fs + "|>" & "—ﬁ„ «·Â« ›"
    Fs = Fs + "|>" & " «—ÌŒ «·‘ﬂÊÏ"
    Fs = Fs + "|>" & "Êﬁ  «·‘ﬂÊÏ"
    Fs = Fs + "|>" & "⁄‰Ê«‰ «·“»Ê‰"
    Fs = Fs + "|>" & "„” ﬁ»· «·‘ﬂÊÏ"
    Fs = Fs + "|>" & "«·„‰ Ã"
    Fs = Fs + "|>" & "«·⁄ÿ·"
    With FlexGrid
        .Visible = False
        .Cols = 11
        .FormatString = Fs
        SetColWidths Colinfocallno, FlexGrid
        .ColWidth(ColinfocustomerNo) = 0
        SetColWidths Colinfocustomername, FlexGrid
        SetColWidths Colinfoacustomerhomephone, FlexGrid
        SetColWidths Colinfocalldate, FlexGrid
        SetColWidths Colinfocalltime, FlexGrid
        SetColWidths Colinfocustomeraddress, FlexGrid
        SetColWidths ColinfoCallReceiver, FlexGrid
        .ColWidth(ColinfoProductFamilyNo) = 0
        .ColWidth(ColinfoNote) = 0
        .Visible = True
    End With
ElseIf i = 4 Then
    Fs = "|>" & "—ﬁ„ «·„⁄—›…"
    Fs = Fs + "|>" & "≈”„ «·„⁄—›…"
    With FlexGrid
        .Visible = False
        .Cols = 3
        .FormatString = Fs
        .ColWidth(ColViaNo) = 0
        SetColWidths ColViaName, FlexGrid
        .Visible = True
    End With
ElseIf i = 5 Then
    Fs = "|>" & "—ﬁ„ «·„‰ÿﬁ…"
    Fs = Fs + "|>" & "≈”„ «·„‰ÿﬁ…"
    With FlexGrid
        .Visible = False
        .Cols = 3
        .FormatString = Fs
        .ColWidth(ColZoneNo) = 0
        SetColWidths ColZoneName, FlexGrid
        .Visible = True
    End With
End If
End Sub

Sub SetColWidths(ByVal ColNo As Integer, FlexGrid As VSFlexGrid)
    With FlexGrid
        .AutoSize ColNo
    End With
End Sub

Sub FillCombos()
    Dim rs As New ADODB.Recordset
    sqlText = "Select ProdFamNo  , ProdFamNamea From AdhamProductFamily "
    Set rs = de.con.Execute(sqlText)
    Set ComboProductFamily.RowSource = rs
    ComboProductFamily.listField = "ProdFamNamea"
    ComboProductFamily.BoundColumn = "ProdFamNo"
    
    sqlText = "Select Id, PaymentType From CoMaintCallPaymentType"
    Set rs = de.con.Execute(sqlText)
    Set ComboPaymentTYpe.RowSource = rs
    ComboPaymentTYpe.listField = "PaymentType"
    ComboPaymentTYpe.BoundColumn = "Id"
    
End Sub
'Sub FillMainCallGrid()
'gridok = False
'    Sqltext = "Select callno , adhamno  , adhamname , adhamphon , calldate , calltime , adhamadress , CallReceiver , ModNo , Notes  from fn_GetMaintCallInfo('" & ConvertControlDate(Date) & "')"
'    Set Rs = de.con.Execute(Sqltext)
'    Set GridCall.DataSource = Rs
'    FillFormating GridCall, 3
'gridok = True
'End Sub
Sub init()

'ReadIniFile App.Path & "\init.ini", ";"
'ConnectString = "Provider=SQLOLEDB.1;Password=" & PWD & ";Persist Security Info=True;User ID=" & UID & " ;Initial Catalog=" & DataBase & ";Data Source=" & ServerName
'If de.con.State <> adStateOpen Then de.con.Open ConnectString, "user1", GetPass
top = 0
left = 0
FillCombos
'FillMainCallGrid
GRidMaintInfo.Rows = 1
FillFormating GRidMaintInfo, 2
'SSTab1.Tab = 1
LDailyCount.Caption = GetMaintCall(Date)
LyearlyCount.Caption = GetMaintCall("01/01/" + LTrim(RTrim(Str(Year(Date)))))
'GridCall.Row = 1
'GridCall_RowColChange
Flag = True
ComboOk = True
OkRepNo = True
End Sub

Function FillVariables(Voption As Integer) As Boolean
On Error GoTo errorhandler
Select Case Voption
    Case 1
        If Val(ComboProductFamily.BoundText) = 0 Or IsEmpty(TxtREpName.Text) Or Val(TxtCustomerName.Tag) = 0 Then
            FillVariables = False
            Exit Function
        End If
    Case 2
End Select
FillVariables = True
Exit Function
errorhandler:
FillVariables = False
End Function

Function FillStructure(Voption As Integer) As Boolean
On Error GoTo errorhandler
    Select Case Voption
        Case 1 ' Call
            With MaintCallRec
                If FillVariables(Voption) Then
                    .ModNo = Val(ComboProductFamily.BoundText)
                    .CallDEscription = TxtREpName.Text
                    .cliNo = Val(TxtCustomer.Tag)
                    .Defindname = Val(TxtCallVia.Tag)
                    .PaymentTYpeId = Val(ComboPaymentTYpe.BoundText)
                    .Notes = txtCallNotes.Text
                    .CallReceiverEmpNo = empNo
                    
                    FillStructure = True
                End If
            End With
        Case 2 ' Customer
        With CustomerRec
            If FillVariables(Voption) Then
                .AdhamName = TxtCustomerName.Text
                 If Val(TxtCustomerName.Tag) <> 0 Then
                    .adhamNo = TxtCustomerName.Tag
                 Else
                    .adhamNo = getmaxCustomer()
                 End If
                .AdhamAdress = TxtAddress.Text
                .AdhamPhon = TxtHomePhone.Text
                .MobilePhone = TxtMobilePhone.Text
                .Workphone = TxtWorkPhone.Text
                .Zone = Val(TxtZoneName.Tag)
                .Notes = TxtNotes.Text
                .Email = TxtEmail.Text
                FillStructure = True
            End If
        End With
    End Select
Exit Function
errorhandler:
FillStructure = False
End Function
Function getmaxCustomer() As Double
Dim rs As New ADODB.Recordset
sqlText = "Select isnull(Max(AdhamNo),0) MaxAdhamNo From AdhamView7"
Set rs = de.con.Execute(sqlText)
getmaxCustomer = rs!MaxAdhamNo + 1
End Function
Function GetMaxCallNo() As Double

Dim rs As New ADODB.Recordset
sqlText = "Select isnull(Max(CallNo),0) MaxCallNo From MaintCall Where year(CallDateTime)=" & Year(Date)
Set rs = de.con.Execute(sqlText)
GetMaxCallNo = rs!MaxCallNo + 1
End Function
Function SaveRec(Voption As Integer) As Boolean
On Error GoTo errorhandler
Select Case Voption
    Case 1
        With MaintCallRec
            .CallNo = GetMaxCallNo
            sqlText = "insert Into MaintCall(Id_ComNo, CallNo, ModNo, CliNo, CallDescription, CallStatus, Notes, PaymentTypeId, Via , CallReceiverEmpNo , CallDatetime) Select " & .CompNo & "," & .CallNo & "," & .ModNo & "," & .cliNo & ",'" & .CallDEscription & "',0,'" & .Notes & "'," & .PaymentTYpeId & ",'" & .Defindname & "'," & .CallReceiverEmpNo & ",GETDATE()"
            de.con.Execute (sqlText)
            
            SaveRec = True
        End With
    Case 2
        With CustomerRec
            If Val(TxtCustomerName.Tag) = 0 Then
                sqlText = "Insert into AdhamView7(Id_ComNo, AdhamNo, adhamname, adhamphon, WorkPhone, MobilePhone, adhamadress, Email, zone, Notes )Values(" & .CompNo & "," & .adhamNo & ",'" & .AdhamName & "','" & .AdhamPhon & "','" & .Workphone & "','" & .MobilePhone & "','" & .AdhamAdress & "','" & .Email & "'," & .Zone & ",'" & .Notes & "')"
                de.con.Execute (sqlText)
                SaveRec = True
            Else
                sqlText = "Update AdhamView7 set adhamname='" & .AdhamName & "',adhamphon='" & .AdhamPhon & "',WorkPhone='" & .Workphone & "',MobilePhone='" & .MobilePhone & "',adhamadress='" & .AdhamAdress & "',Email='" & .Email & "',Zone=" & .Zone & ",Notes='" & .Notes & "' Where AdhamNo = " & .adhamNo
                de.con.Execute (sqlText)
                SaveRec = True
            End If
        End With
    End Select
Exit Function
errorhandler:
SaveRec = False
MsgBox Err.Description
End Function

Function GetEmployeeName(CallNo As Double) As String
On Error GoTo errorhandler
Dim rs As New ADODB.Recordset
sqlText = "Select CallReceiverEmpNo , Fullname from maintcall m1 left outer join empfullname e1 on m1.CallReceiverEmpNo = e1.empno Where CallNo = " & CallNo
Set rs = de.con.Execute(sqlText)
GetEmployeeName = rs!FullName
Exit Function
errorhandler:
MsgBox Err.Description
End Function

Sub FillGrid(Grid As VSFlexGrid, CustomerNo As Integer)
On Error GoTo errorhandler
Dim rs As New ADODB.Recordset
sqlText = "Select AdhamName , AdhamPhon , adhamadress From Adhamview7 Where AdhamNo=" & CustomerNo
Set rs = de.con.Execute(sqlText)
With Grid
    Vrow = .Rows - 1
   .TextMatrix(Vrow, Colinfocallno) = MaintCallRec.CallNo
   .TextMatrix(Vrow, ColinfocustomerNo) = MaintCallRec.cliNo
   .TextMatrix(Vrow, Colinfocustomername) = rs!AdhamName & ""
   .TextMatrix(Vrow, Colinfoacustomerhomephone) = rs!AdhamPhon & ""
   .TextMatrix(Vrow, Colinfocalldate) = MaintCallRec.CallDateTime
   .TextMatrix(Vrow, Colinfocalltime) = MaintCallRec.CallDateTime
   .TextMatrix(Vrow, Colinfocustomeraddress) = rs!AdhamAdress & ""
   .TextMatrix(Vrow, ColinfoCallReceiver) = GetEmployeeName(MaintCallRec.CallNo)
   .TextMatrix(Vrow, ColinfoProductFamilyNo) = MaintCallRec.ModNo
   .TextMatrix(Vrow, ColinfoNote) = MaintCallRec.Notes
End With
Exit Sub
errorhandler:
End Sub

Sub insertintoGrid(Voption As Integer)
    FillREparationInfo MaintCallRec.cliNo
'With GRidMaintInfo
'    .AddItem ""
'    FillGrid GRidMaintInfo, MaintCallRec.CallNo
'    .Col = Colcallno
'    .Sort = flexSortNumericDescending
'End With
End Sub
Sub FillLabelsHeader(CallNo As Double)
Dim rs As New ADODB.Recordset
sqlText = "Select CallNo ,  Convert(varchar(5),calldatetime,108) as DAte   ,  Convert(varchar(10),calldatetime,103) as time ,  CallReceiverEmpNo From maintcall Where CAllNo =" & CallNo
Set rs = de.con.Execute(sqlText)
LCallNo.Caption = rs!CallNo
LCallDate.Caption = rs!Date
LCalltime.Caption = rs!Time
LCallrecieverName.Caption = GetEmployeeName(rs!CallNo)
End Sub
Function ExecProcedure(CustomerId As Double) As Boolean
On Error GoTo errorhandler
sqlText = "Exec Sp_GetReparationInfo " & CustomerId
de.con.Execute (sqlText)
ExecProcedure = True
Exit Function
errorhandler:
ExecProcedure = False
End Function
Private Sub CmdPrint_Click()
If ExecProcedure(Val(TxtCustomer.Tag)) Then
    PrintData
End If
End Sub

Private Sub CmdSave_Click()
If FillStructure(1) Then
    If SaveRec(1) Then
    
        LDailyCount.Caption = GetMaintCall(Date)
        LyearlyCount.Caption = GetMaintCall("01/01/" + LTrim(RTrim(Str(Year(Date)))))
        insertintoGrid (1)
        FillLabelsHeader (MaintCallRec.CallNo)
        If MsgBox(" „ ≈œŒ«· «·‘ﬂÊÏ Â·  —Ìœ ÿ»«⁄… «·ﬁ”Ì„…", vbYesNo + vbQuestion, "ÿ»«⁄… «·ﬁ”Ì„…") = vbYes Then
            PrintData
            ComboProductFamily.SetFocus
            ClearLables
        End If
        
        'GridCall_RowColChange
        'ComboProductFamily.SetFocus
    End If
End If
End Sub
Sub UpdateStatus(CallNo As Double)
sqlText = "Update MaintCall Set CallStatus=1 Where CallNo= " & CallNo
de.con.Execute (sqlText)
End Sub
Sub PrintData()
With cr1
    UpdateStatus (MaintCallRec.CallNo)
    .Connect = ConnectName("")
    .ReportFileName = App.Path + "\Reports\repMaintInfoDetails.rpt"
    .DiscardSavedData = True
    .WindowState = crptMaximized
    .Action = 1
End With

End Sub
Private Sub ComboProductFamily_Change()
If ComboOk Then
    ClearLables
End If
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub Form_Load()
init
End Sub

Function GetRepName(ByVal CallNo As Integer) As String
Dim rs As New ADODB.Recordset

On Error GoTo errorhandler
sqlText = "Select Callid , Calldescription From mntcallscode m1 inner join adhamproductfamily a1 on m1.prodfamno = a1.prodfamno Where Callid = " & CallNo & "  and m1.ProdFamNo =" & Val(ComboProductFamily.BoundText)
Set rs = de.con.Execute(sqlText)

If rs.RecordCount > 0 Then
    GetRepName = rs!CallDEscription
Else
    GetRepName = ""
End If
Exit Function
errorhandler:
GetRepName = ""
End Function
Sub MoveGrid(VTop As Integer, VLeft As Integer, VWidth As Integer, Grid As VSFlexGrid)
With Grid
    .top = VTop
    .Width = VWidth
    .left = VLeft
End With
End Sub

Private Sub GridCall_RowColChange()
'On Error Resume Next
If gridok Then
    With GridCall
        
        LCallNo.Caption = .TextMatrix(.Row, Colinfocallno)
        LCallDate.Caption = .TextMatrix(.Row, Colinfocalldate)
        LCalltime.Caption = .TextMatrix(.Row, Colinfocalltime)
        LCallrecieverName.Caption = .TextMatrix(.Row, ColinfoCallReceiver)
        ComboOk = False
        ComboProductFamily.BoundText = .TextMatrix(.Row, ColinfoProductFamilyNo)
        ComboOk = True
        TxtREpName.Text = .TextMatrix(.Row, ColinfoNote)
        'Lcount.Caption = GRidMaintInfo.Rows - 1
        Flag = False
        TxtCustomer.Tag = .TextMatrix(.Row, ColinfocustomerNo)
        TxtCustomer.Text = .TextMatrix(.Row, Colinfocustomername)
        Flag = True
        FillLabels_Controls Val(.TextMatrix(.Row, ColinfocustomerNo))
    End With
End If
End Sub

Private Sub GRidMaintInfo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim FirstRow As Integer, LastRow As Integer, Vrow As Integer
With GRidMaintInfo
If KeyCode = vbKeyDelete Then
    If MsgBox("Â· √‰  „ √ﬂœ „‰ Õ–› «·ﬁ”Ì„…", vbYesNo + vbDefaultButton2, "Õ–› ﬁ”«∆„") = vbYes Then
        If .Row >= .RowSel Then
            FirstRow = .Row
            LastRow = .RowSel
        Else
            FirstRow = .RowSel
            LastRow = .Row
        End If
        For i = FirstRow To LastRow Step -1
            Vrow = i
            
            If DeleteRow(GRidMaintInfo, Vrow, ColCallNO, "MaintCall", "Callno") Then
                .RemoveItem Vrow
            End If
        Next
    End If
End If
End With

End Sub

Private Sub TxtCallVia_Change()
If Flag Then
    Dim sqlText As String
    If Trim(TxtCallVia.Text) = "" Then
         TxtCallVia.Tag = 0
        Grid.Visible = False
        Exit Sub
        End If
        sqlText = "select top 10 accNo , AccName From MaintCallVia Where AccName like " & LikeExpression(TxtCallVia.Text)
        FillList sqlText, "AccNo", "AccName", Grid1, 4
        MoveGrid TxtCallVia.top + TxtCallVia.Height, TxtCallVia.left, TxtCallVia.Width, Grid1
End If
End Sub

Private Sub TxtCallVia_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = False
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid1
    Flag = True
End Sub

Private Sub TxtCallVia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
With Grid1
    If .Visible Then
        TxtCallVia.Tag = .TextMatrix(.Row, ColViaNo)
        Flag = False
        TxtCallVia.Text = .TextMatrix(.Row, ColViaName)
        Flag = True
        .Visible = False
    Else
        Flag = False
        TxtCallVia.Tag = 0
        TxtCallVia.Text = ""
        Flag = True
    End If
End With
End If

End Sub

Private Sub TxtCustomer_Change()
If Flag Then
    Dim sqlText As String
    If Trim(TxtCustomer.Text) = "" Then
        TxtCustomer.Tag = 0
        Grid.Visible = False
        Exit Sub
        End If
        sqlText = "Select top 10 [adhamno] , [AdhamName] , [AdhamPhon] , [MobilePhone] , [WorkPhone] ,  [adhamadress]  From adhamview7 Where   AdhamName like" & LikeExpression(TxtCustomer.Text) & " Or Adhamphon Like '" & TxtCustomer.Text & "%' or WorkPhone like '" & TxtCustomer.Text & "%' or MobilePhone like '" & TxtCustomer.Text & "%'"
        FillList sqlText, "AdhamNo", "AdhamName", Grid, 1
End If
End Sub

Private Sub TxtCustomer_GotFocus()
ChangeToArabic
End Sub

Private Sub TxtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = False
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True
End Sub
Sub ClearLables()
    Flag = False
        TxtCustomer.Text = ""
        TxtCustomer.Tag = ""
        TxtREpNo.Text = ""
        TxtREpName.Text = ""
        
        TxtCustomerName.Text = ""
        TxtCustomerName.Tag = 0
        TxtHomePhone.Text = ""
        TxtMobilePhone.Text = ""
        TxtWorkPhone.Text = ""

        TxtAddress.Text = ""
        TxtEmail.Text = ""
        TxtZoneName.Tag = ""
        TxtZoneName.Text = ""
        TxtNotes.Text = ""
    Flag = True
        
        TxtREpName.Text = ""
        
        LcustomerName.Caption = ""
        LHomePhone.Caption = ""
        LWorkPhone.Caption = ""
        LMobilePhone.Caption = ""
        LAddress.Caption = ""
        LNotes.Caption = ""
        LZoneName.Caption = ""
        LCallNo.Caption = ""
        LCallDate.Caption = ""
        LCalltime.Caption = ""
        LCallrecieverName.Caption = ""
        Lcount.Caption = ""
        GRidMaintInfo.Rows = 1
End Sub
Sub Colorclass(ByVal FlexGrid As VSFlexGrid)
With FlexGrid
    For i = 1 To .Rows - 1
        .Row = i
        If .TextMatrix(.Row, ColClass) = 1 Then
        For J = 1 To .Cols - 1
            
            .Col = J
            .CellBackColor = &H4B03E0
        Next
    End If
    Next
End With
End Sub

Sub FillREparationInfo(ByVal CustomerId As Double)
Dim rs As New ADODB.Recordset
sqlText = "Select CallStatus , CliNo , callno ,teamno , teamname , calldate  ,calltime  ,Repdate , RepTimeBegin , REpTimeEnd , repprice , symbol  , ProdPurchaseDate , notes   ,  Class   From dbo.GetAllMaintInfo(" & CustomerId & ") Order By Year Desc, CallNo desc"
Set rs = de.con.Execute(sqlText)
Set GRidMaintInfo.DataSource = rs
FillFormating GRidMaintInfo, 2
Lcount.Caption = GRidMaintInfo.Rows - 1
Colorclass GRidMaintInfo
End Sub

Sub FillLabels_Controls(ByVal CustomerId As Double)
Dim rs As New ADODB.Recordset
sqlText = "Select AdhamNo, adhamname, adhamphon, WorkPhone, MobilePhone, adhamadress, kind, zone, ZoneName , Notes, defindname, AccNO , email From dbo.adhamview7 a1  left outer join CoZone c1 on a1.zone = c1.zoneNo Where AdhamNo = " & CustomerId
Set rs = de.con.Execute(sqlText)

If rs.RecordCount > 0 Then
    With rs
        Flag = False
        TxtCustomerName.Text = !AdhamName & ""
        TxtCustomerName.Tag = IIf(IsNull(!adhamNo), 0, !adhamNo)
        
        TxtHomePhone.Text = !AdhamPhon & ""
        TxtMobilePhone.Text = !MobilePhone & ""
        TxtWorkPhone.Text = !Workphone & ""
        TxtAddress.Text = !AdhamAdress & ""
        Flag = True
        TxtEmail.Text = !Email & ""
        Flag = False
        TxtZoneName.Tag = IIf(IsNull(!Zone), 0, !Zone)
        TxtZoneName.Text = !ZoneName & ""
        Flag = True
        TxtNotes.Text = !Notes & ""
        
        LcustomerName.Caption = !AdhamName & ""
        LHomePhone.Caption = !AdhamPhon & ""
        LWorkPhone.Caption = !Workphone & ""
        LMobilePhone.Caption = !MobilePhone & ""
        LAddress.Caption = !AdhamAdress & ""
        LNotes.Caption = !Notes & ""
        
        LZoneName.Caption = !ZoneName & ""
        
    End With
    FillREparationInfo (Val(TxtCustomerName.Tag))
Else
    ClearLables
End If

End Sub

Private Sub TxtCustomer_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
With Grid
    If .Visible Then
        TxtCustomer.Tag = .TextMatrix(.Row, ColCustomerId)
        Flag = False
        TxtCustomer.Text = .TextMatrix(.Row, ColCustomerName)
        FillLabels_Controls Val(TxtCustomer.Tag)
        Flag = True
        .Visible = False
    Else
        Flag = False
        TxtCustomer.Tag = 0
        TxtCustomer.Text = ""
        Flag = True
    End If
End With
End If
End Sub

Private Sub TxtCustomerName_Change()
If Flag Then
    Dim sqlText As String
    If Trim(TxtCustomerName.Text) = "" Then
        TxtCustomer.Tag = 0
        Grid.Visible = False
        Exit Sub
        End If
        sqlText = "Select top 10 [adhamno] , [AdhamName] , [AdhamPhon] , [MobilePhone] , [WorkPhone] ,  [adhamadress]  From adhamview7 Where   AdhamName like" & LikeExpression(TxtCustomerName.Text)
        FillList sqlText, "AdhamNo", "AdhamName", Grid1, 1
        MoveGrid TxtCustomerName.top + TxtCustomerName.Height, TxtCustomerName.left, TxtCustomerName.Width, Grid1
End If
End Sub

Private Sub TxtCustomerName_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = False
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid1
    Flag = True
End Sub

Function GetCustomerName(CustomerId As Double) As String
Dim rs As New ADODB.Recordset
sqlText = "Select adhamName from AdhamView7 Where AdhamNo=" & CustomerId
Set rs = de.con.Execute(sqlText)
GetCustomerName = rs!AdhamName
End Function



Private Sub TxtCustomerName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

With Grid1
    If .Visible Then
        TxtCustomer.Tag = .TextMatrix(.Row, ColCustomerId)
        Flag = False
        TxtCustomer.Text = .TextMatrix(.Row, ColCustomerName)
        FillLabels_Controls Val(TxtCustomer.Tag)
        Flag = True
        .Visible = False
    ElseIf Val(TxtCustomer.Tag) <> 0 And .Visible = False Then
        Flag = False
        TxtCustomer.Text = GetCustomerName(TxtCustomer.Tag)
        Flag = True
    Else
        Flag = False
        TxtCustomer.Tag = 0
        TxtCustomer.Text = ""
        Flag = True
    End If
End With
End If

End Sub

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If FillStructure(2) Then
        If SaveRec(2) Then
            'SSTab1.Tab = 1
        End If
    End If
End If
End Sub

Private Sub TxtHomePhone_Change()
If Flag Then
    Dim sqlText As String
    If Trim(TxtHomePhone.Text) = "" Then
        TxtHomePhone.Tag = 0
        Grid.Visible = False
        Exit Sub
        End If
        sqlText = "Select top 10 [adhamno] , [AdhamName] , [AdhamPhon] , [MobilePhone] , [WorkPhone] ,  [adhamadress]  From adhamview7 Where   Adhamphon Like '" & TxtHomePhone.Text & "%'"
        FillList sqlText, "AdhamNo", "AdhamName", Grid1, 1
        MoveGrid TxtHomePhone.top + TxtHomePhone.Height, TxtHomePhone.left - 7900, TxtHomePhone.Width + 7900, Grid1
End If
End Sub

Private Sub TxtHomePhone_KeyDown(KeyCode As Integer, Shift As Integer)
Flag = False
If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid1
Flag = True

End Sub
Sub ClearCustomerInfo(Voption As Integer)
        Flag = False
        Select Case Voption
            Case 1
                    TxtHomePhone.Text = ""
            Case 2
                TxtWorkPhone.Text = ""
        End Select
'        TxtMobilePhone.Text = ""
'        TxtWorkPhone.Text = ""
        TxtAddress.Text = ""
        TxtEmail.Text = ""
        TxtZoneName.Tag = ""
        TxtZoneName.Text = ""
        TxtNotes.Text = ""
        Flag = True
End Sub
Private Sub TxtHomePhone_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
With Grid1
    If .Visible Then
        If MsgBox("Â·  —Ìœ «·„Õ«›Ÿ… ⁄·Ï «·»Ì«‰«  «·ﬁœÌ„… √„ ≈⁄ „«œ “»Ê‰ ÃœÌœ", vbQuestion + vbDefaultButton2 + vbYesNo, "„⁄·Ê„«  «·“»Ê‰") = vbYes Then
            'ClearCustomerInfo (2)
            TxtHomePhone.Text = .TextMatrix(.Row, ColCustomerHomephone)
            .Visible = False
            SendKeys "{tab}"
         Else
            TxtCustomer.Tag = .TextMatrix(.Row, ColCustomerId)
            Flag = False
            TxtCustomer.Text = .TextMatrix(.Row, ColCustomerName)
            FillLabels_Controls Val(TxtCustomer.Tag)
            Flag = True
            .Visible = False
         
        End If
    ElseIf Val(TxtCustomer.Tag) <> 0 And .Visible = False Then
        Flag = False
        TxtCustomer.Text = GetCustomerName(TxtCustomer.Tag)
        Flag = True
    Else
        Flag = False
        TxtCustomer.Tag = 0
        TxtCustomer.Text = ""
        Flag = True
    End If
End With
End If
End Sub

Private Sub TxtMobilePhone_Change()
If Flag Then
    Dim sqlText As String
    If Trim(TxtCustomer.Text) = "" Then
        TxtCustomer.Tag = 0
        Grid.Visible = False
        Exit Sub
        End If
        sqlText = "Select top 10 [adhamno] , [AdhamName] , [AdhamPhon] , [MobilePhone] , [WorkPhone] ,  [adhamadress]  From adhamview7 Where   WorkPhone like '" & TxtCustomer.Text & "%' or MobilePhone like '" & TxtMobilePhone.Text & "%'"
        FillList sqlText, "AdhamNo", "AdhamName", Grid1, 1
        MoveGrid TxtMobilePhone.top + sstab1.top + TxtMobilePhone.Height, TxtMobilePhone.left + sstab1.left, TxtMobilePhone.Width, Grid1
End If
End Sub

Private Sub TxtMobilePhone_KeyDown(KeyCode As Integer, Shift As Integer)
Flag = False
If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid1
Flag = True

End Sub

Private Sub TxtMobilePhone_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
With Grid1
    If .Visible Then
        If MsgBox("Â·  —Ìœ «·„Õ«›Ÿ… ⁄·Ï «·»Ì«‰«  «·ﬁœÌ„… √„ ≈⁄ „«œ “»Ê‰ ÃœÌœ", vbQuestion + vbDefaultButton2 + vbYesNo, "„⁄·Ê„«  «·“»Ê‰") = vbYes Then
            Flag = False
                TxtMobilePhone.Text = .TextMatrix(.Row, ColCustomerMobilePhone)
            Flag = True
            .Visible = False
            SendKeys "{tab}"
         Else
            TxtCustomer.Tag = .TextMatrix(.Row, ColCustomerId)
            Flag = False
            TxtCustomer.Text = .TextMatrix(.Row, ColCustomerName)
            FillLabels_Controls Val(TxtCustomer.Tag)
            Flag = True
            .Visible = False
        End If
    ElseIf Val(TxtCustomer.Tag) <> 0 And .Visible = False Then
        Flag = False
        TxtCustomer.Text = GetCustomerName(TxtCustomer.Tag)
        Flag = True
    Else
        Flag = False
        TxtCustomer.Tag = 0
        TxtCustomer.Text = ""
        Flag = True
    End If
End With
End If
End Sub

Private Sub TxtREpNo_Change()
If OkRepNo Then
    TxtREpName = GetRepName(Val(TxtREpNo.Text))
End If
End Sub

Sub FillList(sqlText As String, Field1 As String, Field2 As String, List As VSFlexGrid, ByVal Switch As Integer)
    
    Set rs = de.con.Execute(sqlText)
    If rs.RecordCount > 0 Then
        Set List.DataSource = rs
        FillFormating List, Switch
        List.Row = 1
        List.Col = 1
        List.ColSel = List.Cols - 1
        List.Visible = True
    Else
        List.Rows = 1
        ActiveControl.Tag = 0
        List.Visible = False

    End If
End Sub

Private Sub TxtZoneName_Change()
If Flag Then
    Dim sqlText As String
    If Trim(TxtZoneName.Text) = "" Then
         TxtZoneName.Tag = 0
        Grid.Visible = False
        Exit Sub
        End If
        sqlText = "Select ZoneNo , ZoneName From CoZone Where ZoneName like " & LikeExpression(TxtZoneName.Text)
        FillList sqlText, "ZoneNo", "ZoneName", Grid1, 5
        MoveGrid TxtZoneName.top + sstab1.top + TxtZoneName.Height, TxtZoneName.left + sstab1.left, TxtZoneName.Width, Grid1
End If
End Sub

Private Sub TxtZoneName_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = False
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True
End Sub

Private Sub TxtZoneName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
With Grid1
    If .Visible Then
        TxtZoneName.Tag = .TextMatrix(.Row, ColZoneNo)
        Flag = False
        TxtZoneName.Text = .TextMatrix(.Row, ColZoneName)
        Flag = True
        .Visible = False
    ElseIf Val(TxtZoneName.Tag) = 0 Then
        Flag = False
        TxtZoneName.Tag = 0
        TxtZoneName.Text = ""
        Flag = True
    End If
End With
End If
End Sub
