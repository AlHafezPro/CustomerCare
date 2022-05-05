VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmBills 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "›Ê« Ì— Œœ„… «·„” Â·ﬂ"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   11685
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   2085
      Left            =   6540
      TabIndex        =   18
      Top             =   4050
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
   Begin Crystal.CrystalReport cr1 
      Left            =   4950
      Top             =   2430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2745
      Left            =   30
      TabIndex        =   19
      Top             =   3210
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   4842
      _Version        =   131074
      Begin VB.ComboBox ComboDiscount 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmBills.frx":0000
         Left            =   1650
         List            =   "FrmBills.frx":0002
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   270
         Width           =   1515
      End
      Begin VB.TextBox TxtQty 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   270
         Width           =   855
      End
      Begin VB.TextBox TxtitemName 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   9120
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   270
         Width           =   2505
      End
      Begin VSFlex8Ctl.VSFlexGrid flexGrid 
         Height          =   1965
         Left            =   90
         TabIndex        =   20
         Top             =   720
         Width           =   11565
         _cx             =   20399
         _cy             =   3466
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
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "”⁄— «·„” Â·ﬂ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   660
         TabIndex        =   30
         Top             =   30
         Width           =   960
      End
      Begin VB.Label LPrice 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   300
         Width           =   1515
      End
      Begin VB.Label LBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   4590
         TabIndex        =   28
         Top             =   270
         Width           =   705
      End
      Begin VB.Label LItemName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5310
         TabIndex        =   27
         Top             =   270
         Width           =   3765
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·ﬂ„Ì…"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   4140
         TabIndex        =   26
         Top             =   30
         Width           =   405
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·‘—Õ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   8670
         TabIndex        =   25
         Top             =   30
         Width           =   420
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·—’Ìœ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   4620
         TabIndex        =   24
         Top             =   30
         Width           =   435
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·—ﬁ„ «·„Œ“‰Ì"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   10665
         TabIndex        =   23
         Top             =   30
         Width           =   930
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·Õ”„"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   23
         Left            =   3330
         TabIndex        =   22
         Top             =   30
         Width           =   405
      End
      Begin VB.Label LDiscount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3210
         TabIndex        =   21
         Top             =   300
         Width           =   495
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3165
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   5583
      _Version        =   131074
      Begin VB.TextBox TxtModelName 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   4830
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   6795
      End
      Begin VB.TextBox TxtClientName 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   270
         Width           =   2925
      End
      Begin VB.CheckBox ChkBill 
         Alignment       =   1  'Right Justify
         Caption         =   "›« Ê—… „ƒﬁ Â"
         Height          =   375
         Left            =   3390
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1560
         Width           =   1755
      End
      Begin VB.TextBox txtModelQty 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   960
         Width           =   675
      End
      Begin VB.CommandButton CmdNewCustomer 
         Caption         =   "“»‹‹‹‹‹‹«∆‰"
         Height          =   345
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   270
         Width           =   1185
      End
      Begin VB.TextBox TxtFeesQty 
         Alignment       =   2  'Center
         Height          =   405
         Left            =   6930
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1530
         Width           =   885
      End
      Begin VB.TextBox TxtOtherFeesPrice 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   2730
         Width           =   615
      End
      Begin VB.TextBox TxtFeesDescription 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   2730
         Width           =   1815
      End
      Begin VB.TextBox TxtOthersFeesQty 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1860
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   2730
         Width           =   465
      End
      Begin MSMask.MaskEdBox TxtDate 
         Height          =   345
         Left            =   9630
         TabIndex        =   1
         Top             =   270
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo ComboType 
         Height          =   360
         Left            =   7290
         TabIndex        =   3
         Top             =   270
         Width           =   1305
         _ExtentX        =   2302
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
      Begin MSDataListLib.DataCombo ComboFees 
         Height          =   360
         Left            =   7830
         TabIndex        =   8
         Top             =   1560
         Width           =   3795
         _ExtentX        =   6694
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
      Begin MSDataListLib.DataCombo ComboOperationType 
         Height          =   360
         Left            =   8640
         TabIndex        =   2
         Top             =   270
         Width           =   975
         _ExtentX        =   1720
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
      Begin MSDataListLib.DataCombo ComboPayment 
         Height          =   360
         Left            =   5970
         TabIndex        =   4
         Top             =   270
         Width           =   1305
         _ExtentX        =   2302
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
      Begin MSMask.MaskEdBox TxtFixBillDate 
         Height          =   345
         Left            =   2220
         TabIndex        =   10
         Top             =   1590
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VSFlex8Ctl.VSFlexGrid FlexFees 
         Height          =   1155
         Left            =   5160
         TabIndex        =   35
         Top             =   1950
         Width           =   6495
         _cx             =   11456
         _cy             =   2037
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
      Begin MSDataListLib.DataCombo ComboDestination 
         Height          =   360
         Left            =   4830
         TabIndex        =   5
         Top             =   270
         Width           =   1125
         _ExtentX        =   1984
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·ÃÂ‹‹‹‹…"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   5340
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   30
         Width           =   540
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·√ÃÊ—"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   11130
         TabIndex        =   59
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label LClientType 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1260
         TabIndex        =   58
         Top             =   270
         Width           =   615
      End
      Begin VB.Label LSymbol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   990
         Width           =   2895
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„ÊœÌ·"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   10980
         TabIndex        =   56
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Â« › /  “»Ê‰"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   0
         Left            =   3810
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   30
         Width           =   900
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "‰Ê⁄ «·’Ì«‰…"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   4
         Left            =   7815
         TabIndex        =   54
         Top             =   30
         Width           =   780
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «·›« Ê—…"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   9780
         TabIndex        =   53
         Top             =   30
         Width           =   960
      End
      Begin VB.Label LBillNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   10770
         TabIndex        =   52
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·›« Ê—…"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   10725
         TabIndex        =   51
         Top             =   30
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "‰Ê⁄ «·⁄„·Ì…"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   6
         Left            =   8865
         TabIndex        =   50
         Top             =   30
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   3
         Left            =   6450
         TabIndex        =   49
         Top             =   30
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·⁄œœ"
         Height          =   195
         Index           =   3
         Left            =   825
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   960
         Width           =   330
      End
      Begin VB.Label LFixBillDateCaption 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «· À»Ì "
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2430
         TabIndex        =   47
         Top             =   1380
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label LClientPhoneNBR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1890
         TabIndex        =   46
         Top             =   660
         Width           =   2895
      End
      Begin VB.Label LTransfered 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   60
         TabIndex        =   45
         Top             =   1590
         Width           =   2115
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "«·⁄œœ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   16
         Left            =   7485
         TabIndex        =   44
         Top             =   1320
         Width           =   330
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "√ÃÊ— „Œ ·›…"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   17
         Left            =   4260
         TabIndex        =   43
         Top             =   2730
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·‘‹‹‹‹‹‹‹—Õ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   18
         Left            =   3390
         TabIndex        =   42
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·”⁄—"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   19
         Left            =   1455
         TabIndex        =   41
         Top             =   2400
         Width           =   390
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·⁄œœ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   20
         Left            =   2010
         TabIndex        =   40
         Top             =   2400
         Width           =   330
      End
      Begin VB.Label LOthersFees 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   45
         TabIndex        =   39
         Top             =   2760
         Width           =   1125
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "≈Ã„«·Ì «·”⁄—"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   22
         Left            =   180
         TabIndex        =   38
         Top             =   2400
         Width           =   930
      End
      Begin VB.Label LFeesAmount 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1560
         Width           =   1725
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "«·”⁄— «·«›—«œÌ ··„” Â·ﬂ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   5130
         TabIndex        =   36
         Top             =   1320
         Width           =   1665
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   495
      Left            =   0
      TabIndex        =   60
      Top             =   6000
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   873
      _Version        =   131074
      Begin Threed.SSCommand CmdSearch 
         Height          =   435
         Left            =   1395
         TabIndex        =   67
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "»ÕÀ"
      End
      Begin Threed.SSCommand CmdAdd 
         Height          =   435
         Left            =   10320
         TabIndex        =   0
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÃœÌœ"
      End
      Begin Threed.SSCommand CmdDelete 
         Height          =   435
         Left            =   7650
         TabIndex        =   66
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Õ–›"
      End
      Begin Threed.SSCommand CmdCancel 
         Height          =   435
         Left            =   4980
         TabIndex        =   65
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
         _Version        =   131074
         ForeColor       =   8388608
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " —«Ã⁄"
      End
      Begin Threed.SSCommand CmdEdit 
         Height          =   435
         Left            =   8985
         TabIndex        =   64
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
         _Version        =   131074
         ForeColor       =   8388608
         PictureAnimationDelay=   66
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " ⁄œÌ·"
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   435
         Left            =   6315
         TabIndex        =   17
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
         _Version        =   131074
         ForeColor       =   8388608
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Õ›Ÿ"
      End
      Begin Threed.SSCommand CmdExit 
         Height          =   435
         Left            =   60
         TabIndex        =   63
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
         _Version        =   131074
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Œ—ÊÃ"
      End
      Begin Threed.SSCommand CmdPrint 
         Height          =   435
         Left            =   2730
         TabIndex        =   62
         Top             =   30
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   767
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÿ»«⁄…"
      End
      Begin Threed.SSCommand CmdPreview 
         Height          =   435
         Left            =   3855
         TabIndex        =   61
         Top             =   30
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   767
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "„⁄«Ì‰…"
      End
   End
   Begin Threed.SSFrame SSFrame10 
      Height          =   375
      Left            =   9270
      TabIndex        =   68
      Top             =   6510
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   661
      _Version        =   131074
      Begin VB.CommandButton CmdLast 
         Height          =   285
         Left            =   2190
         Picture         =   "FrmBills.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   72
         TabStop         =   0   'False
         ToolTipText     =   "Last"
         Top             =   60
         Width           =   255
      End
      Begin VB.CommandButton CmdNext 
         Height          =   285
         Left            =   1920
         Picture         =   "FrmBills.frx":0536
         Style           =   1  'Graphical
         TabIndex        =   71
         TabStop         =   0   'False
         ToolTipText     =   "Next"
         Top             =   60
         Width           =   255
      End
      Begin VB.CommandButton CmdPrevious 
         Height          =   285
         Left            =   330
         Picture         =   "FrmBills.frx":0630
         Style           =   1  'Graphical
         TabIndex        =   70
         TabStop         =   0   'False
         ToolTipText     =   "Previous"
         Top             =   60
         Width           =   255
      End
      Begin VB.CommandButton CmdFirst 
         Height          =   285
         Left            =   60
         Picture         =   "FrmBills.frx":072A
         Style           =   1  'Graphical
         TabIndex        =   69
         TabStop         =   0   'False
         ToolTipText     =   "First"
         Top             =   60
         Width           =   255
      End
      Begin VB.Label LNavigator 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   600
         TabIndex        =   73
         Top             =   60
         Width           =   1305
      End
   End
   Begin VB.Label LTotal 
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
      Height          =   345
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   81
      Top             =   6510
      Width           =   1545
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "«·≈Ã„«·Ì"
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
      Height          =   435
      Index           =   14
      Left            =   1710
      TabIndex        =   80
      Top             =   6510
      Width           =   825
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·√ﬁ·«„"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   15
      Left            =   8745
      TabIndex        =   79
      Top             =   6540
      Width           =   480
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
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   78
      Top             =   6510
      Width           =   615
   End
   Begin VB.Label LTotalFees 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6690
      TabIndex        =   77
      Top             =   6510
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·√ÃÊ—"
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
      Left            =   7590
      TabIndex        =   76
      Top             =   6510
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ﬁÌ„… «·„Ê«œ"
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
      Left            =   5745
      TabIndex        =   75
      Top             =   6510
      Width           =   885
   End
   Begin VB.Label LTotalItems 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4740
      TabIndex        =   74
      Top             =   6510
      Width           =   975
   End
End
Attribute VB_Name = "FrmBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RsNavigator As New ADODB.Recordset
Dim Ok As Boolean, Flag As Boolean
Dim Pos As Integer, RecNum   As Integer
Dim TypeRec As Boolean


Const ColNo = 1
Const ColName = 2
Const col3 = 3
Const Col4 = 4

Const ColId = 1
Const ColStkNo = 2
Const ColStkName = 3

Const ColQty = 4
Const ColPrice = 5
'Const colPriceTypeId = 6
'Const ColPriceTypeName = 7
Const ColDiscount = 6

Const ColFeesSer = 1
Const ColFeesBillNo = 2
Const ColFeesTypeName = 3
Const ColFeesQty = 4
'Const ColFeesCliPrice = 5
'Const ColFeesPriceName = 5
Const ColFeesAmount = 5


Dim MvMaintPaymentRec As MvMaintPaymentRecTYpe, MvMaintPaymentRecTypeDetailsRec As MvMaintPaymentRecTypeDetails

Dim BeforeOperationTYpeID As Integer, CurrentOperationTYpeID   As Integer

Function GetFeesAmount() As Double
On Error GoTo errorhandler
Dim FeesAmount As Double
FeesAmount = 0
With FlexFees
    For i = 1 To .Rows - 1
        FeesAmount = FeesAmount + Val(.TextMatrix(i, ColFeesAmount) * Val(.TextMatrix(i, ColFeesQty)))
    Next
End With
GetFeesAmount = FeesAmount
Exit Function
errorhandler:
GetFeesAmount = 0
End Function

Function GetTotalPrice(IDBill As Double, Vindex As Integer) As Double
On Error GoTo errorhandler
Dim rsTotalPrice As New ADODB.Recordset
sqlText = "select sum(TotPrice) as TotPrice From MvMaintPaymentsQry Where Stat<>6 And billno = " & IDBill
If Vindex = 0 Then
    sqlText = sqlText & " and row=1"
End If
Set rsTotalPrice = de.con.Execute(sqlText)
'GetTotalPrice = rsTotalPrice!TotPrice + IIf(vindex = 0, 0, GetFeesAmount)
GetTotalPrice = rsTotalPrice!TotPrice
Exit Function
errorhandler:
GetTotalPrice = -1
End Function

Function GetTotalPriceBeforDiscount(IDBill As Double, Vindex As Integer) As Double
On Error GoTo errorhandler
Dim rsTotalPrice As New ADODB.Recordset
sqlText = "select isnull(sum(isnull(price,0) *isnull(qty,0)),0) as TotPrice From MvMaintPaymentsQry where Stat<>6 and billno = " & IDBill
If Vindex = 0 Then
    sqlText = sqlText & " and row=1"
End If
Set rsTotalPrice = de.con.Execute(sqlText)
'GetTotalPrice = rsTotalPrice!TotPrice + IIf(vindex = 0, 0, GetFeesAmount)
GetTotalPriceBeforDiscount = rsTotalPrice!TotPrice
Exit Function
errorhandler:
GetTotalPriceBeforDiscount = -1
End Function

Function ChkBalance(stkno As String) As Boolean
On Error GoTo errohandler
ChkBalance = False
Dim rsStkNoTeamBalance  As New ADODB.Recordset
        sqlText = "select isnull(sum(fnlqnt),0)fnlqnt from stkinf s1 where stkno='" & stkno & "' and strid=" & GetStrId(systemConfigration.MainStoreNo)
        Set rsStkNoTeamBalance = de.con.Execute(sqlText)
        If rsStkNoTeamBalance!fnlqnt >= 0 Then
            ChkBalance = True
        End If
Exit Function
errohandler:
ChkBalance = False
MsgBox Err.Description
End Function

Function DeleteRec(BillNo As Double) As Boolean
On Error GoTo errorhandler
de.con.BeginTrans
    With FlexGrid
     For i = 1 To .Rows - 1
          RemoveFromMvStock .TextMatrix(i, ColId)
          If Not ChkBalance(.TextMatrix(i, ColStkNo)) Then
             de.con.RollbackTrans
             MsgBox "·«Ì„ﬂ‰ «·Õ–› ·«‰ «·—’Ìœ «·„” Êœ⁄ ·„ Ì⁄œ ﬂ«›Ì«", vbExclamation, "«·—’Ìœ ”«·»"
             Exit Function
          End If
     Next
    End With
    sqlText = "Delete From MvMaintFees Where BillNo=" & BillNo
    sqlText = sqlText & " Delete From MvMaintPaymentsDetails Where BillNo=" & BillNo
    sqlText = sqlText & " Delete From MvMaintPayments Where BillNo=" & BillNo
    de.con.Execute (sqlText)
de.con.CommitTrans
DeleteRec = True
Exit Function
errorhandler:
DeleteRec = False
MsgBox Err.Description
de.con.RollbackTrans
End Function

Function MaxRec() As Double
Dim RsMax As New ADODB.Recordset
    sqlText = "Select isnull(Max(BillNo),0)MaxBillNo From MvMaintPayments"
    Set RsMax = de.con.Execute(sqlText)
    MaxRec = RsMax!MaxBillNo
End Function

Function InsertNewClient(ClientName As String) As Double
    On Error GoTo errorhandler
    Dim RsMaxi As New ADODB.Recordset
    sqlText = "Insert Into Coclient (ClientName) Values('" & ClientName & "')"
    de.con.Execute (sqlText)
    sqlText = "select max(ClientId) MaxClientId From CoClient"
    Set RsMaxi = de.con.Execute(sqlText)
    InsertNewClient = RsMaxi!MaxClientId
    Exit Function
errorhandler:
    InsertNewClient = 0
    MsgBox Err.Description
End Function

Sub ClearControls()
Ok = False
    LBillNo.Caption = ""
    LBillNo.Tag = ""
    TxtDate.Text = Format(Now, "dd/mm/yyyy")
    TxtFixBillDate.Text = Format(Now, "dd/mm/yyyy")
    ChkBill.Value = 0
'    ComboOperationType.BoundText = ""
'    ComboPayment.BoundText = ""
'    ComboType.BoundText = ""
'    TxtClientName.Tag = 0
'    TxtClientName.Text = ""
'    txtClientPhoneNBR.Text = ""
'    LClientType.Caption = ""
'    TxtModelName.Tag = 0
'    TxtModelName.Text = ""
'    LSymbol.Caption = ""
'    TxtDollar.Text = ""
    txtModelQty.Text = 1
    TxtFeesDescription.Text = ""
    TxtOtherFeesPrice.Text = ""
    TxtOthersFeesQty.Text = ""
    TxtItemName.Tag = "0"
    TxtItemName.Text = ""
    LItemName.Caption = ""
    LBalance.Caption = ""
    TxtQty.Text = ""
    'ComboCurrencyType.BoundText = ""
    ComboFees.BoundText = ""
    'ComboFeesPriceType.BoundText = ""
    TxtFeesQty.Text = 1
    LFeesAmount.Caption = ""
    'LfeesClientPrice.Caption = ""
    LPrice.Caption = ""
    LTotal.Caption = ""
    LDiscount.Caption = ""
    LCount.Caption = ""
    LTotalItems.Caption = ""
    LTotalFees.Caption = ""
    LTransfered.Caption = ""
    FlexGrid.Rows = 1
    FlexFees.Rows = 1
    
    
    With MvMaintPaymentRecTypeDetailsRec
        .BillNo = 0
        .stkno = ""
        .Qty = 0
        .PriceTYpe = 0
        .Price = 0
    End With
Ok = True
End Sub

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
    fs = "|>" + ""
    fs = fs + "|>" + ""
    fs = fs + "|>" + ""
    fs = fs + "|>" + ""
    With Grid
        .FormatString = fs
        .Cols = 5
        If Pos = 3 Then
            SetColWidths ColNo, Grid
        Else
            .ColWidth(ColNo) = 0
        End If
        If Pos = 1 Or Pos = 2 Then
            SetColWidths col3, Grid
        Else
            .ColWidth(col3) = 0
        End If
        SetColWidths ColName, Grid
        .ColWidth(Col4) = 0
    End With
ElseIf i = 2 Then
    fs = "|>" + "Id"
    fs = fs + "|>" + "«·—ﬁ„ «·„Œ“‰Ì"
    fs = fs + "|>" + "«·‘—Õ"
    fs = fs + "|>" + "«·ﬂ„Ì…"
    fs = fs + "|>" + "«·”⁄— «·≈›—«œÌ ··„” Â·ﬂ"
'    fs = fs + "|>" + "—ﬁ„ ‰Ê⁄ «·”⁄—"
'    fs = fs + "|>" + "‰Ê⁄ «·”⁄—"
    fs = fs + "|>" + "‰”»… «·Õ”„"
    With FlexGrid
        .FormatString = fs
        .Cols = 7
        .ColWidth(ColId) = 0
        SetColWidths ColStkNo, FlexGrid
        SetColWidths ColStkName, FlexGrid
        SetColWidths ColQty, FlexGrid
        SetColWidths ColPrice, FlexGrid
'        .ColWidth(colPriceTypeId) = 0
'        SetColWidths ColPriceTypeName, FlexGrid
        SetColWidths ColDiscount, FlexGrid
'        .ColWidth(ColPaymentTypeId) = 0
'        SetColWidths ColPaymentTypeName, FlexGrid
    End With
ElseIf i = 3 Then
    fs = "|>" + "Id"
    fs = fs + "|>" + "—ﬁ„ «·›« Ê—…"
    fs = fs + "|>" + "‰Ê⁄ «·≈ÃÊ—"
    fs = fs + "|>" + "«·⁄œœ"
'    fs = fs + "|>" + "‰Ê⁄ «·”⁄—"
    fs = fs + "|>" + "«·”⁄— «·≈›—«œÌ ··„” Â·ﬂ"
    With FlexFees
        .FormatString = fs
        .Cols = 6
        .ColWidth(ColFeesSer) = 0
        .ColWidth(ColFeesBillNo) = 0
'        SetColWidths ColFeesSer, FlexFees
'        SetColWidths ColFeesBillNo, FlexFees
        SetColWidths ColFeesTypeName, FlexFees
        SetColWidths ColFeesQty, FlexFees
'        SetColWidths ColFeesPriceName, FlexFees
        SetColWidths ColFeesAmount, FlexFees
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
    With txtClientName
       Grid.top = SSFrame1.top + .top + .Height
       Grid.left = SSFrame1.left + .left
       Grid.Width = .Width
    End With
ElseIf X = 2 Then
    With TxtModelName
       Grid.top = SSFrame1.top + .top + .Height
       Grid.left = SSFrame1.left + .left
       Grid.Width = .Width
End With
ElseIf X = 3 Then
    With TxtItemName
       Grid.top = SSFrame2.top + .top + .Height
       Grid.left = SSFrame2.left + .left
       Grid.Width = .Width
End With
ElseIf X = 4 Then
    With TxtRecipient
       Grid.top = sstab1.top + SSFrame2.top + .top + .Height
       Grid.left = SSFrame2.left + .left
       Grid.Width = .Width
End With
ElseIf X = 5 Then
    With TxtFamNo
       Grid.top = SSFrameModel.top + sstab1.top + .top + .Height
       Grid.left = SSFrameModel.left + .left
       Grid.Width = .Width
End With
ElseIf X = 6 Then
    With TxtModelName
       Grid.top = .top + .Height
       Grid.left = .left
       Grid.Width = .Width
End With
ElseIf X = 7 Then
    With TxtUnExecutedReason
       Grid.top = sstab1.top + .top + .Height
       Grid.left = .left
       Grid.Width = .Width
End With
ElseIf X = 8 Then
    With TxtError
       Grid.top = .top + .Height
       Grid.left = .left
       Grid.Width = .Width
End With
ElseIf X = 9 Then
    With TxtItemName
       Grid.top = .top + .Height
       Grid.left = .left
       Grid.Width = .Width
End With
ElseIf X = 10 Then
    With txtCallNo
       Grid.top = sstab1.top + .top + .Height
       Grid.left = sstab1.left + .left
       Grid.Width = .Width
End With
End If
End Sub

Sub EnableCmds(FAdd As Boolean, FEdit As Boolean, FDelete As Boolean, FSave As Boolean, FUndo As Boolean, FFirst As Boolean, FNext As Boolean, FPrevious As Boolean, FLast As Boolean, FPreviow As Boolean, Fprint As Boolean)
    CmdAdd.Enabled = FAdd
    CmdEdit.Enabled = FEdit
    CmdDelete.Enabled = FDelete
    cmdSave.Enabled = FSave
    CmdCancel.Enabled = FUndo
    CmdFirst.Enabled = FFirst
    CmdLast.Enabled = FLast
    CmdNext.Enabled = FNext
    CmdPrevious.Enabled = FPrevious
    CmdPreview.Enabled = FPreviow
    cmdPrint.Enabled = Fprint
    
End Sub

Sub EnableControls(FControl As Boolean)
Dim ctrl As Control
For Each ctrl In Me.Controls
    If TypeOf ctrl Is TextBox Or TypeOf ctrl Is MaskEdBox Or TypeOf ctrl Is CheckBox Or TypeOf ctrl Is VSFlexGrid Or TypeOf ctrl Is DataCombo Or TypeOf ctrl Is ComboBox Then
        ctrl.Enabled = FControl
    End If
Next
End Sub


Sub FillCombos()
    
    Dim RsOperationType As New ADODB.Recordset
    If OperationEmpStr = "" Or PaymentEmpStr = "" Or MaintTYpeEmpStr = "" Then Exit Sub
    sqlText = "select OpNo , OpName  from operkind Where OpNo in (" & OperationEmpStr & ")"
    Set RsOperationType = de.con.Execute(sqlText)
    Set ComboOperationType.RowSource = RsOperationType
    ComboOperationType.listField = "OpName"
    ComboOperationType.BoundColumn = "OpNo"
    
    
    Dim rsPayment As New ADODB.Recordset
    sqlText = "Select No , Name  From PayMethod Where No in (" & PaymentEmpStr & ")"
    Set rsPayment = de.con.Execute(sqlText)
    Set ComboPayment.RowSource = rsPayment
    ComboPayment.listField = "Name"
    ComboPayment.BoundColumn = "No"
    
    Dim rsType As New ADODB.Recordset
    sqlText = "select no , stat from dbo.maintypestat where No in (" & MaintTYpeEmpStr & ")"
    Set rsType = de.con.Execute(sqlText)
    Set ComboType.RowSource = rsType
    ComboType.listField = "Stat"
    ComboType.BoundColumn = "No"
    
    
    Dim RsDestination As New ADODB.Recordset
    sqlText = "Select Id , Destination From CoDestination"
    Set RsDestination = de.con.Execute(sqlText)
    Set ComboDestination.RowSource = RsDestination
    ComboDestination.listField = "Destination"
    ComboDestination.BoundColumn = "Id"
    
    Dim rsFees As New ADODB.Recordset
    sqlText = "Select FeesId , FeesName  From  CoMaintFees Where isnull(CliPriceafterdiscount,0) <> 0 or isnull(DealPriceafterdiscount,0) <> 0 or isnull(DistPriceafterdiscount,0) <> 0  "
    Set rsFees = de.con.Execute(sqlText)
    Set ComboFees.RowSource = rsFees
    ComboFees.listField = "FeesName"
    ComboFees.BoundColumn = "FeesId"

'    Dim rsCurrency As New ADODB.Recordset
'    sqlText = "select PriceNo , PriceTYpe , col   from dbo.PriceTypes where PriceNo in (3)" ' 1,2
'    Set rsCurrency = de.con.Execute(sqlText)
'    Set ComboCurrencyType.RowSource = rsCurrency
'    ComboCurrencyType.listField = "PriceTYpe"
'    ComboCurrencyType.BoundColumn = "PriceNo"
    
'        Set ComboFeesPriceType.RowSource = rsCurrency
'    ComboFeesPriceType.listField = "PriceTYpe"
'    ComboFeesPriceType.BoundColumn = "PriceNo"
End Sub
Sub InitNavigator()
    If OperationEmpStr = "" Or PaymentEmpStr = "" Or MaintTYpeEmpStr = "" Then Exit Sub
    sqlText = "Select   billno , billdate , isnull(FixBillDate,'') as FixBillDate, OperationTYpe   , mainttype , PaymentTYpeId , DestinationId , clientid , class , modno, modelQty , FeesDescription , OtherFeesQty  , OtherFeesPrice , OtherFeesAmount , IsFixed , IsTransfered  from mvmaintpayments Where OperationTYpe in (" & OperationEmpStr & ") and  mainttype in(" & MaintTYpeEmpStr & ") order by billno"
    Set RsNavigator = de.con.Execute(sqlText)
End Sub
Function GetModelName(ModNo As Integer) As String
On Error GoTo errorhandler
Dim RsModelName As New ADODB.Recordset

sqlText = "Select Name    from adhammodels  Where ModNo=" & ModNo
Set RsModelName = de.con.Execute(sqlText)
If RsModelName.RecordCount > 0 Then
    GetModelName = RsModelName!name
Else
    GetModelName = ""
End If
Exit Function
errorhandler:
GetModelName = ""
End Function

Function GetSymbol(ModNo As Integer) As String
On Error GoTo errorhandler
Dim RsSymbol As New ADODB.Recordset

sqlText = "Select Symbol    from adhammodels   Where ModNo=" & ModNo
Set RsSymbol = de.con.Execute(sqlText)
If RsSymbol.RecordCount > 0 Then
    GetSymbol = RsSymbol!Symbol
Else
    GetSymbol = ""
End If
Exit Function
errorhandler:
GetSymbol = ""
End Function
Sub FillGrid(BillNo As Double)

On Error GoTo errorhandler
    Dim RsDetails As New ADODB.Recordset
    sqlText = "Select Id , StkNo , StkName , Qty , Price , Discount  From MvMaintPaymentsDetailsQry Where BillNo=" & BillNo
    Set RsDetails = de.con.Execute(sqlText)
    Set FlexGrid.DataSource = RsDetails
    FillFormating 2
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

Sub FillFees(BillNo As Double)
Dim rs As New ADODB.Recordset
'Sqltext = "select Id , BillNo , FeesTypeId , FeesName , FeesPriceTYpe , PriceType ,  FeesAmount     from mvmaintfeesQry "
sqlText = "select Id , BillNo , FeesTypeId , FeesQty  ,  FeesAmount     from mvmaintfeesQry where BillNo= " & BillNo & " and FeesTypeId <> 0 "
Set rs = de.con.Execute(sqlText)
Set FlexFees.DataSource = rs
FillFormating 3
End Sub
Sub FillControlsFromSql(rs As Recordset)

Ok = False
With rs
    CurrentOperationTYpeID = 0
    BeforeOperationTYpeID = 0
    LBillNo.Tag = rs!BillNo
    LBillNo.Caption = rs!BillNo
    TxtDate.Text = ConvertSqlDate(rs!Billdate)
    ChkBill.Value = Val(Abs(rs!IsFixed))
    TxtFixBillDate.Text = ConvertSqlDate(rs!FixBillDate)
    ComboOperationType.BoundText = rs!OperationType
    ComboType.BoundText = rs!MaintType
    ComboPayment.BoundText = IIf(IsNull(rs!PaymentTYpeId), -1, rs!PaymentTYpeId)
    ComboDestination.BoundText = IIf(IsNull(rs!DestinationId), -1, rs!DestinationId)
    
    txtClientName.Tag = rs!clientId
    txtClientName.Text = GetClientName(rs!clientId, rs!Class)
    LClientPhoneNBR.Caption = GetPhoneNbr(rs!clientId, rs!Class)
    LClientType.Tag = rs!Class
    LClientType.Caption = GetClassName(rs!Class)
    
'    TxtDollar.text = IIf(IsNull(rs!Roe), 0, rs!Roe)
    TxtModelName.Tag = rs!ModNo
    TxtModelName.Text = GetModelName(rs!ModNo)
    LSymbol.Caption = GetSymbol(rs!ModNo)
    txtModelQty.Text = IIf(IsNull(rs!ModelQty), 0, rs!ModelQty)
    TxtFeesDescription.Text = rs!FeesDescription
    TxtOthersFeesQty.Text = IIf(IsNull(rs!OtherFeesQty), 0, rs!OtherFeesQty)
    TxtOtherFeesPrice.Text = IIf(IsNull(rs!OtherFeesPrice), 0, rs!OtherFeesPrice)
    LOthersFees.Caption = IIf(rs!OtherFeesAmount = 0, "", rs!OtherFeesAmount)
    
'    ComboFees.BoundText = Rs!FeesTYpeId
'    ComboFeesPriceType.BoundText = Rs!FeesPriceType
'    LFeesAmount.Caption = Rs!FeesAmount
    
    LTransfered.Caption = IIf(rs!IsTransfered = False, "", "„—Õ‹‹‹‹· ≈·Ï «·„Õ«”‹‹‹‹»…")
    LTransfered.Tag = rs!IsTransfered
    If ChkBill.Value = 0 Then
        TxtFixBillDate.Visible = False
        TxtFixBillDate.Visible = False
    Else
        TxtFixBillDate.Visible = True
        TxtFixBillDate.Visible = True
    End If
    FillFees rs!BillNo
    FillGrid rs!BillNo
    LTotalFees.Caption = GetFeesAmount
    LTotalItems.Caption = GetTotalPrice(rs!BillNo, 0)
    LTotal.Caption = GetTotalPrice(rs!BillNo, 1)
    'LTotalBeforDiscount.Caption = GetTotalPriceBeforDiscount(Rs!BillNo, 1)
    LCount.Caption = FlexGrid.Rows - 1
    
    'TxtitemName.Tag = 0
    'TxtitemName.Text = ""
    'LItemName.Caption = ""
    'LBalance.Caption = ""
    'TxtQty.Text = ""
    'ComboCurrencyType.BoundText = ""
    'LPrice.Caption = ""
End With
Ok = True
End Sub

Sub MoveNavigator(ByVal i As Integer)
Dim RSTemp As New ADODB.Recordset
On Error GoTo errorhandler
If OperationEmpStr = "" Or PaymentEmpStr = "" Or MaintTYpeEmpStr = "" Then Exit Sub
If RsNavigator.RecordCount = 0 Then Exit Sub
Select Case i
    Case 1 'First
        RsNavigator.MoveFirst
        RecNum = 1
    Case 2 'Previous
        RsNavigator.MovePrevious
        If RsNavigator.BOF Then
            RsNavigator.MoveFirst
            RecNum = 1
        Else
            RecNum = RecNum - 1
        End If
    Case 3 'Next
        RsNavigator.MoveNext
        If RsNavigator.EOF Then
            RsNavigator.MoveLast
            RecNum = RsNavigator.RecordCount
        Else
            RecNum = RecNum + 1
        End If
    Case 4 'Last
        RsNavigator.MoveLast
        RecNum = RsNavigator.RecordCount
End Select
FillControlsFromSql RsNavigator
LNavigator.Caption = LTrim(RTrim(Str(RecNum))) & "/" & LTrim(RTrim(Str(RsNavigator.RecordCount)))
Exit Sub
errorhandler:
MsgBox "Error In Navigator"
End Sub
Sub MoveToRec(IDBill As Double)
With RsNavigator
If RsNavigator.RecordCount = 0 Then
    ClearControls
    RecNum = 0
    LNavigator.Caption = ""
    Exit Sub
End If
.MoveLast
RecNum = RsNavigator.RecordCount
    Do While Not .BOF
       If !BillNo <> IDBill Then
            .MovePrevious
            RecNum = RecNum - 1
        Else
            FillControlsFromSql RsNavigator
            LNavigator.Caption = LTrim(RTrim(Str(RecNum))) & "/" & LTrim(RTrim(Str(RsNavigator.RecordCount)))
            Exit Sub
       End If
    Loop
End With
End Sub
Sub FillCombosInGrid(FlexGrid As VSFlexGrid, ByVal Col As Integer, Vindex As Integer)
Dim RsClass As New ADODB.Recordset
Dim Lst As String
    Select Case Vindex
        Case 1
            sqlText = "select PriceNo , PriceTYpe , col   from dbo.PriceTypes where PriceNo in (1,2,3)"
            Set RsClass = de.con.Execute(sqlText)
            If RsClass.RecordCount > 0 Then
                With FlexGrid
                    Lst = .BuildComboList(RsClass, "PriceTYpe", "PriceNo", vbYellow)
                    .ColComboList(Col) = Lst
                End With
            Else
                With FlexGrid
                    .Rows = 1
                End With
            End If
        Case 2
            sqlText = "Select No , Name  From PayMethod"
            Set RsClass = de.con.Execute(sqlText)
            If RsClass.RecordCount > 0 Then
                With FlexGrid
                    Lst = .BuildComboList(RsClass, "Name", "No", vbYellow)
                    .ColComboList(Col) = Lst
                End With
            Else
                With FlexGrid
                    .Rows = 1
                End With
            End If
      Case 3
        sqlText = "Select FeesId , FeesName  From  CoMaintFees Where isnull(CliPriceafterdiscount,0) <> 0 or isnull(DealPriceafterdiscount,0) <> 0 or isnull(DistPriceafterdiscount,0) <> 0  "
        Set RsClass = de.con.Execute(sqlText)
        If RsClass.RecordCount > 0 Then
            With FlexGrid
                Lst = .BuildComboList(RsClass, "FeesName", "FeesId", vbYellow)
                .ColComboList(Col) = Lst
            End With
        Else
            With FlexGrid
                .Rows = 1
            End With
        End If
    End Select
End Sub
Sub init()
    top = 0
    left = 0
    FillCombos
    InitNavigator
     EnableControls False
    Ok = True
    FlexGrid.Rows = 1
    FlexFees.Rows = 1
    
    If LoadForm Then
        LoadForm = False
        MoveToRec IDBill
    Else
        MoveNavigator 4    'Move Last
    End If
    
    FillCombosInGrid FlexGrid, ColPriceTypeName, 1
    FillCombosInGrid FlexGrid, ColPaymentTypeName, 2
    FillCombosInGrid FlexFees, ColFeesPriceName, 1
    FillCombosInGrid FlexFees, ColFeesTypeName, 3
    FlexGrid.Editable = flexEDKbdMouse
    FlexFees.Editable = flexEDKbdMouse
    FillFormating 2
    FillFormating 3

End Sub

Private Sub ChkBill_Click()
If ChkBill.Value Then
    ChkBill.Caption = "›« Ê—… „À» …"
    TxtFixBillDate.Text = Format(Now, "dd/mm/yyyy")
    TxtFixBillDate.Visible = True
Else
    ChkBill.Caption = "›« Ê—… „ƒﬁ Â"
    TxtFixBillDate.Visible = False
End If
End Sub

Private Sub ChkBill_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ChkBill.Value Then
        TxtFixBillDate.SetFocus
        SendKeys "{home}+{end}"
    Else
        TxtFeesDescription.SetFocus
        TxtFeesDescription.SelStart = 0
        TxtFeesDescription.SelLength = Len(TxtFeesDescription.Text)
    End If
End If
End Sub

Private Sub CmdAdd_Click()
TypeRec = True
EnableCmds False, False, False, True, True, False, False, False, False, False, False
EnableControls True
ClearControls
TxtDate.SetFocus
SendKeys "{home}+{end}"
End Sub

Private Sub CmdCancel_Click()
On Error GoTo errorhandler
    EnableCmds True, True, True, False, False, True, True, True, True, True, True
    EnableControls False
    If RsNavigator!BillNo = Null Then
        MoveToRec Val(LBillNo.Tag)
    Else
        MoveToRec Val(RsNavigator!BillNo)
    End If
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

Private Sub CmdDelete_Click()
On Error GoTo errorhandler
If Not RsNavigator!IsTransfered Then
    If MsgBox("Â· √‰  „ √ﬂœ „‰ Õ–› «·›« Ê—…", vbYesNo + vbDefaultButton2, "Õ–›") = vbYes Then
        If DeleteRec(RsNavigator!BillNo) Then
            InitNavigator
            MoveToRec MaxRec
            EnableCmds True, True, True, False, False, True, True, True, True, True, True
        End If
    End If
Else
    MsgBox "«·›« Ê—… „—Õ·….... ·«Ì„ﬂ‰ﬂ Õ–› «·›« Ê—…", vbInformation, "«·›« Ê—… „—Õ·…"
End If
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub
Sub ClearItems()
        TxtItemName.Tag = "0"
        TxtItemName.Text = ""
        LItemName.Caption = ""
        LBalance.Caption = ""
        TxtQty.Text = ""
        LDiscount.Caption = ""
        ComboDiscount.Clear
End Sub
Private Sub CmdEdit_Click()
    If Not RsNavigator!IsTransfered Then
        ClearItems
        TypeRec = False
        EnableCmds False, False, False, True, True, False, False, False, False, False, False
        EnableControls True
        ComboOperationType.Enabled = GetOperationTypeEnableValue
        CmdNewCustomer.Enabled = True
        TxtDate.SetFocus
        SendKeys "{home}+{end}"
    Else
        MsgBox "«·›« Ê—… „—Õ·….... ·«Ì„ﬂ‰ﬂ «· ⁄œÌ· √Ê «·≈÷«›… ⁄·ÌÂ«", vbInformation, "«·›« Ê—… „—Õ·…"
    End If
End Sub

Function FillVariables(Vindex As Integer) As Boolean
On errro GoTo errorhandler
Select Case Vindex
    Case 1
            If Not IsDate(TxtDate.Text) Or Val(ComboType.BoundText) = 0 Or Val(txtClientName.Tag) = 0 Or Val(ComboOperationType.BoundText) = 0 Or ComboPayment.BoundText = "" Then
                FillVariables = False
                Exit Function
            End If
    Case 2
            If TxtItemName.Tag = "0" Or Val(TxtQty.Text) = 0 Then
                FillVariables = False
                Exit Function
            ElseIf Not isOkQty(TxtItemName.Tag, Val(TxtQty.Text)) Then
                FillVariables = False
            MsgBox "«·ﬂ„ÌÂ √ﬂ»— „‰ «·—’Ìœ", vbCritical, "«·—’Ìœ ·«Ì”„Õ"
                Exit Function
            End If
End Select
FillVariables = True
Exit Function
errorhandler:
FillVariables = False
End Function

Function GetMaxId() As Double
On Error GoTo errorhandler
    Dim RsMax As New ADODB.Recordset
    sqlText = "Select Max(Id) as MaxId From MvMaintPaymentsdetails"
    Set RsMax = de.con.Execute(sqlText)
    GetMaxId = RsMax!maxId
Exit Function
errorhandler:
GetMaxId = -1
End Function

Function GetMaxBillNo() As Double
On Error GoTo errorhandler
    Dim RsMax As New ADODB.Recordset
    sqlText = "Select Max(BillNo) as MaxBillNo From MvMaintPayments"
    Set RsMax = de.con.Execute(sqlText)
    GetMaxBillNo = RsMax!MaxBillNo
Exit Function
errorhandler:
GetMaxBillNo = -1
End Function
'Function GetSerByYear(Vyear As Integer) As Integer
'On Error GoTo errorhandler
'    Dim RsMax As New ADODB.Recordset
'    Sqltext = "Select isnull(Dbo.GetSerByyear(" & Vyear & "),0) as SerByYear"
'    Set RsMax = de.con.Execute(Sqltext)
'    GetSerByYear = RsMax!SerByyear + 1
'Exit Function
'errorhandler:
'GetSerByYear = -1
'End Function
Function FillStructure(Vindex As Integer) As Boolean
On Error GoTo errorhandler
    Select Case Vindex
        Case 1
            If FillVariables(1) Then
               With MvMaintPaymentRec
                .Billdate = ConvertControlDate(TxtDate.Text)
                '.SerByyear = GetSerByYear(Year(.Billdate))
                .OperationType = Val(ComboOperationType.BoundText)
                .MaintType = Val(ComboType.BoundText)
                DoEvents
                DoEvents
                .PaymentTYpeId = IIf(ComboPayment.BoundText = "", -1, Val(ComboPayment.BoundText))
                .DestinationId = ComboDestination.BoundText
                .clientId = Val(txtClientName.Tag)
                .Class = Val(LClientType.Tag)
                .ModNo = Val(TxtModelName.Tag)
                .ModelQty = Val(txtModelQty.Text)
                .FeesDescription = TxtFeesDescription.Text
                .OtherFeesPrice = Val(TxtOtherFeesPrice.Text)
                .OtherFeesQty = Val(TxtOthersFeesQty.Text)
                .IsFixed = ChkBill.Value
                If .IsFixed = 1 Then
                    .FixBillDate = ConvertControlDate(TxtFixBillDate.Text)
                Else
                    .FixBillDate = ConvertControlDate(TxtDate.Text)
                End If
           End With
        Else
            FillStructure = False
            Exit Function
         End If
        Case 2
            If FillVariables(2) Then
                With MvMaintPaymentRecTypeDetailsRec
                    .BillNo = MvMaintPaymentRec.BillNo
                    .stkno = TxtItemName.Tag
                    .Qty = Val(TxtQty.Text)
                    .PriceTYpe = 3 ' CliPrice
                    .Price = Val(LPrice.Caption)
                    .discount = Val(LDiscount.Caption)
                    .DestinationStoreId = txtClientName.Tag
                    .Class = LClientType.Tag
                    .OperationType = Val(ComboOperationType.BoundText)
                End With
            Else
                FillStructure = False
                Exit Function
            End If
    End Select
FillStructure = True
Exit Function
errorhandler:
FillStructure = False
End Function

Function InsertIntoMvStock(stkno As String, BillNo As Double, Qty As Double, OperationType As Integer, clientId As Double, Class As Integer) As Boolean
On Error GoTo errorhandler
Dim rs As New ADODB.Recordset
Dim QtyType As Integer

'sqlText = "Select Stkrelatedno  , Qty  From CoMaintitemrelated Where StkNo = '" & stkno & "'"
'Set rs = de.con.Execute(sqlText)
If OperationType = 1 Then
    QtyType = 1
Else
    QtyType = 0
End If

    sqlText = "Insert Into Stmov(ByanId , StkId  , StrId , Movdate , DocType , DocNum ,  Qty , QtyType)Values("
    sqlText = sqlText & NewRec & "," & GetStkId(stkno) & "," & GetStrId(systemConfigration.MainStoreNo) & ",Convert(varchar(10),getdate(),101)," & IIf(Class = 4, 30, 18) & "," & BillNo & "," & Qty & "," & QtyType & ")"
    de.con.Execute (sqlText)
    
   If Class = 4 Then
            sqlText = " Insert Into Stmov(ByanId , StkId  , StrId , Movdate , DocType , DocNum ,  Qty , QtyType)Values("
            sqlText = sqlText & NewRec + 1 & "," & GetStkId(stkno) & "," & clientId & " ,Convert(varchar(10),getdate(),101),30," & BillNo & "," & Qty & "," & IIf(QtyType = 1, 0, 1) & ")"
            de.con.Execute (sqlText)
    End If
    
    If GetBalance(stkno, GetStrId(systemConfigration.MainStoreNo)) < 0 Or GetBalance(stkno, clientId) < 0 Then
        InsertIntoMvStock = False
    Else
        InsertIntoMvStock = True
    End If
Exit Function
errorhandler:
InsertIntoMvStock = False
End Function

Sub FillDetails(BillId As Double)
On Error GoTo errorhandler
    With MvMaintPaymentRecTypeDetailsRec
        de.con.BeginTrans
            sqlText = "Insert Into MvMaintPaymentsDetails(BillNo, StkNo, discount , Qty, PriceTYpe, Price , Empno )Values( "
            sqlText = sqlText & .BillNo & ",'" & .stkno & "'," & .discount & "," & .Qty & "," & .PriceTYpe & "," & .Price & "," & empNo & ")"
            de.con.Execute (sqlText)
            If InsertIntoMvStock(.stkno, .BillNo, .Qty, .OperationType, .DestinationStoreId, .Class) Then
                de.con.CommitTrans
            Else
                de.con.RollbackTrans
                MsgBox "«·—’Ìœ ·«Ì”„Õ" & Chr(13) & " Ì—ÃÏ «· «ﬂœ „‰ —’Ìœ «·„” Êœ⁄", vbInformation + vbExclamation, "«·—’Ì‹‹‹‹œ ·«Ì”‹‹‹„Õ"
                Exit Sub
            End If
        .Id = GetMaxId
        AddToGrid 1, FlexGrid
    End With
Exit Sub
errorhandler:
de.con.RollbackTrans
MsgBox Err.Description
End Sub

Function foundStkNo(stkno As String, BillNo As Double) As Boolean
On Error GoTo errorhandler
Dim rs As New ADODB.Recordset
sqlText = "Select Count(*) As CountRec From MvMaintPaymentsDetails Where Stkno='" & stkno & "' and billNo =" & BillNo
Set rs = de.con.Execute(sqlText)
If rs!CountRec = 1 Then
    foundStkNo = True
Else
    foundStkNo = False
End If
Exit Function
errorhandler:
foundStkNo = True
MsgBox Err.Description
End Function

Sub FillFeesDetails(BillNo As Double)
With FlexFees
    sqlText = "Delete From mvmaintfees Where BillNo=" & BillNo
    de.con.Execute (sqlText)
    For i = 1 To .Rows - 1
        sqlText = "insert into mvmaintfees(BillNo , FeesTypeId , FeesQty , FeesPriceTYpe , FeesAmount,EmpNo)Values("
        sqlText = sqlText & BillNo & "," & Val(.Cell(flexTextFlat, i, ColFeesTypeName, i, ColFeesTypeName)) & "," & .TextMatrix(i, ColFeesQty) & ",3," & .TextMatrix(i, ColFeesAmount) & "," & empNo & ")"
        de.con.Execute (sqlText)
    Next
End With
End Sub

Function SaveRec() As Boolean
On Error GoTo errorhandler


If TypeRec Then   ' New Rec
    If LBillNo.Tag = "" Then
        If FillStructure(1) Then
            With MvMaintPaymentRec
                de.con.BeginTrans
                    sqlText = "Insert Into MvMaintPayments( Billdate, FixBillDate , OperationType , MaintType, PaymentTypeId , DestinationId , ClientId, Class ,  ModNo, ModelQty, FeesDescription ,   OtherFeesQty, OtherFeesPrice , IsFixed , EmpNo )Values("
                    sqlText = sqlText & "'" & .Billdate & "','" & .FixBillDate & "'," & .OperationType & "," & .MaintType & "," & .PaymentTYpeId & "," & .DestinationId & "," & .clientId & "," & .Class & "," & .ModNo & "," & .ModelQty & ",'" & .FeesDescription & "'," & .OtherFeesQty & "," & .OtherFeesPrice & "," & .IsFixed & "," & empNo & ")"
                    de.con.Execute (sqlText)
                    .BillNo = GetMaxBillNo
                    FillFeesDetails .BillNo
                de.con.CommitTrans
                If FillStructure(2) Then
                    FillDetails .BillNo
                End If
            End With
        Else
            SaveRec = False
            Exit Function
        End If
        InitNavigator
        MoveToRec MvMaintPaymentRec.BillNo
    Else
        If FillStructure(2) Then
            If Not foundStkNo(MvMaintPaymentRecTypeDetailsRec.stkno, MvMaintPaymentRecTypeDetailsRec.BillNo) Then
                FillDetails MvMaintPaymentRecTypeDetailsRec.BillNo
            End If
        End If
    End If
Else ' Update
     If FillStructure(1) Then
        With MvMaintPaymentRec
'            Sqltext = "Update MvMaintPayments Set Billdate='" & .Billdate & "',FixBillDate='" & .FixBillDate & "', OperationType=" & .OperationType & ",MaintType=" & .MaintType & ",PaymentTypeId=" & .PaymentTYpeId & ",ClientId= " & .ClientId & ", Class=" & .Class & ", ModNo=" & .ModNo & ",ModelQty=" & .ModelQty & ", FeesTYpeId=" & .FeesTYpeId & ",FeesQty=" & .FeesQty & ",FeesPriceType=" & .FeesPriceType & ", FeesAmount=" & .FeesAmount & ",IsFixed=" & .IsFixed & "  Where BillNo=" & RsNavigator!BillNo
'            de.con.Execute (Sqltext)
            de.con.BeginTrans
                sqlText = "Update MvMaintPayments Set Billdate='" & .Billdate & "',FixBillDate='" & .FixBillDate & "', OperationType=" & .OperationType & ",MaintType=" & .MaintType & ",PaymentTypeId=" & .PaymentTYpeId & ",DestinationId=" & .DestinationId & ",ClientId= " & .clientId & ", Class=" & .Class & ",ModNo=" & .ModNo & ",ModelQty=" & .ModelQty & ",OtherFeesQty =" & .OtherFeesQty & ",OtherFeesPrice=" & .OtherFeesPrice & ",FeesDescription='" & .FeesDescription & "',IsFixed=" & .IsFixed & "  Where BillNo=" & RsNavigator!BillNo
                de.con.Execute (sqlText)
                FillFeesDetails RsNavigator!BillNo
            de.con.CommitTrans
            IDBill = RsNavigator!BillNo
            InitNavigator
            MoveToRec IDBill
            MvMaintPaymentRec.BillNo = IDBill
        If FillStructure(2) Then
            If Not foundStkNo(MvMaintPaymentRecTypeDetailsRec.stkno, IDBill) Then
                FillDetails IDBill
            End If
        End If
        End With
     Else
        SaveRec = False
        Exit Function
     End If
End If
LCount.Caption = FlexGrid.Rows - 1
LTotalItems.Caption = GetTotalPrice(Val(LBillNo.Tag), 0)
LTotal.Caption = GetTotalPrice(Val(LBillNo.Tag), 1)
'LTotalBeforDiscount.Caption = GetTotalPriceBeforDiscount(Val(LBillNo.Caption), 1)
SaveRec = True
Exit Function
errorhandler:
de.con.RollbackTrans
SaveRec = False
MsgBox Err.Description
End Function

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdNewItem_Click()
TxtItemName.SetFocus
SendKeys "{Home}+{End}"
End Sub
Function showErrorPrices() As Boolean
On Error GoTo errorhandler
Dim rs As New ADODB.Recordset
sqlText = "SELECT  Count(*)CountRec  FROM    MvMaintPaymentsQry Where  Stat <> 6 and  billno =" & Val(LBillNo.Tag) & " and NativePrice <> Price "
Set rs = de.con.Execute(sqlText)
If rs!CountRec >= 1 Then
    showErrorPrices = True
Else
    showErrorPrices = False
End If
'
Exit Function
errorhandler:
MsgBox Err.Description
End Function
Sub PrintData(Vindex As Integer)
On Error GoTo errorhandler
    With Cr1
        showErrorPrices
        .Connect = ConnectName("")
        .SQLQuery = "SELECT    billno, billdate , FixBillDate , MaintTYpe, ClientName, ClientPhoneNBr, Name, Symbol, stkno, stkname, qty, price, TotPrice   FROM    MvMaintPaymentsQry Where stat<>6 and  billno =" & Val(LBillNo.Tag) & " Order By Row , StkNo"
        .ReportFileName = App.Path + "\Reports\RepBill2.rpt"
        .DiscardSavedData = True
        .WindowState = crptMaximized
        Select Case Vindex
            Case 0
                 If showErrorPrices Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì √”⁄«— «·„Ê«œ....", vbCritical, "Œÿ√ ›Ì √”⁄«— «·„Ê«œ"
                    .WindowShowPrintBtn = False
                Else
                    .WindowShowPrintBtn = True
                End If
                .Destination = crptToWindow
            Case 1
                 If showErrorPrices() Then
                    If MsgBox("ÌÊÃœ Œÿ√ ›Ì √”⁄«— «·„Ê«œ...." & Chr(13) & "Â·  —Ìœ «·ÿ»«⁄…ø", vbCritical + vbYesNo + vbDefaultButton2, "Œÿ√ ›Ì √”⁄«— «·„Ê«œ") = vbYes Then
                        .Destination = crptToPrinter
                    End If
                Else
                    .Destination = crptToPrinter
                End If
        End Select
        .Action = 1
    End With
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub
Function GetSerial(SerByyear As Integer) As Double
On Error GoTo errorhandler
Dim rs As New ADODB.Recordset
    sqlText = "Select billno from MvMaintPayments where SerByyear =" & SerByyear & " and year(billdate)=" & Year(Now)
    Set rs = de.con.Execute(sqlText)
    GetSerial = rs!BillNo
Exit Function
errorhandler:
MsgBox Err.Description
GetSerial = -1

End Function
Sub AddItemsFromModelList(BillNo As Double)
On Error GoTo errorhandler
    sqlText = "Exec Sp_AddMaintItems " & BillNo & "," & empNo
    de.con.Execute (sqlText)
    MoveToRec BillNo
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub


'Private Sub CmdItems_Click()
'On Error GoTo ErrorHandler
'If ComboOperationType.BoundText <> 2 Then
'    If SaveRec Then
'        AddItemsFromModelList RsNavigator!BillNo
'    End If
'Else
'   MsgBox "·«Ì„ﬂ‰ ≈÷«€… „Ê«œ ·⁄„·Ì… «·„— Ã⁄", vbInformation + vbExclamation
'End If
'Exit Sub
'ErrorHandler:
'MsgBox Err.Description
'End Sub

Private Sub CmdNewCustomer_Click()
On Error GoTo errorhandler
    If Val(LClientType.Tag) = 0 Then
        clientId = Val(txtClientName.Tag)
        ClientName = txtClientName.Text
        ClientPhoneNBr = LClientPhoneNBR.Caption
        frmNewCustomer.Show 1
        Ok = False
        txtClientName.Text = ClientName
        txtClientName.Tag = clientId
        LClientType.Tag = 2
        LClientPhoneNBR.Caption = ClientPhoneNBr
        
        TxtModelName.SetFocus
        TxtModelName.SelStart = 0
        TxtModelName.SelLength = Len(TxtModelName.Text)
        Ok = True
    End If
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

Private Sub CmdPreview_Click()
PrintData 0
End Sub

Private Sub CmdPrint_Click()
PrintData 1
End Sub

Private Sub CmdSave_Click()
If SaveRec Then
    EnableCmds True, True, True, False, False, True, True, True, True, True, True
    EnableControls False
    CmdAdd.SetFocus
    MsgBox " „ Õ›Ÿ «·›« Ê—… »‰Ã«Õ", vbInformation, "Õ›Ÿ «·›« Ê—…"
Else
    MsgBox "·„ Ì „ Õ›ÿ «·›« Ê—…...... Ì—ÃÏ «· √ﬂœ „‰ «·ÕﬁÊ· «·›«—€…", vbInformation, "Œÿ√ ›Ì «·≈œŒ«·"
End If
End Sub

Private Sub CmdFirst_Click()
With RsNavigator
    MoveNavigator 1
End With
End Sub

Private Sub CmdLast_Click()
With RsNavigator
    MoveNavigator 4
End With
End Sub

Private Sub CmdNext_Click()
With RsNavigator
    MoveNavigator 3
End With

End Sub

Private Sub CmdPrevious_Click()
With RsNavigator
    MoveNavigator 2
End With
End Sub



Function GetPrice(stkno As String, Index As Integer) As Double
Dim rs As New ADODB.Recordset, RsPrice   As New ADODB.Recordset, Col As String

'If VCurrency = 0 Then
'GetPrice = 0
'Exit Function
'End If

Select Case Index
    Case 1
        Col = "CliPrice"

'       Select Case VCurrency
'        Case 1
'            Col = "DealPriceafterdiscount"
'        Case 2
'            Col = "DistPriceafterdiscount"
'        Case 3
'            Col = "CliPriceafterdiscount"
'       End Select

    Case 2
        Col = "CliPrice"
'       Select Case VCurrency
'        Case 1
'            Col = "DealPrice"
'        Case 2
'            Col = "DistPrice"
'        Case 3
'            Col = "CliPrice"
'       End Select
End Select

'Sqltext = "Select Col from dbo.PriceTypes Where PriceNo=" & VCurrency
'Set rs = de.con.Execute(Sqltext)

    sqlText = "Select " & Col & " as Price from Costock Where StkNo='" & stkno & "'"
    Set RsPrice = de.con.Execute(sqlText)
    If RsPrice.RecordCount > 0 Then
        GetPrice = RsPrice!Price
    Else
        GetPrice = 0
    End If
End Function

'Function GetDiscount(stkno As String) As Double
'Dim rs As New ADODB.Recordset, RsPrice   As New ADODB.Recordset
'sqlText = "Select discount from costock Where stkno='" & stkno & "'"
'Set rs = de.con.Execute(sqlText)
'
'If rs.RecordCount > 0 Then
'    GetDiscount = rs!discount
'Else
'    GetDiscount = 0
'End If
'End Function

Function GetDiscount() As Double
Dim rs As New ADODB.Recordset, RsPrice   As New ADODB.Recordset
sqlText = "Select DiscountPercentage discount from maintusers Where empno =" & empNo
Set rs = de.con.Execute(sqlText)

If rs.RecordCount > 0 Then
    GetDiscount = rs!discount
Else
    GetDiscount = 0
End If
End Function

Function SearchRec() As Double
On Error GoTo errorhandler
Dim i As Double
i = InputBox("√œŒ· —ﬁ„ «·›« Ê—…", "«·»ÕÀ ⁄‰ ›« Ê—…")
If Val(i) <> 0 Then
    SearchRec = i
Else
    SearchRec = -1
End If
Exit Function
errorhandler:
SearchRec = -1
End Function
Private Sub CmdSearch_Click()
MoveToRec SearchRec
'FrmSeatchBills.Show
End Sub

'Private Sub ComboCurrencyType_Change()
'    LPrice.Caption = GetPrice(Val(ComboCurrencyType.BoundText), TxtitemName.Tag, 1)
'    LDiscount.Caption = GetDiscount()
'End Sub

'Private Sub ComboCurrencyType_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    ComboDiscount.SetFocus
'    SendKeys "{home}+{end}"
'End If
'End Sub

'Private Sub ComboDestination_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    TxtClientName.SetFocus
'    Sendkeys "{home}+{end}"
'End If
'End Sub

Function GetStkNoPrice(Price As Double, discount As Integer) As Double
On Error GoTo errorhandler
Dim rs As New ADODB.Recordset
sqlText = "Select dbo.fn_GetStkNoPrice (" & discount & "," & Price & ") as Price "
Set rs = de.con.Execute(sqlText)
GetStkNoPrice = IIf(IsNull(rs!Price), Price, rs!Price)
Exit Function
errorhandler:
GetStkNoPrice = Price
MsgBox Err.Description
End Function




Private Sub ComboDestination_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtClientName.SetFocus
    SendKeys "{home}+{end}"
End If

End Sub

Private Sub ComboDiscount_GotFocus()
'ComboDiscount.ListIndex = 0
End Sub

Private Sub ComboDiscount_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errorhandler
Dim Price As Double, discount As Integer
Price = GetPrice(TxtItemName.Tag, 2)
discount = Val(ComboDiscount.Text)
LPrice.Caption = GetStkNoPrice(Price, discount)
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

Private Sub ComboDiscount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TxtItemName.Tag = "0" Then
        cmdSave.SetFocus
    Else
        If SaveRec Then
        End If
        ClearItems
        TxtItemName.SetFocus
        SendKeys "{home}+{end}"
    End If
End If
End Sub

Private Sub ComboFees_Change()
Dim feesClientPrint As Integer
If Val(TxtFeesQty.Text) > 0 Then
    feesClientPrice = GetFeesPrice(Val(ComboFees.BoundText))
    'LfeesClientPrice.Caption = feesClientPrice
    LFeesAmount.Caption = feesClientPrice
    LTotal.Caption = GetTotalPrice(Val(LBillNo.Tag), 1)
End If
End Sub

Private Sub ComboFees_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtFeesQty.SetFocus
    TxtFeesQty.SelStart = 0
    TxtFeesQty.SelLength = Len(TxtFeesQty.Text)
End If
End Sub

Function GetOperationTypeEnableValue() As Boolean
If FlexGrid.Rows > 1 Then
GetOperationTypeEnableValue = False
Else
GetOperationTypeEnableValue = True
End If
End Function


Sub AddToGrid(Vindex As Integer, FlexGrid As VSFlexGrid)
Dim Vrow As Integer
If Vindex = 1 Then
    With FlexGrid
        .AddItem ""
        Vrow = .Rows - 1
        .TextMatrix(Vrow, ColId) = MvMaintPaymentRecTypeDetailsRec.Id
        .TextMatrix(Vrow, ColStkNo) = MvMaintPaymentRecTypeDetailsRec.stkno
        .TextMatrix(Vrow, ColStkName) = LItemName.Caption
        .TextMatrix(Vrow, ColQty) = MvMaintPaymentRecTypeDetailsRec.Qty
        .TextMatrix(Vrow, colPriceTypeId) = MvMaintPaymentRecTypeDetailsRec.PriceTYpe
'        .TextMatrix(Vrow, ColPriceTypeName) = ComboCurrencyType.Text
        .TextMatrix(Vrow, ColPrice) = MvMaintPaymentRecTypeDetailsRec.Price
        .TextMatrix(Vrow, ColDiscount) = MvMaintPaymentRecTypeDetailsRec.discount
        ComboOperationType.Enabled = GetOperationTypeEnableValue
        FillCombosInGrid FlexGrid, ColPriceTypeName, 1
        FillFormating 2
    End With
ElseIf Vindex = 2 Then
    With FlexGrid
        .AddItem ""
        Vrow = .Rows - 1
        .TextMatrix(Vrow, ColFeesTypeName) = ComboFees.BoundText
'        .TextMatrix(Vrow, ColFeesPriceName) = ComboFeesPriceType.BoundText
        .TextMatrix(Vrow, ColFeesAmount) = LFeesAmount.Caption
        .TextMatrix(Vrow, ColFeesQty) = TxtFeesQty.Text
        FillCombosInGrid FlexFees, ColFeesPriceName, 1
        FillCombosInGrid FlexFees, ColFeesTypeName, 3
        FillFormating 3
        'LTotal.Caption = GetTotalPrice(Val(LBillNo.Tag), 1)
        LTotalFees.Caption = GetFeesAmount
    End With
    
End If
End Sub
Function GetFeesPrice(FeesId As Integer) As Double
'Dim rs As New ADODB.Recordset, RsPrice   As New ADODB.Recordset
'sqlText = "Select Col from dbo.PriceTypes Where PriceNo=" & VCurrency
'Set rs = de.con.Execute(sqlText)

'If rs.RecordCount > 0 Then
    sqlText = "Select CliPriceafterdiscount as Price from CoMaintFees Where FeesId=" & FeesId
    Set RsPrice = de.con.Execute(sqlText)
    If RsPrice.RecordCount > 0 Then
        GetFeesPrice = RsPrice!Price
    Else
       GetFeesPrice = 0
    End If
'End If
End Function
'Private Sub ComboFeesPriceType_Change()
'    TxtFeesAmount.Text = GetFeesPrice(Val(ComboFeesPriceType.BoundText), Val(ComboFees.BoundText))
'    LTotal.Caption = GetTotalPrice(Val(LBillNo.Tag), 1)
'End Sub
'
'Private Sub ComboFeesPriceType_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    TxtFeesAmount.SelStart = 0
'    TxtFeesAmount.SelLength = Len(TxtFeesAmount.Text)
'    TxtFeesAmount.SetFocus
'End If
'End Sub
Sub ChangeItmes(BillNo As Double)
On Error GoTo errorhandler
    sqlText = "Sp_ChangeItemData " & BillNo & "," & empNo
    de.con.Execute sqlText
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub
Private Sub comboOperationTYpe_Change()
'If ComboOperationType.BoundText = "" Then Exit Sub
'CurrentOperationTYpeID = ComboOperationType.BoundText
'With FlexGrid
'If BeforeOperationTYpeID = 0 Then Exit Sub
'If BeforeOperationTYpeID <> CurrentOperationTYpeID Then
'    If .Rows > 1 And Val(LBillNo.tag) <> 0 Then
'        SQlMessage = "·«Ì„ﬂ‰ «· €ÌÌ— »”»» ÊÃÊœ „Ê«œ „Œ“‰Ì… ⁄·Ï «·›« Ê—…"
'        SQlMessage = SQlMessage & Chr(13)
'        SQlMessage = SQlMessage & "ÌÃ» Õ–› «·„Ê«œ ﬁ»·  €ÌÌ— ‰Ê⁄ «·⁄„·Ì…!!!"
'        MsgBox SQlMessage, vbInformation + vbExclamation, " €ÌÌ— ‰Ê⁄ «·⁄„·Ì…"
'        ComboOperationType.BoundText = BeforeOperationTYpeID
'    End If
'End If
'End With

'Dim SQlMessage As String
'With FlexGrid
'    If .Rows > 1 And Val(LBillNo.Caption) <> 0 Then
'        SQlMessage = "·«Ì„ﬂ‰ «· €ÌÌ— »”»» ÊÃÊœ „Ê«œ „Œ“‰Ì… ⁄·Ï «·›« Ê—…"
'        SQlMessage = SQlMessage & Chr(13)
'        SQlMessage = SQlMessage & "ÌÃ» Õ–› «·„Ê«œ ﬁ»·  €ÌÌ— ‰Ê⁄ «·⁄„·Ì…!!!"
'
'        If MsgBox(SQlMessage, vbYesNo + vbDefaultButton2 + vbQuestion, " €ÌÌ— ‰Ê⁄ «·⁄„·Ì…") = vbYes Then
'            ChangeItmes RsNavigator!BillNo
'        End If
'    End If
'End With
End Sub

Private Sub ComboOperationType_GotFocus()
'BeforeOperationTYpeID = ComboOperationType.BoundText

End Sub

Private Sub ComboOperationType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ComboType.SetFocus
End If
End Sub

Private Sub ComboPayment_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ComboDestination.SetFocus
End If
End Sub

Private Sub ComboType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ComboPayment.SetFocus

End If
End Sub

Private Sub FlexFees_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo errorhandler
Dim RsUpdate As New ADODB.Recordset
With FlexFees
    Select Case Col
        Case ColFeesPriceName, ColFeesTypeName, ColFeesQty
            .TextMatrix(Row, ColFeesAmount) = GetFeesPrice(Val(.Cell(flexTextFlat, Row, ColFeesTypeName, Row, ColFeesTypeName)))
             If .TextMatrix(Row, ColFeesSer) <> "" Then
                sqlText = "Update MvMaintFees Set FeesTypeId=" & Val(.Cell(flexTextFlat, Row, ColFeesTypeName, Row, ColFeesTypeName)) & ",FeesPriceTYpe=" & Val(.Cell(flexTextFlat, Row, ColFeesPriceName, Row, ColFeesPriceName)) & ",FeesQty =" & .TextMatrix(Row, ColFeesQty) & ",FeesAmount=" & .TextMatrix(Row, ColFeesAmount) & " Where Id=" & .TextMatrix(Row, ColFeesSer)
                de.con.Execute (sqlText)
             End If
    End Select
    LTotalFees.Caption = GetFeesAmount
    'LTotal.Caption = GetTotalPrice(Val(LBillNo.Tag), 1)
End With
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

Private Sub FlexFees_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
If Col = ColFeesAmount Then cancel = True
End Sub

Private Sub FlexFees_KeyDown(KeyCode As Integer, Shift As Integer)
Dim FirstRow As Integer, LastRow As Integer, Vrow As Integer
With FlexFees
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
            If .Rows = 1 Then
            Else
              If .TextMatrix(i, ColId) <> "" Then
                If DeleteRow(FlexFees, Vrow, ColFeesSer, "MvMaintFees", "Id") Then
                    ClearHeader
                    .RemoveItem Vrow
                    LTotalFees.Caption = GetFeesAmount
                    LTotal.Caption = GetTotalPrice(Val(LBillNo.Tag), 1)
                End If
              Else
                    ClearHeader
                    .RemoveItem Vrow
                    LTotalFees.Caption = GetFeesAmount
                    LTotal.Caption = GetTotalPrice(Val(LBillNo.Tag), 1)
            End If
          End If
        Next
        .Col = ColMopName
        .SetFocus
    End If
End If
End With

End Sub

Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo errorhandler
Dim Price As Double, discount As Integer
Dim RsUpdate As New ADODB.Recordset
With FlexGrid
    Select Case Col
        Case ColPriceTypeName
            Price = GetPrice(.TextMatrix(Row, ColStkNo), 2)
'            discount = GetDiscount(.TextMatrix(Row, ColStkNo))
            discount = .TextMatrix(.Row, ColDiscount)
            If Price <> 0 Then
                .TextMatrix(Row, ColPrice) = GetStkNoPrice(Price, discount)
                '.TextMatrix(Row, ColDiscount) = discount
            Else
                MsgBox "«·”⁄— Ì”«ÊÌ «·’›— ... ·«Ì„ﬂ‰  ⁄œÌ· «·›« Ê—…", vbInformation + vbExclamation, "«·”⁄— Ì”«ÊÌ ’›—"
                Exit Sub
            End If
            sqlText = "Update  MvMaintPaymentsDetails Set updatefrom=1,PriceTYpe=" & Val(.Cell(flexTextFlat, Row, ColPriceTypeName, Row, ColPriceTypeName)) & " ,Price=" & .TextMatrix(Row, ColPrice) & ",discount=" & .TextMatrix(Row, ColDiscount) & " Where id=" & .TextMatrix(Row, ColId)
            Set RsUpdate = de.con.Execute(sqlText)
        Case ColPaymentTypeName
            sqlText = "Update  MvMaintPaymentsDetails Set updatefrom=2,PaymentTypeId=" & Val(.Cell(flexTextFlat, Row, ColPaymentTypeName, Row, ColPaymentTypeName)) & " Where id=" & .TextMatrix(Row, ColId)
            Set RsUpdate = de.con.Execute(sqlText)
'        Case ColQty
'            Sqltext = "Update  MvMaintPaymentsDetails Set Qty=" & .TextMatrix(Row, ColQty) & " Where id=" & .TextMatrix(Row, Colid)
'            Set RsUpdate = de.con.Execute(Sqltext)
    End Select
LCount.Caption = .Rows - 1
LTotalItems.Caption = GetTotalPrice(Val(LBillNo.Tag), 0)

LTotal.Caption = GetTotalPrice(Val(LBillNo.Tag), 1)
'LTotalBeforDiscount.Caption = GetTotalPriceBeforDiscount(Val(LBillNo.Caption), 1)

End With
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

Private Sub FlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
If Col = ColPrice Or Col = ColStkName Or Col = ColStkNo Or Col = ColQty Or Col = ColDiscount Then cancel = True
End Sub

Private Sub flexGrid_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
On Error GoTo errorhandler
Dim Price  As Double
    Dim RsUpdate As New ADODB.Recordset
With FlexGrid
    Select Case Col
        Case ColPriceTypeName
            Price = GetPrice(.TextMatrix(Row, ColStkNo), 1)
            If Price <> 0 Then
                .TextMatrix(Row, ColPrice) = Price
            Else
                MsgBox "«·”⁄— Ì”«ÊÌ «·’›— ... ·«Ì„ﬂ‰  ⁄œÌ· «·›« Ê—…", vbInformation + vbExclamation, "«·”⁄— Ì”«ÊÌ ’›—"
                Exit Sub
            End If
            .TextMatrix(Row, ColPrice) = Price
            sqlText = "Update  MvMaintPaymentsDetails Set updatefrom =3 , PriceTYpe=" & Val(.Cell(flexTextFlat, Row, ColPriceTypeName, Row, ColPriceTypeName)) & " ,Price=" & .TextMatrix(Row, ColPrice) & " Where id=" & .TextMatrix(Row, ColId)
            Set RsUpdate = de.con.Execute(sqlText)
        Case ColPaymentTypeName
            sqlText = "Update  MvMaintPaymentsDetails Set updatefrom =4 , PaymentTypeId=" & Val(.Cell(flexTextFlat, Row, ColPaymentTypeName, Row, ColPaymentTypeName)) & " Where id=" & .TextMatrix(Row, ColId)
            Set RsUpdate = de.con.Execute(sqlText)
'        Case ColQty
'            Sqltext = "Update  MvMaintPaymentsDetails Set Qty=" & .TextMatrix(Row, ColQty) & " Where id=" & .TextMatrix(Row, Colid)
'            Set RsUpdate = de.con.Execute(Sqltext)
    End Select
    LTotalItems.Caption = GetTotalPrice(Val(LBillNo.Tag), 0)
    LTotal.Caption = GetTotalPrice(Val(LBillNo.Tag), 1)
    'LTotalBeforDiscount.Caption = GetTotalPriceBeforDiscount(Val(LBillNo.Caption), 1)
End With
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub
Sub UpdateRecords(i As Integer)
With FlexGrid
    sqlText = "Update MvMaintPayments Set StkNo='',"
End With
End Sub
Sub ClearHeader()
Ok = False
TxtItemName.Tag = ""
TxtItemName.Text = ""
LItemName.Caption = ""
'ComboCurrencyType.BoundText = ""
LBalance.Caption = ""
LPrice.Caption = ""
TxtQty.Text = ""
Ok = True
End Sub
Sub RemoveFromMvStock(Id As Double)
On Error GoTo errorahndler
With FlexGrid
    sqlText = "Sp_Delete_MvStock " & Id & "," & GetStrId(systemConfigration.MainStoreNo)
    de.con.Execute (sqlText)
End With
Exit Sub
errorahndler:
MsgBox Err.Description
End Sub
Private Sub FlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim FirstRow As Integer, LastRow As Integer, Vrow As Integer
With FlexGrid
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
            If .Rows = 1 Then
                'UpdateRecords i
            Else
                  RemoveFromMvStock .TextMatrix(i, ColId)
                  If DeleteRow(FlexGrid, Vrow, ColId, "MvMaintPaymentsDetails", "Id") Then
                    ClearHeader
                    .RemoveItem Vrow
                    LCount.Caption = .Rows - 1
                    LTotalItems.Caption = GetTotalPrice(Val(LBillNo.Tag), 0)
                    LTotal.Caption = GetTotalPrice(Val(LBillNo.Tag), 1)
                    'LTotalBeforDiscount.Caption = GetTotalPriceBeforDiscount(Val(LBillNo.Caption), 1)
                End If
            End If
        Next
        ComboOperationType.Enabled = GetOperationTypeEnableValue
        .Col = ColMopName
        .SetFocus
    End If
End If
End With
End Sub

'Private Sub FlexGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If Button And vbRightButton Then
'            PopupMenu mnufile
'    End If
'End Sub

Private Sub Grid_RowColChange()
On Error GoTo errorhandler
If Flag Then
    Ok = False
    With Grid
       Select Case Pos
        Case 1
            txtClientName.Tag = .TextMatrix(.Row, ColNo)
            txtClientName.Text = .TextMatrix(.Row, ColName)
            LClientType.Caption = .TextMatrix(.Row, col3)
            LClientType.Tag = .TextMatrix(.Row, Col4)
            LClientPhoneNBR.Caption = GetPhoneNbr(txtClientName.Tag, LClientType.Tag)
        Case 2
            TxtModelName.Tag = .TextMatrix(.Row, ColNo)
            TxtModelName.Text = .TextMatrix(.Row, col3)
            LSymbol.Caption = .TextMatrix(.Row, ColName)
        Case 3
            TxtItemName.Tag = .TextMatrix(.Row, ColNo)
            TxtItemName.Text = .TextMatrix(.Row, ColNo)
            LItemName.Caption = .TextMatrix(.Row, ColName)
            LBalance.Caption = ""
            
        Case 4
            TxtRecipient.Tag = .TextMatrix(.Row, ColNo)
            TxtRecipient.Text = .TextMatrix(.Row, ColName)
        Case 5
            TxtFamNo.Tag = .TextMatrix(.Row, ColNo)
            TxtFamNo.Text = .TextMatrix(.Row, ColName)
        Case 6
            TxtModelName.Tag = .TextMatrix(.Row, ColNo)
            TxtModelName.Text = .TextMatrix(.Row, ColName)
        Case 7
            TxtUnExecutedReason.Tag = .TextMatrix(.Row, ColNo)
            TxtUnExecutedReason.Text = .TextMatrix(.Row, ColName)
        Case 8
            TxtError.Tag = .TextMatrix(.Row, ColNo)
            TxtError.Text = .TextMatrix(.Row, ColName)
        Case 9
            TxtItemName.Tag = .TextMatrix(.Row, ColNo)
            TxtItemName.Text = .TextMatrix(.Row, ColName)
       End Select
    End With
    Ok = True
End If
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub
Private Sub Form_Load()
init
End Sub


Private Sub TxtBeginDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtEndDate.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txtClientPhoneNBR_KeyPress(KeyAscii As Integer)
On Error GoTo errorhandler
If KeyAscii = 13 Then
    If txtClientPhoneNBR.Text <> "" And Val(txtClientName.Tag) <> 0 Then
        If Val(LClientType.Tag) = 2 Then ' “»Ê‰
            sqlText = "Update CoClient Set ClientPhoneNBr='" & txtClientPhoneNBR.Text & "' Where ClientId=" & txtClientName.Tag
            de.con.Execute (sqlText)
        End If
    End If
    TxtModelName.SetFocus
    SendKeys "{home}+{end}"
End If
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub
'Sub ChangePriceType(Vindex As Integer)
'Dim FirstRow As Integer, LastRow As Integer
'On Error GoTo ErrorHandler
'    With flexGrid
'        If .Row >= .RowSel Then
'            FirstRow = .Row
'            LastRow = .RowSel
'        Else
'            FirstRow = .RowSel
'            LastRow = .Row
'        End If
'        For i = FirstRow To LastRow Step -1
'
'            Price = GetPrice(.TextMatrix(i, ColStkNo), 1)
'            If Price <> 0 Then
'                .TextMatrix(Row, ColPrice) = Price
'            Else
'                MsgBox "«·”⁄— Ì”«ÊÌ «·’›— ... ·«Ì„ﬂ‰  ⁄œÌ· «·›« Ê—…", vbInformation + vbExclamation, "«·”⁄— Ì”«ÊÌ ’›—"
'                Exit Sub
'            End If
'            .TextMatrix(i, ColPrice) = Price
'            sqlText = "Update  MvMaintPaymentsDetails Set UpdateFrom=5 ,  PriceTYpe=" & mnu(Vindex).HelpContextID & " ,Price=" & .TextMatrix(i, ColPrice) & " Where id=" & .TextMatrix(i, ColId)
'            de.con.Execute (sqlText)
'        Next
'        MoveToRec RsNavigator!BillNo
'    End With
'Exit Sub
'ErrorHandler:
'MsgBox Err.Description
'End Sub



'Private Sub mnu_Click(Index As Integer)
'ChangePriceType Index
'End Sub

Private Sub TxtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ComboOperationType.Enabled Then
        ComboOperationType.SetFocus
    Else
        ComboType.SetFocus
    End If
End If
End Sub

Private Sub TxtEndDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtItemName.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub TxtDollar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtModelName.SetFocus
    TxtModelName.SelStart = 0
    TxtModelName.SelLength = Len(TxtModelName.Text)
End If
End Sub

Private Sub TxtFeesAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(ComboFees.MatchedWithList) And Val(LFeesAmount.Caption) <> 0 Then
        AddToGrid 2, FlexFees
    End If
    ChkBill.SetFocus
End If
End Sub

Private Sub TxtFeesDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtOthersFeesQty.SetFocus
    TxtOthersFeesQty.SelStart = 0
    TxtOthersFeesQty.SelLength = Len(TxtOthersFeesQty.Text)
End If
End Sub

Private Sub TxtFeesQty_Change()
Dim feesClientPrint As Integer

feesClientPrice = GetFeesPrice(Val(ComboFees.BoundText))
'LfeesClientPrice.Caption = feesClientPrice
LFeesAmount.Caption = feesClientPrice

If TypeRec Then
    LTotal.Caption = 0
Else
    LTotal.Caption = GetTotalPrice(Val(LBillNo.Tag), 1)
End If
End Sub

Private Sub TxtFeesQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(ComboFees.BoundText) <> 0 And Val(LFeesAmount.Caption) <> 0 Then
        AddToGrid 2, FlexFees
    End If
    ChkBill.SetFocus
End If


'If KeyAscii = 13 Then

'    TxtFeesAmount.SelStart = 0
'    TxtFeesAmount.SelLength = Len(TxtFeesAmount.text)
'    TxtFeesAmount.SetFocus
'End If
End Sub

Private Sub TxtFixBillDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtFeesDescription.SetFocus
    TxtFeesDescription.SelStart = 0
    TxtFeesDescription.SelLength = Len(TxtFeesDescription.Text)
End If
End Sub

Private Sub TxtItemName_Change()
On Error GoTo errorhandler
Dim RsSearch As New ADODB.Recordset
If TxtItemName.Text = "" Then
    TxtItemName.Tag = "0"
    Grid.Visible = False
    Exit Sub
End If

If Ok Then
    Flag = False
'    Sqltext = "Select Top 10 c1.StkNo , StkName  , s1.FnlQnt from CoStock c1 inner join stkinfQry s1 on c1.id = s1.stkid  where StkName Like" & LikeExpression(TxtitemName.Text) & " or c1.StkNo like '" & TxtitemName.Text & "%'"
    sqlText = "Select Top 10 c1.StkNo , StkName  from CoStock c1 where StkName Like" & LikeExpression(TxtItemName.Text) & " or c1.StkNo like '" & TxtItemName.Text & "%'"
    Set RsSearch = de.con.Execute(sqlText)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        Grid.Row = 0
        FillFormating 1
        ChangeCursor 3
        Grid.Visible = True
    Else
        TxtItemName.Tag = "0"
        Grid.Visible = False
    End If
    Flag = True
End If
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

Private Sub TxtItemName_GotFocus()
Pos = 3
End Sub


Private Sub TxtitemName_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True

End Sub
Sub fillcombodiscount(discount As Integer)
On Error GoTo errorhandler
ComboDiscount.Clear
If discount = 0 Then
    
    ComboDiscount.AddItem (0)
Else
    For i = 0 To discount
        ComboDiscount.AddItem (i)
    Next
End If
ComboDiscount.ListIndex = 0
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub
Function GetBalance(stkno As String, Strid As Double)
On Error GoTo errorhandler
Dim rs As New ADODB.Recordset
sqlText = "select fnlqnt from stkinf s1  where StkNo = '" & LTrim(RTrim(stkno)) & "' and StrId=" & Strid
Set rs = de.con.Execute(sqlText)
If rs.RecordCount > 0 Then
    GetBalance = rs!fnlqnt
Else
    GetBalance = 0
End If
Exit Function
errorhandler:
GetBalance = 0
End Function
Private Sub TxtitemName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim Balance As Double
     If Grid.Visible Then
        Balance = GetBalance(Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo), GetStrId(systemConfigration.MainStoreNo))
    End If
    If Grid.Visible And Balance > 0 Then
        Ok = False
        TxtItemName.Tag = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
        TxtItemName.Text = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
        LItemName.Caption = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColName)
        LBalance.Caption = Balance
         LPrice.Caption = GetPrice(TxtItemName.Tag, 2)
        LDiscount.Caption = GetDiscount()
        fillcombodiscount (Val(LDiscount.Caption))
        Ok = True
    
    ElseIf Grid.Visible = False And TxtItemName.Text <> "" And TxtItemName.Tag <> "0" Then
        TxtQty.SetFocus
        TxtQty.SelStart = 0
        TxtQty.SelLength = Len(TxtQty.Text)
        Exit Sub
    Else
        Ok = False
        TxtItemName.Tag = "0"
        TxtItemName.Text = ""
        LItemName.Caption = ""
        LBalance.Caption = ""
        LDiscount.Caption = ""
        fillcombodiscount (Val(LDiscount.Caption))
        TxtItemName.SetFocus
        
        Ok = True
        Exit Sub
    End If
    Grid.Visible = False
    
    TxtQty.SetFocus
    TxtQty.SelStart = 0
    TxtQty.SelLength = Len(TxtQty.Text)
End If
End Sub

Private Sub TxtModelName_Change()
On Error GoTo errorhandler
Dim RsSearch As New ADODB.Recordset
If TxtModelName.Text = "" Then
    TxtModelName.Tag = 0
    Grid.Visible = False
    Exit Sub
End If
If Ok Then
    Flag = False
    sqlText = "select ModNo , Symbol , Name    from adhammodels Where Symbol    Like" & LikeExpression(TxtModelName.Text) & " or Name    like " & LikeExpression(TxtModelName.Text)
    Set RsSearch = de.con.Execute(sqlText)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        Grid.Row = 0
        FillFormating 1
        ChangeCursor 2
        Grid.Visible = True
    Else
        TxtModelName.Tag = 0
        Grid.Visible = False
    End If
    Flag = True
End If
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub


Private Sub TxtModelName_GotFocus()
ChangeToArabic
Pos = 2
End Sub

Private Sub TxtModelName_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True
End Sub

Private Sub TxtModelName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible Then
        Ok = False
        TxtModelName.Tag = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
        TxtModelName.Text = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), col3)
        LSymbol.Caption = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColName)
        Ok = True
    ElseIf Grid.Visible = False And TxtModelName.Text <> "" And Val(TxtModelName.Tag) <> 0 Then
        txtModelQty.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    Else
        Ok = False
        TxtModelName.Tag = 0
        TxtModelName.Text = ""
        LSymbol.Caption = ""
        Ok = True
    End If
    Grid.Visible = False
    txtModelQty.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub
'----------------------------------------------





Private Sub txtClientName_Change()
On Error GoTo errorhandler
Dim RsSearch As New ADODB.Recordset
If txtClientName.Text = "" Then
    txtClientName.Tag = 0
    Grid.Visible = False
    Exit Sub
End If
If Ok Then
    Flag = False
    sqlText = "Select Top 10 [ClientId] , [ClientName]  , [ClientTypeName] , Class  From ClientQry Where ClientName like" & LikeExpression(txtClientName.Text) & " or ClientPhoneNBr like " & LikeExpression(txtClientName.Text)
    Set RsSearch = de.con.Execute(sqlText)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        Grid.Row = 0
        FillFormating 1
        ChangeCursor 1
        Grid.Visible = True
    Else
        txtClientName.Tag = 0
        Grid.Visible = False
    End If
    Flag = True
End If
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub


Private Sub TxtClientName_GotFocus()
ChangeToArabic
Pos = 1
End Sub

Private Sub txtClientName_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True
End Sub

Function GetClassName(Class As Integer)
Select Case Class
    Case 1
        GetClassName = " «Ã—"
    Case 2
        GetClassName = "“»Ê‰"
    Case 3
        GetClassName = "„ÊŸ›"
    Case 4
        GetClassName = "’«·Â"
    Case Else
        GetClassName = "€Ì— „⁄—Ê›"
End Select
End Function
Function GetClientName(clientId As Double, Class As Integer) As String
On Error GoTo errorhandler
Dim RsClientName As New ADODB.Recordset
sqlText = "Select ClientName From ClientQry Where ClientId=" & clientId & " and Class =" & Class
Set RsClientName = de.con.Execute(sqlText)
If RsClientName.RecordCount > 0 Then
    GetClientName = RsClientName!ClientName
Else
    GetClientName = ""
End If
Exit Function
errorhandler:
GetClientName = ""
End Function

Function GetPhoneNbr(clientId As Double, Class As Integer) As String
On Error GoTo errorhandler
Dim RsPhone As New ADODB.Recordset
sqlText = "Select isnull(ClientPhonenbr,'') ClientPhonenbr From ClientQry Where ClientId=" & clientId & " and Class =" & Class
Set RsPhone = de.con.Execute(sqlText)
If RsPhone.RecordCount > 0 Then
    GetPhoneNbr = RsPhone!ClientPhoneNBr
Else
    GetPhoneNbr = ""
End If
Exit Function
errorhandler:
GetPhoneNbr = ""
End Function

Private Sub txtClientName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible Then
        Ok = False
        txtClientName.Tag = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
        txtClientName.Text = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColName)
        LClientType.Caption = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), col3)
        LClientType.Tag = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), Col4)
        LClientPhoneNBR.Caption = GetPhoneNbr(txtClientName.Tag, LClientType.Tag)
        Ok = True
    ElseIf Grid.Visible = False And txtClientName.Text <> "" And Val(txtClientName.Tag) <> 0 Then
        TxtModelName.SetFocus
        TxtModelName.SelStart = 0
        TxtModelName.SelLength = Len(TxtModelName.Text)
        Exit Sub
    ElseIf Grid.Visible = False And txtClientName.Text <> "" And Val(txtClientName.Tag) = 0 Then
        Ok = False
        CmdNewCustomer.SetFocus
        txtClientName.Tag = 0
        LClientType.Tag = 0
        LClientType.Caption = ""
        LClientPhoneNBR.Caption = ""
        Exit Sub
        Ok = True
    Else
        Ok = False
        txtClientName.Tag = 0
        txtClientName.Text = ""
        Ok = True
    End If
    Grid.Visible = False
    TxtModelName.SetFocus
    TxtModelName.SelStart = 0
    TxtModelName.SelLength = Len(TxtModelName.Text)
End If
End Sub

Private Sub txtModelQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ComboFees.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub TxtOtherFees_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtItemName.SetFocus
    TxtItemName.SelStart = 0
    TxtItemName.SelLength = Len(TxtItemName.Text)
End If
End Sub

Private Sub TxtOtherFeesPrice_Change()
On Error GoTo errorhandler
LOthersFees.Caption = Val(TxtOtherFeesPrice.Text) * Val(TxtOthersFeesQty.Text)
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

Private Sub TxtOtherFeesPrice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtItemName.SetFocus
    TxtItemName.SelStart = 0
    TxtItemName.SelLength = Len(TxtItemName.Text)
End If
End Sub

Private Sub TxtOthersFeesQty_Change()
On Error GoTo errorhandler
    LOthersFees.Caption = Val(TxtOtherFeesPrice.Text) * Val(TxtOthersFeesQty.Text)
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

Private Sub TxtOthersFeesQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtOtherFeesPrice.SetFocus
    TxtOtherFeesPrice.SelStart = 0
    TxtOtherFeesPrice.SelLength = Len(TxtOtherFeesPrice.Text)
End If

End Sub

Function isOkQty(ÚStkid As String, Qty As Integer) As Boolean
On Error GoTo errorhandler
Dim rs As New ADODB.Recordset
sqlText = "Select FnlQnt from Stkinf where Stkno='" & ÚStkid & "' and strid=" & GetStrId(systemConfigration.MainStoreNo)
Set rs = de.con.Execute(sqlText)
If rs.RecordCount <> 0 Then
    If Qty > rs!fnlqnt Then
        isOkQty = False
        Exit Function
    End If
Else
    isOkQty = False
    Exit Function
End If
isOkQty = True
Exit Function
errorhandler:
MsgBox Err.Description
isOkQty = True
End Function

Private Sub TxtQty_Change()
On Error GoTo errorhandler
    LPrice.Caption = GetPrice(TxtItemName.Tag, 1)
    LDiscount.Caption = GetDiscount()
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

Private Sub TxtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        ComboDiscount.SetFocus
        SendKeys "{home}+{end}"
End If
End Sub

