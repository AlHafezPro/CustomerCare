VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMaintCallManager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "≈œ«—… «·„” ÊÌ« "
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   14655
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Slider Slider1 
      Height          =   345
      Left            =   7530
      TabIndex        =   48
      Top             =   6480
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   609
      _Version        =   393216
      Max             =   20
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexOperation 
      Height          =   1635
      Left            =   7590
      TabIndex        =   32
      Top             =   960
      Width           =   3255
      _cx             =   5741
      _cy             =   2884
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
      RightToLeft     =   0   'False
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
      Height          =   6045
      Left            =   10860
      TabIndex        =   18
      Top             =   720
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   10663
      _Version        =   131074
      Begin VB.ListBox List1 
         ForeColor       =   &H00000080&
         Height          =   3765
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2220
         Width           =   3615
      End
      Begin VB.TextBox TxtEmp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   30
         Width           =   2445
      End
      Begin VSFlex8Ctl.VSFlexGrid MSHFlexGrid1 
         Height          =   1875
         Left            =   60
         TabIndex        =   21
         Top             =   330
         Width           =   3645
         _cx             =   6429
         _cy             =   3307
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
         Ellipsis        =   1
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·„” À„—"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   45
         Width           =   1065
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   780
      Left            =   7890
      TabIndex        =   17
      Top             =   7200
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   1376
      _Version        =   131074
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   27
         Left            =   2250
         TabIndex        =   44
         Tag             =   "0"
         Top             =   7890
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ﬂ‘› Õ”«» «·„Ê«œ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   26
         Left            =   2430
         TabIndex        =   43
         Tag             =   "0"
         Top             =   7650
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "—’Ìœ «·„Ê«œ «·„Œ“‰ÌÂ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   25
         Left            =   2340
         TabIndex        =   42
         Tag             =   "0"
         Top             =   7380
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Õ—ﬂÂ „«œÂ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   24
         Left            =   2280
         TabIndex        =   41
         Tag             =   "0"
         Top             =   7080
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "‰ﬁ· „‰ „” Êœ⁄ ≈·Ï „” Êœ⁄"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   23
         Left            =   2280
         TabIndex        =   40
         Tag             =   "0"
         Top             =   6750
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«·≈œŒ«· Ê«·≈Œ—«Ã"
         Alignment       =   1
      End
      Begin Threed.SSCheck ChkPrintBills 
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   39
         Tag             =   "6"
         Top             =   5010
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   450
         _Version        =   131074
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÿ»«⁄… ›« Ê— «’·Ì…"
      End
      Begin Threed.SSCheck ChkExportItemsBills 
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   38
         Tag             =   "6"
         Top             =   2280
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   450
         _Version        =   131074
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " ’œÌ— ≈‘⁄«—«  «·„Ê«œ «·√Ê·Ì…"
      End
      Begin Threed.SSCheck ChkTransferToMvStock 
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   37
         Tag             =   "6"
         Top             =   1410
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   450
         _Version        =   131074
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " —ÕÌ· ≈· «·„” Êœ⁄« "
      End
      Begin Threed.SSCheck ChkCancelTRansfer 
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   36
         Tag             =   "6"
         Top             =   1980
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   450
         _Version        =   131074
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "≈·€«¡ «· À»Ì "
      End
      Begin Threed.SSCheck ChkAllowTransfer 
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   35
         Tag             =   "6"
         Top             =   1710
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   450
         _Version        =   131074
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " —ÕÌ· ≈·Ï «·„Õ«”»…"
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   22
         Left            =   2250
         TabIndex        =   28
         Tag             =   "0"
         Top             =   6105
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " ’œÌ— «·„⁄·Ê„«  «·«”«”ÌÂ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   21
         Left            =   2250
         TabIndex        =   27
         Tag             =   "0"
         Top             =   5745
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "—›⁄ «”⁄«— «·„Ê«œ «·«Ê·ÌÂ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   20
         Left            =   2250
         TabIndex        =   26
         Tag             =   "8"
         Top             =   2010
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " —„Ì“ ﬁ«∆„…   «„ÊœÌ·« "
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   19
         Left            =   2190
         TabIndex        =   25
         Tag             =   "29"
         Top             =   5445
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«·»«—ﬂÊœ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   18
         Left            =   2220
         TabIndex        =   24
         Tag             =   "28"
         Top             =   5160
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "—»ÿ «·„ÊœÌ·« "
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   17
         Left            =   2220
         TabIndex        =   23
         Tag             =   "27"
         Top             =   4890
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ﬂÊﬂ« ﬂÊ·«"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   16
         Left            =   2250
         TabIndex        =   7
         Tag             =   "9"
         Top             =   2310
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " —„Ì“  «»⁄Ì… «·„ÊœÌ·« "
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   15
         Left            =   2250
         TabIndex        =   9
         Tag             =   "11"
         Top             =   2910
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   14
         Left            =   2250
         TabIndex        =   14
         Tag             =   "26"
         Top             =   4620
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "√—‘›… «·»Ì«‰« "
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   13
         Left            =   2250
         TabIndex        =   13
         Tag             =   "25"
         Top             =   4305
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "≈Õ’«∆Ì… ⁄‰«ÊÌ‰ «·’Ì«‰… «·Œ«—ÃÌ…"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   10
         Left            =   2250
         TabIndex        =   12
         Tag             =   "24"
         Top             =   3945
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÿ»«⁄Â Õ—ﬂ«  «·„” Êœ⁄« "
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   9
         Left            =   2250
         TabIndex        =   8
         Tag             =   "10"
         Top             =   2580
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   12
         Left            =   2250
         TabIndex        =   16
         Tag             =   "23"
         Top             =   8190
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "≈œ«—… «·„” ÊÌ« "
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   2
         Left            =   2250
         TabIndex        =   2
         Tag             =   "3"
         Top             =   540
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   6
         Left            =   2250
         TabIndex        =   10
         Tag             =   "21"
         Top             =   3285
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ã—œ «·„Ê«œ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   3
         Left            =   2250
         TabIndex        =   3
         Tag             =   "4"
         Top             =   795
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "›Ê« Ì— Œœ„… «·„” Â·ﬂ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   1
         Left            =   3840
         TabIndex        =   1
         Tag             =   "2"
         Top             =   270
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   5
         Left            =   2250
         TabIndex        =   5
         Tag             =   "6"
         Top             =   1395
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«·»ÕÀ ⁄‰ ›« Ê—…"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   7
         Left            =   2250
         TabIndex        =   11
         Tag             =   "22"
         Top             =   3615
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«·⁄‰«ÊÌ‰ «·„ﬂ——…"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   0
         Left            =   2250
         TabIndex        =   0
         Tag             =   "1"
         Top             =   30
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   8
         Left            =   2250
         TabIndex        =   6
         Tag             =   "7"
         Top             =   1695
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   11
         Left            =   2250
         TabIndex        =   15
         Tag             =   "0"
         Top             =   6435
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«·„” Êœ⁄« "
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   4
         Left            =   2250
         TabIndex        =   4
         Tag             =   "5"
         Top             =   1095
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2130
      Top             =   0
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
            Picture         =   "FrmMaintCallManager.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":4E7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":7777
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":A126
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":C64B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":EE03
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":11817
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":14169
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":16EDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":1972C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":1C5D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":1F32D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":21CCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":24C2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":27A8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":2A4B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":2CE6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":2F79F
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":323DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":34CE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":3774B
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":3A6FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":3D025
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":3F95A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":42509
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":44C49
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":475B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":49E40
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":4C04F
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":4E9AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":511D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":53C8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":566C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":591DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":5C176
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":5EF42
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMaintCallManager.frx":61BBC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexMaintType 
      Height          =   1635
      Left            =   7560
      TabIndex        =   33
      Top             =   2880
      Width           =   3255
      _cx             =   5741
      _cy             =   2884
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
      RightToLeft     =   0   'False
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
   Begin VSFlex8Ctl.VSFlexGrid FlexPayment 
      Height          =   1635
      Left            =   7560
      TabIndex        =   34
      Top             =   4800
      Width           =   3225
      _cx             =   5689
      _cy             =   2884
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
      RightToLeft     =   0   'False
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
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   6075
      Left            =   0
      TabIndex        =   45
      Top             =   720
      Width           =   7455
      _cx             =   13150
      _cy             =   10716
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   14655
      _ExtentX        =   25850
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
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Label LDiscountPercentage 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   8610
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   6510
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "‰”»Â «·Õ”„ ⁄·Ï «·„Ê«œ"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   9240
      TabIndex        =   46
      Top             =   6510
      Width           =   1545
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "‰Ê⁄ «·’Ì«‰…"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   7560
      TabIndex        =   31
      Top             =   2670
      Width           =   780
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "‰Ê⁄ «·⁄„·Ì…"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   6
      Left            =   7620
      TabIndex        =   30
      Top             =   750
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   7560
      TabIndex        =   29
      Top             =   4560
      Width           =   795
   End
End
Attribute VB_Name = "FrmMaintCallManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ColNo = 1
Const ColName = 2
Const ColChk = 3


Const ColTagId = 1
Const ColTagName = 2
Const coltagchk = 3




Const ColEmpNo = 1
Const Colfullname = 2

Dim rs As New ADODB.Recordset
Dim X As Double
Dim Flag As Boolean
Function GetTypes(flexGrid As VSFlexGrid, Col As Integer) As String
Dim Str As String
Str = ""
With flexGrid
    For i = 1 To .Rows - 1
        If .TextMatrix(i, ColChk) Then
            Str = Str & "," & .TextMatrix(i, Col)
        End If
    Next
    If Str <> "" Then
        GetTypes = Mid(Str, 2)
    Else
        GetTypes = ""
    End If
End With
End Function

Function FillCompStr() As String
Dim Str As String
Str = ""
With flexGrid
    For i = 1 To .Rows - 1
        If .TextMatrix(i, ColChk) = "-1" Then
            Str = Str & "," & .TextMatrix(i, ColCompNo)
        End If
    Next
End With
Str = Mid(Str, 2)
FillCompStr = Str
End Function

Sub FillInitGrid()

Dim RsOperationType As New ADODB.Recordset
sqlText = "select OpNo , OpName , 0 Chk  from operkind"
Set RsOperationType = de.con.Execute(sqlText)
Set FlexOperation.DataSource = RsOperationType
FillFormating FlexOperation
FlexOperation.Editable = flexEDKbdMouse

    Dim rsPayment As New ADODB.Recordset
    sqlText = "Select No , Name , 0 Chk  From PayMethod"
    Set rsPayment = de.con.Execute(sqlText)
    Set FlexPayment.DataSource = rsPayment
    FillFormating FlexPayment
    FlexPayment.Editable = flexEDKbdMouse
    
    Dim rsType As New ADODB.Recordset
    sqlText = "select no , stat , 0 Chk from dbo.maintypestat"
    Set rsType = de.con.Execute(sqlText)
    Set FlexMaintType.DataSource = rsType
    FillFormating FlexMaintType
    FlexMaintType.Editable = flexEDKbdMouse
    
End Sub
Sub init()
Dim rs As New ADODB.Recordset
    Me.top = 0
    Me.left = 0
    
    sqlText = "Select TagId , TagName , 0 chk From comainttag"
    Set rs = de.con.Execute(sqlText)
    Set flexGrid.DataSource = rs
    FillFormating flexGrid
    flexGrid.Editable = flexEDKbdMouse
    
    sqlText = "Select Top 5 EmpNo , FullName From Dbo.FullName where EmpNo=0"
    Set rs = de.con.Execute(sqlText)
    Set MSHFlexGrid1.DataSource = rs
    MSHFlexGrid1.FormatString = FillFs
    SetColWidths ColEmpNo, MSHFlexGrid1
    SetColWidths Colfullname, MSHFlexGrid1
    FillList
    FillInitGrid
    msgFlag = True
End Sub
Sub FillFormating(flexGrid As VSFlexGrid)
    fs = "|>" + "TagId"
    fs = fs + "|>" + "Tag Name"
    fs = fs + "|>" + "Chk"
   With flexGrid
        .FormatString = fs
        .Cols = 4
        .ColWidth(coltagchk) = 400
        .ColDataType(coltagchk) = flexDTBoolean
        .ColWidth(ColTagId) = 0
        SetColWidths ColTagName, flexGrid
   End With
End Sub
Function GetAllow(ByVal Vindex As Integer) As Double
X = 0
Select Case Vindex
    Case 1
        For i = ChkAllowTransfer.LBound To ChkAllowTransfer.UBound
            If ChkAllowTransfer(i).Value Then
                X = X + (2 ^ ChkAllowTransfer(i).Tag)
            End If
        Next
   Case 2
        For i = ChkCancelTRansfer.LBound To ChkCancelTRansfer.UBound
            If ChkCancelTRansfer(i).Value Then
                X = X + (2 ^ ChkCancelTRansfer(i).Tag)
            End If
        Next
   Case 3
        For i = ChkTransferToMvStock.LBound To ChkTransferToMvStock.UBound
            If ChkTransferToMvStock(i).Value Then
                X = X + (2 ^ ChkTransferToMvStock(i).Tag)
            End If
        Next
   Case 4
        For i = ChkExportItemsBills.LBound To ChkExportItemsBills.UBound
            If ChkExportItemsBills(i).Value Then
                X = X + (2 ^ ChkExportItemsBills(i).Tag)
            End If
        Next
   Case 5
        For i = ChkPrintBills.LBound To ChkPrintBills.UBound
            If ChkPrintBills(i).Value Then
                X = X + (2 ^ ChkPrintBills(i).Tag)
            End If
        Next
End Select
GetAllow = X
End Function
Sub SaveRec()
On Error GoTo ERORHANDLER

Dim X As Double
X = SetPermission

Dim OperationStr As String
Dim MaintTypeStr As String
Dim PaymentStr As String
Dim AlowTransfer As Double
Dim CancelAccount As Double
Dim AllowTransferStock As Double
Dim AllowExportItems As Double
Dim AllowPrintBills As Double
Dim DiscountPrecentag As Integer

OperationStr = GetTypes(FlexOperation, ColNo)
MaintTypeStr = GetTypes(FlexMaintType, ColNo)
PaymentStr = GetTypes(FlexPayment, ColNo)
DiscountPrecentag = Val(LDiscountPercentage.Caption)


AlowTransfer = GetAllow(1)
CancelAccount = GetAllow(2)
AllowTransferStock = GetAllow(3)
AllowExportItems = GetAllow(4)
AllowPrintBills = GetAllow(5)


With List1
    If InTable(.ItemData(.ListIndex)) Then
        Update .ItemData(.ListIndex), X, OperationStr, MaintTypeStr, PaymentStr, AlowTransfer, CancelAccount, AllowTransferStock, AllowExportItems, AllowPrintBills, DiscountPrecentag
    Else
        InsertAccount .ItemData(.ListIndex), X, OperationStr, MaintTypeStr, PaymentStr, AlowTransfer, CancelAccount, AllowTransferStock, AllowExportItems, AllowPrintBills, DiscountPrecentag
    End If
End With
MsgBox " „  Œ“Ì‰ «·’·«ÕÌ«  ", vbMsgBoxRight + vbSystemModal + vbApplicationModal + vbInformation, " Œ“Ì‰"

Exit Sub
ERORHANDLER:
MsgBox Err.Description
End Sub
Function FoundEmpNo(TEmpNo As Integer) As Boolean
sqlText = "Select * From MaintUsers where empNo= " & TEmpNo
Set rs = de.con.Execute(sqlText)
If rs.EOF And rs.BOF Then
    FoundEmpNo = False
Else
    FoundEmpNo = True
End If
End Function

Sub ClearChk()
For i = 0 To Check.UBound
    Check(i).Value = ssCBUnchecked
Next
End Sub

Sub FillList()
    sqlText = "Select a.EmpNo , Permission , FullName  From MaintUsers a, dbo.FullName f where a.empno = f.empno "
    Set rs = de.con.Execute(sqlText)
    If rs.EOF And rs.BOF Then Exit Sub
    rs.MoveFirst
    Do While Not rs.EOF
        With List1
           .AddItem rs!FullName
           .ItemData(.NewIndex) = rs!empNo
           rs.MoveNext
        End With
    Loop
    List1.ListIndex = 0
End Sub

Function Found(Vrow As Integer) As Boolean
With List1
    For i = 0 To .ListCount - 1
        If .ItemData(i) = MSHFlexGrid1.TextMatrix(Vrow, ColEmpNo) Then
            Found = True
            Exit Function
        Else
            Found = False
        End If
    Next i
End With
End Function

Function FillFs() As String
    fs = "|<" + "—/„"
    fs = fs + "|<" + "≈”„ «·„” À„—"
    
    FillFs = fs
End Function


Sub SetColWidths(ColNo As Integer, MSHFlexGrid1 As VSFlexGrid)
    With MSHFlexGrid1
        .AutoSize (ColNo)
    End With
End Sub

Private Sub CmdDetailsFix_Click()
    msgFlag = False
    CmdSave_Click
    msgFlag = True
    FrmSubDetails.Show 1
End Sub
Private Sub CmdDetailsReport_Click()
                msgFlag = False
                CmdSave_Click
                msgFlag = True
                FrmSubDetailsReport.Show 1
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Function SetPermission() As Double
Dim rs As New ADODB.Recordset
Dim X As Double
    X = 0
    For i = 0 To Check.UBound
        If Check(i).Value = True Then
            X = X + (2 ^ Check(i).Tag)
        End If
    Next i
    SetPermission = X
End Function
Sub Update(empNo As Integer, Permission As Double, OperationStr As String, MaintTypeStr As String, PaymentStr As String, AllowTransfer As Double, CancelAccount As Double, AllowTransferStock As Double, AllowExportItems As Double, AllowPrintBills As Double, DiscountPrecentage As Integer)
    sqlText = "Update MaintUsers Set Permission=" & Permission & ",OperationStr='" & OperationStr & "',MaintTypeStr='" & MaintTypeStr & "',PaymentStr='" & PaymentStr & "',AllowTransfer=" & AllowTransfer & ",CancelAccount=" & CancelAccount & ",AllowTransferStock=" & AllowTransferStock & ",AllowExportItems =" & AllowExportItems & ",AllowPrintBills=" & AllowPrintBills & ",DiscountPercentage=" & DiscountPrecentage & "  Where EmpNo=" & empNo
    de.con.Execute (sqlText)
End Sub
Sub InsertAccount(empNo As Integer, Permission As Double, OperationStr As String, MaintTypeStr As String, PaymentStr As String, AllowTransfer As Double, CancelAccount As Double, AllowTransferStock As Double, AllowExportItems As Double, AllowPrintBills As Double, DiscountPrecentage As Integer)
    sqlText = "Insert into MaintUsers (EmpNo ,Permission , OperationStr , MaintTYpeStr , PaymentStr , AllowTransfer  ,CancelAccount,AllowTransferStock,AllowExportItems , AllowPrintBills,DiscountPercentage) Values( " & empNo & "," & Permission & ",'" & OperationStr & "','" & MaintTypeStr & "','" & PaymentStr & "'," & AllowTransfer & "," & CancelAccount & "," & AllowTransferStock & "," & AllowExportItems & "," & AllowPrintBills & "," & DiscountPrecentage & ")"
     de.con.Execute (sqlText)
End Sub
Function InTable(empNo As Integer) As Boolean
    sqlText = "Select * From MaintUsers Where EmpNo=" & empNo
    Set rs = de.con.Execute(sqlText)
    If rs.RecordCount = 0 Then
        InTable = False
    Else
        InTable = True
    End If
End Function
Private Sub CmdSave_Click()
SaveRec
End Sub
Private Sub Command1_Click()
MsgBox SubReport
End Sub

Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrorHandler
With flexGrid
    sqlText = "Delete From comaintpermission Where EmpNo =" & List1.ItemData(List1.ListIndex) & " and TagId=" & .TextMatrix(Row, ColNo)
    de.con.Execute (sqlText)
    If .TextMatrix(Row, ColChk) Then
        sqlText = "Insert Into comaintpermission(TagId , EmpNo) Values(" & .TextMatrix(Row, ColNo) & "," & List1.ItemData(List1.ListIndex) & ")"
        de.con.Execute (sqlText)
    End If
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description

End Sub

Private Sub FlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    If Col <> ColChk Then cancel = True
End Sub

Private Sub FlexMaintType_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrorHandler
With FlexMaintType
    sqlText = "Delete From maintypestatPermission Where EmpNo =" & List1.ItemData(List1.ListIndex) & " and NO=" & .TextMatrix(Row, ColNo)
    de.con.Execute (sqlText)
    If .TextMatrix(Row, ColChk) Then
        sqlText = "Insert Into maintypestatPermission(No, EmpNo) Values(" & .TextMatrix(Row, ColNo) & "," & List1.ItemData(List1.ListIndex) & ")"
        de.con.Execute (sqlText)
    End If
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description

End Sub

Private Sub FlexOperation_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrorHandler
With FlexOperation
    sqlText = "Delete From OperKindPermission Where EmpNo =" & List1.ItemData(List1.ListIndex) & " and opno=" & .TextMatrix(Row, ColNo)
    de.con.Execute (sqlText)
    If .TextMatrix(Row, ColChk) Then
        sqlText = "Insert Into OperKindPermission(opno, EmpNo) Values(" & .TextMatrix(Row, ColNo) & "," & List1.ItemData(List1.ListIndex) & ")"
        de.con.Execute (sqlText)
    End If
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description

End Sub

Private Sub FlexPayment_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrorHandler
With FlexPayment
    sqlText = "Delete From PayMethodPermission Where EmpNo =" & List1.ItemData(List1.ListIndex) & " and NO=" & .TextMatrix(Row, ColNo)
    de.con.Execute (sqlText)
    If .TextMatrix(Row, ColChk) Then
        sqlText = "Insert Into PayMethodPermission(No, EmpNo) Values(" & .TextMatrix(Row, ColNo) & "," & List1.ItemData(List1.ListIndex) & ")"
        de.con.Execute (sqlText)
    End If
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description

End Sub

Private Sub Form_Load()
    init
End Sub
Sub ClearPermissionW_R_D(ByVal Vindex As Integer)
Select Case Vindex
    Case 1 ' Allow Transfer Account
        For i = ChkAllowTransfer.LBound To ChkAllowTransfer.UBound
            ChkAllowTransfer(i).Value = ssCBUnchecked
        Next
    Case 3 ' Cancel TRansfer Account
        For i = ChkCancelTRansfer.LBound To ChkCancelTRansfer.UBound
            ChkCancelTRansfer(i).Value = ssCBUnchecked
        Next
    Case 4 ' TRansfer To MvStock
        For i = ChkTransferToMvStock.LBound To ChkTransferToMvStock.UBound
            ChkTransferToMvStock(i).Value = ssCBUnchecked
        Next
    Case 5 ' Export Items Bills
        For i = ChkExportItemsBills.LBound To ChkExportItemsBills.UBound
                ChkExportItemsBills(i).Value = ssCBUnchecked
        Next
End Select
End Sub
Private Sub List1_Click()
Flag = False
With List1
tempempno = .ItemData(.ListIndex)

    sqlText = "Select  TagId , TagName ,chk From fn_MaintPermission(" & tempempno & ")"
    Set rs = de.con.Execute(sqlText)
    Set flexGrid.DataSource = rs
    FillFormating flexGrid


    sqlText = "Select  opno , opname , chk From fn_operkindPermission(" & tempempno & ")"
    Set rs = de.con.Execute(sqlText)
    Set FlexOperation.DataSource = rs
    FillFormating FlexOperation

    sqlText = "Select  no , name , chk From fn_PaymethodPermission(" & tempempno & ")"
    Set rs = de.con.Execute(sqlText)
    Set FlexPayment.DataSource = rs
    FillFormating FlexPayment

    sqlText = "Select  no , stat , chk From fn_maintypestatPermission(" & tempempno & ")"
    Set rs = de.con.Execute(sqlText)
    Set FlexMaintType.DataSource = rs
    FillFormating FlexMaintType


    sqlText = "   select isnull(DiscountPercentage,0)  DiscountPercentage From MaintUsers  Where empno=" & tempempno
    Set rs = de.con.Execute(sqlText)
    If rs.RecordCount > 0 Then
    LDiscountPercentage.Caption = rs!DiscountPercentage
    Slider1.Value = rs!DiscountPercentage
    Else
        LDiscountPercentage.Caption = 0
        Slider1.Value = 0
    End If

End With



Flag = True
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    With List1
        sqlText = "Delete From MaintUsers where EmpNo=" & .ItemData(.ListIndex)
        de.con.Execute (sqlText)
        .RemoveItem .ListIndex
    End With
End If
End Sub

Private Sub MSHFlexGrid1_DblClick()
With MSHFlexGrid1
    If Not Found(.Row) Then
        List1.AddItem .TextMatrix(.Row, Colfullname)
        List1.ItemData(List1.NewIndex) = .TextMatrix(.Row, ColEmpNo)
    End If
End With
End Sub

Private Sub Slider1_Change()
LDiscountPercentage.Caption = Slider1.Value
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        Unload Me
    Case 3
        SaveRec
End Select
End Sub

Private Sub TxtEmp_Change()
'On Error Resume Next
    If IsNumeric(TxtEmp.Text) Then
        sqlText = "Select Top 5 EmpNo , FullName From Dbo.FullName Where EmpNo =" & TxtEmp.Text
        Set rs = de.con.Execute(sqlText)
    Else
        sqlText = "Select Top 5 EmpNo , FullName From Dbo.FullName Where FullName like '%" & IIf(LTrim(RTrim(TxtEmp.Text)) = "", 0, LTrim(RTrim(TxtEmp.Text))) & "%' Order By FullName"
        Set rs = de.con.Execute(sqlText)
    End If
    Set MSHFlexGrid1.DataSource = rs
    MSHFlexGrid1.FormatString = FillFs
    SetColWidths ColEmpNo, MSHFlexGrid1
    SetColWidths Colfullname, MSHFlexGrid1
End Sub

Private Sub TxtEmp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With MSHFlexGrid1
            If .Rows = 1 Then Exit Sub
                If Not Found(1) Then
                    List1.AddItem .TextMatrix(1, Colfullname)
                    List1.ItemData(List1.NewIndex) = .TextMatrix(1, ColEmpNo)
                End If
        End With
    End If
End Sub
