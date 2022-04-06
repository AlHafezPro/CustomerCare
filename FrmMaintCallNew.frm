VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmMaintCallNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·‘ﬂ«ÊÏ"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   2835
      Left            =   5520
      TabIndex        =   36
      Top             =   4950
      Visible         =   0   'False
      Width           =   2415
      _cx             =   4260
      _cy             =   5001
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
   Begin MSDataListLib.DataCombo ComboProductFamily 
      Height          =   480
      Left            =   7650
      TabIndex        =   2
      Top             =   1290
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   847
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport Cr1 
      Left            =   5130
      Top             =   2550
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSFrame ClientFrame 
      DragMode        =   1  'Automatic
      Height          =   2595
      Left            =   60
      TabIndex        =   21
      Top             =   1920
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   4577
      _Version        =   131074
      Begin VB.TextBox TxtCustomerName 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   7590
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   450
         Width           =   4935
      End
      Begin VB.TextBox TxtClientNote 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000FF&
         Height          =   465
         Left            =   4980
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   2070
         Width           =   10125
      End
      Begin VB.TextBox TxtClientDefindname 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2040
         Width           =   4815
      End
      Begin VB.TextBox TxtAddress 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1230
         Width           =   14985
      End
      Begin VB.TextBox TxtZoneName 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   450
         Width           =   4845
      End
      Begin VB.TextBox TxtHomePhone 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   12555
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   450
         Width           =   2595
      End
      Begin VB.TextBox TxtMobilePhone 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   4950
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   450
         Width           =   2595
      End
      Begin VB.Label LRepeatClaimTeamName 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   7620
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   90
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Label LClaimsIsRepeat 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "«·‘ﬂÊÏ „ﬂ——Â"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   10110
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   90
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "≈”„ «·“»Ê‰"
         Height          =   195
         Index           =   11
         Left            =   11805
         RightToLeft     =   -1  'True
         TabIndex        =   53
         ToolTipText     =   "«·»ÕÀ ›ﬁÿ ⁄‰ ÿ—Ìﬁ  —ﬁ„ «·Â« › Ê «·„Ê»«Ì·"
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„·«ÕŸ«  ⁄‰ «·“»Ê‰"
         DragMode        =   1  'Automatic
         Height          =   195
         Index           =   9
         Left            =   13800
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1740
         Width           =   1290
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "»„⁄—›… "
         Height          =   195
         Index           =   10
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1740
         Width           =   510
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·⁄‰Ê«‰"
         Height          =   195
         Index           =   8
         Left            =   14610
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   900
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„‰ÿﬁ…"
         Height          =   195
         Index           =   7
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   150
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·Â« › «·À«» "
         Height          =   195
         Index           =   4
         Left            =   13980
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   150
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·„Ê»«Ì·"
         Height          =   195
         Index           =   5
         Left            =   6705
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   150
         Width           =   825
      End
   End
   Begin VB.TextBox TxtSearch 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   10260
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1290
      Width           =   4965
   End
   Begin VB.TextBox TxtIssueName 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1290
      Width           =   7515
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   885
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   1561
      _Version        =   131074
      Begin VB.Label LastCallNbr 
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
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   1763
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   450
         Width           =   1395
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «Œ— ‘ﬂÊÏ"
         Height          =   195
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   150
         Width           =   1035
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ √Œ— “»«—Â"
         Height          =   195
         Left            =   375
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   150
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «·‘ﬂÊÏ"
         Height          =   195
         Index           =   1
         Left            =   12570
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   150
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Êﬁ  «·‘ﬂÊÏ"
         Height          =   195
         Index           =   2
         Left            =   11130
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   120
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„” ·„ «·‘ﬂÊÏ"
         Height          =   195
         Index           =   3
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "⁄œœ „—«  «·“Ì«—…"
         Height          =   195
         Left            =   3750
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   150
         Width           =   1110
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·‘ﬂÊÏ"
         Height          =   195
         Index           =   5
         Left            =   14295
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Õ«·Â «·‘ﬂÊÏ"
         Height          =   195
         Index           =   4
         Left            =   9570
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   150
         Width           =   900
      End
      Begin VB.Label LastDataVisit 
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
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   450
         Width           =   1395
      End
      Begin VB.Label LMaintCallStateName 
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
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   7982
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   450
         Width           =   2505
      End
      Begin VB.Label LCallNo 
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
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   13905
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   450
         Width           =   1245
      End
      Begin VB.Label LCLientVisitorCount 
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
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   3466
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   450
         Width           =   1395
      End
      Begin VB.Label LCallReceiverName 
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
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   5169
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   450
         Width           =   2505
      End
      Begin VB.Label LCalltime 
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
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   10795
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   450
         Width           =   1245
      End
      Begin VB.Label LCallDate 
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
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   12348
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   450
         Width           =   1245
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   615
      Left            =   60
      TabIndex        =   26
      Top             =   4560
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   1085
      _Version        =   131074
      Begin Threed.SSCommand cmdSaveClient 
         Height          =   555
         Left            =   9540
         TabIndex        =   11
         Top             =   30
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   979
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
         Caption         =   "Õ›Ÿ „⁄·Ê„«  «·“»Ê‰"
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   3834
         TabIndex        =   13
         Top             =   30
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   979
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
         Caption         =   "ÿ»«⁄Â"
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   555
         Left            =   1932
         TabIndex        =   38
         Top             =   30
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   979
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
      Begin Threed.SSCommand CmdExit 
         Height          =   555
         Left            =   30
         TabIndex        =   29
         Top             =   30
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   979
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
      Begin Threed.SSCommand cmdSave 
         Height          =   555
         Left            =   7638
         TabIndex        =   12
         Top             =   30
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   979
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
         Caption         =   "Õ›Ÿ «·‘ﬂÊÏ"
      End
      Begin Threed.SSCommand CmdCancel 
         Height          =   555
         Left            =   5736
         TabIndex        =   28
         Top             =   30
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   979
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
      Begin Threed.SSCommand CmdAdd 
         Height          =   555
         Left            =   13350
         TabIndex        =   0
         Top             =   30
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   979
         _Version        =   131074
         ForeColor       =   8388608
         PictureAnimationDelay=   0
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
      Begin Threed.SSCommand CmdEdit 
         Height          =   555
         Left            =   11442
         TabIndex        =   27
         Top             =   30
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   979
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
   End
   Begin Threed.SSFrame NavigatorFrame 
      Height          =   375
      Left            =   12720
      TabIndex        =   30
      Top             =   5340
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   661
      _Version        =   131074
      Begin VB.CommandButton CmdFirst 
         Height          =   285
         Left            =   60
         Picture         =   "FrmMaintCallNew.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "First"
         Top             =   60
         Width           =   255
      End
      Begin VB.CommandButton CmdPrevious 
         Height          =   285
         Left            =   330
         Picture         =   "FrmMaintCallNew.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Previous"
         Top             =   60
         Width           =   255
      End
      Begin VB.CommandButton CmdNext 
         Height          =   285
         Left            =   1920
         Picture         =   "FrmMaintCallNew.frx":062C
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Next"
         Top             =   60
         Width           =   255
      End
      Begin VB.CommandButton CmdLast 
         Height          =   285
         Left            =   2190
         Picture         =   "FrmMaintCallNew.frx":0726
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Last"
         Top             =   60
         Width           =   255
      End
      Begin VB.Label LNavigator 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   600
         TabIndex        =   35
         Top             =   60
         Width           =   1305
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·»ÕÀ ⁄‰ «·“»Ê‰"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   13845
      RightToLeft     =   -1  'True
      TabIndex        =   20
      ToolTipText     =   "«·»ÕÀ ›ﬁÿ ⁄‰ ÿ—Ìﬁ  —ﬁ„ «·Â« › Ê «·„Ê»«Ì·"
      Top             =   960
      Width           =   1320
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·⁄«∆·…"
      Height          =   195
      Index           =   0
      Left            =   9765
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1020
      Width           =   420
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·⁄ÿ·"
      Height          =   195
      Index           =   1
      Left            =   7170
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1020
      Width           =   390
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint 
         Caption         =   " ﬁ—Ì— «·‘ﬂ«ÊÏ Ê«·«’·«Õ« "
      End
   End
   Begin VB.Menu mnuRefresh 
      Caption         =   "mnuRefresh"
      Visible         =   0   'False
      Begin VB.Menu mnuRefreshMaintCalls 
         Caption         =   " ÕœÌÀ «·»Ì«‰« "
      End
   End
End
Attribute VB_Name = "FrmMaintCallNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsNavigator As New ADODB.Recordset
Dim Ok As Boolean, Flag As Boolean, Pos As Integer, RecNum  As Double, maintCallState As EnumState
Dim EntryFullName As String

Dim maintDataService_ As New MaintDataService
Dim clientValidationResult As New clientValidationResult
Dim maintCallViewModelInfo_ As MaintCallViewModel

Dim FieldsArrayToSearch(4) As New MapClientField

Dim FieldsZoneArrayToSearch(2) As New MapClientField

'Dim FieldReceiverArrayToSearch(2) As MapClientField

'Const ColNo = 1
'Const ColName = 2
Dim Cols(4) As New MapClientField
Dim ZoneCols(2) As New MapClientField
'Dim ReceiverCols(2) As New MapClientField


Const ColNo = 1
Const ColName = 2

'Dim prevDataIsNull As Boolean
'Dim prevSendrLength As Integer
Dim isOkayToSearch As Boolean


Sub ChangeCursor(sender As Control, Optional top As Variant, Optional left As Variant)

    With sender
       Grid.top = top + .top + .Height
       Grid.left = left + .left
       Grid.Width = .Width
    End With

End Sub
Sub FillCombos()
Dim rsproductFamilly As New ADODB.Recordset
sqlText = "select prodFamNo , prodFamName  from AdhamProductFamily"
Set rsproductFamilly = de.con.Execute(sqlText)
Set ComboProductFamily.RowSource = rsproductFamilly
ComboProductFamily.listField = "prodFamName"
ComboProductFamily.BoundColumn = "prodFamNo"
End Sub

Sub EnableControls(FControl As Boolean)
Dim ctrl As Control
For Each ctrl In Me.Controls
    If TypeOf ctrl Is TextBox Or TypeOf ctrl Is MaskEdBox Or TypeOf ctrl Is CheckBox Or TypeOf ctrl Is VSFlexGrid Or TypeOf ctrl Is DataCombo Then
        ctrl.Enabled = FControl
    End If
Next
End Sub

Sub EnableCmds(FAdd As Boolean, FEdit As Boolean, Fprint As Boolean, FSave As Boolean, FClientSave As Boolean, FUndo As Boolean, FSearch As Boolean, FNavigator As Boolean)
    CmdAdd.Enabled = FAdd
    CmdEdit.Enabled = FEdit
    cmdPrint.Enabled = Fprint
    cmdSave.Enabled = FSave
    cmdSaveClient.Enabled = FClientSave
    CmdCancel.Enabled = FUndo
    'CmdPrint.Enabled = Fprint
    cmdSearch.Enabled = FSearch
    NavigatorFrame.Enabled = FNavigator
End Sub


Sub ClearControls()
On Error GoTo ErrorHandler
Ok = False
    LCallNo.Caption = ""
    LCallDate.Caption = Format(maintCallViewModelInfo_.CallDateTime, "dd/MM/yyyy")
    LCalltime.Caption = Format(maintCallViewModelInfo_.CallDateTime, "HH:MM")
    LCallReceiverName.Caption = maintCallViewModelInfo_.CallReceiverName
    LCallReceiverName.Tag = maintCallViewModelInfo_.CallReceiverEmpNo
    LastDataVisit.Caption = ""
    LastCallNbr.Caption = ""
    ComboProductFamily.BoundText = -1
    TxtIssueName.Text = ""
    TxtSearch.Tag = ""
    TxtSearch.Text = ""
    TxtCustomerName.Tag = ""
    TxtCustomerName.Text = ""
    Me.LCLientVisitorCount.Caption = ""
    LClaimsIsRepeat.Visible = False
    LRepeatClaimTeamName.Caption = ""
    LRepeatClaimTeamName.Visible = False
    TxtMobilePhone.Text = ""
    TxtHomePhone.Text = ""
    TxtZoneName.Tag = ""
    TxtZoneName.Text = ""
    TxtAddress.Text = ""
    TxtClientNote.Text = ""
    TxtClientDefindname.Text = ""
'    TxtNotes.Text = ""
'    TxtDefindname.Text = ""
    LMaintCallStateName.Caption = maintDataService_.GetMaintCallState(maintCallViewModelInfo_.callState)
Ok = True
Exit Sub
ErrorHandler:
Ok = True
MsgBox Err.Description
End Sub

Private Sub CmdAdd_Click()
searchClientIsAllow = False
Set maintCallViewModelInfo_ = New MaintCallViewModel
maintCallState = NewRecord
EnableCmds False, False, False, True, True, True, False, False
EnableControls True
ClearControls
searchClientIsAllow = True
TxtSearch.SetFocus
End Sub

Private Sub CmdCancel_Click()
On Error GoTo ErrorHandler
    EnableCmds True, True, True, False, False, False, True, True
    EnableControls False
    
    If maintCallViewModelInfo_.CallNo <> 0 Then
        MoveToRec maintCallViewModelInfo_.CallNo, True
    Else
        MoveToRec Val(RsNavigator!CallNo), True
    End If
Exit Sub
ErrorHandler:
MsgBox Err.Description

End Sub

Private Sub CmdDelete_Click()
On Error GoTo ErrorHandler
If maintCallViewModelInfo_.callState = Repared Or maintCallViewModelInfo_.callState = UnderwayAndPrinted Then
    MsgBox "·«Ì„ﬂ‰  ⁄œÌ· «Ê Õ‹–› «·‘ﬂÊÏ", vbExclamation, "«·‘ﬂÊÏ „‰›‹–Â √Ê  „ ÿ»«⁄ Â«"
    Exit Sub
End If
Dim result As maintCallResult

If MsgBox("Â· √‰  „ √ﬂœ „‰ Õ–› «·‘ﬂÊÏ", vbYesNo + vbDefaultButton2, "Õ–›") = vbYes Then
    If maintCallViewModelInfo_.Id <> 0 Then
        maintCallState = DeleteRecord
        Set result = SaveChanges(True)
        If result.MaintCallResultStatus Then
            InitNavigator
            MoveNavigator 4
            EnableCmds True, True, True, False, False, False, True, True
            MsgBox " „ Õ–› «·‘ﬂÊÏ »‰Ã«Õ", vbInformation, "Õ–› «·‘ﬂÊÏ"
        Else
            MsgBox result.MaintCallResultDescription, vbExclamation + vbMsgBoxRight, "Œÿ√ ›Ì «·Õ–›"
            Dim ctrl As Control
            Set ctrl = Me.GetTheLastFocusControl
            ctrl.SetFocus
            SendKeys "{home}+{end}"
        End If

    End If
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Sub DeleteMaintCall()

End Sub
Private Sub CmdEdit_Click()
searchClientIsAllow = False
    If maintCallViewModelInfo_.callState = Repared Then
        MsgBox "·«Ì„ﬂ‰  ⁄œÌ· «Ê Õ‹–› «·‘ﬂÊÏ", vbExclamation, "«·‘ﬂÊÏ „‰›‹–Â"
        Exit Sub
    End If
    EnableCmds False, False, False, True, True, True, False, False
    EnableControls True
    ComboProductFamily.SetFocus
    maintCallState = DefaultRecord
    maintCallViewModelInfo_.clientInfo.ClientState = Default
    searchClientIsAllow = True
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdFirst_Click()

    MoveNavigator 1

End Sub

Private Sub CmdLast_Click()

    MoveNavigator 4

End Sub

Private Sub CmdNext_Click()

    MoveNavigator 3


End Sub

Private Sub CmdPrevious_Click()

    MoveNavigator 2

End Sub

Sub FillControls()
On Error GoTo ErrorHandler
Ok = False

With maintCallViewModelInfo_
'    LId.Caption = .Id
    LCallNo.Caption = IIf(IsNull(.CallNo), "", .CallNo)
    LCallDate.Caption = Format(.CallDateTime, "dd/mm/yyyy")
    LCalltime.Caption = Format(.CallDateTime, "HH:MM")
    LCallReceiverName.Tag = IIf(IsNull(.CallEntryEmpNo), "", .CallReceiverEmpNo)
    LCallReceiverName.Caption = IIf(IsNull(.CallEntryName), "", .CallReceiverName)
    LastDataVisit.Caption = IIf(IsNull(.LastVisitDate), "", .LastVisitDate)
    LastCallNbr.Caption = IIf(IsNull(.LastCallNo), "", .LastCallNo)
    ChangeStatus (.callState)
    ComboProductFamily.BoundText = .ModNo
    TxtIssueName.Text = .CallDEscription & ""
    
    TxtSearch.Tag = .cliNo
    TxtSearch.Text = .clientInfo.AdhamName & ""
    
    TxtCustomerName.Tag = .cliNo
    TxtCustomerName.Text = .clientInfo.AdhamName & ""
    LCLientVisitorCount.Caption = .ClientVisitorCount
    LClaimsIsRepeat.Visible = .ClaimIsRepeat
    If .ClaimIsRepeat Then
         LRepeatClaimTeamName.Visible = True
         LRepeatClaimTeamName.Caption = .ClaimTeamName
    Else
         LRepeatClaimTeamName.Visible = False
         LRepeatClaimTeamName.Caption = ""
    End If
    TxtMobilePhone.Text = .clientInfo.MobilePhone & ""
    TxtHomePhone.Text = .clientInfo.AdhamPhon & ""
    TxtZoneName.Tag = .clientInfo.Zone
    TxtZoneName.Text = .clientInfo.ZoneName
    TxtAddress.Text = .clientInfo.AdhamAdress & ""
    TxtClientNote.Text = .clientInfo.Notes & ""
    TxtClientDefindname.Text = .clientInfo.Defindname & ""
    
'    TxtNotes.Text = .Notes & ""
'    TxtDefindname.Text = .Defindname & ""
    
'    If Not IsNull(.CallReceiverEmpNo) And .CallReceiverEmpNo <> 0 Then
'        TxtReceiverName.Tag = .CallReceiverEmpNo
'        TxtReceiverName.Text = .CallReceiverName
'    Else
'        TxtReceiverName.Tag = 0
'        TxtReceiverName.Text = ""
'    End If
    If Not IsNull(.CallReceiverEmpNo) And .CallReceiverEmpNo <> 0 Then
        LCallReceiverName.Tag = .CallReceiverEmpNo
        LCallReceiverName.Caption = .CallReceiverName
    Else
        LCallReceiverName.Tag = ""
        LCallReceiverName.Caption = ""
    End If
    
End With
Ok = True
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Sub FillControlsFromSql(rs As Recordset)
On Error GoTo ErrorHandler
Set maintCallViewModelInfo_ = maintDataService_.GetMaintCallInfo(rs!CallNo)
FillControls
Exit Sub
ErrorHandler:
MsgBox Err.Description

End Sub


Sub MoveNavigator(ByVal i As Integer)
Dim RSTemp As New ADODB.Recordset
On Error GoTo ErrorHandler
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
ErrorHandler:
MsgBox "Error In Navigator"
End Sub

Sub FillArray(Index As Integer, arr() As MapClientField)

If Index = 1 Then
    Set arr(1) = New MapClientField
    Set arr(2) = New MapClientField
    Set arr(3) = New MapClientField
    Set arr(4) = New MapClientField
    
    arr(1).filedId = "AdhamNo"
    arr(1).FiledName = "—ﬁ„ «·“»Ê‰"
    arr(1).IsVisible = False
    arr(1).IsListColumn = True
    arr(1).IsListField = False
    arr(1).Order = 1
    arr(1).IsSearchable = False
    
    
    arr(2).filedId = "AdhamName"
    arr(2).FiledName = "≈”„ «·“»Ê‰"
    arr(2).IsVisible = True
    arr(2).IsListColumn = False
    arr(2).IsListField = True
    arr(2).Order = 2
    arr(2).IsNumeric = False
    arr(2).IsSearchable = False
    
    
    arr(3).filedId = "adhamphon"
    arr(3).FiledName = "—ﬁ„ «·Â« ›"
    arr(3).IsVisible = True
    arr(3).IsListColumn = False
    arr(3).IsListField = False
    arr(3).Order = 3
    arr(3).IsNumeric = True
   
   
    arr(4).filedId = "MobilePhone"
    arr(4).FiledName = "—ﬁ„ «·„Ê»«Ì·"
    arr(4).IsVisible = True
    arr(4).IsListColumn = False
    arr(4).IsListField = False
    arr(4).Order = 4
    arr(4).IsNumeric = True
    
ElseIf Index = 2 Then
    Set arr(1) = New MapClientField
    Set arr(2) = New MapClientField
    
    arr(1).filedId = "ZoneNo"
    arr(1).FiledName = "—ﬁ„ «·„‰ÿﬁÂ"
    arr(1).IsVisible = False
    arr(1).IsListColumn = True
    arr(1).IsListField = False
    arr(1).Order = 1
    
    arr(2).filedId = "ZoneName"
    arr(2).FiledName = "≈”„ «·„‰ÿﬁÂ"
    arr(2).IsVisible = True
    arr(2).IsListColumn = False
    arr(2).IsListField = True
    arr(2).Order = 2
ElseIf Index = 3 Then
    Set arr(1) = New MapClientField
    Set arr(2) = New MapClientField
    
    arr(1).filedId = "EmpNo"
    arr(1).FiledName = "—ﬁ„ «·„ÊŸ›"
    arr(1).IsVisible = False
    arr(1).IsListColumn = True
    arr(1).IsListField = False
    arr(1).Order = 1
    
    arr(2).filedId = "FullName"
    arr(2).FiledName = "≈”„ «·„ÊŸ›"
    arr(2).IsVisible = True
    arr(2).IsListColumn = False
    arr(2).IsListField = True
    arr(2).Order = 2
End If
End Sub

Sub FillArrays()
    FillArray 1, FieldsArrayToSearch
    FillArray 1, Cols
    
    FillArray 2, FieldsZoneArrayToSearch
    FillArray 2, ZoneCols
    
'    FillArray 3, FieldReceiverArrayToSearch
'    FillArray 3, ReceiverCols
End Sub

Sub InitNavigator()
    Set RsNavigator = maintDataService_.GetAllMaintCall
End Sub
Sub fillClientInfo()
On Error GoTo ErrorHandler
With maintCallViewModelInfo_
    Ok = False
    TxtSearch.Tag = .clientInfo.adhamNo
    TxtSearch.Text = .clientInfo.AdhamName & ""
    
    TxtCustomerName.Tag = .clientInfo.adhamNo
    TxtCustomerName.Text = .clientInfo.AdhamName & ""
    LCLientVisitorCount.Caption = .ClientVisitorCount
    LClaimsIsRepeat.Visible = .ClaimIsRepeat
    TxtMobilePhone.Text = .clientInfo.MobilePhone & ""
    TxtHomePhone.Text = .clientInfo.AdhamPhon & ""
    TxtZoneName.Tag = .clientInfo.Zone
    TxtZoneName.Text = .clientInfo.ZoneName
    TxtAddress.Text = .clientInfo.AdhamAdress & ""
    TxtClientNote.Text = .clientInfo.Notes & ""
    TxtClientDefindname.Text = .clientInfo.Defindname & ""
    Ok = True
End With

Exit Sub
ErrorHandler:
MsgBox Err.Description

End Sub
Sub CreateNewClaim(ClientNo As Long)
On Error GoTo ErrorHandler
    searchClientIsAllow = False
    Set maintCallViewModelInfo_ = New MaintCallViewModel
    maintCallState = NewRecord
    EnableCmds False, False, False, True, True, True, False, False
    EnableControls True
    ClearControls
    Set maintCallViewModelInfo_.clientInfo = maintDataService_.GetClientInfoById(ClientNo)
    fillClientInfo
    searchClientIsAllow = True
    

Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Sub init()

    
    top = 0
    left = 0
    Pos = 0
    Ok = False
    maintCallState = DefaultRecord
    FillCombos
    FillArrays
    InitNavigator

    If LoadForm Then
        LoadForm = False
        MoveToRec idCallNo, True
        idCallOrder = 0
    Else
        MoveNavigator 4    'Move Last
    End If
    EnableControls False
    
    If ClientNo <> 0 Then
        CreateNewClaim (ClientNo)
        ClientNo = 0
    End If
    
    
    Grid.Rows = 1
    Grid.SelectionMode = flexSelectionListBox
    Ok = True
End Sub


Private Sub CmdPrint_Click()
On Error GoTo ErrorHandler
Dim currentState As EnumMaintCallState
currentState = maintCallViewModelInfo_.callState

If currentState = Repared Then
    MsgBox "«·‘ﬂÊÏ „‰€–Â ·«Ì„ﬂ‰ «·ÿ»«⁄Â", vbExclamation, "·«Ì„ﬂ‰ «·ÿ»«⁄Â"
Exit Sub
End If
maintDataService_.UpdateMaintCallState maintCallViewModelInfo_.CallNo, EnumMaintCallState.UnderwayAndPrinted
ChangeStatus EnumMaintCallState.UnderwayAndPrinted
If maintDataService_.ExecProcedure(maintCallViewModelInfo_.CallNo) Then
    With Cr1
        .Connect = ConnectName("")
        .ReportFileName = App.Path & "\Reports\RepMaintCall.rpt"
        .SQLQuery = "    Select "
        .SQLQuery = .SQLQuery & "    CallNo,"
        .SQLQuery = .SQLQuery & "    adhamname,"
        .SQLQuery = .SQLQuery & "    adhamphon,"
        .SQLQuery = .SQLQuery & "    defindname,"
        .SQLQuery = .SQLQuery & "    CallDescription,"
        .SQLQuery = .SQLQuery & "    notes,"
        .SQLQuery = .SQLQuery & "    adhamadress,"
        .SQLQuery = .SQLQuery & "    ProdFamName,"
        .SQLQuery = .SQLQuery & "    CallDateTime,"
        .SQLQuery = .SQLQuery & "    Region,"
        .SQLQuery = .SQLQuery & "    City,"
        .SQLQuery = .SQLQuery & "    Part,"
        .SQLQuery = .SQLQuery & "    ReceiverName ,"
        .SQLQuery = .SQLQuery & "    CountRec"
        .SQLQuery = .SQLQuery & "    From"
        .SQLQuery = .SQLQuery & "    tCustomerInfo Where CallNo <> 0"
        .SQLQuery = .SQLQuery & " and CallNo in (" & maintCallViewModelInfo_.CallNo & ")"
        .SQLQuery = .SQLQuery & " Order By CallNo"
        .DiscardSavedData = True
        .Destination = crptToPrinter
        .Action = 1

        CmdAdd.SetFocus
    End With
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
maintDataService_.UpdateMaintCallState maintCallViewModelInfo_.CallNo, currentState
ChangeStatus currentState

End Sub

Private Sub CmdSave_Click()
Dim result As maintCallResult
Dim choice As Boolean

If MsgBox("Â· «‰  „ «ﬂœ „‰ Õ›Ÿ «·‘ﬂÊÏ", vbYesNo + vbDefaultButton2, "Õ›Ÿ «·‘ﬂÊÏ") = vbYes Then
    choice = True
Else
    choice = False
End If

Set result = SaveChanges(choice)
    If result.MaintCallResultStatus Then
        EnableCmds True, True, True, False, False, False, True, True
        EnableControls False
       If maintCallState = NewRecord And choice Then
            InitNavigator
            MoveNavigator 4
       End If
        maintCallViewModelInfo_.callState = DefaultRecord
        If choice Then
            cmdPrint.SetFocus
            MsgBox " „ Õ›Ÿ «·‘ﬂÊÏ »‰Ã«Õ", vbInformation, "Õ›Ÿ «·‘ﬂÊÏ"
        Else
                CmdAdd.SetFocus
        End If
    Else
        MsgBox result.MaintCallResultDescription, vbExclamation + vbMsgBoxRight, "Œÿ√ ›Ì «· Œ“Ì‰ √Ê «·≈÷«›‹Â"
        Dim ctrl As Control
        Set ctrl = Me.GetTheLastFocusControl
        ctrl.SetFocus
        SendKeys "{home}+{end}"
    End If
End Sub

Function GetTheLastFocusControl() As Control
    Set GetTheLastFocusControl = Me.ActiveControl
End Function

Function SaveChanges(choice As Boolean) As maintCallResult
    On Error GoTo ErrorHandler
    Dim result As maintCallResult
    Set result = maintDataService_.SaveMaintCallChanges(maintCallViewModelInfo_, maintCallState, choice)
    Set SaveChanges = result
    Exit Function
ErrorHandler:
    Set SaveChanges = result
End Function

Sub ChangeCallState()
If maintCallState <> NewRecord Then
    maintCallState = UpdateRecord
End If
End Sub

Function SearchRec() As Double
On Error GoTo ErrorHandler
Dim i As Double
i = InputBox("√œŒ· —ﬁ„ «·‘ﬂÊÏ", "«·»ÕÀ ⁄‰ «·‘ﬂÊÏ")
If Val(i) <> 0 Then
    SearchRec = i
Else
    SearchRec = -1
End If
Exit Function
ErrorHandler:
SearchRec = -1
End Function

Private Sub cmdSaveClient_Click()
Dim result As maintCallResult
Set result = SaveClientChanges

If result.MaintCallResultStatus Then
    maintCallViewModelInfo_.clientInfo.ClientState = Default
    MsgBox " „ Õ›Ÿ „⁄·Ê„«  «·“»Ê‰", vbInformation, "„⁄·Ê„«  «·“»Ê‰"
    cmdSave.SetFocus
Else
    MsgBox result.MaintCallResultDescription, vbExclamation + vbMsgBoxRight, "Œÿ√ ›Ì „⁄·Ê„«  «·“»Ê‰"
End If
End Sub

Function SaveClientChanges() As maintCallResult
    On Error GoTo ErrorHandler
    Dim result As maintCallResult
    Set result = maintDataService_.SaveClient(maintCallViewModelInfo_)
    Set SaveClientChanges = result
    Exit Function
ErrorHandler:
    Set SaveClientChanges = result
End Function
Private Sub CmdSearch_Click()
MoveToRec SearchRec, False
'FrmSeatchBills.Show
End Sub



Private Sub ComboProductFamily_Change()
ChangeCallState
End Sub

Private Sub ComboProductFamily_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler
If KeyAscii = 13 Then
    TxtIssueName.SetFocus
    SendKeys "{home}+{end}"
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub ComboProductFamily_LostFocus()
If ComboProductFamily.MatchedWithList Then
    maintCallViewModelInfo_.ModNo = ComboProductFamily.BoundText
Else
     maintCallViewModelInfo_.ModNo = Null
End If
End Sub

Private Sub Form_Load()
    init
End Sub

Sub MoveToRec(CallNo As Double, IsOrderId As Boolean)
With RsNavigator
If RsNavigator.RecordCount = 0 Then
    ClearControls
    RecNum = 0
    LNavigator.Caption = ""
    Exit Sub
End If
.MoveLast
RecNum = 1
    Do While Not .BOF
        If IsOrderId Then
             If !CallNo <> CallNo Then
                 .MovePrevious
                 RecNum = RecNum + 1
             Else
                 FillControlsFromSql RsNavigator
                 LNavigator.Caption = LTrim(RTrim(Str(RecNum))) & "/" & LTrim(RTrim(Str(RsNavigator.RecordCount)))
                 Exit Sub
            End If
        Else
            If IsNull(!CallNo) Or !CallNo <> CallNo Then
                 .MovePrevious
                 RecNum = RecNum + 1
             Else
                 FillControlsFromSql RsNavigator
                 LNavigator.Caption = LTrim(RTrim(Str(RecNum))) & "/" & LTrim(RTrim(Str(RsNavigator.RecordCount)))
                 Exit Sub
            End If
        End If
    Loop
End With
End Sub


Sub ChangeStatus(state As EnumMaintCallState)
    maintCallViewModelInfo_.callState = state
    LMaintCallStateName.Caption = maintDataService_.GetMaintCallState(state)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button And vbRightButton Then
            If Gettag(empNo, 42) Then
                PopupMenu mnu
            End If
            PopupMenu mnuRefresh
    End If
End Sub

Private Sub Form_Unload(cancel As Integer)
'MsgBox "close the form"
End Sub

Private Sub Grid_RowColChange()
On Error GoTo ErrorHandler
If Flag Then
    Ok = False
    With Grid
       Select Case Pos
        Case 1
            TxtSearch.Tag = .TextMatrix(.Row, 1)
            TxtSearch.Text = .TextMatrix(.Row, 2)
        Case 2
            TxtZoneName.Tag = .TextMatrix(.Row, 1)
            TxtZoneName.Text = .TextMatrix(.Row, 2)
        Case 3
            TxtReceiverName.Tag = .TextMatrix(.Row, 1)
            TxtReceiverName.Text = .TextMatrix(.Row, 2)
       End Select
    End With
    Ok = True
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Sub PrintRep()
On Error GoTo ErrorHandler


Dim sqlText As String
sqlText = "SELECT  adhamname,RepPrice,Bdate,Edate,Notes,CallNo"
sqlText = sqlText & " ,RegestDate,ZoneName,TeamName "
sqlText = sqlText & " From AdhamClientRepView "
sqlText = sqlText & " where CliNo= " & maintCallViewModelInfo_.cliNo
sqlText = sqlText & " order by repDate asc "
With Cr1
    .Connect = ConnectName("")
    .SQLQuery = sqlText
    .ReportFileName = App.Path & "\Reports\RepClientReparation.rpt"
    .DiscardSavedData = True
    .WindowState = crptMaximized
    .Destination = crptToWindow
    .Action = 1
End With

Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub mnuPrint_Click()
PrintRep
End Sub

Private Sub mnuRefreshMaintCalls_Click()
InitNavigator
MoveNavigator 4
End Sub

'Private Sub timerDelayForClients_Timer()
'isOkayToSearch = True
'
'End Sub

Private Sub TxtAddress_Change()
ChangeClientState
End Sub

Private Sub TxtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtClientNote.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub TxtAddress_LostFocus()
maintCallViewModelInfo_.clientInfo.AdhamAdress = TxtAddress.Text
End Sub

Private Sub TxtClientDefindname_Change()
ChangeClientState
End Sub

Private Sub TxtClientDefindname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdSaveClient.SetFocus
    'SendKeys "{home}+{end}"
End If
End Sub

Private Sub TxtClientDefindname_LostFocus()
maintCallViewModelInfo_.clientInfo.Defindname = TxtClientDefindname.Text
End Sub

Private Sub TxtClientNote_Change()
ChangeClientState
End Sub

Private Sub TxtClientNote_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtClientDefindname.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub TxtClientNote_LostFocus()
maintCallViewModelInfo_.clientInfo.Notes = TxtClientNote.Text
End Sub

Private Sub TxtCustomerName_Change()
ChangeClientState
'    If searchClientIsAllow Then
'    With FrmSearchClients
'
'        .txtClientName.Text = TxtCustomerName.Text
'
'
'        .Show 1
'        searchClientIsAllow = False
'
'
'        TxtCustomerName.Text = customerName
'        TxtCustomerName.Tag = customerNumber
'        SendKeys "{home}+{end}"
'        searchClientIsAllow = True
'
'
'    End With
'    End If
End Sub

Private Sub TxtDefindname_Change()
ChangeCallState
End Sub

Private Sub TxtDefindname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdSaveClient.SetFocus
End If
End Sub

Private Sub TxtDefindname_LostFocus()
maintCallViewModelInfo_.Defindname = TxtDefindname.Text
End Sub



Function GetDataMemebers(listField() As MapClientField) As String

Dim fields As String
fields = ""
For i = 1 To UBound(listField)
    fields = fields & "," & listField(i).filedId
Next
GetDataMemebers = Mid(fields, 2)
End Function

Sub Search(sender As Control, Pos As Integer, tableName As String, listField() As MapClientField, dataMember As String, Optional isChangeCursor = True, Optional top As Variant, Optional left As Variant)
On Error GoTo ErrorHandler
Dim RsSearch As New ADODB.Recordset
Dim sqlWhere As String

If Ok Then
    sender.Tag = ""
'If sender.Text = "" Or (prevDataIsNull And prevSendrLength <= Len(sender.Text)) Then
If sender.Text = "" Then
   ' If sender.Text = "" Then
        Grid.Visible = False
        Exit Sub
    End If
    Flag = False

    sqlText = "Select top 10 " & GetDataMemebers(listField) & " From " & tableName
    sqlWhere = ""
    For i = 1 To UBound(listField)
        If listField(i).IsSearchable Then
            If IsNumeric(sender) Then
                If listField(i).IsNumeric Then
                
'                    If i <> UBound(listField) And sqlWhere = "" Then
'                        sqlWhere = " Or "
'                    End If
                    sqlWhere = sqlWhere & " Or ltrim(rtrim(" & listField(i).filedId & ")) Like '" & LTrim(RTrim(sender.Text)) & "%'"
                End If
            Else
                If Not listField(i).IsNumeric Then
'                    If i <> UBound(listField) And sqlWhere = "" Then
'                        sqlWhere = " Or "
'                    End If
                    sqlWhere = sqlWhere & " Or " & listField(i).filedId & " Like " & LikeExpression(sender.Text)

                End If
            End If

        End If

    Next
    If sqlWhere <> "" Then
        sqlWhere = Mid(sqlWhere, 4, Len(sqlWhere))
        sqlText = sqlText & " Where " & sqlWhere
    Else
        sender.Tag = 0
        Grid.Visible = False
        Exit Sub
    End If
    Set RsSearch = de.con.Execute(sqlText)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        'Grid.Row = 0
        FillFormating Pos, Grid
        If isChangeCursor Then ChangeCursor sender, top, left
        Grid.Visible = True
        'prevDataIsNull = False
    Else
        sender.Tag = 0
        Grid.Visible = False
'        prevDataIsNull = True
'        prevSendrLength = Len(sender.Text)
    End If
    Flag = True
End If


   Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Sub FillFormating(ByVal i As Integer, FlexGrid As VSFlexGrid)
If i = 1 Then
    fs = ""
    For i = 1 To UBound(Cols)
        fs = fs + "|>" + Cols(i).FiledName
    Next

    With FlexGrid
        .FormatString = fs
        .Cols = UBound(Cols) + 1
        For i = 1 To UBound(Cols)
            If Cols(i).IsVisible Then
                SetColWidths Cols(i).Order, FlexGrid
            Else
                .ColWidth(i) = 0
            End If
        Next
    End With
ElseIf i = 2 Then
 For i = 1 To UBound(ZoneCols)
        fs = fs + "|>" + ZoneCols(i).FiledName
    Next

    With FlexGrid
        .FormatString = fs
        .Cols = UBound(ZoneCols) + 1
        For i = 1 To UBound(ZoneCols)
            If ZoneCols(i).IsVisible Then
                SetColWidths ZoneCols(i).Order, FlexGrid
            Else
                .ColWidth(i) = 0
            End If
        Next
    End With
ElseIf i = 3 Then
 For i = 1 To UBound(ReceiverCols)
        fs = fs + "|>" + ReceiverCols(i).FiledName
    Next

    With FlexGrid
        .FormatString = fs
        .Cols = UBound(ReceiverCols) + 1
        For i = 1 To UBound(ReceiverCols)
            If ReceiverCols(i).IsVisible Then
                SetColWidths ReceiverCols(i).Order, FlexGrid
            Else
                .ColWidth(i) = 0
            End If
        Next
    End With
End If
End Sub

Private Sub TxtCustomerName_GotFocus()
ChangeToArabic
Pos = 1
End Sub

Private Sub TxtCustomerName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        TxtMobilePhone.SetFocus
        SendKeys "{Home}+{End}"
End If
End Sub

Sub FillCustomerInfo(clientInfo As AdhamViewModel)
On Error GoTo errorhandlelr
    Ok = False
    With clientInfo
        TxtCustomerName.Text = .AdhamName & ""
        TxtMobilePhone.Text = .MobilePhone & ""
        TxtHomePhone.Text = .AdhamPhon & ""
        TxtAddress.Text = .AdhamAdress & ""
        TxtZoneName.Tag = IIf(IsNull(.Zone), "", .Zone)
        TxtZoneName.Text = IIf(IsNull(.ZoneName), "", .ZoneName)
        TxtClientNote.Text = IIf(IsNull(.Notes), "", .Notes)
        TxtClientDefindname.Text = IIf(IsNull(.Defindname), "", .Defindname)
    End With
    Ok = True
Exit Sub
errorhandlelr:
MsgBox Err.Description
End Sub

Private Sub TxtCustomerName_LostFocus()
    maintCallViewModelInfo_.clientInfo.adhamNo = TxtSearch.Tag
    maintCallViewModelInfo_.clientInfo.AdhamName = TxtCustomerName.Text
End Sub

Sub ClearClientInfo()
Ok = False
        TxtCustomerName.Text = ""
        TxtMobilePhone.Text = ""
        TxtHomePhone.Text = ""
        TxtAddress.Text = ""
        TxtZoneName.Tag = ""
        TxtZoneName.Text = ""
        TxtClientNote.Text = ""
        TxtClientDefindname.Text = ""
Ok = True
End Sub

Sub ChangeClientState()
If maintCallViewModelInfo_.clientInfo.ClientState <> NewClient Then
    maintCallViewModelInfo_.clientInfo.ClientState = UpdateClient
End If
End Sub

Private Sub TxtHomePhone_Change()
ChangeClientState
End Sub

Private Sub TxtHomePhone_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtCustomerName.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub TxtHomePhone_LostFocus()
maintCallViewModelInfo_.clientInfo.AdhamPhon = LTrim(RTrim(TxtHomePhone.Text))
End Sub

Private Sub TxtIssueName_Change()
ChangeCallState
End Sub

Private Sub TxtIssueName_GotFocus()
ChangeToArabic
End Sub

Private Sub TxtIssueName_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler
If KeyAscii = 13 Then
    TxtHomePhone.SetFocus
    SendKeys "{home}+{end}"
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub TxtIssueName_LostFocus()
 maintCallViewModelInfo_.CallDEscription = TxtIssueName.Text
End Sub


Private Sub TxtMobilePhone_Change()
ChangeClientState
End Sub

Private Sub TxtMobilePhone_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtZoneName.SetFocus
    SendKeys "{Home}+{End}"
End If
End Sub

Private Sub TxtMobilePhone_LostFocus()
    maintCallViewModelInfo_.clientInfo.MobilePhone = TxtMobilePhone.Text
End Sub

'Private Sub TxtNotes_Change()
'ChangeCallState
'End Sub
'
'Private Sub txtNotes_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    TxtDefindname.SetFocus
'    SendKeys "{Home}+{End}"
'End If
'End Sub
'
'Private Sub txtNotes_LostFocus()
'maintCallViewModelInfo_.Notes = TxtNotes.Text
'End Sub

Private Sub TxtSearch_Change()
    Search TxtSearch, 1, "AdhamView7", FieldsArrayToSearch, "AdhamNo", True, 0, 0
    
End Sub

Private Sub TxtSearch_GotFocus()
TxtSearch.SelStart = 0
TxtSearch.SelLength = Len(TxtSearch.Text)
SendKeys "{home}+{end}"
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Grid.Visible Then
            Ok = False
            TxtSearch.Tag = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColNo), Grid.TextMatrix(Grid.Row, ColNo))
            TxtSearch.Text = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColName), Grid.TextMatrix(Grid.Row, ColName))
            'LClaimsIsRepeat.Visible = maintDataService_.IsClaimIsRepeatForTheClient(TxtCustomerName.Tag, Now)
            Ok = True
            Grid.Visible = False
'        Else
'            LClaimsIsRepeat.Visible = False
        End If
        ComboProductFamily.SetFocus

End If
End Sub

Private Sub TxtSearch_LostFocus()
Dim lastVisit As Variant
Dim LastCallNo As Variant
    If TxtSearch.Tag <> "" And Val(TxtSearch.Tag) <> 0 Then
        Dim claimsCountForTheClient As Boolean
        If maintCallViewModelInfo_.cliNo <> Val(TxtSearch.Tag) Then
            ChangeClientState
        End If
        maintCallViewModelInfo_.cliNo = Val(TxtSearch.Tag)
        Dim clientInfo As New AdhamViewModel
        Set clientInfo = maintDataService_.GetClientInfoById(maintCallViewModelInfo_.cliNo)
        Set maintCallViewModelInfo_.clientInfo = clientInfo
        FillCustomerInfo clientInfo
        claimsCountForTheClient = maintDataService_.claimsCountForTheClient(maintCallViewModelInfo_.cliNo, maintCallViewModelInfo_.CallDateTime, maintCallState)
        If claimsCountForTheClient Then
            LRepeatClaimTeamName.Caption = maintDataService_.GetTeamNameForRepeationClaim(maintCallViewModelInfo_.cliNo, maintCallViewModelInfo_.CallDateTime)
            LRepeatClaimTeamName.Visible = True
        Else
            LRepeatClaimTeamName.Caption = ""
            LRepeatClaimTeamName.Visible = False
        End If
        claimTeamForClient = maintDataService_.claimsCountForTheClient(maintCallViewModelInfo_.cliNo, maintCallViewModelInfo_.CallDateTime, maintCallState)
        LClaimsIsRepeat.Visible = claimsCountForTheClient ' IIf(claimsCountForTheClient = 1, True, False)
        
        LCLientVisitorCount.Caption = maintDataService_.GetClientVisitorCount(maintCallViewModelInfo_.cliNo, maintCallViewModelInfo_.CallDateTime)
        lastVisit = maintDataService_.GetLastVisitForTheClient(maintCallViewModelInfo_.cliNo, maintCallViewModelInfo_.CallDateTime)
        LastDataVisit.Caption = IIf(IsNull(lastVisit), "", lastVisit)
        LastCallNo = maintDataService_.GetLastCallNoForTheClient(maintCallViewModelInfo_.cliNo, maintCallViewModelInfo_.CallDateTime)
        LastCallNbr.Caption = IIf(IsNull(LastCallNo), "", LastCallNo)
    Else
        maintCallViewModelInfo_.cliNo = Null
        maintCallViewModelInfo_.clientInfo.adhamNo = Null
        maintCallViewModelInfo_.clientInfo.ClientState = NewClient
'        maintCallViewModelInfo_.clientInfo.AdhamName = ""
'        maintCallViewModelInfo_.clientInfo.AdhamAdress = ""
'        maintCallViewModelInfo_.clientInfo.AdhamPhon = ""
'        maintCallViewModelInfo_.clientInfo.AdNo = Null
'        maintCallViewModelInfo_.clientInfo.Defindname = ""
'        maintCallViewModelInfo_.clientInfo.MobilePhone = ""
'        maintCallViewModelInfo_.clientInfo.Notes = ""
'        maintCallViewModelInfo_.clientInfo.Zone = Null
'        maintCallViewModelInfo_.clientInfo.ZoneName = ""
        'ClearClientInfo
        LClaimsIsRepeat.Visible = False
        LCLientVisitorCount.Caption = ""
        LastDataVisit.Caption = ""
        LastCallNbr.Caption = ""
    End If
    ChangeCallState
    prevDataIsNull = False
End Sub

'Private Sub TxtReceiverName_Change()
'Search TxtReceiverName, 3, "EmpFullName", FieldReceiverArrayToSearch, "EmpNo", True, 0, 0
'ChangeCallState
'End Sub
'
'Private Sub TxtReceiverName_GotFocus()
'ChangeToArabic
'Pos = 3
'End Sub
'
'Private Sub TxtReceiverName_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
'End Sub

'Private Sub TxtReceiverName_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'        If Grid.Visible Then
'            OK = False
'            TxtReceiverName.Tag = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, 1), Grid.TextMatrix(Grid.Row, 1))
'            TxtReceiverName.Text = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, 2), Grid.TextMatrix(Grid.Row, 2))
'            OK = True
'            Grid.Visible = False
'        End If
'        cmdSave.SetFocus
'End If
'End Sub

'Private Sub TxtReceiverName_LostFocus()
'If TxtReceiverName.Tag <> "" Or Val(TxtReceiverName.Tag) <> 0 Then
'    maintCallViewModelInfo_.CallReceiverEmpNo = TxtReceiverName.Tag
'Else
'    maintCallViewModelInfo_.CallReceiverEmpNo = Null
'End If
'  prevDataIsNull = False
'End Sub

Private Sub TxtZoneName_Change()
Search TxtZoneName, 2, "CoZone", FieldsZoneArrayToSearch, "ZoneNo", True, ClientFrame.top, ClientFrame.left
ChangeClientState
End Sub

Private Sub TxtZoneName_GotFocus()
ChangeToArabic


Pos = 2
End Sub

Private Sub TxtZoneName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
End Sub

Private Sub TxtZoneName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Grid.Visible Then
            Ok = False
            TxtZoneName.Tag = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, 1), Grid.TextMatrix(Grid.Row, 1))
            TxtZoneName.Text = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, 2), Grid.TextMatrix(Grid.Row, 2))
            Ok = True
            Grid.Visible = False
        End If
        TxtAddress.SetFocus
        SendKeys "{Home}+{End}"
End If
End Sub

Private Sub TxtZoneName_LostFocus()
If TxtZoneName.Tag <> "" And Val(TxtZoneName.Tag) <> 0 Then
    maintCallViewModelInfo_.clientInfo.Zone = TxtZoneName.Tag
Else
    maintCallViewModelInfo_.clientInfo.Zone = Null
End If
  prevDataIsNull = False
End Sub
