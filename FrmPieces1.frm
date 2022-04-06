VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmPieces1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "√—’œ… «·√—ﬁ«„ «·„Œ“‰Ì…"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   8625
   Begin Crystal.CrystalReport cr1 
      Left            =   3195
      Top             =   4335
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileName   =   "C:\Amin.xls"
      PrintFileType   =   19
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   420
      Left            =   30
      TabIndex        =   23
      Top             =   2910
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   741
      _Version        =   131074
      Begin VB.TextBox TxtSearch 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   5610
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   60
         Width           =   2400
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "»ÕÀ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   8220
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   45
         Width           =   330
      End
   End
   Begin Threed.SSFrame SSFrame5 
      Height          =   390
      Left            =   30
      TabIndex        =   11
      Top             =   6780
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   688
      _Version        =   131074
      Begin VB.Label LNum 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   5100
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   90
         Width           =   2115
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "⁄œœ «·„Ê«œ «·„Œ“‰Ì…"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   7260
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   90
         Width           =   1260
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2220
      Left            =   0
      TabIndex        =   5
      Top             =   690
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   3916
      _Version        =   131074
      Begin VB.CheckBox ChkStore 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7590
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   30
         Value           =   1  'Checked
         Width           =   225
      End
      Begin Threed.SSFrame SSFrame6 
         Height          =   615
         Left            =   30
         TabIndex        =   26
         Top             =   1560
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   1085
         _Version        =   131074
         Begin Threed.SSCheck Chk1 
            Height          =   225
            Left            =   2640
            TabIndex        =   32
            Tag             =   "7"
            Top             =   330
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   397
            _Version        =   131074
            ForeColor       =   8388608
            Caption         =   "«·√—ﬁ«„ «·„”«ÊÌ… ··’›—"
            Alignment       =   1
         End
         Begin Threed.SSCheck Chk 
            Height          =   225
            Left            =   2640
            TabIndex        =   27
            Tag             =   "6"
            Top             =   60
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   397
            _Version        =   131074
            ForeColor       =   8388608
            Caption         =   "«·√—ﬁ«„ «·„Œ“‰Ì… «·„ÊÃ»…"
            Alignment       =   1
         End
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   915
         Left            =   30
         TabIndex        =   19
         Top             =   630
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   1614
         _Version        =   131074
         Begin Threed.SSOption Option 
            Height          =   240
            Index           =   0
            Left            =   2565
            TabIndex        =   22
            Tag             =   "3"
            Top             =   90
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   423
            _Version        =   131074
            MarqueeDirection=   1
            ForeColor       =   8388608
            Caption         =   "ﬂ«›…  «·√—ﬁ«„ «·„Œ“‰Ì…"
            Alignment       =   1
            Value           =   -1
         End
         Begin Threed.SSOption Option 
            Height          =   330
            Index           =   1
            Left            =   1980
            TabIndex        =   21
            Tag             =   "4"
            Top             =   270
            Width           =   2580
            _ExtentX        =   4551
            _ExtentY        =   582
            _Version        =   131074
            MarqueeDirection=   1
            ForeColor       =   8388608
            Caption         =   "«·√—ﬁ«„ «·„Œ“‰Ì… «· Ì ·Ì” ·Â« ‘—Õ"
            Alignment       =   1
         End
         Begin Threed.SSOption Option 
            Height          =   330
            Index           =   2
            Left            =   2295
            TabIndex        =   20
            Tag             =   "5"
            Top             =   540
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   582
            _Version        =   131074
            MarqueeDirection=   1
            ForeColor       =   8388608
            Caption         =   "«·√—ﬁ«„ «·„Œ“‰Ì… «· Ì ·Â« ‘—Õ"
            Alignment       =   1
         End
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   585
         Left            =   30
         TabIndex        =   16
         Top             =   45
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   1032
         _Version        =   131074
         Begin Threed.SSOption StockOption 
            Height          =   240
            Index           =   0
            Left            =   2220
            TabIndex        =   18
            Tag             =   "1"
            Top             =   45
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   423
            _Version        =   131074
            ForeColor       =   8388736
            Caption         =   "«· Ã„Ì⁄ Õ”» «·—ﬁ„ «·„Œ“‰Ì"
            Alignment       =   1
            Value           =   -1
         End
         Begin Threed.SSOption StockOption 
            Height          =   240
            Index           =   1
            Left            =   2625
            TabIndex        =   17
            Tag             =   "2"
            Top             =   315
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   423
            _Version        =   131074
            ForeColor       =   8388736
            Caption         =   "«· Ã„Ì⁄ Õ”» «·„” Êœ⁄"
            Alignment       =   1
         End
      End
      Begin VB.TextBox TxtStrNo 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   5745
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   315
         Width           =   2805
      End
      Begin VB.TextBox TxtDescription 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   4695
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1815
         Width           =   3345
      End
      Begin VB.TextBox TxtItemTo 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   4695
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1065
         Width           =   1545
      End
      Begin VB.ComboBox ComboItemTo 
         Height          =   315
         ItemData        =   "FrmPieces1.frx":0000
         Left            =   7215
         List            =   "FrmPieces1.frx":0016
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1065
         Width           =   825
      End
      Begin VB.ComboBox ComboItemFrom 
         Height          =   315
         ItemData        =   "FrmPieces1.frx":002F
         Left            =   7215
         List            =   "FrmPieces1.frx":0045
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   705
         Width           =   825
      End
      Begin VB.TextBox TxtItemFrom 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   4695
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   705
         Width           =   1545
      End
      Begin Threed.SSCommand CmdSearch 
         Height          =   345
         Left            =   4710
         TabIndex        =   29
         Top             =   330
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         Picture         =   "FrmPieces1.frx":005E
      End
      Begin Threed.SSCheck chkChoose 
         Height          =   315
         Left            =   6900
         TabIndex        =   28
         Top             =   1440
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   131074
         ForeColor       =   8388608
         Caption         =   " ÕœÌœ √—ﬁ«„ „Œ“‰Ì…"
         Alignment       =   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„” Êœ⁄"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   7890
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   90
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "≈·Ï"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   7
         Left            =   8220
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1110
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   6
         Left            =   8310
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   750
         Width           =   210
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·‘—Õ"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   8145
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·„«œ…"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   6270
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1110
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·„«œ…"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   6270
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   750
         Width           =   675
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7890
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
            Picture         =   "FrmPieces1.frx":0172
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":2848
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":56A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":7E4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":A746
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":CC6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":F423
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":11E37
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":14789
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":174FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":19D4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":1CBF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":1F94D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":222ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":2524F
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":27C77
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":2A631
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":2CF62
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":2F868
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":322D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":35281
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":37BAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":3A4DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":3CC1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":3F58E
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":41E16
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":44025
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":46982
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":491AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":4BC60
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":4E69C
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":511B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":5414C
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":56F18
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":59B92
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":5CA24
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":5F662
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPieces1.frx":62211
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   1217
      ButtonWidth     =   1191
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   37
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   15
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   34
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid MSHFlexGrid1 
      Height          =   3435
      Left            =   30
      TabIndex        =   30
      Top             =   3330
      Width           =   8595
      _cx             =   15161
      _cy             =   6059
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
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
End
Attribute VB_Name = "FrmPieces1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim OshowIndex As Integer
Const ColPieceStockNo = 1
Const ColPieceStockName = 2
Const ColStrNo = 3
Const ColFnlQnt = 4



Const ColStkno_1 = 1
Const ColStkname_1 = 2
Const ColFnlqnt_1 = 3
Const Col_1 = 4
Const Col_2 = 5
Const Col_3 = 6
Const Col_4 = 7
Const Col_5 = 8
Const Col_6 = 9
Const Col_7 = 10
Const Col_8 = 11
Const Col_9 = 12
Const Col_10 = 13
Const Col_11 = 14
Const Col_12 = 15
Const Col_13 = 16
Const Col_14 = 17
Const Col_15 = 18
Const Col_16 = 19
Const Col_17 = 20
Const Col_18 = 21

Dim CellRow As Integer
Dim CellCol As Integer, Rs As New ADODB.Recordset, Optionindex As Integer, StockOption1 As Integer
Function Fillitems() As String
Dim Str As String
Str = ""
With FrmChooseItems.FGrid
    For i = 0 To .Rows - 1
        Str = Str + ",'" + .TextMatrix(i, 1) & "'"
    Next
End With
Fillitems = Mid(Str, 2)
End Function
Sub ClearGrid()
    Sqltext = "Select PieceStockNo, PieceName , StrNo , Fnlqnt From TempPieces Where  PiecestockNo ='-1' Order By PiecestockNo"
    Set Rs = Dstock.Con.Execute(Sqltext)
    Set MSHFlexGrid1.DataSource = Rs
    FillFormatString (OshowIndex)
End Sub
Function StrText(StrNo As String)
tt = ""
    For i = 1 To Len(StrNo)
        If Right(Left(StrNo, i), 1) = "-" Then
            tt = tt + ","
        Else
            tt = tt & Right(Left(StrNo, i), 1)
        End If
   Next
StrText = tt
End Function

Function CountRecord(WhereStr As String) As Recordset
    Sqltext = "Select Count(*) CountRec From TempPieces Where PieceStockNo <> '-1' "
'    Set Rs = Dstock.Con.Execute(sqltext)
    Set CountRecord = Dstock.Con.Execute(Sqltext)

End Function
Sub FillFormatString(Index As Integer)
Select Case Index
    Case 0
        Fs = "|>" + "—ﬁ„ «·„«œ…"
        Fs = Fs + "|>" + "≈”„ «·„«œ…"
        Fs = Fs + "|>" + "—ﬁ„ «·„” Êœ⁄"
        Fs = Fs + "|>" + "«·—’Ìœ «·‰Â«∆Ì"
        With MSHFlexGrid1
            .FormatString = Fs
            SetColWidths ColPieceStockNo
            SetColWidths ColPieceStockName
            If StockOption1 = 0 Then
                .ColWidth(ColStrNo) = 0
            Else
                SetColWidths ColStrNo
            End If
            SetColWidths ColFnlQnt
        
        End With
      Case 1
        Fs = "|>" + "—ﬁ„ «·„«œ…"
        Fs = Fs + "|>" + "≈”„ «·„«œ…"
        Fs = Fs + "|>" + "«·—’Ìœ «·‰Â«∆Ì"
        Fs = Fs + "|>" + "1"
        Fs = Fs + "|>" + "2"
        Fs = Fs + "|>" + "3"
        Fs = Fs + "|>" + "4"
        Fs = Fs + "|>" + "5"
        Fs = Fs + "|>" + "6"
        Fs = Fs + "|>" + "7"
        Fs = Fs + "|>" + "8"
        Fs = Fs + "|>" + "9"
        Fs = Fs + "|>" + "10"
        Fs = Fs + "|>" + "11"
        Fs = Fs + "|>" + "12"
        Fs = Fs + "|>" + "13"
        Fs = Fs + "|>" + "14"
        Fs = Fs + "|>" + "15"
        Fs = Fs + "|>" + "16"
        Fs = Fs + "|>" + "17"
        Fs = Fs + "|>" + "18"
        With MSHFlexGrid1
            .FormatString = Fs
            SetColWidths ColStkno_1
            SetColWidths ColStkname_1
            SetColWidths ColFnlqnt_1
            SetColWidths Col_1
            SetColWidths Col_2
            SetColWidths Col_3
            SetColWidths Col_4
            SetColWidths Col_5
            SetColWidths Col_6
            SetColWidths Col_7
            SetColWidths Col_8
            SetColWidths Col_9
            SetColWidths Col_10
            SetColWidths Col_11
            SetColWidths Col_12
            SetColWidths Col_13
            SetColWidths Col_14
            SetColWidths Col_15
            SetColWidths Col_16
            SetColWidths Col_17
            SetColWidths Col_18
        End With
 End Select
End Sub



Sub SetColWidths(ColNo As Integer)
    Dim i, J, s, w
    With MSHFlexGrid1
            s = 0
            For i = 0 To .Rows - 1
                w = TextWidth(.TextMatrix(i, ColNo))
                If w > s Then s = w
            Next i
            .ColWidth(ColNo) = s + 100
    End With
End Sub


Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub ExportToExcel()
Dim objXL As Excel.Application
Dim objWB As Excel.Workbook
Dim objWS As Excel.Worksheet
Dim r As Long
Dim c As Long
Set objXL = New Excel.Application
Set objWB = objXL.Workbooks.Add
Set objWS = objWB.Worksheets(1)

With objWS

For r = 0 To MSHFlexGrid1.Rows - 1
For c = 0 To MSHFlexGrid1.Cols - 1
.Cells(r + 1, c + 1) = MSHFlexGrid1.TextMatrix(r, c)
Next
Next
.Cells.Columns.AutoFit
End With
objXL.Visible = True
Set objWS = Nothing
Set objWB = Nothing
Set objXL = Nothing

End Sub
Sub AddRestitems()
Dim RsRest As New ADODB.Recordset

Sqltext = "Select t1.PieceStockNo , t1.PieceName , t1.StrNo , t1.FnlQnt From TempPieces t1 Left outer Join CommonPiecesQry c1  on t1.PieceStockNo = c1.PieceStockNo Where c1.PieceStockNo is null"
Set RsRest = Dstock.Con.Execute(Sqltext)
If RsRest.RecordCount > 0 Then
    With MSHFlexGrid1
    Do While Not RsRest.EOF
        .AddItem ""
        .TextMatrix(.Rows - 1, ColPieceStockNo) = RsRest!PieceStockNo & ""
        .TextMatrix(.Rows - 1, ColPieceStockName) = RsRest!PieceName & ""
        .TextMatrix(.Rows - 1, ColStrNo) = RsRest!StrNo & ""
        .TextMatrix(.Rows - 1, ColFnlQnt) = RsRest!FnlQnt & ""
        For i = 1 To .Cols - 1
            .Row = .Rows - 1
            .Col = i
            .CellBackColor = vbWhite
        Next
        RsRest.MoveNext
    Loop
    End With
End If
End Sub
Private Sub SerachData()
Dim SqlWhere As String, SqlShow As String
Screen.MousePointer = vbHourglass
Sqltext = "Truncate Table TempPieces"
Dstock.Con.Execute (Sqltext)
SqlWhere = ""


If LTrim(RTrim(TxtStrNo.Text)) <> "" Or ChkStore.Value Then
    If ChkStore.Value Then
        SqlWhere = " And Strid in (" & Strids & ")"
    Else
        SqlWhere = " And Convert(int,StrNo) in (" & StrText(TxtStrNo.Text) & ")"
    End If
    
Else
    SqlWhere = " And Convert(int,StrNo) in (0)"
End If

If LTrim(RTrim(TxtItemFrom.Text)) <> "" Then
    SqlWhere = SqlWhere & " And ltrim(rtrim(Stkno)) " & ComboItemFrom.Text & " '" & TxtItemFrom.Text & "'"
End If

If LTrim(RTrim(TxtItemTo.Text)) <> "" Then
    SqlWhere = SqlWhere & " And ltrim(rtrim(Stkno))  " & ComboItemTo.Text & " '" & TxtItemTo.Text & "'"
End If

Select Case StockOption1
    Case 0
        Sqltext = "insert into TempPieces(PiecestockNo , FnlQnt) Select Stkno  , Sum(FnlQnt) From Stkinf Where ltrim(rtrim(Stkno))<>'-1' " & SqlWhere & " Group By StkNo"
    Case 1
        Sqltext = "insert into TempPieces(PiecestockNo , StrNo , FnlQnt) Select Stkno , StrNo , FnlQnt From Stkinf Where ltrim(rtrim(Stkno)) <>'-1'" & SqlWhere
End Select
Dstock.Con.Execute (Sqltext)
Sqltext = "Update TempPieces Set PieceName = StkName From CoStock s Where  TempPieces.PieceStockNo= s.StkNo"
Dstock.Con.Execute (Sqltext)
    
    If LTrim(RTrim(TxtDescription.Text)) <> "" Then
        Sqltext = "Delete From TempPieces  Where PieceName Not like '%" & TxtDescription.Text & "%'"
        Dstock.Con.Execute (Sqltext)
    End If
    
    Select Case Optionindex
        Case 1
            Sqltext = "Delete From TempPieces  Where PieceName <> ''"
        Case 2
            Sqltext = "Delete From TempPieces  Where PieceName = ''"
    End Select
    Dstock.Con.Execute (Sqltext)
    If chkChoose.Value Then
        sTRItems = Fillitems
        If LTrim(RTrim(sTRItems)) <> "" Then
            Sqltext = "Delete From TempPieces Where PiecestockNo not in(" & sTRItems & ")"
        End If
    End If
    Dstock.Con.Execute (Sqltext)
    If Chk.Value Then
        Sqltext = "Delete From TempPieces Where FnlQnt =0 "
    End If
    Dstock.Con.Execute (Sqltext)
    If Chk1.Value Then
        Sqltext = "Delete From TempPieces Where FnlQnt >0 "
    End If
    Dstock.Con.Execute (Sqltext)
    ClearGrid
   
    Sqltext = "Select * From CommonPiecesQry Order by PiecestockNo , Strno"
    Set Rs = Dstock.Con.Execute(Sqltext)
    If Dstock.Con.State <> adStateOpen Then
    Dstock.Con.Open
    End If
    
    
    If Rs.State <> adStateOpen Then Rs.Open "Select * From CommonPiecesQry Order by PiecestockNo , Strno", Dstock.Con
    
    Set MSHFlexGrid1.DataSource = Rs
    MSHFlexGrid1.BackColor = &HFFFFC0
    AddRestitems
    FillFormatString OshowIndex
    MSHFlexGrid1.Col = ColPieceStockNo
    MSHFlexGrid1.Sort = flexSortGenericAscending
    LNum.Caption = MSHFlexGrid1.Rows - 1

    If MSHFlexGrid1.Rows = 1 Then
        Screen.MousePointer = vbDefault
        MsgBox "·«ÌÊÃœ „⁄·Ê„«  ÕÊ· Â–Â «·„Ê«œ", vbExclamation, " ‰»ÌÂ"
        Exit Sub
    End If
    Screen.MousePointer = vbDefault

End Sub

 
Private Sub PrintData()
With cr1
    .Connect = ConnectName
   ' .PrintFileName = "c:\nn.xls"
   ' .PrintFileType = crptExcel50
    .ReportFileName = App.Path + "\Reports\Items.rpt"
    .SQLQuery = "Select PieceStockNo, PieceName, StrNo, FnlQnt From dbo.TempPieces Order By PieceStockNo"
    .DiscardSavedData = True
    .WindowState = crptMaximized
    .Action = 1
End With
End Sub

Private Sub CmdSearch_Click()
FrmChoose.Show 1
TxtStrNo = StrNo
If TxtStrNo.Text <> "" Then
    ChkStore.Value = 0
Else
    ChkStore.Value = 1
End If
End Sub

Private Sub ComboItemFrom_Change()
    ComboItemTo.Text = ComboItemFrom.Text
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tAB}"
        SendKeys "{nUMLOCK}", True
        SendKeys "{hOME}+{eND}", True
    End If
End Sub
Private Sub Form_Load()
'    OShow(0).Value = True
'    OshowIndex = 0
    Dim Rs As New ADODB.Recordset
    If Dstock.Con.State <> adStateOpen Then Dstock.Con.Open
'    LNum.Caption = Format(CountRecord("")!CountRec, "###,###,###,###")
    ClearGrid
    ComboItemFrom.ListIndex = 0
    ComboItemTo.ListIndex = 0
    Optionindex = 0
    Top = 0
    Left = 0
End Sub

Private Sub Option_Click(Index As Integer, Value As Integer)
Optionindex = Index
End Sub


Private Sub chkChoose_Click(Value As Integer)
If chkChoose.Value Then
    FrmChooseItems.Show 1
End If
End Sub

Private Sub StockOption_Click(Index As Integer, Value As Integer)
StockOption1 = Index
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        PrintData
    Case 3
        SerachData
    Case 5
       ExportToExcel
    Case 7
       Unload Me
End Select
End Sub


Private Sub TxtItemFrom_Change()
    TxtItemTo.Text = TxtItemFrom.Text
End Sub

Private Sub TxtSearch_Change()

    Sqltext = "Select PieceStockNo, PieceName, StrNo, FnlQnt  From TempPieces Where PieceStockNo like '" & TxtSearch.Text & "%'"
    Sqltext = Sqltext & " Or StrNo like '%" & TxtSearch.Text & "%'"
    Sqltext = Sqltext & " Or PieceName like '%" & TxtSearch.Text & "%'"
    Sqltext = Sqltext & " Or FnlQnt like '%" & TxtSearch.Text & "%'"
    Set Rs = Dstock.Con.Execute(Sqltext)
    Set MSHFlexGrid1.DataSource = Rs
    FillFormatString (OshowIndex)

End Sub
