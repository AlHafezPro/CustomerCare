VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmDocument 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ﬂ‘› Õ”«»  «·„Ê«œ"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3345
   ScaleWidth      =   6435
   Begin VB.TextBox TxtRoot 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   3570
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1140
      Width           =   2055
   End
   Begin VB.TextBox TxtStrNo 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2490
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   750
      Width           =   3135
   End
   Begin VB.TextBox TxtType 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   3630
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   1995
   End
   Begin VB.TextBox TxtDescription 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2070
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1980
      Width           =   3525
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2595
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   4577
      _Version        =   131074
      ForeColor       =   192
      Caption         =   "ŒÌ«—«   «·»ÕÀ"
      Alignment       =   1
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   510
         Width           =   1935
         Begin Threed.SSOption Option2 
            Height          =   255
            Index           =   1
            Left            =   60
            TabIndex        =   23
            Top             =   420
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   131074
            ForeColor       =   8388608
            Caption         =   "«· Ì ·Â« Õ—ﬂ…"
            Alignment       =   1
         End
         Begin Threed.SSOption Option2 
            Height          =   255
            Index           =   2
            Left            =   60
            TabIndex        =   22
            Top             =   750
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   131074
            ForeColor       =   8388608
            Caption         =   "«· Ì ·Ì” ·Â« Õ—ﬂ…"
            Alignment       =   1
         End
         Begin Threed.SSOption Option2 
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   21
            Top             =   120
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   131074
            ForeColor       =   8388608
            Caption         =   "ﬂ«›… «·Õ—ﬂ« "
            Alignment       =   1
            Value           =   -1
         End
      End
      Begin VB.CheckBox ChkDetail 
         Alignment       =   1  'Right Justify
         Caption         =   "«· ›’Ì·Ì"
         Height          =   255
         Left            =   210
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   270
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin Threed.SSOption Option1 
         Height          =   315
         Index           =   2
         Left            =   30
         TabIndex        =   15
         Top             =   2160
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         _Version        =   131074
         Caption         =   "«·—’Ìœ Ì”«ÊÌ «·’›—"
         Alignment       =   1
      End
      Begin Threed.SSOption Option1 
         Height          =   315
         Index           =   1
         Left            =   30
         TabIndex        =   14
         Top             =   1860
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         _Version        =   131074
         Caption         =   "«·—’Ìœ „ÊÃ»"
         Alignment       =   1
      End
      Begin Threed.SSOption Option1 
         Height          =   315
         Index           =   0
         Left            =   30
         TabIndex        =   13
         Top             =   1590
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         _Version        =   131074
         Caption         =   "ﬂ· «·√—’œ…"
         Alignment       =   1
         Value           =   -1
      End
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   2400
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3090
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
            Picture         =   "FrmDocument.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocument.frx":6209F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6435
      _ExtentX        =   11351
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
            ImageIndex      =   37
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "1"
                  Text            =   "«·„Ê«œ «· Ì ·„   Õ—ﬂ"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   34
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox TxtFromDate 
      Height          =   345
      Left            =   4380
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtToDate 
      Height          =   345
      Left            =   4380
      TabIndex        =   5
      Top             =   2940
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·‘—Õ"
      Height          =   195
      Index           =   3
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2010
      Width           =   420
   End
   Begin Threed.SSCommand CmdType 
      Height          =   375
      Left            =   3210
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1560
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   661
      _Version        =   131074
      PictureFrames   =   1
      Picture         =   "FrmDocument.frx":64A4E
   End
   Begin Threed.SSCommand CmdSearch 
      Height          =   345
      Left            =   2100
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   750
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      _Version        =   131074
      PictureFrames   =   1
      Picture         =   "FrmDocument.frx":64B62
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "‰Ê⁄ «·Õ—ﬂ…"
      Height          =   195
      Index           =   2
      Left            =   5670
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1650
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·Ã–—"
      Height          =   195
      Index           =   2
      Left            =   6075
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1170
      Width           =   345
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·„” Êœ⁄"
      Height          =   195
      Index           =   0
      Left            =   5790
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   750
      Width           =   630
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "≈·Ï  «—ÌŒ"
      Height          =   195
      Index           =   1
      Left            =   5745
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3000
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰  «—ÌŒ"
      Height          =   195
      Index           =   0
      Left            =   5820
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2550
      Width           =   600
   End
End
Attribute VB_Name = "FrmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sub FillCombos()
'    Dim RSStr As New ADODB.Recordset
'    Sqltext = "Select Id , StrNo , StrName From NameStr where id in(" & Strids & ") Order By StrNo"
'    Set RSStr = de.con.Execute(Sqltext)
'    Set ComboStr.RowSource = RSStr
'    ComboStr.ListField = "StrName"
'    ComboStr.BoundColumn = "Id"
'    ComboStr.BoundText = RSStr!Id
'
'
'    Dim RsByanType As New ADODB.Recordset
'    Sqltext = "Select TypeId , TypeName  From CoByanType  where TypeId <> 0 Order BY TypeId "
'    Set RsByanType = de.con.Execute(Sqltext)
'    Set ComboByanType.RowSource = RsByanType
'    ComboByanType.ListField = "TypeName"
'    ComboByanType.BoundColumn = "TypeId"
'End Sub
Sub init()
    'FillCombos
    Top = 0
    Left = 0
    TxtFromDate.Text = "01/01/" & LTrim(RTrim(Str(Year(Date))))
    TxtToDate.Text = "31/12/" & LTrim(RTrim(Str(Year(Date))))
End Sub
Sub ExportToExcel()

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass
Dim Rs As New ADODB.Recordset
Dim Extension  As String
Dim cmd As New ADODB.Command

    'If de.ConExportsDBF.State <> adStateOpen Then de.ConExportsDBF.Open
'SQltext = "SET DEFAULT TO " & App.Path
'de.ConAttendanceDBF.Execute (SQltext)



Extension = "_" & Trim(Str(Year(Now))) + "_" + Trim(Str(Month(Now))) + "_" + Trim(Str(Day(Now))) & "_" & Trim(Str(Hour(Now))) & "_" & Trim(Str(Minute(Now)))
strFileName = App.Path & "\DocumentsByType_" & Extension & ".xlsx"

'DbfFileName = App.Path & "\Exports" & Extension
'Sqltext = "create table " & DbfFileName & "(CommitNo c(254), CommitDate c(254), CommitValue n(20,3), DemandNo c(254), ExporterNo i, ExporterName c(254), DealNo i, dealname c(254), CountryNo i, CountryName c(254), ProductNum i, CarNum i)"
'de.ConExportsDBF.Execute (Sqltext)


'Dim FromDate As String
'Sqltext = "Select CONVERT(varchar(10),DATEADD(dd,-(DAY(GETDATE())-1),GETDATE()),101)AS FromDate"
'
'Set Rs = de.con.Execute(Sqltext)
'FromDate = Rs!FromDate
'
'Dim TillDate As String
'Sqltext = "Select convert(varchar(10),DATEADD(s,-1,DATEADD(mm, DATEDIFF(m,0,GETDATE())+1,0)) ,101)  AS TillDate"
'Set Rs = de.con.Execute(Sqltext)
'TillDate = Rs!TillDate



Sqltext = "Exec GetDocument '" & ConvertControlDate(TxtFromDate.Text) & "','" & ConvertControlDate(TxtToDate.Text) & "','" & Replace(TxtStrNo.Text, "-", ",") & "','" & Replace(TxtType.Text, "-", ",") & "','" & TxtRoot.Text & "','" & Strids & "','" & TxtDescription.Text & "',1"
''Sqltext = "Exec hafez2000dev.dbo.sp_ExportDetails '01/01/2012','12/31/2012'"
cmd.ActiveConnection = de.con
cmd.CommandText = Sqltext
cmd.CommandTimeout = 0
cmd.Execute

Sqltext = "SELECT  [«·—ﬁ„ «·„Œ“‰Ì],[«·‘—Õ]    ,[—’Ìœ √Ê· «·„œ…]      ,[»Ì«‰ «·Õ—ﬂ…]      ,[«·Ã—œ]      ,[‰ﬁ· ]      ,[„‘ —Ì« ]      ,[‘Õ‰ Œ«—ÃÌ]      ,[«„— ’—› —∆Ì”Ì]      ,[‘Õ‰ »Ì—Ê ]      ,[«„— ’—› „— Ã⁄]      ,[ ’œÌ— »Ì—Ê ]      ,[ ’œÌ— Œ«—ÃÌ]      ,[’—› /≈‰ «Ã]      ,[’Ì«‰…]      ,[’Ì«‰…/ ’œÌ—]      ,[ »œÌ·]      ,[‰›«Ì« ]      ,[„ ›—ﬁ« ]      ,[Œœ„… „” Â·ﬂ]      ,[„— Ã⁄ „” Êœ⁄ «·’Ì«‰…]      ,[’Ì«‰… »—‰«„Ã]      ,[„»Ì⁄« ]      ,[„— Ã⁄« ]      ,[≈œŒ«· „‰ „⁄„·]      ,[„— Ã⁄ Ê—‘…]      ,[ „— Ã⁄ „‰ «·„” Êœ⁄]      ,[«·„Ã„Ê⁄]      ,[«·ÊÕœ…]      ,[«·—’Ìœ]      ,[«··“Ê„ «·„ Êﬁ⁄]  FROM [Stock2013].[dbo].[DocumentByTYpeQry]"
Set Rs = de.con.Execute(Sqltext)

'Sqltext = "Select AtdcardNo , Date , Hour , HostName , OperationDatetime RealDate , OperationDatetime RealTime from HafezGeneral.dbo.tt4 where Date >=convert(varchar(10),getdate(),101)"
Dim objXL As Excel.Application
Dim objWB As Excel.Workbook
Dim objWS As Excel.Worksheet
Dim r As Long
Dim c As Long

Set objXL = New Excel.Application
Set objWB = objXL.Workbooks.Add
Set objWS = objWB.Worksheets(1)

With objWS
For i = 0 To Rs.Fields.Count - 1
    .Cells(1, i + 1) = Rs.Fields(i).Name
Next
Rs.MoveFirst
For r = 0 To Rs.RecordCount - 1
For c = 0 To Rs.Fields.Count - 1
    .Cells(r + 2, c + 1) = Rs.Fields(c)
Next
    Rs.MoveNext
    If Rs.EOF Then Exit For
Next
End With
objXL.Visible = True
Set objWS = Nothing
Set objWB = Nothing
Set objXL = Nothing

Screen.MousePointer = vbDefault
Exit Sub
ErrorHandler:
MsgBox Err.Description
Screen.MousePointer = vbDefault

End Sub
Private Sub ChkDetail_Click()

If ChkDetail.Value Then
    ChkDetail.Caption = "«· ›’Ì·Ì"
Else
    ChkDetail.Caption = "«·≈Ã„«·Ì"
End If

End Sub

'Private Sub ChkBalance_Click()
'If ChkBalance.Value Then
'    ChkBalance.Caption = "«·—’Ìœ „ÊÃ»"
'Else
'    ChkBalance.Caption = "«·—’Ìœ Ì”«ÊÌ «·’›—"
'End If
'End Sub

'Private Sub Chk_Click()
'If Chk.Value Then
'    Chk.Caption = "«· Ì ·Â« Õ—ﬂ…"
'Else
'    Chk.Caption = "«· Ì ·Ì” ·Â« Õ—ﬂ…"
'End If
'End Sub

Private Sub CmdSearch_Click()
FrmChoose.Show 1
TxtStrNo = StrNo

End Sub

Private Sub CmdType_Click()
FrmChooseTypes.Show 1
TxtType.Text = ByanType
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    SendKeys "{home}+{End}"
End If
End Sub

Private Sub Form_Load()
init
End Sub
Sub PrintData(Optional TYpeId)
Dim cmd As New ADODB.Command

On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    Sqltext = "Exec GetDocument '" & ConvertControlDate(TxtFromDate.Text) & "','" & ConvertControlDate(TxtToDate.Text) & "','" & Replace(TxtStrNo.Text, "-", ",") & "','" & Replace(TxtType.Text, "-", ",") & "','" & TxtRoot.Text & "','" & Strids & "','" & TxtDescription.Text & "'"
    cmd.ActiveConnection = de.con
    cmd.CommandText = Sqltext
    cmd.CommandTimeout = 0
    cmd.Execute
'    de.con.Execute (SQltext)

    Screen.MousePointer = vbDefault
    
    With cr1
        cr1.Connect = ConnectName("")
        If IsMissing(TYpeId) Then
            If ChkDetail.Value Then
                .Formulas(0) = "FromDate='" & TxtFromDate.Text & "'"
                .Formulas(1) = "ToDate='" & TxtToDate.Text & "'"
                .Formulas(2) = "StrNBr='" & TxtStrNo.Text & "'"
                .ReportFileName = App.Path + "\reports\RepStockDocument.rpt"
            Else
                .Formulas(0) = ""
                .Formulas(1) = ""
                .Formulas(2) = ""
                .ReportFileName = App.Path + "\reports\RepGeneralDocument.rpt"
            End If
            
            If Option2(0).Value Then
                .SQLQuery = "Select  StkId, StkNo, StkName, FB, [In], [Out], InOut, FnlQnt From DocumentQry Where  StkId<>-1 "
            ElseIf Option2(1).Value Then
                .SQLQuery = "Select  StkId, StkNo, StkName, FB, [In], [Out], InOut, FnlQnt From DocumentQry Where ([In] <> 0 Or [Out] <> 0) "
            Else
                .SQLQuery = "Select  StkId, StkNo, StkName, FB, [In], [Out], InOut, FnlQnt From DocumentQry Where [In] = 0 and [Out] = 0 "
            End If
            
            If Option1(0).Value Then ' ﬂ· «·√—’œ…
            ElseIf Option1(1).Value Then '«·—’Ìœ «·„ÊÃ»
                .SQLQuery = .SQLQuery & " And   FnlQnt >0 "
            ElseIf Option1(2).Value Then '«·—’Ìœ «·„”«ÊÌ ’›—
                .SQLQuery = .SQLQuery & " And   FnlQnt =0 "
            End If
            .SQLQuery = .SQLQuery & "Order By ltrim(rtrim(StkNo)) ,  [In] , [Out]"
        Else
            .Formulas(0) = "FromDate='" & TxtFromDate.Text & "'"
            .Formulas(1) = "ToDate='" & TxtToDate.Text & "'"
            .Formulas(2) = "StrNBr='" & TxtStrNo.Text & "'"

            .ReportFileName = App.Path & "\Reports\RepNotWorking.rpt"
            .SQLQuery = "Select YEAR, STKNO, StkName, FB, In, OUT, fnlqnt from T_NOTWORKING order by fnlqnt desc"
            
        End If
        
        .WindowState = crptMaximized
        .Action = 1
    End With
Exit Sub
ErrorHandler:
Screen.MousePointer = vbDefault
MsgBox Err.Description
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        PrintData
    Case 3
        ExportToExcel
    Case 5
        Unload Me
End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Tag
    Case 1
        PrintData 1
End Select

End Sub

Private Sub TxtFromDate_Change()
TxtToDate.Text = TxtFromDate.Text
End Sub
