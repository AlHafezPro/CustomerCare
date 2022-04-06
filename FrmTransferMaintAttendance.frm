VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTransferMaintAttendance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " —ÕÌ· œÊ«„ Œœ„Â «·’Ì«‰Â «·Õ«—ÃÌÂ"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1725
   ScaleWidth      =   5475
   Begin MSMask.MaskEdBox TxtFromDate 
      Height          =   375
      Left            =   3540
      TabIndex        =   0
      Top             =   750
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtTillDate 
      Height          =   375
      Left            =   3540
      TabIndex        =   1
      Top             =   1290
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   30
      Top             =   870
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
            Picture         =   "FrmTransferMaintAttendance.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferMaintAttendance.frx":6209F
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
      Width           =   5475
      _ExtentX        =   9657
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
            Object.ToolTipText     =   " —ÕÌ· œÊ«„ Œœ„Â «·’Ì«‰Â «·Õ«—ÃÌÂ"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Õ–› œÊ«„ Œœ„Â «·’Ì«‰Â «·Õ«—ÃÌÂ"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰  «—ÌŒ"
      Height          =   195
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   780
      Width           =   600
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "≈·Ï  «—ÌŒ"
      Height          =   195
      Left            =   4770
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1380
      Width           =   675
   End
End
Attribute VB_Name = "FrmTransferMaintAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function TransferMaintAttendance(FromDate As String, TillDate As String) As Boolean
On Error GoTo errorhandler
Dim FDate As String, TDate As String
Dim cmd As New ADODB.Command
Screen.MousePointer = vbHourglass
If IsDate(FromDate) Then
    FDate = ConvertControlDate(FromDate)
End If
If IsDate(TxtTillDate.Text) Then
    TDate = ConvertControlDate(TillDate)
End If
'Sqltext = "Exec CalcAtdEmpDay '" & FDate & "', '" & TDate & "', 0, 0 WITH RECOMPILE"

sqltext = "Exec sp_Transfer_Maint_Atend '" & FDate & "', '" & TDate & "'"
cmd.CommandText = sqltext
cmd.ActiveConnection = de.con
cmd.CommandTimeout = 0
cmd.Execute

Screen.MousePointer = vbDefault
TransferMaintAttendance = True
Exit Function
errorhandler:
Screen.MousePointer = vbDefault
TransferMaintAttendance = False
MsgBox Err.Description

End Function
Function RemoveTransferAttendance(FromDate As String, TillDate As String)
On Error GoTo errorhandler
Dim FDate As String, TDate As String
Dim cmd As New ADODB.Command
Screen.MousePointer = vbHourglass
If IsDate(FromDate) Then
    FDate = ConvertControlDate(FromDate)
End If
If IsDate(TxtTillDate.Text) Then
    TDate = ConvertControlDate(TillDate)
End If

sqltext = "Exec sp_Delete_Transfer_Maint_Atend '" & FDate & "', '" & TDate & "'"
cmd.CommandText = sqltext
cmd.ActiveConnection = de.con
cmd.CommandTimeout = 0
cmd.Execute
Screen.MousePointer = vbDefault
RemoveTransferAttendance = True
Exit Function
errorhandler:
Screen.MousePointer = vbDefault
RemoveTransferAttendance = False
MsgBox Err.Description
End Function
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        If MsgBox("Â·  —Ìœ ‰—ÕÌ· œÊ«„ «·’»«‰Â «·Œ«—ÃÌÂø", vbYesNo + vbDefaultButton2) = vbYes Then
            If TransferMaintAttendance(txtFromDate.Text, TxtTillDate.Text) Then
                MsgBox " „  ⁄„·Ì… «·Õ”«» »‰Ã«Õ", vbInformation, " —ÕÌ· œÊ«„ «·’Ì«‰Â «·Œ«—ÃÌÂ"
            End If
        End If
    Case 2
       If MsgBox("Â·  —Ìœ Õ–› «·œÊ«„ «·Õ«’ »«·’Ì«‰Â «·Œ«—ÃÌÂ", vbYesNo + vbDefaultButton2) = vbYes Then
        If RemoveTransferAttendance(txtFromDate.Text, TxtTillDate.Text) Then
            MsgBox " „  ⁄„·Ì…«·Õ–›  »‰Ã«Õ", vbInformation, "Õ–› œÊ«„ «·’Ì«‰Â «·Œ«—ÃÌÂ"
         End If
       End If
    Case 4
        Unload Me
End Select
End Sub
Sub Init()
top = 0
left = 0
txtFromDate.Text = "21/" & Right("00" + Trim(Str(IIf(Month(Date) = 1, 12, Month(Date) - 1))), 2) & "/" & IIf(Month(Date) = 1, Trim(Str(Year(Date) - 1)), Trim(Str(Year(Date))))
TxtTillDate.Text = "20/" & Right("00" + Trim(Str(Month(Date))), 2) & "/" & Trim(Str(Year(Date)))
End Sub
Private Sub Form_Load()
Init
End Sub
