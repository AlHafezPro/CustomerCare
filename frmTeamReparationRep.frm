VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTeamReparationRep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ﬁ—Ì— √⁄„«· «·’Ì«‰Â «·Œ«—ÃÌÂ ·ÊÕœÂ"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2760
   ScaleWidth      =   4680
   Begin MSDataListLib.DataCombo ComboTeam 
      Height          =   360
      Left            =   1290
      TabIndex        =   2
      Top             =   1710
      Width           =   2175
      _ExtentX        =   3836
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
   Begin Crystal.CrystalReport cr1 
      Left            =   1590
      Top             =   690
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
            Picture         =   "frmTeamReparationRep.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamReparationRep.frx":6209F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox TxtFromDate 
      Height          =   345
      Left            =   2130
      TabIndex        =   0
      Top             =   810
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtToDate 
      Height          =   345
      Left            =   2130
      TabIndex        =   1
      Top             =   1260
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSDataListLib.DataCombo ComboPayment 
      Height          =   360
      Left            =   1290
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
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
            Object.ToolTipText     =   " ﬁ—Ì— √⁄„«· «·’Ì«‰Â «·Œ«—ÃÌÂ ··ÊÕœ« "
            ImageIndex      =   37
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "1"
                  Text            =   " ﬁ—Ì— «·«Ã„«·Ì ··ÊÕœ« "
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·ÊÕœÂ"
      Height          =   195
      Left            =   4020
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1710
      Width           =   450
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "‰Ê⁄ «·’Ì«‰…"
      Height          =   195
      Index           =   4
      Left            =   3690
      TabIndex        =   6
      Top             =   2190
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰  «—ÌŒ"
      Height          =   195
      Index           =   0
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "≈·Ï  «—ÌŒ"
      Height          =   195
      Index           =   1
      Left            =   3795
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1290
      Width           =   675
   End
End
Attribute VB_Name = "frmTeamReparationRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim externalReparationByTeamTypeRec As ExternalReparationByTeamType

Private Sub Form_Load()
init
End Sub


Sub init()
    top = 0
    left = 0
    FillCombos
End Sub

Sub FillCombos()
    Dim rsPayment As New ADODB.Recordset
    sqlText = "Select No , Name  From PayMethod "
    Set rsPayment = de.con.Execute(sqlText)
    Set ComboPayment.RowSource = rsPayment
    ComboPayment.listField = "Name"
    ComboPayment.BoundColumn = "No"
    ComboPayment.BoundText = 0
    
    Dim rsTeam As New ADODB.Recordset
    sqlText = "Select TeamNo , TeamName from maintteam "
    Set rsTeam = de.con.Execute(sqlText)
    Set ComboTeam.RowSource = rsTeam
    ComboTeam.listField = "TeamName"
    ComboTeam.BoundColumn = "TeamNo"
    ComboTeam.BoundText = 0
    
End Sub

Friend Sub GetExternalReparationByTeamType()


If Not IsDate(TxtFromDate.Text) Then
    externalReparationByTeamTypeRec.FromDate = "01/01/" & Right("0000" + LTrim(RTrim(Str(Year(Now)))), 4)
Else
    externalReparationByTeamTypeRec.FromDate = TxtFromDate.Text
End If

If Not IsDate(TxtToDate.Text) Then
    externalReparationByTeamTypeRec.TillDate = "31/12/" & Right("0000" + LTrim(RTrim(Str(Year(Now)))), 4)
Else
    externalReparationByTeamTypeRec.TillDate = TxtToDate.Text
    
End If

If Not ComboPayment.MatchedWithList Then
    externalReparationByTeamTypeRec.PaymenrTypeMethod = Null
Else
    externalReparationByTeamTypeRec.PaymenrTypeMethod = ComboPayment.BoundText
    
End If

If Not ComboTeam.MatchedWithList Then
 externalReparationByTeamTypeRec.TeamNo = Null
Else
 externalReparationByTeamTypeRec.TeamNo = ComboTeam.BoundText
End If



End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
            PrintRep 1
    Case 3
        Unload Me
End Select
End Sub

Function ExecuteProcedure() As Boolean
On Error GoTo errorhandler
    sqlText = "sp_summary_teams '"
    sqlText = sqlText & ConvertControlDate(externalReparationByTeamTypeRec.FromDate) & "','"
    sqlText = sqlText & ConvertControlDate(externalReparationByTeamTypeRec.TillDate) & "',"
    sqlText = sqlText & externalReparationByTeamTypeRec.PaymenrTypeMethod & ","
    sqlText = sqlText & IIf(IsNull(externalReparationByTeamTypeRec.TeamNo), "Null", externalReparationByTeamTypeRec.TeamNo)
        
        de.con.Execute (sqlText)
    ExecuteProcedure = True
Exit Function
errorhandler:
ExecuteProcedure = False
MsgBox Err.Description
End Function

Sub PrintRep(Voption As Integer)
On Error GoTo errorhandler
Screen.MousePointer = vbHourglass
GetExternalReparationByTeamType
With cr1
    .Connect = ConnectName("")
        .Formulas(0) = "FromDate=" & "'" & externalReparationByTeamTypeRec.FromDate & "'"
        .Formulas(1) = "TillDate=" & "'" & externalReparationByTeamTypeRec.TillDate & "'"
        
         
        Select Case Voption
        Case 1
                .SQLQuery = "select CallNo, clino, Qty, price, Total, Notes, RepPrice, TeamNo, Cash, repDate, adhamname, name, TeamName, PieceStockNo, PieceName, defindname from ReparationByTeam"
                .SQLQuery = .SQLQuery & " Where CallNo <>0"
                If Not IsNull(externalReparationByTeamTypeRec.FromDate) Then
                .SQLQuery = .SQLQuery & " and regestdate>='" & ConvertControlDate(externalReparationByTeamTypeRec.FromDate) & "'"
                End If
                
                If Not IsNull(externalReparationByTeamTypeRec.TillDate) Then
                .SQLQuery = .SQLQuery & " and regestdate<='" & ConvertControlDate(externalReparationByTeamTypeRec.TillDate) & "'"
                End If
                
                If Not IsNull(externalReparationByTeamTypeRec.PaymenrTypeMethod) Then
                .SQLQuery = .SQLQuery & " and Cash=" & externalReparationByTeamTypeRec.PaymenrTypeMethod
                End If
                
                If Not IsNull(externalReparationByTeamTypeRec.TeamNo) Then
                .SQLQuery = .SQLQuery & " and TeamNo=" & externalReparationByTeamTypeRec.TeamNo
                End If
                
                .ReportFileName = App.Path & "\Reports\RepReparationByTeam.rpt"
        Case 2
                If ExecuteProcedure Then
                    .SQLQuery = "Select teamno , teamname , total , repprice , fee ,"
                    .SQLQuery = .SQLQuery & " PaymentMethodNo, PaymentMethodName From t_summaryTeam Order By TeamName"
                    .ReportFileName = App.Path + "\Reports\SummaryTeamsRpt.rpt"
                End If
        End Select
        .DiscardSavedData = True
        .WindowState = crptMaximized
        .Action = 1
End With
Screen.MousePointer = vbDefault
Exit Sub
errorhandler:
MsgBox Err.Description
Screen.MousePointer = vbDefault
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Tag
    Case 1 'Team Summary
            PrintRep 2
    
    End Select
End Sub
