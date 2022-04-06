VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDailyTeamRep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ﬁ—Ì— √⁄„«· «·’Ì«‰… «·Œ«—ÃÌ… «·ÌÊ„Ì…"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   11685
   Begin Threed.SSFrame SSFrame1 
      Height          =   645
      Left            =   60
      TabIndex        =   9
      Top             =   7380
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   1138
      _Version        =   131074
      Begin VB.Label Lcount 
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
         Left            =   9240
         TabIndex        =   11
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "⁄œœ «·«’·«Õ« "
         Height          =   195
         Left            =   10500
         TabIndex        =   10
         Top             =   180
         Width           =   960
      End
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   1560
      Top             =   810
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   6075
      Left            =   30
      TabIndex        =   2
      Top             =   1230
      Width           =   11625
      _cx             =   20505
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
   Begin VB.TextBox txtteamNo 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   7320
      TabIndex        =   1
      Top             =   780
      Width           =   855
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3210
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
            Picture         =   "frmDailyTeamRep.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyTeamRep.frx":6209F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
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
            ImageIndex      =   37
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "1"
                  Text            =   " ﬁ—Ì— ··»‰ﬂ"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "2"
                  Text            =   " ﬁ—Ì— „’«—Ì› «· ÊÿÌ‰"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   4530
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   0
         Width           =   6825
      End
   End
   Begin MSMask.MaskEdBox mskDate 
      Height          =   405
      Left            =   9300
      TabIndex        =   0
      Top             =   780
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   714
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Left            =   10980
      TabIndex        =   8
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label2 
      Caption         =   "—ﬁ„ «·Ê—‘… "
      Height          =   255
      Left            =   8280
      TabIndex        =   7
      Top             =   840
      Width           =   945
   End
   Begin VB.Label Label3 
      Caption         =   "«”„ «·ÊÕœ…"
      Height          =   255
      Left            =   6270
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   870
      Width           =   855
   End
   Begin VB.Label lblTeamName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Height          =   300
      Left            =   6030
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   810
      Width           =   90
   End
   Begin VB.Menu mnu 
      Caption         =   "file"
      Visible         =   0   'False
      Begin VB.Menu select 
         Caption         =   " ÕœÌœ «·„Õœœ"
      End
      Begin VB.Menu cancel 
         Caption         =   "≈·€«¡ ««·„Õœœ"
      End
   End
End
Attribute VB_Name = "frmDailyTeamRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const ColPicked = 1
Const ColCallNo = 2
Const coladhamName = 3
Const colTeamNo = 4
Const colRepDate = 5
Const colRepPrice = 6
Const colNotes = 7
Const colDescription = 8
Sub FillFormating(ByVal i As Integer, FlexGrid As VSFlexGrid)
If i = 1 Then
    fs = "|>" + "chk"
    fs = fs + "|>" + "—ﬁ„ «·‘ﬂÊÏ"
    fs = fs + "|>" + "≈”„ «·“»Ê‰"
    fs = fs + "|>" + "—ﬁ„ «·Ê—‘Â"
    fs = fs + "|>" + " «—ÌŒ «·«’·«Õ"
    fs = fs + "|>" + "ﬁÌ„Â «·«’·«Õ"
    fs = fs + "|>" + "„·«ÃŸ« "
    fs = fs + "|>" + "«·‘—Õ"
    With FlexGrid
        .FormatString = fs
        .Cols = 9
        .ColWidth(ColPicked) = 400
        .ColDataType(ColPicked) = flexDTBoolean
        SetColWidths ColCallNo, FlexGrid
        SetColWidths coladhamName, FlexGrid
        SetColWidths colTeamNo, FlexGrid
        SetColWidths colRepDate, FlexGrid
        SetColWidths colRepPrice, FlexGrid
        SetColWidths colNotes, FlexGrid
        SetColWidths colDescription, FlexGrid
    End With
End If
End Sub

Sub SetColWidths(ByVal ColNo As Integer, FlexGrid As VSFlexGrid)
    With FlexGrid
        .AutoSize (ColNo)
    End With
End Sub





Private Sub chkPickAll_Click()
If chkPickAll.Value = 1 Then
    For i = 1 To vsGrid.Rows - 1
        vsGrid.TextMatrix(i, 1) = 1
    Next i
Else
    For i = 1 To vsGrid.Rows - 1
        vsGrid.TextMatrix(i, 1) = 0
    Next i
End If
End Sub

Private Sub chkPickAll_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Function GetCallNos() As String
On Error GoTo ErrorHandler
 Dim callNos As String
 callNos = ""
    With vsGrid
        For i = 1 To .Rows - 1
            If .TextMatrix(i, ColPicked) Then
                callNos = callNos & "," & .TextMatrix(i, ColCallNo)
            End If
        Next
    End With
    
callNos = Mid(callNos, 2)
GetCallNos = callNos
Exit Function
ErrorHandler:
callNos = ""
MsgBox Err.Description
End Function

Private Sub PrintRep()
On Error GoTo ErrorHandler
    Dim SQL As String, callNos As String
    If txtteamNo.Text & "" = "" Or Not IsNumeric(txtteamNo) Then
        MsgBox "ÌÃ» ≈œŒ«· —ﬁ„ «·Ê—‘… Õ’—«", vbOKOnly + vbCritical + vbMsgBoxRight, " ‰»ÌÂ"
        txtteamNo.SetFocus
    Else
                callNos = GetCallNos
                SQL = " SELECT distinct "
                SQL = SQL & " MaintSaeedDalyView.CallNo, MaintSaeedDalyView.calltime, MaintSaeedDalyView.teamno , "
                SQL = SQL & " MaintSaeedDalyView.RepDate, MaintSaeedDalyView.RepTimeBegin, MaintSaeedDalyView.RepTimeEnd, "
                SQL = SQL & " MaintSaeedDalyView.PiecesPrice, MaintSaeedDalyView.WorkPrice, MaintSaeedDalyView.MethodName, "
                SQL = SQL & " MaintSaeedDalyView.symbol, MaintSaeedDalyView.ProdFamName, MaintSaeedDalyView.RecieverName, "
                SQL = SQL & " MaintSaeedDalyView.adhamname, MaintSaeedDalyView.adhamphon, MaintSaeedDalyView.adhamadress, "
                SQL = SQL & " MaintSaeedDalyView.LastRepDate, MaintSaeedDalyView.Repetition, MaintSaeedDalyView.Notes, "
                SQL = SQL & " MaintSaeedDalyView.description, MaintSaeedDalyView.VoltBefor, MaintSaeedDalyView.VoltAfter, "
                SQL = SQL & " MaintSaeedDalyView.MaxEndHour, MaintSaeedDalyView.TeamName, MaintSaeedDalyView.ProdPurchaseDate , "
                SQL = SQL & " MaintSaeedDalyView.CallReceiver, MaintSaeedDalyView.CallDescription, MaintSaeedDalyView.InFactory, "
                SQL = SQL & " MaintSaeedDalyView.OutFactory, MaintSaeedDalyView.vehimmNo, MaintSaeedDalyView.VehOwner, "
                SQL = SQL & " MaintSaeedDalyView.LastDayVisit "
                SQL = SQL & "  From "
                SQL = SQL & " MaintSaeedDalyView where MaintSaeedDalyView.RepDate ='" & TransDateToSql(mskDate) & "' and MaintSaeedDalyView.TeamNo = " & txtteamNo.Text
                SQL = SQL & " and CallNo in (" & IIf(callNos = "", -1, callNos) & ")"

            With cr1
                  .ReportFileName = App.Path & "\Reports\SaeedRepFax.rpt"
                  .Connect = ConnectName("")
                  .DiscardSavedData = True
                  .WindowState = crptMaximized
                  .SQLQuery = SQL
                  .Action = 1
            End With
    End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub cancel_Click()
UpdateRow vsGrid, False
End Sub

Private Sub Form_Load()
mskDate = Date - 1
top = 0
left = 0
vsGrid.Rows = 1
FillFormating 1, vsGrid
vsGrid.Editable = flexEDKbdMouse
End Sub

Private Sub mskDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"


End Sub

Sub UpdateRow(FlexGrid As VSFlexGrid, isChk As Boolean)
On Error GoTo ErrorHandler
With FlexGrid
        If .Row >= .RowSel Then
            FirstRow = .Row
            LastRow = .RowSel
        Else
            FirstRow = .RowSel
            LastRow = .Row
        End If
        For i = FirstRow To LastRow Step -1
                .TextMatrix(i, ColPicked) = isChk
        Next i
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub select_Click()
UpdateRow vsGrid, True
End Sub

Private Sub txtteamNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub txtteamNo_LostFocus()
Dim SQL As String
Dim rs As New ADODB.Recordset
If Not (txtteamNo.Text & "" = "") And IsNumeric(txtteamNo.Text) Then
    SQL = "select *  from Maintteam where TeamNo = " & txtteamNo
    'If DeMaint.CnnMaint.state = adStateClosed Then DeMaint.CnnMaint.Open , user, GetPass
    'If rs.state <> adStateClosed Then rs.Close
    rs.Open SQL, de.con, adOpenForwardOnly, adLockReadOnly
    If rs.EOF And rs.BOF Then
        MsgBox "—ﬁ„ «·ÊÕœ… €Ì— „ÊÃÊœ", vbCritical + vbOKOnly + vbMsgBoxRight, " ‰»ÌÂ"
        With txtteamNo
            .SelStart = 0
            .SelLength = Len(txtteamNo.Text)
            .SetFocus
        End With
    Else
        lblTeamName.Caption = rs!TeamName
        'SQL = "select Picked,CallNO,TeamNo,RepDate,RepPrice,Notes,Description  from Reparation where TeamNO = " & txtTeamNO.Text & "and Repdate ='" & TransDateToSql(mskDate.Text) & "'"
        SQL = "Select 1 as Picked , R.CallNO ,A.adhamName , R.TeamNo, R.RepDate, R.RepPrice, R.Notes,"
        SQL = SQL & " R.Description from Reparation R left outer join MaintCall M On R.CallNo = M.CallNo "
        SQL = SQL & " left outer join adhamview7 A "
        SQL = SQL & " On m.CliNo = A.AdhamNo "
        SQL = SQL & " Where R.TeamNO = " & txtteamNo.Text & "and R.Repdate ='" & TransDateToSql(mskDate.Text) & "'"
        'If rs.state <> adStateClosed Then rs.Close
        Set rs = de.con.Execute(SQL)
        'rs.Open SQL, de.con, adOpenForwardOnly, adLockReadOnly
        Set vsGrid.DataSource = rs
        FillFormating 1, vsGrid
        LCount.Caption = rs.RecordCount
    End If
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
        Case 1
            PrintRep
        Case 3
            Unload Me
End Select
End Sub

Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
If Col <> ColPicked Then cancel = True
End Sub

Private Sub vsGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button And vbRightButton Then
            PopupMenu mnu
    End If
End Sub
