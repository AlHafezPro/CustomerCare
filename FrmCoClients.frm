VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{1C81E0B1-BFC2-42C4-A910-97E0FA9F83C9}#1.0#0"; "vsstr8.ocx"
Begin VB.Form FrmCoClients 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " —„Ì“ ««·“»«∆‰"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   16575
   Begin VB.CheckBox ChkWithReparation 
      Alignment       =   1  'Right Justify
      Caption         =   " «· Ì ·Â« «’·«Õ« "
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   30
      Width           =   2535
   End
   Begin Crystal.CrystalReport Cr1 
      Left            =   7530
      Top             =   3540
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   585
      Left            =   90
      TabIndex        =   5
      Top             =   7230
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   1032
      _Version        =   131074
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "⁄œœ «·“»«—« "
         Height          =   195
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   150
         Width           =   825
      End
      Begin VB.Label LCountVisited 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   120
         Width           =   1395
      End
      Begin VB.Label LCount 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   14010
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   90
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "⁄œœ «·“»«∆‰"
         Height          =   195
         Left            =   15600
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   120
         Width           =   750
      End
   End
   Begin VB.TextBox txtTop 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   2670
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "10"
      Top             =   30
      Width           =   945
   End
   Begin VB.TextBox txtClientName 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   4620
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   11355
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   6675
      Left            =   90
      TabIndex        =   1
      Top             =   540
      Width           =   16455
      _cx             =   29025
      _cy             =   11774
      Appearance      =   1
      BorderStyle     =   0
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
   Begin VSStr8LibCtl.VSFlexString VSFlexString1 
      Left            =   4080
      Top             =   690
      Text            =   ""
      Pattern         =   ""
      CaseSensitive   =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "⁄œœ «·”Ã·« "
      Height          =   195
      Left            =   3705
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   90
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·“»Ê‰"
      Height          =   195
      Left            =   16050
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   420
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint 
         Caption         =   " ﬁ—Ì— «·‘ﬂ«ÊÏ Ê«·«’·«Õ« "
      End
      Begin VB.Menu newClaim 
         Caption         =   "≈‰‘«¡ ‘ﬂÊÏ"
      End
   End
End
Attribute VB_Name = "FrmCoClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ok As Boolean, Flag As Boolean, Pos As Integer


Const ColClientNo = 1
Const ColClientName = 2
Const ColClientPhoneNBr = 3
Const ColClientMobilePhoneNbr = 4
Const ColClientAddress = 5

Const ColClientZoneNo = 6
Const ColClientZoneName = 7
Const ColClientNotes = 8
Const ColclientDefineName = 9
Const ColClientVisitedCount = 10



Dim fildsToSearch As FiledColumns

Const ColNo = 1
Const ColName = 2
Dim mainDataService As New MaintDataService

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

Sub FillFormating(ByVal i As Integer, FlexGrid As VSFlexGrid)
If i = 1 Then
   
    fs = "|>" + "«·—ﬁ„"
    fs = fs + "|>" + "«·≈”„"

    With FlexGrid
        .FormatString = fs
        .Cols = 3

        SetColWidths ColNo, FlexGrid
        SetColWidths ColName, FlexGrid

    End With
ElseIf i = 2 Then
    fs = "|>" + "—ﬁ„ «·“»Ê‰"
    fs = fs + "|>" + "≈”„ «·“»Ê‰"
    fs = fs + "|>" + "—ﬁ„ «·Â« ›"
    fs = fs + "|>" + "—ﬁ„ «·„Ê»«Ì·"
    fs = fs + "|>" + "⁄‰Ê«‰ «·“»Ê‰"
    fs = fs + "|>" + "ZoneNo"
    fs = fs + "|>" + "«·„‰ŸﬁÂ"
    fs = fs + "|>" + "„·«ÕŸ« "
    fs = fs + "|>" + "«·„⁄—›Â"
    fs = fs + "|>" + "⁄œœ „—«  «·“Ì«—Â"
    
    
  
    
    With FlexGrid
        .FormatString = fs
        .Cols = 11
        
        SetColWidths ColClientNo, FlexGrid
        SetColWidths ColClientName, FlexGrid
        SetColWidths ColClientPhoneNBr, FlexGrid
        SetColWidths ColClientMobilePhoneNbr, FlexGrid
        SetColWidths ColClientAddress, FlexGrid
        .ColWidth(ColClientZoneNo) = 0
        SetColWidths ColClientZoneName, FlexGrid
        SetColWidths ColClientNotes, FlexGrid
        SetColWidths ColclientDefineName, FlexGrid
        SetColWidths ColClientVisitedCount, FlexGrid
        
    End With
End If
End Sub



Sub ChangeCursor(sender As Control, Optional top As Variant = 0, Optional left As Variant = 0)

    With sender
       Grid.top = top + .top + .Height
       Grid.left = left + .left
       Grid.Width = .Width
    End With

End Sub
Function GetCountClientsVisited() As Integer
On Error GoTo ErrorHandler
Dim countclientsVisited As Integer

With FlexGrid
    For i = 1 To .Rows - 1
         countclientsVisited = countclientsVisited + .TextMatrix(i, ColClientVisitedCount)
    Next
End With

GetCountClientsVisited = countclientsVisited
Exit Function
ErrorHandler:
MsgBox Err.Description
GetCountClientsVisited = -1
End Function

Sub FillGrid(withReparation As Boolean, VTop As Double, searchFields As FiledColumns, Optional searchFormula As String)
On Error GoTo ErrorHandler
Dim sqlText As String
FillFieldsCollectionToSearch

Dim rsClients As New ADODB.Recordset

sqlText = mainDataService.GetClients(withReparation, VTop, searchFields, searchFormula)
If sqlText = "" Then
FlexGrid.Rows = 1
Else

    Set rsClients = de.con.Execute(sqlText)
    Set FlexGrid.DataSource = rsClients
End If
LCount.Caption = FlexGrid.Rows - 1
LCountVisited.Caption = GetCountClientsVisited
FillFormating 2, FlexGrid
AddDefaultCheckBoxesIntoTheGrid
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Sub AddDefaultCheckBoxesIntoTheGrid()



If fildsToSearch Is Nothing Then
    With FlexGrid
        .Cell(flexcpChecked, 0, 1, 0, .Cols - 2) = True
        .Cell(flexcpChecked, 0, 1, 0, .Cols - 2) = flexUnchecked
        .Cell(flexcpChecked, 0, ColClientPhoneNBr, 0, ColClientPhoneNBr) = True
        .Cell(flexcpChecked, 0, ColClientMobilePhoneNbr, 0, ColClientMobilePhoneNbr) = True
    End With
Else
    Dim fieldColumn As fieldColumn
    With FlexGrid
        For i = 1 To fildsToSearch.Count
            If fildsToSearch(i).isChecked Then
                .Cell(flexcpChecked, 0, fildsToSearch(i).colHeaderName, 0, fildsToSearch(i).colHeaderName) = flexChecked
            Else
                .Cell(flexcpChecked, 0, fildsToSearch(i).colHeaderName, 0, fildsToSearch(i).colHeaderName) = flexUnchecked
            End If
         Next i
    End With
End If

End Sub
Sub init()
    top = 0
    left = 0
    Ok = True
        
    AddDefaultCheckBoxesIntoTheGrid
    FillGrid ChkWithReparation.Value, txtTop.Text, fildsToSearch

    
    FlexGrid.Editable = flexEDKbdMouse
End Sub

Private Sub ChkWithReparation_Click()
If ChkWithReparation.Value Then
    ChkWithReparation.Caption = "«·Ã„Ì⁄"
Else
   ChkWithReparation.Caption = "«· Ì ·Â« «’·«Õ« "
End If
FillFieldsCollectionToSearch
FillGrid ChkWithReparation.Value, Val(txtTop.Text), fildsToSearch, txtClientName.Text
End Sub

Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If FlexGrid.Cell(flexcpChecked, Row, Col) = flexChecked Then
   FillGrid ChkWithReparation.Value, Val(txtTop.Text), fildsToSearch, txtClientName.Text
End If

End Sub

Private Sub FlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
If Row <> 0 Then cancel = True
End Sub

Private Sub FlexGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button And vbRightButton Then
   PopupMenu mnu
            
'            If Gettag(empNo, 42) Then
'                PopupMenu mnu
'            End If
'            If Gettag(empNo, 39) Then
'                PopupMenu mnuClaim
'            End If
    End If
End Sub

'Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'On Error GoTo errorhandler
'With FlexGrid
'    If .TextMatrix(Row, ColZoneName) <> oldZoneName Then
'        If MsgBox("”Ì „  €Ì— ≈”„ «·„‰ŸﬁÂ" + Chr(13) + "Â· √‰  „ √ﬂœ", vbQuestion + vbDefaultButton2 + vbYesNo, " €ÌÌ— ≈”„ «·„‰ÿﬁÂ") = vbYes Then
'            mainDataService.ChangeTheZoneName .TextMatrix(Row, ColZoneName), .TextMatrix(Row, ColZoneNo)
'        Else
'            .TextMatrix(Row, ColZoneName) = oldZoneName
'        End If
'    End If
'End With
'Exit Sub
'errorhandler:
'FlexGrid.TextMatrix(Row, ColZoneName) = oldZoneName
'MsgBox Err.Description
'End Sub

'Private Sub FlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
'If Col <> ColZoneName Then cancel = True
'oldZoneName = FlexGrid.TextMatrix(Row, ColZoneName)
'End Sub

Private Sub FlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim FirstRow As Integer, LastRow As Integer, Vrow As Integer
With FlexGrid
If KeyCode = vbKeyDelete Then
    If MsgBox("Â· √‰  „ √ﬂœ „‰ ⁄„·Ì… «·Õ–›", vbYesNo + vbDefaultButton2, "Õ–› «·“»«∆‰ «·„Õœœ…") = vbYes Then
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
                  'RemoveFromMvStock .TextMatrix(i, Colid)
                  If DeleteRow(FlexGrid, Vrow, ColClientNo, "AdhamView7", "AdhamNo") Then
                    .RemoveItem Vrow
                End If
            End If
        Next
        .Col = ColCallNo
        .SetFocus
    End If
End If
End With
End Sub

Private Sub Form_Load()
init
End Sub

Private Sub mnuPrint_Click()
    If Gettag(empNo, 42) Then
        PrintRep
    Else
        MsgBox "·«ÌÊÃœ ’·«ÕÌÂ ·«” ⁄—«÷ «·‘ﬂ«ÊÏ", vbExclamation, "·«ÌÊÃœ ’·«ÕÌÂ"
    End If


End Sub

 Sub FillFieldsCollectionToSearch()
Set fildsToSearch = New FiledColumns

Dim fieldToAdd As fieldColumn



With FlexGrid
'    If FlexGrid.Cell(flexcpChecked, 0, ColClientName) = flexChecked Then
        Set fieldToAdd = New fieldColumn
        fieldToAdd.filedId = "AdhamName"
        fieldToAdd.colHeaderName = ColClientName
        fieldToAdd.isChecked = IIf(FlexGrid.Cell(flexcpChecked, 0, ColClientName) = flexChecked, True, False)
        fildsToSearch.Add fieldToAdd.filedId, fieldToAdd.colHeaderName, fieldToAdd.isChecked, fieldToAdd.filedId
'    End If
    
'    If FlexGrid.Cell(flexcpChecked, 0, ColClientPhoneNBr) = flexChecked Then
        Set fieldToAdd = New fieldColumn
        fieldToAdd.filedId = "adhamPhon"
        fieldToAdd.colHeaderName = ColClientPhoneNBr
        fieldToAdd.isChecked = IIf(FlexGrid.Cell(flexcpChecked, 0, ColClientPhoneNBr) = flexChecked, True, False)
        fildsToSearch.Add fieldToAdd.filedId, fieldToAdd.colHeaderName, fieldToAdd.isChecked, fieldToAdd.filedId
'    End If
'    If FlexGrid.Cell(flexcpChecked, 0, ColClientMobilePhoneNbr) = flexChecked Then
        Set fieldToAdd = New fieldColumn
        fieldToAdd.filedId = "MobilePhone"
        fieldToAdd.colHeaderName = ColClientMobilePhoneNbr
        fieldToAdd.isChecked = IIf(FlexGrid.Cell(flexcpChecked, 0, ColClientMobilePhoneNbr) = flexChecked, True, False)
        fildsToSearch.Add fieldToAdd.filedId, fieldToAdd.colHeaderName, fieldToAdd.isChecked, fieldToAdd.filedId
'    End If
'    If FlexGrid.Cell(flexcpChecked, 0, ColClientAddress) = flexChecked Then
        Set fieldToAdd = New fieldColumn
        fieldToAdd.filedId = "adhamAdress"
        fieldToAdd.colHeaderName = ColClientAddress
        fieldToAdd.isChecked = IIf(FlexGrid.Cell(flexcpChecked, 0, ColClientAddress) = flexChecked, True, False)
        fildsToSearch.Add fieldToAdd.filedId, fieldToAdd.colHeaderName, fieldToAdd.isChecked, fieldToAdd.filedId
'    End If
'    If FlexGrid.Cell(flexcpChecked, 0, ColclientDefineName) = flexChecked Then
        Set fieldToAdd = New fieldColumn
        fieldToAdd.filedId = "defindname"
        fieldToAdd.colHeaderName = ColclientDefineName
        fieldToAdd.isChecked = IIf(FlexGrid.Cell(flexcpChecked, 0, ColclientDefineName) = flexChecked, True, False)
        fildsToSearch.Add fieldToAdd.filedId, fieldToAdd.colHeaderName, fieldToAdd.isChecked, fieldToAdd.filedId
'    End If
'    If FlexGrid.Cell(flexcpChecked, 0, ColclientDefineName) = flexChecked Then
        Set fieldToAdd = New fieldColumn
        fieldToAdd.filedId = "zoneName"
        fieldToAdd.colHeaderName = ColClientZoneName
        fieldToAdd.isChecked = IIf(FlexGrid.Cell(flexcpChecked, 0, ColClientZoneName) = flexChecked, True, False)
        fildsToSearch.Add fieldToAdd.filedId, fieldToAdd.colHeaderName, fieldToAdd.isChecked, fieldToAdd.filedId
        
        Set fieldToAdd = New fieldColumn
        fieldToAdd.filedId = "Notes"
        fieldToAdd.colHeaderName = ColClientNotes
        fieldToAdd.isChecked = IIf(FlexGrid.Cell(flexcpChecked, 0, ColClientNotes) = flexChecked, True, False)
        fildsToSearch.Add fieldToAdd.filedId, fieldToAdd.colHeaderName, fieldToAdd.isChecked, fieldToAdd.filedId
'    End If
End With

End Sub

Private Sub newClaim_Click()
    If Gettag(empNo, 39) Then
        ClientNo = FlexGrid.TextMatrix(FlexGrid.Row, ColClientNo)
        FrmMaintCallNew.Show
    Else
        MsgBox "·«ÌÊÃœ ’·«ÕÌÂ ·≈‰‘«¡ ‘ﬂÊÏ ÃœÌœÂ", vbExclamation, "·«ÌÊÃœ ’·«ÕÌÂ"
    End If
End Sub

Private Sub txtClientName_Change()

FillGrid ChkWithReparation.Value, Val(txtTop.Text), fildsToSearch, txtClientName.Text

End Sub

Private Sub TxtClientName_GotFocus()
ChangeToArabic
End Sub

Sub PrintRep()
On Error GoTo ErrorHandler


Dim sqlText As String
sqlText = "SELECT  adhamname,RepPrice,Bdate,Edate,Notes,CallNo"
sqlText = sqlText & " ,RegestDate,ZoneName,TeamName "
sqlText = sqlText & " From AdhamClientRepView "
sqlText = sqlText & " where CliNo= " & FlexGrid.TextMatrix(FlexGrid.Row, ColClientNo)
sqlText = sqlText & " order by repDate asc "
With cr1
    .Connect = ConnectName("")
    .SQLQuery = sqlText
    .ReportFileName = App.Path & "\Reports\RepClientReparation.rpt"
    .DiscardSavedData = True
    .WindowState = crptMaximized
    .Action = 1
End With

Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
'Sub AddZoneToTheGrid(ZoneName As String, ZoneNo As Long, FlexGrid As VSFlexGrid)
'On Error GoTo errorhandler
'Dim Vrow As Integer
'With FlexGrid
'    .AddItem ""
'    Vrow = .Rows - 1
'    .TextMatrix(Vrow, ColZoneNo) = ZoneNo
'    .TextMatrix(Vrow, ColZoneName) = ZoneName
'    .TextMatrix(Vrow, ColZoneClientsCount) = 0
'End With
'Exit Sub
'errorhandler:
'MsgBox Err.Description
'End Sub
'Private Sub TxtClientName_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If FlexGrid.Rows = 1 Then
'        Dim ZoneNo As Long
'        ZoneNo = mainDataService.InsertNewZone(txtClientName.Text)
'        If ZoneNo <> -1 Then
'            AddZoneToTheGrid txtClientName.Text, ZoneNo, FlexGrid
'            FillFormating 2, FlexGrid
'            txtClientName.SetFocus
'            SendKeys "{home}+{end}"
'        End If
'    End If
'End If
'End Sub
Private Sub txtTop_Change()
    FillFieldsCollectionToSearch
    FillGrid ChkWithReparation.Value, CDbl(txtTop.Text), fildsToSearch, txtClientName.Text
End Sub
