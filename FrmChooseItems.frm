VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form FrmChooseItems 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Õœœ «·«—ﬁ«„ «·„Œ“‰Ì…"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   5835
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   10292
      _Version        =   131074
      Begin VSFlex8Ctl.VSFlexGrid Grid 
         Height          =   5145
         Left            =   60
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   660
         Visible         =   0   'False
         Width           =   3945
         _cx             =   6959
         _cy             =   9075
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
      Begin VB.TextBox TxtStkNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   330
         Width           =   3915
      End
      Begin VSFlex8Ctl.VSFlexGrid FGrid 
         Height          =   5145
         Left            =   60
         TabIndex        =   3
         Top             =   660
         Width           =   3945
         _cx             =   6959
         _cy             =   9075
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
         Rows            =   2
         Cols            =   2
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·„«œ…"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3270
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   90
         Width           =   675
      End
   End
End
Attribute VB_Name = "FrmChooseItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ColStkId = 1
Const ColStkNo = 2
Const ColStkName = 3

Dim Flag As Boolean

Function DataOk(Stkno As String) As Boolean
Dim RsFind As New ADODB.Recordset
Sqltext = "Select StkNo From CoStock Where StkNo ='" & Stkno & "'"
Set RsFind = de.con.Execute(Sqltext)
If RsFind.RecordCount > 0 Then
   DataOk = True
Else
    DataOk = False
End If
End Function

Sub FillActiveControl(List As VSFlexGrid)
    With List
        If ActiveControl.text <> "" Then
            If Not ActiveControl.DataChanged Then Exit Sub
            Flag = False
            ActiveControl.text = IIf(.Visible = False, "", .TextMatrix(.Row, ColStkNo))
            Flag = True
            ActiveControl.Tag = IIf(.Visible = False, "", .TextMatrix(.Row, ColStkId))
        Else
            ActiveControl.text = ""
            ActiveControl.Tag = ""
        End If
        List.Visible = False
        ActiveControl.DataChanged = False
    End With
End Sub

Sub MoveCursor(KeyCode As Integer)
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

'Sub SetColWidthsVS(ByVal ColNo As Integer, FlexGrid As VSFlexGrid)
'    Dim i, J, s, w
'    With FlexGrid
'            s = 0
'            For i = 0 To .Rows - 1
'                w = TextWidth(.TextMatrix(i, ColNo))
'                If w > s Then s = w
'            Next i
'            .ColWidth(ColNo) = s + 300
'    End With
'End Sub

Sub FillFormatVSFlex(FlexGrid As VSFlexGrid, Switch As Boolean)
If Switch Then
    Fs = "|ID"
    Fs = Fs + "|>" + "«·—ﬁ„ «·„Œ“‰Ì"
    Fs = Fs + "|>" + "«·≈”„"
    With FlexGrid
        .Visible = False
        .FormatString = Fs
            .ColWidth(ColStkId) = 0
            SetColWidths ColStkNo, FlexGrid
            SetColWidths ColStkName, FlexGrid
            .Visible = True
    End With
End If
End Sub

Sub FillList(Sqltext As String, Field1 As String, Field2 As String, List As VSFlexGrid, ByVal Switch As Boolean)
    Set rs = de.con.Execute(Sqltext)
    If rs.RecordCount > 0 Then
        Set List.DataSource = rs
        FillFormatVSFlex List, Switch
        List.Row = 1
        List.Col = 1
        List.ColSel = List.Cols - 1
        List.Visible = True
        TxtStkNo.SetFocus
    Else
'        List.Text = ""
        List.Rows = 1
        TxtStkNo.Tag = 0
        MsgBox "«·„«œ… €Ì— „ÊÃÊœ… ÷„‰ ﬁ«∆„… «·√—ﬁ«„ «·„Œ“‰Ì…", vbExclamation, " ‰»ÌÂ"
        List.Visible = False
        TxtStkNo.SelStart = 0
        TxtStkNo.SelLength = Len(TxtStkNo.text)
        TxtStkNo.SetFocus
    End If
End Sub

Function Found(Str As String) As Boolean
    With FGrid
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 1) = LTrim(RTrim(Str)) Then
                Found = True
                Exit Function
            End If
        Next
        Found = False
    End With
End Function

Sub InsertIntoGrid()
With FGrid
    .AddItem ""
    .TextMatrix(.Rows - 1, 1) = LTrim(RTrim(TxtStkNo.text))
End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Me.Hide
End If
End Sub

Private Sub Form_Load()
    FGrid.ColWidth(0) = 0
    FGrid.Col = 0
    FGrid.Rows = 0
    Flag = True
End Sub

Private Sub Form_Unload(cancel As Integer)
Form_KeyDown vbKeyEscape, 0
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyDelete Then
    With Grid
        If .Row = 0 And .Rows = 1 Then
            .Rows = 0
        Else
        .RemoveItem .Row
        End If
    End With
End If
End Sub

Private Sub TxtStkNo_Change()
If Flag Then
    Dim Sqltext As String
    If Trim(TxtStkNo.text) = "" Then
        TxtStkNo.Tag = ""
        Grid.Visible = False
        Exit Sub
        End If
    Sqltext = "Select top 15 Id , ltrim(rtrim(StkNo))StkNo , ltrim(rtrim(StkName))StkName From CoStock  Where StkNo like " & LikeExpression(TxtStkNo.text) & " Or Stkname Like " & LikeExpression(TxtStkNo.text) & " Order By len(ltrim(rtrim(StkNo))) , ltrim(rtrim(StkNo))"
    FillList Sqltext, "StkNo", "StkNo", Grid, True
End If
End Sub

Private Sub TxtStkNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = False
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode
    Flag = True
End Sub

Private Sub txtStkNo_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler
    If KeyAscii = 13 Then
        If DataOk(TxtStkNo.text) Then  'Chk StkItem If Found in CoStock
            FillActiveControl Grid
            If Not Found(TxtStkNo.text) Then
                InsertIntoGrid
            End If

        Else
            Grid.Visible = False
            Sendkeys "{home}+{end}"
            MsgBox "«·„«œ… €Ì— „ÊÃÊœ… ÷„‰ ﬁ«∆„… «·√—ﬁ«„ «·„Œ“‰Ì…", vbExclamation, " ‰»ÌÂ"
            Exit Sub
        End If
        Sendkeys "{home}+{end}"
    End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

'Private Sub Txtitem_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If Not Found(Txtitem) Then
'        InsertIntoGrid
'    End If
'    Txtitem.SelStart = 0
'    Txtitem.SelLength = Len(Txtitem.Text)
'    Txtitem.SetFocus
'End If
'End Sub
