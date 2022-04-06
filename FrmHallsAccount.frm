VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrmHallsAccount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈Ì—«œ«  Ê „’«—Ì› Œœ„… «·„” Â·ﬂ „Õ«›Ÿ« "
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   11895
   Begin MSDataListLib.DataCombo ComboHall 
      Height          =   360
      Left            =   6030
      TabIndex        =   1
      Top             =   330
      Width           =   4485
      _ExtentX        =   7911
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
   Begin VB.TextBox TxtNotes 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1050
      Width           =   11715
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   4815
      Left            =   60
      TabIndex        =   5
      Top             =   1440
      Width           =   11685
      _cx             =   20611
      _cy             =   8493
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
   Begin MSMask.MaskEdBox TxtDate 
      Height          =   375
      Left            =   10560
      TabIndex        =   0
      Top             =   300
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   330
      Width           =   1395
   End
   Begin MSDataListLib.DataCombo ComboClass 
      Height          =   360
      Left            =   1470
      TabIndex        =   2
      Top             =   330
      Width           =   4515
      _ExtentX        =   7964
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "„·«ÕŸ« "
      Height          =   195
      Index           =   4
      Left            =   11190
      TabIndex        =   10
      Top             =   810
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "«· «—ÌŒ"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   3
      Left            =   11340
      TabIndex        =   9
      Top             =   60
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "«·„»·€"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   2
      Left            =   1080
      TabIndex        =   8
      Top             =   60
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "«· ’‰Ì›"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   1
      Left            =   5460
      TabIndex        =   7
      Top             =   60
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "«·„Õ«›Ÿ…"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   9870
      TabIndex        =   6
      Top             =   60
      Width           =   600
   End
End
Attribute VB_Name = "FrmHallsAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Colid = 1
Const ColDate = 2
Const ColAmount = 3
Const ColHallId = 4
Const ColHallName = 5
Const ColClassId = 6
Const ColClassName = 7
Const ColNotes = 8

Dim HallRec As HallsType

Sub FillFormating(ByVal i As Integer)
If i = 1 Then
    Fs = "|>" + "Id"
    Fs = Fs + "|>" + " «—ÌŒ «·Õ—ﬂ…"
    Fs = Fs + "|>" + "«·„»·€"
    Fs = Fs + "|>" + "—ﬁ„ «·’«·…"
    Fs = Fs + "|>" + "«·’«·…"
    Fs = Fs + "|>" + "‰Ê⁄ «· ’‰Ì›"
    Fs = Fs + "|>" + "«· ’‰Ì›"
    Fs = Fs + "|>" + "„·«ÕŸ« "
    
    With FlexGrid
        .FormatString = Fs
        .Cols = 9
        .ColWidth(Colid) = 0
        SetColWidths ColDate, FlexGrid
        SetColWidths ColAmount, FlexGrid
        .ColWidth(ColHallId) = 0
        SetColWidths ColHallName, FlexGrid
        .ColWidth(ColClassId) = 0
        SetColWidths ColClassName, FlexGrid
        SetColWidths ColNotes, FlexGrid
    End With
End If
End Sub

Sub SetColWidths(ByVal ColNo As Integer, FlexGrid As VSFlexGrid)
    With FlexGrid
        .AutoSize (ColNo)
    End With
End Sub
Sub FillCombos()
    Dim Rs As New ADODB.Recordset
    Sqltext = "Select Code , CodeDescription From HafezDeveloper.dbo.CoExpensiveType"
    Set Rs = de.con.Execute(Sqltext)
    Set ComboClass.RowSource = Rs
    ComboClass.ListField = "CodeDescription"
    ComboClass.BoundColumn = "Code"

    Sqltext = "select HallId , HallName  from CoMaintenanceHalls Where IsVisible=1"
    Set Rs = de.con.Execute(Sqltext)
    Set ComboHall.RowSource = Rs
    ComboHall.ListField = "HallName"
    ComboHall.BoundColumn = "HallId"
End Sub
Sub FillGrid()
On Error GoTo ERRORHANDLER
Dim Rs As New ADODB.Recordset

Sqltext = "Select Id, date , Amount , HallId, hallname, ClassId, Codedescription, notes From t_hallsAccountQry"
Set Rs = de.con.Execute(Sqltext)
Set FlexGrid.DataSource = Rs
FillFormating 1
Exit Sub
ERRORHANDLER:
MsgBox Err.Description
End Sub
Sub init()
    Top = 0
    Left = 0
    FillCombos
    TxtDate.Text = Format(Now, "dd/mm/yyyy")
    FillGrid
    FlexGrid.Editable = flexEDKbdMouse
End Sub

Private Sub comboClass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtamount.SelStart = 0
txtamount.SelLength = Len(txtamount.Text)
txtamount.SetFocus
End If

End Sub

Private Sub ComboHall_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
ComboClass.SetFocus
End If

End Sub

Private Sub flexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With FlexGrid
    Sqltext = "Update t_HallsAccount Set Amount=" & .TextMatrix(Row, ColAmount) & ",date='" & ConvertControlDate(.TextMatrix(Row, ColDate)) & "',notes ='" & .TextMatrix(.Row, ColNotes) & "' Where id=" & .TextMatrix(Row, Colid)
    de.con.Execute (Sqltext)
End With
End Sub

Private Sub FlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With FlexGrid
If .Col <> ColDate And .Col <> ColAmount And .Col <> ColNotes Then Cancel = True
End With
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
            Else
              If .TextMatrix(i, Colid) <> "" Then
                If DeleteRow(FlexGrid, Vrow, Colid, "t_hallsAccount", "Id") Then
                    .RemoveItem Vrow
                End If
              Else
                    .RemoveItem Vrow
            End If
          End If
        Next
    End If
End If
End With
End Sub

Private Sub Form_Load()
    init
End Sub

Private Sub TxtAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtNotes.SelStart = 0
    TxtNotes.SelLength = Len(TxtNotes.Text)
    TxtNotes.SetFocus
End If
End Sub

Private Sub TxtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ComboHall.SetFocus
End If
End Sub
Function FillVariables() As Boolean
On Error GoTo ERRORHANDLER
If Not IsDate(TxtDate.Text) Or ComboHall.BoundText = "" Or ComboClass.BoundText = "" Or Val(txtamount.Text) = 0 Then
    FillVariables = False
    Exit Function
End If
FillVariables = True
Exit Function
ERRORHANDLER:
FillVariables = False
End Function
Function fillstructure() As Boolean
On Error GoTo ERRORHANDLER
    If FillVariables Then
        With HallRec
            .Amount = txtamount.Text
            .ClassId = ComboClass.BoundText
            .HallId = ComboHall.BoundText
            .Date = TxtDate.Text
            .Notes = TxtNotes.Text
        End With
    End If
fillstructure = True
Exit Function
ERRORHANDLER:
fillstructure = False
End Function
Function GetMaxId() As Double
Dim Rs As New ADODB.Recordset
On Error GoTo ERRORHANDLER
Sqltext = "Select Max(Id) as MaxId From t_HallsAccount"
Set Rs = de.con.Execute(Sqltext)
GetMaxId = Rs!MaxId
Exit Function
ERRORHANDLER:
GetMaxId = -1
MsgBox Err.Description
End Function
Function SaveRec() As Boolean
On Error GoTo ERRORHANDLER
If fillstructure Then
    With HallRec
        Sqltext = "Insert Into t_HallsAccount( HallId, ClassId, Amount, date, notes )Values("
        Sqltext = Sqltext & .HallId & "," & .ClassId & "," & .Amount & ",'" & ConvertControlDate(.Date) & "','" & .Notes & "')"
        de.con.Execute (Sqltext)
        .Id = GetMaxId
    End With
End If
SaveRec = True
Exit Function
ERRORHANDLER:
SaveRec = False
MsgBox Err.Description
End Function
Function GetHallName(HallId As Integer) As String
On Error GoTo ERRORHANDLER
Dim Rs As New ADODB.Recordset
Sqltext = "Select HallName From CoMaintenanceHalls Where HallId=" & HallId
Set Rs = de.con.Execute(Sqltext)
If Rs.RecordCount > 0 Then
    GetHallName = Rs!HallName
Else
    GetHallName = ""
End If
Exit Function
ERRORHANDLER:
GetHallName = ""
MsgBox Err.Description
End Function

Function GetClassName(ClassId As Integer) As String
On Error GoTo ERRORHANDLER
Dim Rs As New ADODB.Recordset
Sqltext = "Select CodeDescription From HafezDeveloper.dbo.CoExpensiveType Where Code =" & ClassId
Set Rs = de.con.Execute(Sqltext)
If Rs.RecordCount > 0 Then
    GetClassName = Rs!codedescription
Else
    GetClassName = ""
End If
Exit Function
ERRORHANDLER:
GetClassName = ""
MsgBox Err.Description
End Function


Sub insertintoGrid()
Dim Vrow As Integer
With FlexGrid
    .AddItem ""
    Vrow = .Rows - 1
    .TextMatrix(Vrow, Colid) = HallRec.Id
    .TextMatrix(Vrow, ColDate) = HallRec.Date
    .TextMatrix(Vrow, ColAmount) = HallRec.Amount
    .TextMatrix(Vrow, ColHallId) = HallRec.HallId
    .TextMatrix(Vrow, ColHallName) = GetHallName(HallRec.HallId)
    .TextMatrix(Vrow, ColClassId) = HallRec.Id
    .TextMatrix(Vrow, ColClassName) = GetClassName(HallRec.ClassId)
    .TextMatrix(Vrow, ColNotes) = HallRec.Notes
    FillFormating 1
    
End With
End Sub
Private Sub TxtNotes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If SaveRec Then
        insertintoGrid
    End If
    TxtDate.SelStart = 0
    TxtDate.SelLength = Len(TxtDate.Text)
    TxtDate.SetFocus

End If

End Sub
