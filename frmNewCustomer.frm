VERSION 5.00
Begin VB.Form frmNewCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "“»Ê‰ ÃœÌœ"
   ClientHeight    =   705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   705
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtClientPhoneNBR 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   270
      Width           =   1995
   End
   Begin VB.TextBox TxtClientName 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2070
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   2505
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Â« › «·“»Ê‰"
      Height          =   195
      Index           =   2
      Left            =   1245
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   30
      Width           =   825
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "≈”„ «·“»Ê‰"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   3915
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   30
      Width           =   735
   End
End
Attribute VB_Name = "frmNewCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub init()
txtClientPhoneNBR.Text = ClientPhoneNBr
If IsNumeric(ClientName) Then
    txtClientPhoneNBR.Text = ClientName
    SendKeys "{home}+{end}"
Else
    TxtClientName.Text = ClientName
    SendKeys "{home}+{end}"
End If
End Sub
Private Sub Form_Load()
init
End Sub

Private Sub TxtClientName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtClientPhoneNBR.SetFocus
    txtClientPhoneNBR.SelStart = 0
    txtClientPhoneNBR.SelLength = Len(txtClientPhoneNBR.Text)
End If
End Sub
Function InsertNewClient(ClientName As String, ClientPhoneNBr As String) As Double
    On Error GoTo errorhandler
    Dim RsMaxi As New ADODB.Recordset
    If TxtClientName.Text = "" Then
        InsertNewClient = 0
        Exit Function
    End If
    Sqltext = "Insert Into Coclient (ClientName,ClientPhoneNBR) Values('" & ClientName & "','" & ClientPhoneNBr & "')"
    de.con.Execute (Sqltext)
    Sqltext = "Select Max(ClientId) MaxClientId From CoClient"
    Set RsMaxi = de.con.Execute(Sqltext)
    InsertNewClient = RsMaxi!MaxClientId
    Exit Function
errorhandler:
    InsertNewClient = 0
    MsgBox Err.Description
End Function

Private Sub txtClientPhoneNBR_KeyPress(KeyAscii As Integer)
On Error GoTo errorhandler
    If KeyAscii = 13 Then
        If ClientId <> 0 Then
            If TxtClientName.Text <> "" Then
                Sqltext = "Update Coclient set ClientName='" & TxtClientName.Text & "',ClientPhoneNBR='" & txtClientPhoneNBR.Text & "' Where ClientId=" & ClientId
                de.con.Execute (Sqltext)
            End If
        Else
            ClientId = InsertNewClient(TxtClientName.Text, txtClientPhoneNBR.Text)
        End If
        ClientName = TxtClientName.Text
        ClientPhoneNBr = txtClientPhoneNBR.Text
        Unload Me
    End If
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub
