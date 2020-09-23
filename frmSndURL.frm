VERSION 5.00
Begin VB.Form frmSndURL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send URL"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbRecepient 
      Height          =   315
      ItemData        =   "frmSndURL.frx":0000
      Left            =   1365
      List            =   "frmSndURL.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   105
      Width           =   4215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   330
      Left            =   3675
      TabIndex        =   6
      Top             =   3885
      Width           =   960
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   330
      Left            =   4620
      TabIndex        =   5
      Top             =   3885
      Width           =   960
   End
   Begin VB.TextBox txtURLDesc 
      Height          =   2850
      Left            =   1365
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   945
      Width           =   4215
   End
   Begin VB.TextBox txtUrlAddress 
      Height          =   285
      Left            =   1365
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   525
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "&Recepient"
      Height          =   225
      Left            =   105
      TabIndex        =   4
      Top             =   210
      Width           =   1170
   End
   Begin VB.Label lblUrlDescription 
      Caption         =   "URL &Description"
      Height          =   225
      Left            =   105
      TabIndex        =   2
      Top             =   1050
      Width           =   1170
   End
   Begin VB.Label lblUrlAddress 
      Caption         =   "URL &Address"
      Height          =   225
      Left            =   105
      TabIndex        =   0
      Top             =   630
      Width           =   1695
   End
End
Attribute VB_Name = "frmSndURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdSend_Click()
  Dim RecvUin As Long
  If IcqUdp.SocketState <> icqConnected Then
    MsgBox "Communicator is not currently connected to the ICQ Server. Please connect by choosing File->Connect on the file menu", vbOKOnly, "Not Connected"
    Exit Sub
  End If
  
  If cmbRecepient.ListIndex = -1 Then
    MsgBox "You need to specify the recepient of your message. To do this click on the recepient combo box and choose the nickname of the recepient", vbOKOnly, "No Recepient"
  End If
  
  cmdSend.Enabled = False
  RecvUin = cmbRecepient.ItemData(cmbRecepient.ListIndex)
  SetAck IcqUdp.sendurl(RecvUin, txtUrlAddress.Text, txtURLDesc.Text), enumFrmSndUrl, 0
End Sub

Private Sub Form_Load()
  Dim i As Integer
  Me.Show
  
  ContRS.Filter = "bOnContact = TRUE"
  Set ContRSFilter = ContRS.OpenRecordset
  With ContRSFilter
    .Edit
    ContactTotal = .RecordCount - 1
    .MoveFirst
    Do While .EOF = False
      cmbRecepient.AddItem !DisplayName
      cmbRecepient.ItemData(i) = !uin
      i = i + 1
      .MoveNext
    Loop
  End With
End Sub

Public Sub EventUrlRecvAck()
  cmdSend.Enabled = True
End Sub

Public Sub SetUrlRecepient(uin As Long, strname As String)
  Dim i As Integer
  For i = 0 To cmbRecepient.ListCount
    If cmbRecepient.ItemData(i) = uin Then
      cmbRecepient.ListIndex = i
      Exit Sub
    End If
  Next i
  
  cmbRecepient.AddItem strname
  cmbRecepient.ItemData(cmbRecepient.ListCount) = uin
  cmbRecepient.ListIndex = cmbRecepient.ListCount
End Sub
