VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSndContact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send Contact"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lstContact 
      Height          =   4635
      Left            =   105
      TabIndex        =   4
      Top             =   945
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   8176
      View            =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   330
      Left            =   3150
      TabIndex        =   2
      Top             =   5670
      Width           =   960
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   330
      Left            =   2205
      TabIndex        =   1
      Top             =   5670
      Width           =   960
   End
   Begin VB.ComboBox cmbRecepient 
      Height          =   315
      ItemData        =   "frmSndContact.frx":0000
      Left            =   105
      List            =   "frmSndContact.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   315
      Width           =   4005
   End
   Begin VB.Label Label2 
      Caption         =   "&Contact to Send"
      Height          =   225
      Left            =   105
      TabIndex        =   5
      Top             =   735
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "&Recepient"
      Height          =   225
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Width           =   1170
   End
End
Attribute VB_Name = "frmSndContact"
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
  Dim UinList As Variant
  Dim NickList As Variant
  Dim TotalContact As Integer
  Dim i As Integer
  
  If IcqUdp.SocketState <> icqConnected Then
    MsgBox "Communicator is not currently connected to the ICQ Server. Please connect by choosing File->Connect on the file menu", vbOKOnly, "Not Connected"
    Exit Sub
  End If
  
  If cmbRecepient.ListIndex = -1 Then
    MsgBox "You need to specify the recepient of your message. To do this click on the recepient combo box and choose the nickname of the recepient", vbOKOnly, "No Recepient"
  End If
  
  cmdSend.Enabled = False
  RecvUin = cmbRecepient.ItemData(cmbRecepient.ListIndex)
  TotalContact = -1
  
  With lstContact
    For i = 1 To .ListItems.Count
      If .ListItems(i).Checked = True Then
        TotalContact = TotalContact + 1
        If TotalContact > 0 Then
          ReDim Preserve UinList(TotalContact) As Long
          ReDim Preserve NickList(TotalContact) As String
        Else
          ReDim UinList(0) As Long
          ReDim NickList(0) As String
        End If
          
        UinList(TotalContact) = Val(Right$(.ListItems(i).Key, Len(.ListItems(i).Key) - 1))
        NickList(TotalContact) = .ListItems(i).Text
      End If
    Next i
  End With
  SetAck IcqUdp.SendContact(RecvUin, UinList, NickList), enumFrmSndContact, 0
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
      cmbRecepient.AddItem Uin2Name(!uin)
      cmbRecepient.ItemData(i) = !uin
      
      lstContact.ListItems.Add , "o" & Trim$(Str$(!uin)), Uin2Name(!uin)
      i = i + 1
      .MoveNext
    Loop
  End With
End Sub

Public Sub EventContRecvAck()
  cmdSend.Enabled = True
End Sub

Public Sub SetContRecepient(uin As Long, strname As String)
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

