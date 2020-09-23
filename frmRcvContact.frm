VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRcvContact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Receive Contact"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   364
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInfo 
      Caption         =   "&Info"
      Height          =   330
      Left            =   1575
      TabIndex        =   3
      Top             =   5040
      Width           =   960
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   330
      Left            =   2520
      TabIndex        =   2
      Top             =   5040
      Width           =   1065
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   330
      Left            =   3675
      TabIndex        =   1
      Top             =   5040
      Width           =   1065
   End
   Begin MSComctlLib.ListView lstContact 
      Height          =   4950
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   8731
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "UIN"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nickname"
         Object.Width           =   5556
      EndProperty
   End
End
Attribute VB_Name = "frmRcvContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
  Dim TempUin As Long
  With lstContact
    If .ListItems.Count = 0 Then
      MsgBox "There is no contact on the list anymore.", vbOKOnly, "No more Contact"
      Exit Sub
    End If
    
    TempUin = CLng(.SelectedItem)
    
    If ContactExist(TempUin) Then
      .ListItems.Remove (.SelectedItem.Index)
      Exit Sub
    End If
    ContactNewUser TempUin, .SelectedItem.ListSubItems(1)
    ContactTotal = ContactTotal + 1
    Cont_Add TempUin, .SelectedItem.ListSubItems(1), ContactTotal - 1
    'icqUdp.InfoRequestBasic TempUin
    IcqUdp.ContactAdd TempUin
    .ListItems.Remove (.SelectedItem.Index)
  End With
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Public Sub EventRcvContact(uinlist, nicklist)
  Dim TempList As ListItem

  For i = 0 To UBound(uinlist)
    Set TempList = lstContact.ListItems.Add(, , uinlist(i))
    TempList.ListSubItems.Add , , nicklist(i)
  Next i
End Sub

Private Sub cmdInfo_Click()
  Dim TempUin As Long
  Dim TempFrmIndex As Byte
    
  TempUin = Val(lstContact.SelectedItem.Text)
  If TempUin = 0 Then Exit Sub
  TempFrmIndex = frmGetIndex(enumfrmUserDetail, TempUin)
    
  If TempFrmIndex = 0 Then
    TempFrmIndex = frmLoaded(enumfrmUserDetail)
    If TempFrmIndex = 0 Then
      MsgBox "There is not enough memory to open more window, please close some of your unused windows.", vbCritical, "Error - Too many window"
      Exit Sub
    End If
      
    Set frmUserDetail(TempFrmIndex) = New frmUserInfo
    With frmUserDetail(TempFrmIndex)
      .Show
      .FrmIndex = TempFrmIndex
      .Caption = "User Detail of " & lstContact.SelectedItem.SubItems(1)
      TempUINBuffer = TempUin
    End With
  Else
    frmUserDetail(TempFrmIndex).Show
  End If
End Sub

Private Sub Form_Load()
  Me.Show
End Sub
