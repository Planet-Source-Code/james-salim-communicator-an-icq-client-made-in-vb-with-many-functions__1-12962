VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find/Add Contact"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   330
      Left            =   4305
      TabIndex        =   16
      Top             =   5355
      Width           =   1065
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "&Info"
      Height          =   330
      Left            =   3360
      TabIndex        =   15
      Top             =   5355
      Width           =   960
   End
   Begin MSComctlLib.ListView lstSearch 
      Height          =   2850
      Left            =   210
      TabIndex        =   14
      Top             =   2415
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   5027
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ICQ #"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nick"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "First"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Last"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Email"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Authorize"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   330
      Left            =   5460
      TabIndex        =   13
      Top             =   5355
      Width           =   1065
   End
   Begin VB.CommandButton cmdSearchUin 
      Caption         =   "&ICQ #"
      Height          =   330
      Left            =   5460
      TabIndex        =   12
      Top             =   1890
      Width           =   1065
   End
   Begin VB.CommandButton cmdSearchEmail 
      Caption         =   "&Email"
      Height          =   330
      Left            =   5460
      TabIndex        =   11
      Top             =   1470
      Width           =   1065
   End
   Begin VB.CommandButton cmdSearchName 
      Caption         =   "Search &Name"
      Height          =   1170
      Left            =   5460
      TabIndex        =   10
      Top             =   210
      Width           =   1065
   End
   Begin VB.TextBox txtUin 
      Height          =   285
      Left            =   1470
      MaxLength       =   10
      TabIndex        =   9
      Top             =   1890
      Width           =   3900
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1470
      MaxLength       =   60
      TabIndex        =   7
      Top             =   1470
      Width           =   3900
   End
   Begin VB.TextBox txtLast 
      Height          =   285
      Left            =   1470
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1050
      Width           =   3900
   End
   Begin VB.TextBox txtFirst 
      Height          =   285
      Left            =   1470
      MaxLength       =   20
      TabIndex        =   3
      Top             =   630
      Width           =   3900
   End
   Begin VB.TextBox txtNick 
      Height          =   285
      Left            =   1470
      MaxLength       =   20
      TabIndex        =   1
      Top             =   210
      Width           =   3900
   End
   Begin VB.Label lblUin 
      Caption         =   "ICQ #"
      Height          =   225
      Left            =   210
      TabIndex        =   8
      Top             =   1890
      Width           =   1275
   End
   Begin VB.Label lblEmail 
      Caption         =   "Email Address:"
      Height          =   225
      Left            =   210
      TabIndex        =   6
      Top             =   1470
      Width           =   1275
   End
   Begin VB.Label lblLast 
      Caption         =   "Last Name:"
      Height          =   225
      Left            =   210
      TabIndex        =   4
      Top             =   1050
      Width           =   1380
   End
   Begin VB.Label lblFirst 
      Caption         =   "First Name:"
      Height          =   225
      Left            =   210
      TabIndex        =   2
      Top             =   630
      Width           =   1275
   End
   Begin VB.Label lblNick 
      Caption         =   "Nick Name:"
      Height          =   225
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   1275
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
  Dim TempUin As Long
  Dim TempNick As String
  Dim TempFirst As String
  Dim TempLast As String
  Dim TempEmail As String
  
  With lstSearch
    If .ListItems.Count = 0 Then
      MsgBox "There is no contact on the list anymore.", vbOKOnly, "No more Contact"
      Exit Sub
    End If
    
    TempUin = CLng(.SelectedItem)
    
    If ContactExist(TempUin) Then
      .ListItems.Remove (.SelectedItem.Index)
      Exit Sub
    End If
    
    TempNick = .SelectedItem.ListSubItems(1)
    TempFirst = .SelectedItem.ListSubItems(2)
    TempLast = .SelectedItem.ListSubItems(3)
    TempEmail = .SelectedItem.ListSubItems(4)

    ContactNewUser TempUin, TempNick, TempFirst, TempLast, TempEmail
    ContactTotal = ContactTotal + 1
    Cont_Add TempUin, .SelectedItem.ListSubItems(1), ContactTotal - 1
    'IcqUdp.inforequestbasic TempUin
    IcqUdp.ContactAdd TempUin
    .ListItems.Remove (.SelectedItem.Index)
  End With
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdInfo_Click()
  Dim TempUin As Long
  Dim TempFrmIndex As Byte
    
  TempUin = Val(lstSearch.SelectedItem.Text)
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
      .Caption = "User Detail of " & lstSearch.SelectedItem.SubItems(1)
      TempUINBuffer = TempUin
    End With
  Else
    frmUserDetail(TempFrmIndex).Show
  End If
End Sub

Private Sub cmdSearchEmail_Click()
  cmdSearchName.Enabled = False
  cmdSearchEmail.Enabled = False
  cmdSearchUin.Enabled = False
  lstSearch.ListItems.Clear
  IcqUdp.SearchEmail txtEmail
End Sub

Private Sub cmdSearchName_Click()
  cmdSearchName.Enabled = False
  cmdSearchEmail.Enabled = False
  cmdSearchUin.Enabled = False
  lstSearch.ListItems.Clear
  IcqUdp.SearchName txtNick, txtFirst, txtLast
End Sub

Private Sub cmdSearchUin_Click()
  cmdSearchName.Enabled = False
  cmdSearchEmail.Enabled = False
  cmdSearchUin.Enabled = False
  lstSearch.ListItems.Clear
  IcqUdp.SearchUin CLng(txtUin)
End Sub

Public Sub Event_SearchReply( _
  uin As Long, nick As String, first As String, _
  last As String, email As String, bAuth As Boolean, _
  SearchResult As IcqUdpCtl.enumSearchResult)
  
  Dim TempLst As ListItem
  
  Select Case SearchResult
    Case icqSearchUserFound
      Set TempLst = lstSearch.ListItems.Add(, , Trim$(Str$(uin)))
      TempLst.ListSubItems.Add , , nick
      TempLst.ListSubItems.Add , , first
      TempLst.ListSubItems.Add , , last
      TempLst.ListSubItems.Add , , email
      TempLst.ListSubItems.Add , , IIf(bAuth, "Authorize", "Always")
    Case icqSearchDone
      cmdSearchName.Enabled = True
      cmdSearchEmail.Enabled = True
      cmdSearchUin.Enabled = True
    Case icqSearchTooMany
      cmdSearchName.Enabled = True
      cmdSearchEmail.Enabled = True
      cmdSearchUin.Enabled = True
      MsgBox "Search return too many result, try narrowing down your search by filling in more variable", vbExclamation Or vbOKOnly, "Search too many"
  End Select
End Sub

