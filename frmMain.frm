VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A22D979F-2684-11D2-8E21-10B404C10000}#1.4#0"; "cPopMenu.ocx"
Object = "{452A044D-A6BA-4EE7-94ED-C10266A84B17}#1.0#0"; "IcqUdpv5.ocx"
Begin VB.Form frmMain 
   Caption         =   "Communicator 2001"
   ClientHeight    =   5430
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3420
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   3420
   StartUpPosition =   3  'Windows Default
   Begin IcqUdpCtl.IcqUdp IcqUdpF 
      Left            =   1470
      Top             =   3360
      _ExtentX        =   900
      _ExtentY        =   900
   End
   Begin cPopMenu.PopMenu ctlMenu 
      Left            =   735
      Top             =   3360
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   5145
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   5530
            Text            =   "Offline"
            TextSave        =   "Offline"
            Key             =   "State"
            Object.ToolTipText     =   "Online Status, Click to change status"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListF 
      Left            =   1365
      Top             =   3990
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "StOnline"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C1C
            Key             =   "StAway"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F6E
            Key             =   "StNA"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12C0
            Key             =   "StOccupied"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1612
            Key             =   "StDND"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1964
            Key             =   "StChat"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CB6
            Key             =   "StInvisible"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2008
            Key             =   "StOffline"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":235A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26AE
            Key             =   "SndMsg"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A00
            Key             =   "SndURL"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D52
            Key             =   "SndFile"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30A4
            Key             =   "SndChat"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33F6
            Key             =   "SndContact"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3748
            Key             =   "IconUser"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A9A
            Key             =   "IconConnect"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DEE
            Key             =   "IconDisconnect"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4142
            Key             =   "ICQLogo0"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4494
            Key             =   "ICQLogo3"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":47E6
            Key             =   "IconUserTick"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B3A
            Key             =   "IconUserGroup"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E8E
            Key             =   "ICQLogo1"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":51E0
            Key             =   "IconUserInfo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwContactF 
      Height          =   4740
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   8361
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   617
      LabelEdit       =   1
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "ImgListF"
      Appearance      =   1
   End
   Begin VB.Menu mnu1 
      Caption         =   "&File"
      Begin VB.Menu mnu1Connect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnu1Status 
         Caption         =   "&Status"
         Begin VB.Menu mnu1StOnline 
            Caption         =   "&Online"
         End
         Begin VB.Menu mnu1StAway 
            Caption         =   "&Away"
         End
         Begin VB.Menu mnu1StNA 
            Caption         =   "&N/A"
         End
         Begin VB.Menu mnu1StOccupied 
            Caption         =   "&Occupied"
         End
         Begin VB.Menu mnu1StDND 
            Caption         =   "&DND"
         End
         Begin VB.Menu mnu1StChat 
            Caption         =   "&Chat"
         End
         Begin VB.Menu mnu1StInvisible 
            Caption         =   "&Invisible"
         End
      End
      Begin VB.Menu mnu1sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu1Option 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnu1EditInfo 
         Caption         =   "&Edit Info"
      End
      Begin VB.Menu mnu1sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu1Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu2 
      Caption         =   "&Contact"
      Begin VB.Menu mnu2Add 
         Caption         =   "A&dd User"
      End
      Begin VB.Menu mnu2Import 
         Caption         =   "&Import List"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu2Remove 
         Caption         =   "&Remove User"
      End
      Begin VB.Menu mnu2sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu2SndMsg 
         Caption         =   "Send &Message"
      End
      Begin VB.Menu mnu2SndURL 
         Caption         =   "Send &URL"
      End
      Begin VB.Menu mnu2SndFile 
         Caption         =   "Send &File"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnu2SndChat 
         Caption         =   "Send &Chat"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnu2SndContact 
         Caption         =   "Send &Contact List"
      End
      Begin VB.Menu mnu2sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu2UserInfo 
         Caption         =   "&User Detail"
      End
   End
   Begin VB.Menu mnu3 
      Caption         =   "&Help"
      Begin VB.Menu mnu3About 
         Caption         =   "&About Communicator"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Load frmRcvContact
  frmRcvContact.Show
End Sub

'#######################################################################
'     F O R M    I N I T I A L I Z A T I O N
'#######################################################################
Private Sub Form_Load()
  Set ImgList = ImgListF
  Set tvwContact = tvwContactF
  Set IcqUdp = IcqUdpF
  Set IcqUtility = New clsIcqUtilities
  Set ContDB = OpenDatabase("Contact.mdb")
  Set ContRS = ContDB.OpenRecordset("Contact", dbOpenDynaset)
  Load frmDebug
  
  With ctlMenu
    .ImageList = ImgList
    .SubClassMenu Me
    
    'Menu 1 - File
    .ItemIcon("mnu1Connect") = ImgList.ListImages("IconConnect").Index - 1
    .ItemIcon("mnu1StOnline") = ImgList.ListImages("StOnline").Index - 1
    .ItemIcon("mnu1StAway") = ImgList.ListImages("StAway").Index - 1
    .ItemIcon("mnu1StNA") = ImgList.ListImages("StNA").Index - 1
    .ItemIcon("mnu1StOccupied") = ImgList.ListImages("StOccupied").Index - 1
    .ItemIcon("mnu1StDND") = ImgList.ListImages("StDND").Index - 1
    .ItemIcon("mnu1StChat") = ImgList.ListImages("StChat").Index - 1
    .ItemIcon("mnu1StInvisible") = ImgList.ListImages("StInvisible").Index - 1
    
    'Menu 2 - Contact
    .ItemIcon("mnu2SndMsg") = ImgList.ListImages("SndMsg").Index - 1
    .ItemIcon("mnu2SndURL") = ImgList.ListImages("SndURL").Index - 1
    .ItemIcon("mnu2SndContact") = ImgList.ListImages("SndContact").Index - 1
    .ItemIcon("mnu2UserInfo") = ImgList.ListImages("IconUserInfo").Index - 1
  End With
  
  With tvwContact
    Set NodeOnline = .Nodes.Add(, , , "Online", "StOnline")
    Set NodeOffline = .Nodes.Add(, , , "Offline", "StOffline")
    NodeOnline.Expanded = True
    NodeOffline.Expanded = True
  End With
  
  StatusBar.Panels(1).Picture = ImgList.ListImages("StOffline").ExtractIcon
  Initialize
End Sub
Private Sub Form_Resize()
  StatusBar.Refresh
  tvwContact.Width = frmMain.Width - 100
  If StatusBar.Top > 0 Then tvwContact.Height = StatusBar.Top
End Sub



'#######################################################################
'     I C Q   C O N T R O L   E V E N T S
'#######################################################################
Private Sub IcqUdpF_Connected()
  StatusBar_ChangeState IcqUdp.OnlineState
  Connect_SendContact
  'Connect_SendInvisible
  'Connect_SendVisible
End Sub

Private Sub IcqUdpF_ContactOffline(uin As Long)
  Cont_Change uin, icqOffline
End Sub

Private Sub IcqUdpF_ContactOnline(uin As Long, OnlineState As IcqUdpCtl.enumOnlineState, IntIP As String, ExtIP As String, ExtPort As Integer, bTcpCapable As Boolean, TcpVersion As Long)
  Cont_Change uin, OnlineState
End Sub

Private Sub IcqUdpF_ContactStatusChange(uin As Long, state As IcqUdpCtl.enumOnlineState)
  Cont_Change uin, state
End Sub

Private Sub IcqUdpF_DebugOut(DebugTxt As String)
  frmDebug.DebugOut DebugTxt
End Sub

Private Sub IcqUdpF_Disconnected()
  StatusBar_ChangeState icqOffline
End Sub

Private Sub IcqUdpF_ErrorFound(Number As IcqUdpCtl.enumErrorConstant, Description As String)
  Debug.Print "** Error " + Str$(Number)
End Sub

Private Sub IcqUdpF_InfoReply(InfoType As IcqUdpCtl.enumInfoType, Info As IcqUdpCtl.typContactInfo)
  Dim TempFrmIndex As Byte
  UserInfoSet InfoType, Info
  If Info.lngUIN = IcqUdp.UserUIN Then frmUserInfoUpdate.EventGetInfo CInt(InfoType): Exit Sub
  TempFrmIndex = frmGetIndex(enumfrmUserDetail, Info.lngUIN)
  If TempFrmIndex > 0 Then frmUserDetail(TempFrmIndex).EventGetInfo CInt(InfoType)
End Sub

Private Sub IcqUdpF_MessageReceived(uin As Long, MsgDate As Date, MsgTime As String, MsgType As IcqUdpCtl.enumMessageType, MsgText As String, URLAddress As String, URLDescription As String, authNick As String, authFirst As String, authLast As String, authEmail As String, authReason As String, contNick As Variant, contUin As Variant)
  Select Case MsgType
    Case icqMsgText
      TempFrmIndex = frmGetIndex(enumFrmSndMessage, uin)
      
      If TempFrmIndex = 0 Then
        TempFrmIndex = frmLoaded(enumFrmSndMessage)
        If TempFrmIndex = 0 Then
          MsgBox "There is not enough memory to open more window, please close some of your unused windows.", vbCritical, "Error - Too many window"
          Exit Sub
        End If
      
        Set frmMessage(TempFrmIndex) = New frmSndMessage
        With frmMessage(TempFrmIndex)
          .UserUIN = uin
          .FrmIndex = TempFrmIndex
          .Caption = Uin2Name(uin) & " Message"
        End With
      End If
      
      frmMessage(TempFrmIndex).Show
      frmMessage(TempFrmIndex).EventRecvMsg Uin2Name(uin), MsgText
      
    Case icqMsgUrl
      TempFrmIndex = frmGetIndex(enumFrmSndMessage, uin)
      
      If TempFrmIndex = 0 Then
        TempFrmIndex = frmLoaded(enumFrmSndMessage)
        If TempFrmIndex = 0 Then
          MsgBox "There is not enough memory to open more window, please close some of your unused windows.", vbCritical, "Error - Too many window"
          Exit Sub
        End If
      
        Set frmMessage(TempFrmIndex) = New frmSndMessage
        With frmMessage(TempFrmIndex)
          .UserUIN = uin
          .FrmIndex = TempFrmIndex
          .Caption = Uin2Name(uin) & " Message"
        End With
      End If
      
      frmMessage(TempFrmIndex).Show
      frmMessage(TempFrmIndex).EventRecvURL Uin2Name(uin), URLAddress, URLDescription
    Case icqMsgContact
      Load frmRcvContact
      frmRcvContact.EventRcvContact contUin, contNick
      
    Case icqMsgAuthReq
      Dim MsgBoxResult As VbMsgBoxResult
      MsgBoxResult = MsgBox("User " & authNick & "(#" & Trim$(Str$(uin)) & ") have requested your authorization to add you to his/her list. Would you authorize him/her?", vbYesNoCancel, "Authorization Request")
      Select Case MsgBoxResult
        Case vbYes: IcqUdp.sendauthaccept uin
        Case vbNo: IcqUdp.sendauthdecline uin, ""
      End Select
      
    Case icqMsgExpress
      Load frmRcvMessage
      frmRcvMessage.Show
      frmRcvMessage.EventRecvOther "Exress Message from " & authNick & "(" & authFirst & " " & authLast & "). Email " & authEmail, MsgText
    
    Case icqMsgWebpager
      Load frmRcvMessage
      frmRcvMessage.Show
      frmRcvMessage.EventRecvOther "WebPager Message from " & authNick & "(" & authFirst & " " & authLast & "). Email " & authEmail, MsgText
  End Select
End Sub

Private Sub IcqUdpF_PacketAcknowledge(PacketSeq As Integer)
    CheckAck PacketSeq
End Sub

Private Sub IcqUdpF_SearchReply(uin As Long, Nick As String, First As String, Last As String, Email As String, bAuth As Boolean, SearchResult As IcqUdpCtl.enumSearchResult)
  frmAddUser.Event_SearchReply uin, Nick, First, Last, Email, bAuth, SearchResult
End Sub

'#######################################################################
'     M E N U   A N D   S T A T U S B A R
'#######################################################################
Private Sub mnu1Connect_Click()
  If mnu1Connect.Caption = "&Connect" Then
    UDPConnect
  Else
    UDPDisconnect
  End If
End Sub

Private Sub mnu1EditInfo_Click()
  Load frmUserInfoUpdate
  frmUserInfoUpdate.Show
End Sub

Private Sub mnu1Exit_Click()
  If MsgBox("Are you sure you wanted to quite Communicator 2001?", vbYesNo, "Quit") = vbYes Then
    If IcqUdp.SocketState <> icqDisconnected Then UDPDisconnect
    End
  End If
End Sub

Private Sub mnu1Option_Click()
  Load frmOption
End Sub

Private Sub mnu1StAway_Click()
  IcqUdp.OnlineState = icqAway
  If IcqUdp.SocketState = Disconnected Then
    UDPConnect
  ElseIf IcqUdp.SocketState = icqConnected Then
    StatusBar_ChangeState icqAway
  End If
End Sub

Private Sub mnu1StDND_Click()
  IcqUdp.OnlineState = icqDND
  If IcqUdp.SocketState = Disconnected Then
    UDPConnect
  ElseIf IcqUdp.SocketState = icqConnected Then
    StatusBar_ChangeState icqDND
  End If
End Sub

Private Sub mnu1StInvisible_Click()
  IcqUdp.OnlineState = icqInvisible
  If IcqUdp.SocketState = Disconnected Then
    UDPConnect
  ElseIf IcqUdp.SocketState = icqConnected Then
    StatusBar_ChangeState icqInvisible
  End If
End Sub

Private Sub mnu1StNA_Click()
  IcqUdp.OnlineState = icqNa
  If IcqUdp.SocketState = Disconnected Then
    UDPConnect
  ElseIf IcqUdp.SocketState = icqConnected Then
    StatusBar_ChangeState icqNa
  End If
End Sub

Private Sub mnu1StOccupied_Click()
  IcqUdp.OnlineState = icqOccupied
  If IcqUdp.SocketState = Disconnected Then
    UDPConnect
  ElseIf IcqUdp.SocketState = icqConnected Then
    StatusBar_ChangeState icqOccupied
  End If
End Sub

Private Sub mnu1StOnline_Click()
  IcqUdp.OnlineState = icqOnline
  If IcqUdp.SocketState = Disconnected Then
    UDPConnect
  ElseIf IcqUdp.SocketState = icqConnected Then
    StatusBar_ChangeState icqOnline
  End If
End Sub

Private Sub mnu1StChat_Click()
  IcqUdp.OnlineState = icqChat
  If IcqUdp.SocketState = Disconnected Then
    UDPConnect
  ElseIf IcqUdp.SocketState = icqConnected Then
    StatusBar_ChangeState icqChat
  End If
End Sub

Private Sub mnu2Add_Click()
  Load frmAddUser
  frmAddUser.Show
End Sub

Private Sub mnu2Remove_Click()
  Dim TempKey As String
  Dim TempFrmIndex As Byte
  TempKey = tvwContact.SelectedItem.Key
  
  If Left$(TempKey, 1) = "o" Then
    tvwContact.Nodes.Remove (TempKey)
    ContactDelUser CLng(Right$(TempKey, Len(TempKey) - 1))
  End If
End Sub

Private Sub mnu2SndContact_Click()
  Load frmSndContact
  Dim TempKey As String
  Dim TempUin As Long
  TempKey = tvwContact.SelectedItem.Key
  
  Load frmSndContact
  
  If Left$(TempKey, 1) = "o" Then
    TempUin = Val(Right$(TempKey, Len(TempKey) - 1))
    frmSndContact.SetContRecepient TempUin, tvwContact.SelectedItem.Text
  End If
End Sub

Private Sub mnu2SndMsg_Click()
  Dim TempKey As String
  Dim TempFrmIndex As Byte
  TempKey = tvwContact.SelectedItem.Key
  
  If Left$(TempKey, 1) = "o" Then
    TempKey = Right$(TempKey, Len(TempKey) - 1)
    TempFrmIndex = frmGetIndex(enumFrmSndMessage, Val(TempKey))
    
    If TempFrmIndex = 0 Then
      TempFrmIndex = frmLoaded(enumFrmSndMessage)
      If TempFrmIndex = 0 Then
        MsgBox "There is not enough memory to open more window, please close some of your unused windows.", vbCritical, "Error - Too many window"
        Exit Sub
      End If
      
      Set frmMessage(TempFrmIndex) = New frmSndMessage
      With frmMessage(TempFrmIndex)
        .Show
        .UserUIN = Val(TempKey)
        .FrmIndex = TempFrmIndex
        .Caption = tvwContact.SelectedItem.Text & " Message"
      End With
    Else
      frmMessage(TempFrmIndex).Show
    End If
  End If
End Sub

Private Sub mnu2SndURL_Click()
  Dim TempKey As String
  Dim TempUin As Long
  TempKey = tvwContact.SelectedItem.Key
  
  Load frmSndURL
  
  If Left$(TempKey, 1) = "o" Then
    TempUin = Val(Right$(TempKey, Len(TempKey) - 1))
    frmSndURL.SetUrlRecepient TempUin, tvwContact.SelectedItem.Text
  End If
End Sub

Private Sub mnu2UserInfo_Click()
  Dim TempKey As String
  Dim TempFrmIndex As Byte
  TempKey = tvwContact.SelectedItem.Key
  If Left$(TempKey, 1) = "o" Then
    TempKey = Right$(TempKey, Len(TempKey) - 1)
    
    TempUINBuffer = Val(TempKey)
    TempFrmIndex = frmGetIndex(enumfrmUserDetail, TempUINBuffer)
    
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
        .Caption = "User Detail of " & tvwContact.SelectedItem.Text
      End With
    Else
      frmUserDetail(TempFrmIndex).Show
    End If
  End If
End Sub

Private Sub mnu3About_Click()
  Load frmAbout
End Sub

Private Sub StatusBar_PanelClick(ByVal Panel As MSComctlLib.Panel)
  If Panel = StatusBar.Panels(1) Then
    PopupMenu mnu1Status
  End If
End Sub

Public Sub StatusBar_ChangeState(state As Long)
  Dim StateName As String, StateVarName As String
  Select Case state
    Case icqOnline: StateName = "Online": StateVarName = "Online"
    Case icqAway: StateName = "Away": StateVarName = "Away"
    Case icqNa: StateName = "Extended Away": StateVarName = "NA"
    Case icqOccupied: StateName = "Occupied": StateVarName = "Occupied"
    Case icqDND: StateName = "Do Not Disturb": StateVarName = "DND"
    Case icqChat: StateName = "Free for Chat": StateVarName = "Chat"
    Case icqInvisible: StateName = "Invisible": StateVarName = "Invisible"
    Case Else       'Offline
      StateName = "Offline": StateVarName = "Offline"
  End Select
  StatusBar.Panels(1).Text = StateName
  StatusBar.Panels(1).Picture = ImgList.ListImages("St" + StateVarName).ExtractIcon
End Sub

Private Sub tvwContactF_NodeClick(ByVal Node As MSComctlLib.Node)
  PopupMenu mnu2
End Sub
