Attribute VB_Name = "modConnection"
Option Explicit

Sub UDPConnect()
  With frmMain
    With Account
    .uin = GetINI("Comm2001.ini", "Connect", "UIN")
    .Password = GetINI("Comm2001.ini", "Connect", "Password")
    End With
  
    IcqUdp.RemoteHost = GetINI("Comm2001.ini", "Connect", "RemoteHost")
    IcqUdp.RemotePort = Val(GetINI("Comm2001.ini", "Connect", "RemotePort", 5))
    IcqUdp.UserUIN = Account.uin
    IcqUdp.UserPassword = Account.Password
  
    frmMain.Caption = Trim$(Str$(Account.uin)) & " - Communicator"
    
    IcqUdp.Connect
    .mnu1Connect.Caption = "&Disconnect"
    .ctlMenu.Caption("mnu1Connect") = "&Disconnect"
    .ctlMenu.ItemIcon("mnu1Connect") = ImgList.ListImages("IconDisconnect").Index - 1
    .StatusBar.Panels(1).Text = "Connecting..."
    


  End With
End Sub

Sub UDPDisconnect()
  Dim TempNode As Node
  Dim TempKey As String
  Dim i As Integer

  With frmMain
    IcqUdp.Disconnect
    .mnu1Connect.Caption = "&Connect"
    .ctlMenu.Caption("mnu1Connect") = "&Connect"
    .ctlMenu.ItemIcon("mnu1Connect") = ImgList.ListImages("IconConnect").Index - 1
  End With
  
  Do While NodeOnline.Children > 0
    Cont_Change Val(Mid$(NodeOnline.Child.Key, 2, 20)), icqOffline
  Loop
End Sub

Sub Connect_SendContact()
  Dim uinlist As Variant
  Dim i As Integer
  
  ReDim uinlist(ContactTotal) As Long
  
  For i = 0 To ContactTotal
    uinlist(i) = Contact(i).uin
  Next i
  
  IcqUdp.ContactAdd uinlist
End Sub
