Attribute VB_Name = "modInitialize"
Option Explicit

Sub Initialize()
  Dim TempUin As Long
  Dim TempValue As String
  Dim i As Integer
  
  With Account
    TempValue = GetINI("Comm2001.ini", "Connect", "UIN")
    If TempValue = "" Then .uin = 0 Else .uin = CLng(TempValue)
    .Password = GetINI("Comm2001.ini", "Connect", "Password")
  End With
  
  If Account.uin = 0 Then
    Account.uin = Val(InputBox("Please enter the User Identification Number (UIN) you would like to use.", "Enter UIN"))
    Account.Password = InputBox("Please enter the password for UIN" & Trim$(Str$(Account.uin)), "Enter Password")
    WriteINI "Comm2001.ini", "Connect", "UIN", Trim$(Str$(Account.uin))
    WriteINI "Comm2001.ini", "Connect", "Password", Account.Password
  End If
  
  IcqUdp.RemoteHost = GetINI("Comm2001.ini", "Connect", "RemoteHost")
  TempValue = GetINI("Comm2001.ini", "Connect", "RemotePort", 5)
  If TempValue = "" Then IcqUdp.RemotePort = 0 Else IcqUdp.RemotePort = CLng(TempValue)
  
  If IcqUdp.RemoteHost = "" And IcqUdp.RemotePort = 0 Then
    WriteINI "Comm2001.ini", "Connect", "RemoteHost", "icq.mirabilis.com"
    WriteINI "Comm2001.ini", "Connect", "RemotePort", "4000"
  End If
  
  IcqUdp.UserUIN = Account.uin
  IcqUdp.UserPassword = Account.Password
    
  ContRS.Filter = "bOnContact = TRUE"
  Set ContRSFilter = ContRS.OpenRecordset
  
  
  With ContRSFilter
    If .RecordCount = 0 Then GoTo NoRecordFound
    .Edit
    ContactTotal = .RecordCount - 1
    .MoveFirst
    Do While .EOF = False
      Contact(i).uin = !uin
      Contact(i).OnlineState = icqOffline
      Contact(i).DisplayName = !Nickname
      i = i + 1
      .MoveNext
    Loop
  End With
  
  Cont_Init
  
NoRecordFound:
  If GetINI("Comm2001.ini", "Connect", "AutoConnect") = "1" Then UDPConnect
End Sub
