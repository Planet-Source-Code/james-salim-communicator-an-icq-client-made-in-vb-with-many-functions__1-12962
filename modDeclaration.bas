Attribute VB_Name = "modDeclaration"
Public Type typAccountDetail
  uin As Long
  Password As String
  OnlineState As enumOnlineState
  Index As Long
  NewAccount As Boolean
  TCPLocalHost As String
  TCPLocalPort As Integer
  UDPRemoteHost As String
  UDPRemotePort As Integer
  UserInfo As typContactInfo
End Type

Public Type typContactDetail
  DisplayName As String
  uin As Long
  OnlineState As Integer
  NodePos As Node
  TCPIntHost As String
  TCPExtHost As String
  TCPPort As Integer
  bTcpCapable As Boolean
  TcpVersion As Long
End Type

Public Account As typAccountDetail
Public Contact(255) As typContactDetail
Public ContactTotal As Integer

Public NodeOnline As Node
Public NodeOffline As Node

Public Const icqOffline = &HEEEE

Public ImgList As ImageList
Public tvwContact As TreeView
Public IcqUdp As IcqUdp
Public IcqUtility As clsIcqUtilities

Public ContDB As Database
Public ContRS As Recordset
Public ContRSFilter As Recordset

Public TempUINBuffer As Long
