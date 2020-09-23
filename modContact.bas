Attribute VB_Name = "modTvwContact"
Option Explicit

Public Sub Cont_Init()
  Dim TempDisplayName As String
  Dim i  As Integer
  
  For i = 0 To ContactTotal
    With Contact(i)
      Cont_Add .uin, .DisplayName, i
    End With
  Next i
End Sub

Public Sub Cont_Add(uin As Long, DisplayName As String, Optional ContactPos As Integer = 0)
  Dim TempDisplayName As String
  
  If DisplayName = "" Then
    TempDisplayName = Trim$(Str$(uin))
  Else
    TempDisplayName = DisplayName
  End If
  
  Set Contact(ContactPos).NodePos = tvwContact.Nodes.Add( _
    NodeOffline, tvwChild, "o" & Trim$(Str$(uin)), TempDisplayName, "IconUser")
  Contact(ContactPos).NodePos.Tag = Trim$(Str$(ContactPos))
  Contact(ContactPos).OnlineState = icqOffline
End Sub

Public Sub Cont_Remove(uin As Long, Optional ContactPos As Integer = -1)
  If ContactPos = -1 Then
    tvwContact.Nodes.Remove ("o" + Trim$(Str$(uin)))
  Else
    tvwContact.Nodes.Remove ("o" + Trim$(Str$(Contact(ContactPos).uin)))
  End If
End Sub

Public Sub Cont_Change(uin As Long, state As Long)
  Dim TempImg As String
  Dim TempUIN As String
  Dim TempDisplayName As String
  Dim TempIndex As Long
  
  TempUIN = Trim$(Str$(uin))
  
  With tvwContact.Nodes("o" + TempUIN)
    If .Parent = NodeOnline Then
      '########################################
      If state = icqOffline Then
        TempDisplayName = .Text
        TempIndex = Val(.Tag)
        tvwContact.Nodes.Remove ("o" + TempUIN)
        Set Contact(TempIndex).NodePos = tvwContact.Nodes.Add(NodeOffline, tvwChild, "o" + TempUIN, TempDisplayName, "IconUser")
      Else
        Select Case state
          Case icqOnline: TempImg = "StOnline"
          Case icqAway: TempImg = "StAway"
          Case icqNa: TempImg = "StNA"
          Case icqOccupied: TempImg = "StOccupied"
          Case icqDND: TempImg = "StDND"
          Case icqChat: TempImg = "StChat"
          Case icqInvisible: TempImg = "StInvisible"
        End Select
        .Image = TempImg
      End If
      '########################################
    Else
      '########################################
      If state <> icqOffline Then
        TempDisplayName = .Text
        TempIndex = Val(.Tag)
        tvwContact.Nodes.Remove ("o" + TempUIN)
        
        Select Case state
          Case icqOnline: TempImg = "StOnline"
          Case icqAway: TempImg = "StAway"
          Case icqNa: TempImg = "StNA"
          Case icqOccupied: TempImg = "StOccupied"
          Case icqDND: TempImg = "StDND"
          Case icqChat: TempImg = "StChat"
          Case icqInvisible: TempImg = "StInvisible"
        End Select
        Set Contact(TempIndex).NodePos = tvwContact.Nodes.Add(NodeOnline, tvwChild, "o" + TempUIN, TempDisplayName, TempImg)
      End If
      '########################################
    End If
  End With
  Contact(TempIndex).OnlineState = state
End Sub
