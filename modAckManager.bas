Attribute VB_Name = "modAckManager"
Public Type typNotifyAck
  FrmType As TypFrmType
  FrmIndex As Byte
End Type

Public SlotNotifyAck(255) As typNotifyAck

Public Sub SetAck(Ack As Integer, FrmType As TypFrmType, FrmIndex As Byte)
  With SlotNotifyAck(Ack Mod 256)
    .FrmType = FrmType
    .FrmIndex = FrmIndex
  End With
End Sub

Public Sub CheckAck(Ack As Integer)
  With SlotNotifyAck(Ack Mod 256)
    If .FrmType = enumFrmNull Then Exit Sub
    If .FrmType = enumFrmSndMessage Then frmMessage(.FrmIndex).EventMsgRecvAck
    If .FrmType = enumFrmSndUrl Then frmSndURL.EventUrlRecvAck
    If .FrmType = enumFrmSndContact Then frmSndContact.EventContRecvAck
    
    .FrmType = enumFrmNull
    .FrmIndex = 0
  End With
End Sub
