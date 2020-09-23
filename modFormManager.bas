Attribute VB_Name = "modFormManager"
Option Explicit
Public frmUserDetail(255) As frmUserInfo
Public frmMessage(255) As frmSndMessage

Public Enum TypFrmType
  enumFrmNull = 0
  enumfrmUserDetail = 1
  enumFrmSndMessage = 2
  enumFrmSndUrl = 3
  enumFrmSndContact = 4
End Enum

Public SlotfrmUserDetail As String * 255
Public SlotfrmMessage As String * 255

Public Function frmLoaded(FrmType As TypFrmType) As Byte
  Select Case FrmType
    Case enumfrmUserDetail
      frmLoaded = InStr(1, SlotfrmUserDetail, Chr$(0))
      If frmLoaded > 0 Then Mid$(SlotfrmUserDetail, frmLoaded) = Chr$(1)
    Case enumFrmSndMessage
      frmLoaded = InStr(1, SlotfrmMessage, Chr$(0))
      If frmLoaded > 0 Then Mid$(SlotfrmMessage, frmLoaded) = Chr$(1)
  End Select
End Function

Public Sub frmUnloaded(FrmType As TypFrmType, FrmIndex As Byte)
  Select Case FrmType
    Case enumfrmUserDetail:      Mid$(SlotfrmUserDetail, FrmIndex) = Chr$(0)
    Case enumFrmSndMessage:      Mid$(SlotfrmMessage, FrmIndex) = Chr$(0)
  End Select
End Sub

Public Function frmGetIndex(FrmType As TypFrmType, lngUIN As Long) As Byte
  Dim i As Byte
  i = 0
  
  Select Case FrmType
    Case enumfrmUserDetail
      Do
        i = i + 1
        i = InStr(i, SlotfrmUserDetail, Chr$(1))
        If i = 0 Then Exit Do
        If frmUserDetail(i).UserUIN = lngUIN Then
          frmGetIndex = i
          Exit Function
        End If
      Loop Until i = 255 Or i = 0
    Case enumFrmSndMessage
      Do
        i = i + 1
        i = InStr(i, SlotfrmMessage, Chr$(1))
        If i = 0 Then Exit Do
        If frmMessage(i).UserUIN = lngUIN Then
          frmGetIndex = i
          Exit Function
        End If
      Loop Until i = 255 Or i = 0
  End Select
End Function
