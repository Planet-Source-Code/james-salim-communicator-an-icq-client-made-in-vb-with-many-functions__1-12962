VERSION 5.00
Object = "{EF59A10B-9BC4-11D3-8E24-44910FC10000}#10.0#0"; "vbalEdit.ocx"
Begin VB.Form frmSndMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send Message"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   467
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   416
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmRcvMsg 
      Height          =   4845
      Left            =   15
      TabIndex        =   5
      Top             =   -60
      Width           =   6210
      Begin vbalEdit.vbalRichEdit TxtRcvMsg 
         Height          =   4695
         Left            =   15
         TabIndex        =   4
         Top             =   105
         Width           =   6150
         _ExtentX        =   10848
         _ExtentY        =   8281
         Version         =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         ViewMode        =   0
         Border          =   0   'False
         ScrollBars      =   3
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   330
      Left            =   3990
      TabIndex        =   3
      Top             =   6615
      Width           =   1065
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   330
      Left            =   5040
      TabIndex        =   1
      Top             =   6615
      Width           =   1170
   End
   Begin VB.Frame frmSndMsg 
      Height          =   1800
      Left            =   15
      TabIndex        =   2
      Top             =   4725
      Width           =   6210
      Begin vbalEdit.vbalRichEdit txtSndMsg 
         Height          =   1650
         Left            =   15
         TabIndex        =   0
         Top             =   105
         Width           =   6150
         _ExtentX        =   10848
         _ExtentY        =   2910
         Version         =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         ViewMode        =   0
         Border          =   0   'False
         TextLimit       =   450
         TextOnly        =   -1  'True
         ScrollBars      =   3
      End
   End
End
Attribute VB_Name = "frmSndMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public UserUIN As Long
Public FrmIndex As Byte

Private Sub cmdClose_Click()
  frmUnloaded enumFrmSndMessage, FrmIndex
  Unload Me
End Sub

Private Sub cmdSend_Click()
  If IcqUdp.SocketState <> icqConnected Then
    MsgBox "Communicator is not currently connected to the ICQ Server. Please connect by choosing File->Connect on the file menu", vbOKOnly, "Not Connected"
    Exit Sub
  End If
  
  cmdSend.Enabled = False
  WriteTxt txtSndMsg.Text
  SetAck IcqUdp.sendtext(UserUIN, txtSndMsg.Text), enumFrmSndMessage, FrmIndex
  
  txtSndMsg.Text = ""
End Sub

Private Sub Form_Load()
  txtRcvMsg.ReadOnly = True
End Sub

Public Sub WriteTxt(Message As String)
  Dim IStart As Long, IEnd As Long
  With txtRcvMsg
    .GetSelection IStart, IEnd
    .SetSelection -1, -1
    .FontColour = vbBlue
    .FontItalic = True
    .InsertContents SF_TEXT, "You says" & vbCrLf
    
    .FontColour = vbBlack
    .FontItalic = False
    .InsertContents SF_TEXT, Message & vbCrLf
    .SetSelection IStart, IEnd
  End With
End Sub

Public Sub EventRecvMsg(strname As String, Message As String)
  Dim IStart As Long, IEnd As Long
  With txtRcvMsg
    .GetSelection IStart, IEnd
    .SetSelection -1, -1
    
    .FontColour = vbBlue
    .FontItalic = True
    .InsertContents SF_TEXT, strname & " says" & vbCrLf
    
    .FontColour = vbBlack
    .FontItalic = False
    .InsertContents SF_TEXT, Message & vbCrLf
    
    .SetSelection IStart, IEnd
  End With
End Sub

Public Sub EventRecvURL(strname As String, URLAddress As String, URLDescription As String)
  Dim IStart As Long, IEnd As Long
  With txtRcvMsg
    .GetSelection IStart, IEnd
    .SetSelection -1, -1
    
    .FontColour = vbBlue
    .FontItalic = True
    .InsertContents SF_TEXT, strname & " send a URL" & vbCrLf
    
    .FontColour = vbBlack
    .FontItalic = False
    .InsertContents SF_TEXT, "URL Address : " & URLAddress & vbCrLf
    .InsertContents SF_TEXT, "URL Description : " & URLDescription & vbCrLf
    
    .SetSelection IStart, IEnd
  End With
End Sub

Public Sub EventRecvOther(HeadText As String, BodyText As String)
  Dim IStart As Long, IEnd As Long
  With txtRcvMsg
    .GetSelection IStart, IEnd
    .SetSelection -1, -1
    
    .FontColour = vbBlue
    .FontItalic = True
    .InsertContents SF_TEXT, HeadText & vbCrLf
    
    .FontColour = vbBlack
    .FontItalic = False
    .InsertContents SF_TEXT, BodyText & vbCrLf
    
    .SetSelection IStart, IEnd
  End With
End Sub

Public Sub EventMsgRecvAck()
  cmdSend.Enabled = True
End Sub
