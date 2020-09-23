VERSION 5.00
Object = "{EF59A10B-9BC4-11D3-8E24-44910FC10000}#10.0#0"; "vbalEdit.ocx"
Begin VB.Form frmRcvMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Receive Message"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   396
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "cmdClose"
      Height          =   330
      Left            =   4830
      TabIndex        =   2
      Top             =   4410
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   4425
      Left            =   15
      TabIndex        =   0
      Top             =   -60
      Width           =   5895
      Begin vbalEdit.vbalRichEdit txtRcvMsg 
         Height          =   4270
         Left            =   45
         TabIndex        =   1
         Top             =   105
         Width           =   5800
         _ExtentX        =   10239
         _ExtentY        =   7541
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
         AutoURLDetect   =   0   'False
         ScrollBars      =   0
      End
   End
End
Attribute VB_Name = "frmRcvMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  txtRcvMsg.ReadOnly = True
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

