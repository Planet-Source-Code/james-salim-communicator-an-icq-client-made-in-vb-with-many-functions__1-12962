VERSION 5.00
Object = "{EF59A10B-9BC4-11D3-8E24-44910FC10000}#10.0#0"; "vbalEdit.ocx"
Begin VB.Form frmAbout 
   Caption         =   "About Communicator"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin vbalEdit.vbalRichEdit txtAbout 
      Height          =   1590
      Left            =   525
      TabIndex        =   2
      Top             =   2100
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   2805
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ViewMode        =   0
      Border          =   0   'False
      ScrollBars      =   3
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   1680
      TabIndex        =   0
      Top             =   3675
      Width           =   1065
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1860
      Left            =   840
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   1800
      ScaleWidth      =   2880
      TabIndex        =   1
      Top             =   105
      Width           =   2940
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Show
  txtAbout.ReadOnly = True
  With txtAbout
    .ParagraphAlignment = ercParaCentre
    .InsertContents SF_TEXT, _
      "Communicator version 0.5.0" & vbCrLf & _
      "Coded by James Salim" & vbCrLf & vbCrLf & _
      "Comments, Suggestion and/or Question" & vbCrLf & _
      "direct it to jamessalim@optushome.com.au" & vbCrLf & _
      "or visit http://www.medievilz.com"
  End With
End Sub
