VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EF59A10B-9BC4-11D3-8E24-44910FC10000}#10.0#0"; "vbalEdit.ocx"
Begin VB.Form frmDebug 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DebugWindow"
   ClientHeight    =   3675
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin vbalEdit.vbalRichEdit txtDebug 
      Height          =   3690
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   6509
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
      BackColor       =   0
      ForeColor       =   16777215
      ViewMode        =   0
      Border          =   0   'False
      AutoURLDetect   =   0   'False
      TextOnly        =   -1  'True
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   120
      Left            =   0
      TabIndex        =   0
      Top             =   3555
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   212
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu1File 
      Caption         =   "&File"
      Begin VB.Menu mnu1Save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu1Hide 
         Caption         =   "&Hide Window"
      End
   End
   Begin VB.Menu mnu2Debug 
      Caption         =   "&Debug"
      Begin VB.Menu mnu2SelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnu2SelectNone 
         Caption         =   "Select &None"
      End
      Begin VB.Menu Seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu2Cut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnu2Copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu2Clear 
         Caption         =   "C&lear"
      End
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Me.Show
  txtDebug.FontColour = vbWhite
End Sub


Public Sub DebugOut(DebugTxt As String)
  Dim IStart As Long, IEnd As Long
  With txtDebug
    .GetSelection IStart, IEnd
    .SetSelection -1, -1
    
    .InsertContents SF_TEXT, DebugTxt & vbCrLf
    
    '.SetSelection IStart, IEnd
  End With
End Sub

Private Sub mnu1Hide_Click()
  Me.Hide
End Sub

Private Sub mnu1Save_Click()
  txtDebug.SaveToFile "Debug.log", SF_TEXT
End Sub

Private Sub mnu2Clear_Click()
  txtDebug.Text = ""
End Sub

Private Sub mnu2Copy_Click()
  txtDebug.Copy
End Sub

Private Sub mnu2Cut_Click()
  txtDebug.Cut
End Sub

Private Sub mnu2SelectAll_Click()
  txtDebug.SelectAll
End Sub

Private Sub mnu2SelectNone_Click()
  txtDebug.SelectNone
End Sub
