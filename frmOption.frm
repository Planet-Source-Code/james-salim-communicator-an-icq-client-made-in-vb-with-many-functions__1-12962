VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Communicator Options"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   7995
   Begin VB.Frame frmOpt1 
      BorderStyle     =   0  'None
      Height          =   5790
      Left            =   2625
      TabIndex        =   8
      Top             =   945
      Width           =   5265
      Begin VB.TextBox txtRemoteHost 
         Height          =   285
         Left            =   2100
         TabIndex        =   15
         Top             =   2310
         Width           =   2325
      End
      Begin VB.TextBox txtRemotePort 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2100
         MaxLength       =   8
         TabIndex        =   14
         Top             =   2730
         Width           =   2325
      End
      Begin VB.CheckBox chkAutoConnect 
         Caption         =   "Connect &Automatically"
         Height          =   225
         Left            =   2100
         TabIndex        =   13
         Top             =   1575
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2100
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   1155
         Width           =   2325
      End
      Begin VB.TextBox txtUIN 
         Height          =   285
         Left            =   2100
         MaxLength       =   11
         TabIndex        =   10
         Top             =   735
         Width           =   2325
      End
      Begin VB.Label lblRemoteHost 
         Alignment       =   1  'Right Justify
         Caption         =   "Server &Address:"
         Height          =   225
         Left            =   630
         TabIndex        =   17
         Top             =   2310
         Width           =   1170
      End
      Begin VB.Label lblRemotePort 
         Alignment       =   1  'Right Justify
         Caption         =   "Server &Port:"
         Height          =   225
         Left            =   630
         TabIndex        =   16
         Top             =   2730
         Width           =   1170
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "&Password:"
         Height          =   225
         Left            =   630
         TabIndex        =   11
         Top             =   1155
         Width           =   1170
      End
      Begin VB.Label lblUIN 
         Alignment       =   1  'Right Justify
         Caption         =   "User &UIN:"
         Height          =   225
         Left            =   630
         TabIndex        =   9
         Top             =   735
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   6720
      TabIndex        =   6
      Top             =   7140
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   5565
      TabIndex        =   5
      Top             =   7140
      Width           =   1170
   End
   Begin VB.PictureBox BoxHeader 
      BackColor       =   &H00FFFFFF&
      Height          =   960
      Left            =   -100
      ScaleHeight     =   900
      ScaleWidth      =   8145
      TabIndex        =   0
      Top             =   -100
      Width           =   8200
      Begin VB.Label lblblHeadDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose what you would like to set from the category on the left."
         Height          =   330
         Left            =   735
         TabIndex        =   2
         Top             =   525
         Width           =   7050
      End
      Begin VB.Label lblHeader 
         BackStyle       =   0  'Transparent
         Caption         =   "Set Options and Preferences"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   420
         TabIndex        =   1
         Top             =   210
         Width           =   5265
      End
   End
   Begin VB.Frame frmTvwHolder 
      Height          =   5805
      Left            =   105
      TabIndex        =   3
      Top             =   945
      Width           =   2430
      Begin MSComctlLib.TreeView tvwCategory 
         Height          =   5670
         Left            =   10
         TabIndex        =   7
         Top             =   100
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   10001
         _Version        =   393217
         Indentation     =   529
         LineStyle       =   1
         Style           =   7
         Appearance      =   0
      End
   End
   Begin VB.Frame frmLineSeperator 
      Height          =   115
      Left            =   0
      TabIndex        =   4
      Top             =   6825
      Width           =   7995
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  WriteINI "Comm2001.ini", "Connect", "UIN", txtUin.Text
  WriteINI "Comm2001.ini", "Connect", "Password", txtPassword.Text
  WriteINI "Comm2001.ini", "Connect", "RemoteHost", txtRemoteHost.Text
  WriteINI "Comm2001.ini", "Connect", "RemotePort", txtRemotePort.Text

  
  If chkAutoConnect.Value = Checked Then
    WriteINI "Comm2001.ini", "Connect", "AutoConnect", "1"
  Else
    WriteINI "Comm2001.ini", "Connect", "AutoConnect", "0"
  End If

  Unload Me
End Sub

Private Sub Form_Load()
  Me.Show
  txtUin = GetINI("Comm2001.ini", "Connect", "UIN")
  txtPassword = GetINI("Comm2001.ini", "Connect", "Password")
  txtRemoteHost = GetINI("Comm2001.ini", "Connect", "RemoteHost")
  txtRemotePort = GetINI("Comm2001.ini", "Connect", "RemotePort")
  If GetINI("Comm2001.ini", "Connect", "AutoConnect") = "1" Then
    chkAutoConnect.Value = Checked
  Else
    chkAutoConnect.Value = Unchecked
  End If
End Sub
