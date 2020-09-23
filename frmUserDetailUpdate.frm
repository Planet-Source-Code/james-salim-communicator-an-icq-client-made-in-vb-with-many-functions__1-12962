VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUserInfoUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Detail"
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
   Begin VB.Frame frmOpt6 
      BorderStyle     =   0  'None
      Caption         =   "!Birthdate"
      Height          =   5790
      Left            =   2625
      TabIndex        =   108
      Top             =   945
      Visible         =   0   'False
      Width           =   5265
      Begin VB.TextBox txtInterestCategory 
         Height          =   330
         Index           =   3
         Left            =   210
         MaxLength       =   5
         TabIndex        =   76
         Top             =   1680
         Width           =   1275
      End
      Begin VB.TextBox txtInterestCategory 
         Height          =   330
         Index           =   2
         Left            =   210
         MaxLength       =   5
         TabIndex        =   74
         Top             =   1260
         Width           =   1275
      End
      Begin VB.TextBox txtInterestCategory 
         Height          =   330
         Index           =   1
         Left            =   210
         MaxLength       =   5
         TabIndex        =   72
         Top             =   840
         Width           =   1275
      End
      Begin VB.TextBox txtInterestCategory 
         Height          =   330
         Index           =   0
         Left            =   210
         MaxLength       =   5
         TabIndex        =   70
         Top             =   420
         Width           =   1275
      End
      Begin VB.TextBox txtBackgroundName 
         Height          =   330
         Index           =   2
         Left            =   1575
         MaxLength       =   60
         TabIndex        =   91
         Top             =   5355
         Width           =   3585
      End
      Begin VB.TextBox txtBackgroundName 
         Height          =   330
         Index           =   1
         Left            =   1575
         MaxLength       =   60
         TabIndex        =   89
         Top             =   4935
         Width           =   3585
      End
      Begin VB.TextBox txtBackgroundName 
         Height          =   330
         Index           =   0
         Left            =   1575
         MaxLength       =   60
         TabIndex        =   87
         Top             =   4515
         Width           =   3585
      End
      Begin VB.TextBox txtGroupName 
         Height          =   330
         Index           =   2
         Left            =   1575
         MaxLength       =   60
         TabIndex        =   84
         Top             =   3465
         Width           =   3585
      End
      Begin VB.TextBox txtGroupName 
         Height          =   330
         Index           =   1
         Left            =   1575
         MaxLength       =   60
         TabIndex        =   82
         Top             =   3045
         Width           =   3585
      End
      Begin VB.TextBox txtGroupName 
         Height          =   330
         Index           =   0
         Left            =   1575
         MaxLength       =   60
         TabIndex        =   80
         Top             =   2625
         Width           =   3585
      End
      Begin VB.TextBox txtInterestName 
         Height          =   330
         Index           =   3
         Left            =   1575
         MaxLength       =   60
         TabIndex        =   77
         Top             =   1680
         Width           =   3585
      End
      Begin VB.TextBox txtInterestName 
         Height          =   330
         Index           =   2
         Left            =   1575
         MaxLength       =   60
         TabIndex        =   75
         Top             =   1260
         Width           =   3585
      End
      Begin VB.ComboBox cmbGroupCategory 
         Height          =   315
         Index           =   0
         Left            =   210
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   2625
         Width           =   1275
      End
      Begin VB.ComboBox cmbGroupCategory 
         Height          =   315
         Index           =   1
         Left            =   210
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   3045
         Width           =   1275
      End
      Begin VB.ComboBox cmbGroupCategory 
         Height          =   315
         Index           =   2
         Left            =   210
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   3465
         Width           =   1275
      End
      Begin VB.ComboBox cmbBackgroundCategory 
         Height          =   315
         Index           =   0
         Left            =   210
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   86
         Top             =   4515
         Width           =   1275
      End
      Begin VB.ComboBox cmbBackgroundCategory 
         Height          =   315
         Index           =   1
         Left            =   210
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   4935
         Width           =   1275
      End
      Begin VB.ComboBox cmbBackgroundCategory 
         Height          =   315
         Index           =   2
         Left            =   210
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   5355
         Width           =   1275
      End
      Begin VB.TextBox txtInterestName 
         Height          =   330
         Index           =   0
         Left            =   1575
         MaxLength       =   60
         TabIndex        =   71
         Top             =   420
         Width           =   3585
      End
      Begin VB.TextBox txtInterestName 
         Height          =   330
         Index           =   1
         Left            =   1575
         MaxLength       =   60
         TabIndex        =   73
         Top             =   840
         Width           =   3585
      End
      Begin VB.Label lblPastBackground 
         Caption         =   "&Past Background"
         Height          =   225
         Left            =   210
         TabIndex        =   85
         Top             =   4200
         Width           =   1485
      End
      Begin VB.Label lblInterest 
         Caption         =   "Personal &Interests"
         Height          =   225
         Left            =   210
         TabIndex        =   69
         Top             =   105
         Width           =   1485
      End
      Begin VB.Label lblGroup 
         Caption         =   "Organization, Affiliations, Groups"
         Height          =   225
         Left            =   210
         TabIndex        =   78
         Top             =   2310
         Width           =   1905
      End
   End
   Begin VB.Frame frmOpt5 
      BorderStyle     =   0  'None
      Height          =   5790
      Left            =   2625
      TabIndex        =   107
      Top             =   945
      Visible         =   0   'False
      Width           =   5265
      Begin VB.ComboBox cmbLanguage 
         Height          =   315
         Index           =   2
         Left            =   840
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   3150
         Width           =   3900
      End
      Begin VB.ComboBox cmbLanguage 
         Height          =   315
         Index           =   1
         Left            =   840
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   2730
         Width           =   3900
      End
      Begin VB.ComboBox cmbLanguage 
         Height          =   315
         Index           =   0
         Left            =   840
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   2310
         Width           =   3900
      End
      Begin VB.ComboBox cmbGender 
         Height          =   315
         ItemData        =   "frmUserDetailUpdate.frx":0000
         Left            =   1365
         List            =   "frmUserDetailUpdate.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   735
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker datBirthDate 
         Height          =   330
         Left            =   1365
         TabIndex        =   62
         Top             =   1575
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   23855104
         CurrentDate     =   36849
      End
      Begin VB.TextBox txtAboutInfo 
         Height          =   1590
         Left            =   735
         MaxLength       =   470
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   68
         Top             =   4095
         Width           =   4005
      End
      Begin VB.TextBox txtAge 
         Height          =   285
         Left            =   1365
         MaxLength       =   3
         TabIndex        =   60
         Top             =   1155
         Width           =   3375
      End
      Begin VB.TextBox txtHomeURL 
         Height          =   285
         Left            =   1365
         MaxLength       =   100
         TabIndex        =   56
         Top             =   315
         Width           =   3375
      End
      Begin VB.Label lblAboutInfo 
         Caption         =   "&About Info"
         Height          =   225
         Left            =   420
         TabIndex        =   67
         Top             =   3780
         Width           =   1170
      End
      Begin VB.Label lblLanguage 
         Caption         =   "&Language"
         Height          =   225
         Left            =   420
         TabIndex        =   63
         Top             =   2100
         Width           =   1485
      End
      Begin VB.Label lblAge 
         Caption         =   "&Age"
         Height          =   225
         Left            =   420
         TabIndex        =   59
         Top             =   1155
         Width           =   1485
      End
      Begin VB.Label lblBirthDate 
         Caption         =   "&Birth Date"
         Height          =   225
         Left            =   420
         TabIndex        =   61
         Top             =   1575
         Width           =   1065
      End
      Begin VB.Label lblHomepageURL 
         Caption         =   "&Home URL"
         Height          =   225
         Left            =   420
         TabIndex        =   55
         Top             =   315
         Width           =   1485
      End
      Begin VB.Label lblGender 
         Caption         =   "&Gender"
         Height          =   225
         Left            =   420
         TabIndex        =   57
         Top             =   735
         Width           =   1065
      End
   End
   Begin VB.Frame frmOpt4 
      BorderStyle     =   0  'None
      Height          =   5790
      Left            =   2625
      TabIndex        =   106
      Top             =   945
      Visible         =   0   'False
      Width           =   5265
      Begin VB.ComboBox cmbCompanyOccupation 
         Height          =   315
         Left            =   1365
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   1890
         Width           =   3375
      End
      Begin VB.TextBox txtCompanyURL 
         Height          =   285
         Left            =   1365
         MaxLength       =   100
         TabIndex        =   54
         Top             =   2310
         Width           =   3375
      End
      Begin VB.TextBox txtCompanyPosition 
         Height          =   285
         Left            =   1365
         MaxLength       =   30
         TabIndex        =   50
         Top             =   1470
         Width           =   3375
      End
      Begin VB.TextBox txtCompanyName 
         Height          =   285
         Left            =   1365
         MaxLength       =   30
         TabIndex        =   46
         Top             =   630
         Width           =   3375
      End
      Begin VB.TextBox txtCompanyDepartment 
         Height          =   285
         Left            =   1365
         MaxLength       =   40
         TabIndex        =   48
         Top             =   1050
         Width           =   3375
      End
      Begin VB.Label lblCompanyURL 
         Caption         =   "&Website"
         Height          =   225
         Left            =   420
         TabIndex        =   53
         Top             =   2310
         Width           =   1065
      End
      Begin VB.Label lblCompanyPosition 
         Caption         =   "&Position"
         Height          =   225
         Left            =   420
         TabIndex        =   49
         Top             =   1470
         Width           =   1485
      End
      Begin VB.Label lblCompanyOccupation 
         Caption         =   "&Occupation"
         Height          =   225
         Left            =   420
         TabIndex        =   51
         Top             =   1890
         Width           =   1065
      End
      Begin VB.Label lblCompanyName 
         Caption         =   "&Name"
         Height          =   225
         Left            =   420
         TabIndex        =   45
         Top             =   630
         Width           =   1485
      End
      Begin VB.Label lblCompanyDepartment 
         Caption         =   "&Department"
         Height          =   225
         Left            =   420
         TabIndex        =   47
         Top             =   1050
         Width           =   1065
      End
   End
   Begin VB.Frame frmOpt3 
      BorderStyle     =   0  'None
      Height          =   5790
      Left            =   2625
      TabIndex        =   104
      Top             =   945
      Visible         =   0   'False
      Width           =   5265
      Begin VB.ComboBox cmbWorkCountry 
         Height          =   315
         Left            =   1365
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   3045
         Width           =   3375
      End
      Begin VB.TextBox txtWorkState 
         Height          =   285
         Left            =   1365
         MaxLength       =   3
         TabIndex        =   36
         Top             =   2205
         Width           =   3375
      End
      Begin VB.TextBox txtWorkCity 
         Height          =   285
         Left            =   1365
         MaxLength       =   30
         TabIndex        =   34
         Top             =   1785
         Width           =   3375
      End
      Begin VB.TextBox txtWorkAddress 
         Height          =   1080
         Left            =   420
         MaxLength       =   30
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   525
         Width           =   4320
      End
      Begin VB.TextBox txtWorkZip 
         Height          =   285
         Left            =   1365
         MaxLength       =   10
         TabIndex        =   38
         Top             =   2625
         Width           =   3375
      End
      Begin VB.TextBox txtWorkFax 
         Height          =   285
         Left            =   1365
         MaxLength       =   30
         TabIndex        =   44
         Top             =   4935
         Width           =   3375
      End
      Begin VB.TextBox txtWorkPhone 
         Height          =   285
         Left            =   1365
         MaxLength       =   30
         TabIndex        =   42
         Top             =   4515
         Width           =   3375
      End
      Begin VB.Frame Frame2 
         Height          =   120
         Left            =   315
         TabIndex        =   105
         Top             =   3990
         Width           =   4530
      End
      Begin VB.Label lblWorkState 
         Caption         =   "&State"
         Height          =   225
         Left            =   420
         TabIndex        =   35
         Top             =   2205
         Width           =   1065
      End
      Begin VB.Label lblWorkCity 
         Caption         =   "&City"
         Height          =   225
         Left            =   420
         TabIndex        =   33
         Top             =   1785
         Width           =   1485
      End
      Begin VB.Label lblWorkAddress 
         Caption         =   "&Work Address"
         Height          =   225
         Left            =   420
         TabIndex        =   31
         Top             =   210
         Width           =   1380
      End
      Begin VB.Label lblWorkCountry 
         Caption         =   "&Country"
         Height          =   225
         Left            =   420
         TabIndex        =   39
         Top             =   3045
         Width           =   1065
      End
      Begin VB.Label lblWorkZip 
         Caption         =   "&ZIP"
         Height          =   225
         Left            =   420
         TabIndex        =   37
         Top             =   2625
         Width           =   1485
      End
      Begin VB.Label lblWorkFax 
         Caption         =   "&Facsimile"
         Height          =   225
         Left            =   420
         TabIndex        =   43
         Top             =   4935
         Width           =   1065
      End
      Begin VB.Label lblWorkPhone 
         Caption         =   "&Telephone"
         Height          =   225
         Left            =   420
         TabIndex        =   41
         Top             =   4515
         Width           =   1485
      End
   End
   Begin VB.Frame frmOpt2 
      BorderStyle     =   0  'None
      Height          =   5790
      Left            =   2625
      TabIndex        =   102
      Top             =   945
      Visible         =   0   'False
      Width           =   5265
      Begin VB.ComboBox cmbTimeZone 
         Height          =   315
         Left            =   1365
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3465
         Width           =   3375
      End
      Begin VB.ComboBox cmbHomeCountry 
         Height          =   315
         Left            =   1365
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3045
         Width           =   3375
      End
      Begin VB.Frame Seperator2 
         Height          =   120
         Left            =   315
         TabIndex        =   103
         Top             =   3990
         Width           =   4530
      End
      Begin VB.TextBox txtHomeCellular 
         Height          =   285
         Left            =   1365
         MaxLength       =   30
         TabIndex        =   30
         Top             =   5355
         Width           =   3375
      End
      Begin VB.TextBox txtHomePhone 
         Height          =   285
         Left            =   1365
         MaxLength       =   30
         TabIndex        =   26
         Top             =   4515
         Width           =   3375
      End
      Begin VB.TextBox txtHomeFax 
         Height          =   285
         Left            =   1365
         MaxLength       =   30
         TabIndex        =   28
         Top             =   4935
         Width           =   3375
      End
      Begin VB.TextBox txtHomeZip 
         Height          =   285
         Left            =   1365
         MaxLength       =   10
         TabIndex        =   20
         Top             =   2625
         Width           =   3375
      End
      Begin VB.TextBox txtHomeAddress 
         Height          =   1080
         Left            =   420
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   525
         Width           =   4320
      End
      Begin VB.TextBox txtHomeCity 
         Height          =   285
         Left            =   1365
         MaxLength       =   40
         TabIndex        =   16
         Top             =   1785
         Width           =   3375
      End
      Begin VB.TextBox txtHomeState 
         Height          =   285
         Left            =   1365
         MaxLength       =   3
         TabIndex        =   18
         Top             =   2205
         Width           =   3375
      End
      Begin VB.Label lblTimezone 
         Caption         =   "&Time zone"
         Height          =   225
         Left            =   420
         TabIndex        =   23
         Top             =   3465
         Width           =   960
      End
      Begin VB.Label lblHomeCellular 
         Caption         =   "&Cellular"
         Height          =   225
         Left            =   420
         TabIndex        =   29
         Top             =   5355
         Width           =   1485
      End
      Begin VB.Label lblHomePhone 
         Caption         =   "&Telephone"
         Height          =   225
         Left            =   420
         TabIndex        =   25
         Top             =   4515
         Width           =   1485
      End
      Begin VB.Label lblHomeFax 
         Caption         =   "&Facsimile"
         Height          =   225
         Left            =   420
         TabIndex        =   27
         Top             =   4935
         Width           =   1065
      End
      Begin VB.Label lblHomeZip 
         Caption         =   "&ZIP"
         Height          =   225
         Left            =   420
         TabIndex        =   19
         Top             =   2625
         Width           =   1485
      End
      Begin VB.Label lblHomeCountry 
         Caption         =   "&Country"
         Height          =   225
         Left            =   420
         TabIndex        =   21
         Top             =   3045
         Width           =   1065
      End
      Begin VB.Label lblHomeAddress 
         Caption         =   "&Home Address"
         Height          =   225
         Left            =   420
         TabIndex        =   13
         Top             =   210
         Width           =   1380
      End
      Begin VB.Label lblHomeCity 
         Caption         =   "&City"
         Height          =   225
         Left            =   420
         TabIndex        =   15
         Top             =   1785
         Width           =   1485
      End
      Begin VB.Label lblHomeState 
         Caption         =   "&State"
         Height          =   225
         Left            =   420
         TabIndex        =   17
         Top             =   2205
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdSaveInfo 
      Caption         =   "&Save Info"
      Height          =   330
      Left            =   4305
      TabIndex        =   92
      ToolTipText     =   "Get the latest information from user."
      Top             =   7140
      Width           =   1170
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   6720
      TabIndex        =   94
      ToolTipText     =   "Close this window"
      Top             =   7140
      Width           =   1170
   End
   Begin VB.CommandButton cmdGetInfo 
      Caption         =   "&Get Info"
      Height          =   330
      Left            =   5460
      TabIndex        =   93
      ToolTipText     =   "Get the latest information from user."
      Top             =   7140
      Width           =   1170
   End
   Begin VB.PictureBox BoxHeader 
      BackColor       =   &H00FFFFFF&
      Height          =   960
      Left            =   -100
      ScaleHeight     =   900
      ScaleWidth      =   8145
      TabIndex        =   96
      Top             =   -100
      Width           =   8200
      Begin VB.Label lblHeadDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose what you would like to see from the category on the left."
         Height          =   330
         Left            =   735
         TabIndex        =   98
         Top             =   525
         Width           =   7050
      End
      Begin VB.Label lblHeader 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Details"
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
         TabIndex        =   97
         Top             =   210
         Width           =   5265
      End
   End
   Begin VB.Frame frmTvwHolder 
      Height          =   5805
      Left            =   105
      TabIndex        =   99
      Top             =   945
      Width           =   2430
      Begin MSComctlLib.TreeView tvwCategory 
         Height          =   5670
         Left            =   10
         TabIndex        =   95
         Top             =   100
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   10001
         _Version        =   393217
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   0
      End
   End
   Begin VB.Frame frmLineSeperator 
      Height          =   115
      Left            =   0
      TabIndex        =   100
      Top             =   6825
      Width           =   7995
   End
   Begin VB.Frame frmOpt1 
      BorderStyle     =   0  'None
      Height          =   5790
      Left            =   2625
      TabIndex        =   101
      Top             =   945
      Width           =   5265
      Begin VB.CheckBox chkPublishIP 
         Caption         =   "PublishIP"
         Height          =   225
         Left            =   420
         TabIndex        =   12
         Top             =   4620
         Width           =   2640
      End
      Begin VB.CheckBox chkWebPresence 
         Caption         =   "&Web Presence"
         Height          =   225
         Left            =   420
         TabIndex        =   11
         Top             =   4305
         Width           =   3060
      End
      Begin VB.CheckBox chkAuthorize 
         Caption         =   "Require &Authorization"
         Height          =   225
         Left            =   420
         TabIndex        =   10
         Top             =   3990
         Width           =   3375
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Index           =   2
         Left            =   420
         MaxLength       =   60
         TabIndex        =   9
         ToolTipText     =   "Click to copy email to clipboard."
         Top             =   3360
         Width           =   4425
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Index           =   1
         Left            =   420
         MaxLength       =   60
         TabIndex        =   8
         ToolTipText     =   "Click to copy email to clipboard."
         Top             =   2940
         Width           =   4425
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Index           =   0
         Left            =   420
         MaxLength       =   60
         TabIndex        =   7
         ToolTipText     =   "Click to copy email to clipboard."
         Top             =   2520
         Width           =   4425
      End
      Begin VB.TextBox txtLastname 
         Height          =   285
         Left            =   1470
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1470
         Width           =   3375
      End
      Begin VB.TextBox txtFirstname 
         Height          =   285
         Left            =   1470
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1050
         Width           =   3375
      End
      Begin VB.TextBox txtNickname 
         Height          =   285
         Left            =   1470
         MaxLength       =   20
         TabIndex        =   1
         Top             =   630
         Width           =   3375
      End
      Begin VB.Label lblEmail 
         Caption         =   "&E-mail Address"
         Height          =   225
         Left            =   420
         TabIndex        =   6
         Top             =   2205
         Width           =   1695
      End
      Begin VB.Label lblLastname 
         Caption         =   "&Last name"
         Height          =   225
         Left            =   420
         TabIndex        =   4
         Top             =   1575
         Width           =   1065
      End
      Begin VB.Label lblFirstname 
         Caption         =   "&First name"
         Height          =   225
         Left            =   420
         TabIndex        =   2
         Top             =   1155
         Width           =   1485
      End
      Begin VB.Label lblNickname 
         Caption         =   "&Nick name"
         Height          =   225
         Left            =   420
         TabIndex        =   0
         Top             =   735
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frmUserInfoUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public UserUIN As Long
Public FrmIndex As Byte
Private InfoReceived As Byte
Private InfoUpdated As Byte

Private Sub cmdGetInfo_Click()
  If IcqUdp.SocketState <> icqConnected Then
    MsgBox "Communicator is not currently connected to the ICQ Server. Please connect by choosing File->Connect on the file menu", vbOKOnly, "Not Connected"
    Exit Sub
  End If
  
  cmdGetInfo.Enabled = False
  InfoReceived = 0
  IcqUdp.inforequestall UserUIN
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

'#############################################################################
'Saving Module
Private Sub cmdSaveInfo_Click()
  Dim UserInfo As typContactInfo
  With UserInfo
    .strNickname = txtNickname
    .strFirstName = txtFirstname
    .strLastName = txtLastname
    .strEmail = txtEmail(0)
    .strEmail2 = txtEmail(1)
    .strEmail3 = txtEmail(2)
        
    .strStreet = txtHomeAddress
    .strCity = txtHomeCity
    .intCountryCode = GetCmbData(cmbHomeCountry)
    .strState = txtHomeState
    .lngZip = txtHomeZip
    .byteTimeZone = GetCmbData(cmbTimeZone)

    .intAge = CInt(txtAge)
    .byteGender = GetCmbData(cmbGender)
    .strPhone = txtHomePhone
    .strFax = txtHomeFax
    .strCellular = txtHomeCellular
    .strHomepageURL = txtHomeURL
    .strAboutInfo = txtAboutInfo
      
    .strWorkPhone = txtWorkPhone
    .strWorkFax = txtWorkFax
    .strWorkAddress = txtWorkAddress
    .strWorkCity = txtWorkCity
    .strWorkState = txtWorkState
    .lngWorkZip = txtWorkZip
    .intWorkCountry = GetCmbData(cmbWorkCountry)
        
    .strWorkName = txtCompanyName
    .strWorkDepartment = txtCompanyDepartment
    .strWorkPosition = txtCompanyPosition
    .intWorkOccupation = GetCmbData(cmbCompanyOccupation)
    .strWorkWebURL = txtCompanyURL
        
    If IsNull(datBirthDate.Value) Then
      .byteBirthDay = 0
      .byteBirthMonth = 0
      .byteBirthYear = 0
    Else
      .byteBirthDay = datBirthDate.Day
      .byteBirthMonth = datBirthDate.Month
      .byteBirthYear = CByte(Format(datBirthDate, "yy"))
    End If
    
    .byteLanguage1 = GetCmbData(cmbLanguage(0))
    .byteLanguage2 = GetCmbData(cmbLanguage(1))
    .byteLanguage3 = GetCmbData(cmbLanguage(2))
      
    .byteInterestTotal = 4
    .intInterestCategory(0) = txtInterestCategory(0).Text
    .intInterestCategory(1) = txtInterestCategory(1).Text
    .intInterestCategory(2) = txtInterestCategory(2).Text
    .intInterestCategory(3) = txtInterestCategory(3).Text
    .strInterestName(0) = txtInterestName(0).Text
    .strInterestName(1) = txtInterestName(1).Text
    .strInterestName(2) = txtInterestName(2).Text
    .strInterestName(3) = txtInterestName(3).Text

    .byteBackgroundTotal = 3
    .intBackgroundCategory(0) = GetCmbData(cmbBackgroundCategory(0))
    .intBackgroundCategory(1) = GetCmbData(cmbBackgroundCategory(1))
    .intBackgroundCategory(2) = GetCmbData(cmbBackgroundCategory(2))
    .strBackgroundName(0) = txtBackgroundName(0).Text
    .strBackgroundName(1) = txtBackgroundName(1).Text
    .strBackgroundName(2) = txtBackgroundName(2).Text
    
    .byteOrganizationTotal = 3
    .intOrganizationCategory(0) = GetCmbData(cmbGroupCategory(0))
    .intOrganizationCategory(1) = GetCmbData(cmbGroupCategory(1))
    .intOrganizationCategory(2) = GetCmbData(cmbGroupCategory(2))
    .strOrganizationName(0) = txtGroupName(0).Text
    .strOrganizationName(1) = txtGroupName(1).Text
    .strOrganizationName(2) = txtGroupName(2).Text
    
    .bAuthorize = IIf(chkAuthorize.Value = Checked, True, False)
    .bWebPresence = IIf(chkWebPresence.Value = Checked, True, False)
    .bPublishIP = IIf(chkPublishIP.Value = Checked, True, False)
  End With
  
  cmdSaveInfo.Enabled = False
  IcqUdp.InfoUpdate icqall, UserInfo
  UserInfoSet icqMain, UserInfo
  UserInfoSet icqMetaMore, UserInfo
  UserInfoSet icqAbout, UserInfo
  UserInfoSet icqAffiliations, UserInfo
  UserInfoSet icqInterest, UserInfo
  UserInfoSet icqWork, UserInfo
  
  cmdSaveInfo.Enabled = True
End Sub
'#############################################################################

Private Sub Form_Load()
  Dim i As Integer, j As Integer, TempName As String, TempValue As Integer
  tvwCategory.Nodes.Add , , "Main", "Main"
  tvwCategory.Nodes.Add , , "Home", "Home"
  tvwCategory.Nodes.Add , , "Work", "Work"
  tvwCategory.Nodes.Add , , "Company", "Company"
  tvwCategory.Nodes.Add , , "More", "More"
  tvwCategory.Nodes.Add , , "Others", "Others"

  UserUIN = IcqUdp.UserUIN
  
  'Fill combo box value for Country (Home/Work)
  For i = 0 To 121
    TempValue = IcqUtility.getcountrycode(i)
    TempName = IcqUtility.GetCountryName(TempValue)
    With cmbHomeCountry
      .AddItem TempName
      .ItemData(.NewIndex) = TempValue
    End With
    With cmbWorkCountry
      .AddItem TempName
      .ItemData(.NewIndex) = TempValue
    End With
  Next i
  
  'Fill combo box value for TimeZone
  For i = 24 To 0 Step -1
    TempValue = i
    TempName = IcqUtility.GetTimeZone(TempValue)
    With cmbTimeZone
      .AddItem TempName
      .ItemData(.NewIndex) = TempValue
    End With
  Next i
  For i = 255 To 230 Step -1
    TempValue = i
    TempName = IcqUtility.GetTimeZone(TempValue)
    With cmbTimeZone
      .AddItem TempName
      .ItemData(.NewIndex) = TempValue
    End With
  Next i
  
  'Fill Combobox value for Occupation
  For i = 0 To 27
    TempValue = IcqUtility.GetOccupationCode(i)
    TempName = IcqUtility.GetOccupationName(TempValue)
    With cmbCompanyOccupation
      .AddItem TempName
      .ItemData(.NewIndex) = TempValue
    End With
  Next i
  
  'Fill Combobox value for Occupation
  For i = 0 To 33
    TempValue = i
    TempName = IcqUtility.GetLangName(i)
    For j = 0 To 2
      With cmbLanguage(j)
        .AddItem TempName
        .ItemData(.NewIndex) = TempValue
      End With
    Next j
  Next i

  'Fill Combobox value for Group, Organization, Affiliations
  For i = 0 To 19
    TempValue = IcqUtility.GetAffiliationsCode(i)
    TempName = IcqUtility.GetAffiliationsName(TempValue)
    For j = 0 To 2
      With cmbGroupCategory(j)
        .AddItem TempName
        .ItemData(.NewIndex) = TempValue
      End With
    Next j
  Next i
  
  'Fill Combobox value for Past Backgrounds
  For i = 0 To 7
    TempValue = IcqUtility.GetPastBackgroundCode(i)
    TempName = IcqUtility.GetPastBackgroundname(TempValue)
    For j = 0 To 2
      With cmbBackgroundCategory(j)
        .AddItem TempName
        .ItemData(.NewIndex) = TempValue
      End With
    Next j
  Next i

  GetUserInfo

End Sub

Private Sub tvwCategory_NodeClick(ByVal Node As MSComctlLib.Node)
  Select Case Node.Key
    Case "Main"
      frmOpt1.Visible = True
      frmOpt2.Visible = False
      frmOpt3.Visible = False
      frmOpt4.Visible = False
      frmOpt5.Visible = False
      frmOpt6.Visible = False
    Case "Home"
      frmOpt1.Visible = False
      frmOpt2.Visible = True
      frmOpt3.Visible = False
      frmOpt4.Visible = False
      frmOpt5.Visible = False
      frmOpt6.Visible = False
    Case "Work"
      frmOpt1.Visible = False
      frmOpt2.Visible = False
      frmOpt3.Visible = True
      frmOpt4.Visible = False
      frmOpt5.Visible = False
      frmOpt6.Visible = False
    Case "Company"
      frmOpt1.Visible = False
      frmOpt2.Visible = False
      frmOpt3.Visible = False
      frmOpt4.Visible = True
      frmOpt5.Visible = False
      frmOpt6.Visible = False
    Case "More"
      frmOpt1.Visible = False
      frmOpt2.Visible = False
      frmOpt3.Visible = False
      frmOpt4.Visible = False
      frmOpt5.Visible = True
      frmOpt6.Visible = False
    Case "Others"
      frmOpt1.Visible = False
      frmOpt2.Visible = False
      frmOpt3.Visible = False
      frmOpt4.Visible = False
      frmOpt5.Visible = False
      frmOpt6.Visible = True
  End Select
End Sub

Public Sub EventGetInfo(InfoType As Integer)
  InfoReceived = InfoReceived + 1
  If InfoReceived = 7 Then cmdGetInfo.Enabled = True:  GetUserInfo
End Sub

Public Sub GetUserInfo()
  ContRS.Filter = "UIN =" + Str$(UserUIN)
  Set ContRSFilter = ContRS.OpenRecordset
  
  With ContRSFilter
    If .RecordCount = 0 Then
      .AddNew
      !uin = UserUIN
      .Update
      ContRS.Filter = "UIN =" + Str$(UserUIN)
      Set ContRSFilter = ContRS.OpenRecordset
      .MoveFirst
      .Edit
    Else
      .MoveFirst
      .Edit
    End If
    
    txtNickname = !Nickname
    txtFirstname = !FirstName
    txtLastname = !LastName
    txtEmail(0) = !Email1
    txtEmail(1) = !Email2
    txtEmail(2) = !Email3
        
    txtHomeAddress = !Street
    txtHomeCity = !City
    SetCmbId cmbHomeCountry, !Country
    txtHomeState = !state
    txtHomeZip = Trim$(Str$(!ZIP))
    SetCmbId cmbTimeZone, !TimeZone

    txtAge = Trim$(Str$(!Age))
    cmbGender.ListIndex = !Gender
    txtHomePhone = !Phone
    txtHomeFax = !Fax
    txtHomeCellular = !Cellular
    txtHomeURL = !URL
    txtAboutInfo = !AboutInfo
      
    txtWorkPhone = !WorkPhone
    txtWorkFax = !WorkFax
    txtWorkAddress = !WorkAddress
    txtWorkCity = !WorkCity
    txtWorkState = !WorkState
    txtWorkZip = !WorkZip
    SetCmbId cmbWorkCountry, !WorkCountry
        
    txtCompanyName = !WorkName
    txtCompanyDepartment = !WorkDepartment
    txtCompanyPosition = !WorkPosition
    SetCmbId cmbCompanyOccupation, !WorkOccupation
    txtCompanyURL = !WorkURL
        
    If IsNull(!BirthDate) Then datBirthDate.Value = Null Else datBirthDate.Value = !BirthDate
    
    SetCmbId cmbLanguage(0), !Language1
    SetCmbId cmbLanguage(1), !Language2
    SetCmbId cmbLanguage(2), !Language3
      
    txtInterestCategory(0).Text = Trim$(Str$(!InterestCategory0))
    txtInterestCategory(1).Text = Trim$(Str$(!InterestCategory1))
    txtInterestCategory(2).Text = Trim$(Str$(!InterestCategory2))
    txtInterestCategory(3).Text = Trim$(Str$(!InterestCategory3))
    txtInterestName(0).Text = !InterestName0
    txtInterestName(1).Text = !InterestName1
    txtInterestName(2).Text = !InterestName2
    txtInterestName(3).Text = !InterestName3

    SetCmbId cmbBackgroundCategory(0), !BackgroundCategory0
    SetCmbId cmbBackgroundCategory(1), !BackgroundCategory1
    SetCmbId cmbBackgroundCategory(2), !BackgroundCategory2
    txtBackgroundName(0).Text = !BackgroundName0
    txtBackgroundName(1).Text = !BackgroundName1
    txtBackgroundName(2).Text = !BackgroundName2
    
    SetCmbId cmbGroupCategory(0), !OrganizationCategory0
    SetCmbId cmbGroupCategory(1), !OrganizationCategory1
    SetCmbId cmbGroupCategory(2), !OrganizationCategory2
    txtGroupName(0).Text = !OrganizationName0
    txtGroupName(1).Text = !OrganizationName1
    txtGroupName(2).Text = !OrganizationName2
    
    chkAuthorize.Value = IIf(!bAuthorize, 1, 0)
    chkWebPresence.Value = IIf(!bWebPresence, 1, 0)
    chkPublishIP.Value = IIf(!bPublishIP, 1, 0)
  End With
End Sub

Public Sub SetCmbId(cmbBox As ComboBox, ItemData As Long)
  Dim i As Integer, Id As Long
  Id = -1
  For i = 0 To cmbBox.ListCount - 1
    If cmbBox.ItemData(i) = ItemData Then Id = i: Exit For
  Next i
  
  cmbBox.ListIndex = Id
End Sub

Public Function GetCmbData(cmbBox As ComboBox) As Long
  If cmbBox.ListIndex = -1 Then GetCmbData = 0: Exit Function
  GetCmbData = cmbBox.ItemData(cmbBox.ListIndex)
End Function

