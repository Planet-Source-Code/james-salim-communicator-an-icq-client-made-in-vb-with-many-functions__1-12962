VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserInfo 
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
      Height          =   5790
      Left            =   2625
      TabIndex        =   90
      Top             =   945
      Visible         =   0   'False
      Width           =   5265
      Begin VB.TextBox txtAffiliations 
         Height          =   1335
         Left            =   210
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   72
         Top             =   2310
         Width           =   4635
      End
      Begin VB.TextBox txtInterests 
         Height          =   1440
         Left            =   210
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   70
         Top             =   315
         Width           =   4635
      End
      Begin VB.TextBox txtPastBackground 
         Height          =   1335
         Left            =   210
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   74
         Top             =   4200
         Width           =   4635
      End
      Begin VB.Label lblGroup 
         Caption         =   "Organization, Affiliations, Groups"
         Height          =   225
         Left            =   210
         TabIndex        =   71
         Top             =   2100
         Width           =   1905
      End
      Begin VB.Label lblInterest 
         Caption         =   "Personal &Interests"
         Height          =   225
         Left            =   210
         TabIndex        =   69
         Top             =   105
         Width           =   1485
      End
      Begin VB.Label lblPastBackground 
         Caption         =   "&Past Background"
         Height          =   225
         Left            =   210
         TabIndex        =   73
         Top             =   3990
         Width           =   1485
      End
   End
   Begin VB.Frame frmOpt5 
      BorderStyle     =   0  'None
      Height          =   5790
      Left            =   2625
      TabIndex        =   89
      Top             =   945
      Visible         =   0   'False
      Width           =   5265
      Begin VB.TextBox txtAboutInfo 
         Height          =   1590
         Left            =   735
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   68
         Top             =   4095
         Width           =   4005
      End
      Begin VB.TextBox txtLanguage 
         Height          =   285
         Index           =   2
         Left            =   735
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   3255
         Width           =   4005
      End
      Begin VB.TextBox txtLanguage 
         Height          =   285
         Index           =   1
         Left            =   735
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   2835
         Width           =   4005
      End
      Begin VB.TextBox txtLanguage 
         Height          =   285
         Index           =   0
         Left            =   735
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   2415
         Width           =   4005
      End
      Begin VB.TextBox txtAge 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   1155
         Width           =   3375
      End
      Begin VB.TextBox txtBirthDate 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   1575
         Width           =   3375
      End
      Begin VB.TextBox txtHomeURL 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   315
         Width           =   3375
      End
      Begin VB.TextBox txtGender 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   735
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
      TabIndex        =   88
      Top             =   945
      Visible         =   0   'False
      Width           =   5265
      Begin VB.TextBox txtCompanyURL 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   2310
         Width           =   3375
      End
      Begin VB.TextBox txtCompanyPosition 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1470
         Width           =   3375
      End
      Begin VB.TextBox txtCompanyOccupation 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   1890
         Width           =   3375
      End
      Begin VB.TextBox txtCompanyName 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   630
         Width           =   3375
      End
      Begin VB.TextBox txtCompanyDepartment 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
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
      TabIndex        =   86
      Top             =   945
      Visible         =   0   'False
      Width           =   5265
      Begin VB.TextBox txtWorkState 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2205
         Width           =   3375
      End
      Begin VB.TextBox txtWorkCity 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1785
         Width           =   3375
      End
      Begin VB.TextBox txtWorkAddress 
         Height          =   1080
         Left            =   420
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   525
         Width           =   4320
      End
      Begin VB.TextBox txtWorkCountry 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   3045
         Width           =   3375
      End
      Begin VB.TextBox txtWorkZip 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2625
         Width           =   3375
      End
      Begin VB.TextBox txtWorkFax 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   4935
         Width           =   3375
      End
      Begin VB.TextBox txtWorkPhone 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   4515
         Width           =   3375
      End
      Begin VB.Frame Frame2 
         Height          =   120
         Left            =   315
         TabIndex        =   87
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
      TabIndex        =   84
      Top             =   945
      Visible         =   0   'False
      Width           =   5265
      Begin VB.Frame Seperator2 
         Height          =   120
         Left            =   315
         TabIndex        =   85
         Top             =   3990
         Width           =   4530
      End
      Begin VB.TextBox txtTimezone 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   3465
         Width           =   3375
      End
      Begin VB.TextBox txtHomeCellular 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   5355
         Width           =   3375
      End
      Begin VB.TextBox txtHomePhone 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   4515
         Width           =   3375
      End
      Begin VB.TextBox txtHomeFax 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   4935
         Width           =   3375
      End
      Begin VB.TextBox txtHomeZip 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2625
         Width           =   3375
      End
      Begin VB.TextBox txtHomeCountry 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   3045
         Width           =   3375
      End
      Begin VB.TextBox txtHomeAddress 
         Height          =   1080
         Left            =   420
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   525
         Width           =   4320
      End
      Begin VB.TextBox txtHomeCity 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1785
         Width           =   3375
      End
      Begin VB.TextBox txtHomeState 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
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
   Begin VB.Frame frmOpt1 
      BorderStyle     =   0  'None
      Height          =   5790
      Left            =   2625
      TabIndex        =   83
      Top             =   945
      Width           =   5265
      Begin VB.CheckBox chkPublishIP 
         Caption         =   "PublishIP"
         Height          =   225
         Left            =   420
         TabIndex        =   78
         Top             =   4620
         Width           =   2640
      End
      Begin VB.CheckBox chkWebPresence 
         Caption         =   "&Web Presence"
         Height          =   225
         Left            =   420
         TabIndex        =   12
         Top             =   4305
         Width           =   3060
      End
      Begin VB.CheckBox chkAuthorize 
         Caption         =   "Require &Authorization"
         Height          =   225
         Left            =   420
         TabIndex        =   11
         Top             =   3990
         Width           =   3375
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Index           =   2
         Left            =   420
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Click to copy email to clipboard."
         Top             =   3360
         Width           =   4425
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Index           =   1
         Left            =   420
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Click to copy email to clipboard."
         Top             =   2940
         Width           =   4425
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Index           =   0
         Left            =   420
         Locked          =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "Click to copy email to clipboard."
         Top             =   2520
         Width           =   4425
      End
      Begin VB.TextBox txtLastname 
         Height          =   285
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1470
         Width           =   3375
      End
      Begin VB.TextBox txtFirstname 
         Height          =   285
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1050
         Width           =   3375
      End
      Begin VB.TextBox txtNickname 
         Height          =   285
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   630
         Width           =   3375
      End
      Begin VB.Label lblEmail 
         Caption         =   "&E-mail Address"
         Height          =   225
         Left            =   420
         TabIndex        =   7
         Top             =   2205
         Width           =   1695
      End
      Begin VB.Label lblLastname 
         Caption         =   "&Last name"
         Height          =   225
         Left            =   420
         TabIndex        =   5
         Top             =   1575
         Width           =   1065
      End
      Begin VB.Label lblFirstname 
         Caption         =   "&First name"
         Height          =   225
         Left            =   420
         TabIndex        =   3
         Top             =   1155
         Width           =   1485
      End
      Begin VB.Label lblNickname 
         Caption         =   "&Nick name"
         Height          =   225
         Left            =   420
         TabIndex        =   1
         Top             =   735
         Width           =   1380
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   6720
      TabIndex        =   75
      ToolTipText     =   "Close this window"
      Top             =   7140
      Width           =   1170
   End
   Begin VB.CommandButton cmdGetInfo 
      Caption         =   "&Get Info"
      Height          =   330
      Left            =   5565
      TabIndex        =   76
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
      TabIndex        =   0
      Top             =   -100
      Width           =   8200
      Begin VB.Label lblHeadDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose what you would like to see from the category on the left."
         Height          =   330
         Left            =   735
         TabIndex        =   80
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
         TabIndex        =   79
         Top             =   210
         Width           =   5265
      End
   End
   Begin VB.Frame frmTvwHolder 
      Height          =   5805
      Left            =   105
      TabIndex        =   81
      Top             =   945
      Width           =   2430
      Begin MSComctlLib.TreeView tvwCategory 
         Height          =   5670
         Left            =   10
         TabIndex        =   77
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
      TabIndex        =   82
      Top             =   6825
      Width           =   7995
   End
End
Attribute VB_Name = "frmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public UserUIN As Long
Public FrmIndex As Byte
Private InfoReceived As Byte

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
  frmUnloaded enumfrmUserDetail, FrmIndex
  Unload Me
End Sub

Private Sub Form_Load()
  tvwCategory.Nodes.Add , , "Main", "Main"
  tvwCategory.Nodes.Add , , "Home", "Home"
  tvwCategory.Nodes.Add , , "Work", "Work"
  tvwCategory.Nodes.Add , , "Company", "Company"
  tvwCategory.Nodes.Add , , "More", "More"
  tvwCategory.Nodes.Add , , "Others", "Others"

  UserUIN = TempUINBuffer
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
    txtHomeCountry = IcqUtility.GetCountryName(!Country)
    txtHomeState = !state
    txtHomeZip = Trim$(Str$(!ZIP))
    txtTimezone = IcqUtility.GetTimeZone(!TimeZone)

    txtAge = Trim$(Str$(!Age))
    Select Case !Gender
      Case icqMale:    txtGender = "Male"
      Case icqFemale:    txtGender = "Female"
      Case Else:    txtGender = "Not specified"
    End Select
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
    txtWorkCountry = IcqUtility.GetCountryName(!WorkCountry)
        
    txtCompanyName = !WorkName
    txtCompanyDepartment = !WorkDepartment
    txtCompanyPosition = !WorkPosition
    txtCompanyOccupation = IcqUtility.GetOccupationName(IcqUtility.GetOccupationIndex(!WorkOccupation))
    txtCompanyURL = !WorkURL
        
    txtBirthDate = Format(!BirthDate, "dd mmmm, yyyy")
    txtLanguage(0) = IcqUtility.GetLangName(!Language1)
    txtLanguage(1) = IcqUtility.GetLangName(!Language2)
    txtLanguage(2) = IcqUtility.GetLangName(!Language3)
      
    txtInterests.Text = _
      Trim$(Str$(!InterestCategory0)) & " : " & !InterestName0 & vbCrLf & _
      Trim$(Str$(!InterestCategory1)) & " : " & !InterestName1 & vbCrLf & _
      Trim$(Str$(!InterestCategory2)) & " : " & !InterestName2 & vbCrLf & _
      Trim$(Str$(!InterestCategory3)) & " : " & !InterestName3

    txtPastBackground.Text = _
      IcqUtility.GetPastBackgroundname(!BackgroundCategory0) & " : " & !BackgroundName0 & vbCrLf & _
      IcqUtility.GetPastBackgroundname(!BackgroundCategory1) & " : " & !BackgroundName1 & vbCrLf & _
      IcqUtility.GetPastBackgroundname(!BackgroundCategory2) & " : " & !BackgroundName2 & vbCrLf & _
      IcqUtility.GetPastBackgroundname(!BackgroundCategory3) & " : " & !BackgroundName3

    txtAffiliations.Text = _
      IcqUtility.GetAffiliationsName(!OrganizationCategory0) & " : " & !OrganizationName0 & vbCrLf & _
      IcqUtility.GetAffiliationsName(!OrganizationCategory1) & " : " & !OrganizationName1 & vbCrLf & _
      IcqUtility.GetAffiliationsName(!OrganizationCategory2) & " : " & !OrganizationName2 & vbCrLf & _
      IcqUtility.GetAffiliationsName(!OrganizationCategory3) & " : " & !OrganizationName3
      
    chkAuthorize.Value = IIf(!bAuthorize, 1, 0)
    chkWebPresence.Value = IIf(!bWebPresence, 1, 0)
    chkPublishIP.Value = IIf(!bPublishIP, 1, 0)

  End With
End Sub

