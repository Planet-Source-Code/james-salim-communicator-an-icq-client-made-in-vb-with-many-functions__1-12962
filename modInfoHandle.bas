Attribute VB_Name = "modInfoHandle"
Option Explicit

Public Sub UserInfoSet(InfoType As IcqUdpCtl.enumInfoType, Info As IcqUdpCtl.typContactInfo)
  ContRS.Filter = "UIN =" + Str$(Info.lngUIN)
  Set ContRSFilter = ContRS.OpenRecordset
  
  With ContRSFilter
    If .RecordCount = 0 Then
      .AddNew
      !uin = Info.lngUIN
      !bOnContact = False
    Else
      .MoveFirst
      .Edit
    End If
    
    Select Case InfoType
      Case icqBasic
        !DisplayName = IIf(Info.strNickname = "", Trim$(Str$(Info.lngUIN)), Info.strNickname)
        !Nickname = Info.strNickname
        !Firstname = Info.strFirstName
        !Lastname = Info.strLastName
        !email1 = Info.strEmail
        !bAuthorize = Info.bAuthorize
        !bWebPresence = Info.bWebPresence
        !bPublishIP = Info.bPublishIP
      Case icqMore
        !City = Info.strCity
        !Country = Info.intCountryCode
        !TimeZone = Info.byteTimeZone
        !state = Info.strState
        !Age = Info.intAge
        !Gender = Info.byteGender
        !Phone = Info.strPhone
        !URL = Info.strHomepageURL
        !AboutInfo = Info.strAboutInfo
      
      Case icqMain
        !Nickname = Info.strNickname
        !Firstname = Info.strFirstName
        !Lastname = Info.strLastName
        !email1 = Info.strEmail
        !Email2 = Info.strEmail2
        !Email3 = Info.strEmail3
        !City = Info.strCity
        !state = Info.strState
        !Phone = Info.strPhone
        !Fax = Info.strFax
        !Street = Info.strStreet
        !Cellular = Info.strCellular
        !ZIP = Info.lngZip
        !Country = Info.intCountryCode
        !TimeZone = Info.byteTimeZone
        !bAuthorize = Info.bAuthorize
        !bWebPresence = Info.bWebPresence
        !bPublishIP = Info.bPublishIP
        
      Case icqWork
        !WorkCity = Info.strWorkCity
        !WorkState = Info.strWorkState
        !WorkPhone = Info.strWorkPhone
        !WorkFax = Info.strWorkFax
        !WorkAddress = Info.strWorkAddress
        !WorkZip = Info.lngWorkZip
        !WorkCountry = Info.intWorkCountry
        !WorkName = Info.strWorkName
        !WorkDepartment = Info.strWorkDepartment
        !WorkPosition = Info.strWorkPosition
        !WorkOccupation = Info.intWorkOccupation
        !WorkURL = Info.strWorkWebURL
        
      Case icqMetaMore
        !Age = Info.intAge
        !Gender = Info.byteGender
        !URL = Info.strHomepageURL
        If IsDate(Str$(Info.byteBirthDay) + Str$(Info.byteBirthMonth) + Str$(Info.byteBirthYear)) = True Then
          !BirthDate = CDate(Str$(Info.byteBirthDay) + Str$(Info.byteBirthMonth) + Str$(Info.byteBirthYear))
        End If
        !Language1 = Info.byteLanguage1
        !Language2 = Info.byteLanguage2
        !Language3 = Info.byteLanguage3
      
      Case icqAbout
        !AboutInfo = Info.strAboutInfo
      
      Case icqInterest
        !InterestTotal = Info.byteInterestTotal
        !InterestCategory0 = Info.intInterestCategory(0)
        !InterestName0 = Info.strInterestName(0)
        !InterestCategory1 = Info.intInterestCategory(1)
        !InterestName1 = Info.strInterestName(1)
        !InterestCategory2 = Info.intInterestCategory(2)
        !InterestName2 = Info.strInterestName(2)
        !InterestCategory3 = Info.intInterestCategory(3)
        !InterestName3 = Info.strInterestName(3)
      
      Case icqAffiliations
        !BackgroundTotal = Info.byteBackgroundTotal
        !BackgroundCategory0 = Info.intBackgroundCategory(0)
        !BackgroundName0 = Info.strBackgroundName(0)
        !BackgroundCategory1 = Info.intBackgroundCategory(1)
        !BackgroundName1 = Info.strBackgroundName(1)
        !BackgroundCategory2 = Info.intBackgroundCategory(2)
        !BackgroundName2 = Info.strBackgroundName(2)
        !BackgroundCategory3 = Info.intBackgroundCategory(3)
        !BackgroundName3 = Info.strBackgroundName(3)
        
        !OrganizationTotal = Info.byteOrganizationTotal
        !OrganizationCategory0 = Info.intOrganizationCategory(0)
        !OrganizationName0 = Info.strOrganizationName(0)
        !OrganizationCategory1 = Info.intOrganizationCategory(1)
        !OrganizationName1 = Info.strOrganizationName(1)
        !OrganizationCategory2 = Info.intOrganizationCategory(2)
        !OrganizationName2 = Info.strOrganizationName(2)
        !OrganizationCategory3 = Info.intOrganizationCategory(3)
        !OrganizationName3 = Info.strOrganizationName(3)
    
      Case icqHPCategory
        !HPCategoryTotal = Info.byteHPCategoryTotal
        !HPCategoryCategory0 = Info.intHPCategoryCategory(0)
        !HPCategoryName0 = Info.strHPCategoryName(0)
        !HPCategoryCategory1 = Info.intHPCategoryCategory(1)
        !HPCategoryName1 = Info.strHPCategoryName(1)
        !HPCategoryCategory2 = Info.intHPCategoryCategory(2)
        !HPCategoryName2 = Info.strHPCategoryName(2)
        !HPCategoryCategory3 = Info.intHPCategoryCategory(3)
        !HPCategoryName3 = Info.strHPCategoryName(3)
    End Select
    
    If Info.strNickname <> "" Then !DisplayName = Info.strNickname
    .Update
  End With
End Sub

Public Function Uin2Name(uin As Long) As String
  ContRS.Filter = "UIN =" + Str$(uin)
  Set ContRSFilter = ContRS.OpenRecordset
  
  With ContRSFilter
    If .RecordCount = 0 Then
      Uin2Name = Trim$(Str$(uin))
    Else
      .MoveFirst
      .Edit
      Uin2Name = !Nickname
      If Uin2Name = "" Then Uin2Name = Trim$(Str$(uin))
    End If
  End With
End Function

Public Sub ContactNewUser(uin As Long, nick As String, Optional first As String, Optional last As String, Optional email As String)
  ContRS.Filter = "UIN =" + Str$(uin)
  Set ContRSFilter = ContRS.OpenRecordset
  
  With ContRSFilter
    If .RecordCount = 0 Then
      .AddNew
      !uin = uin
      !Nickname = nick
      !Firstname = first
      !Lastname = last
      !email1 = email
      !bOnContact = True
    Else
      .MoveFirst
      .Edit
      !bOnContact = True
    End If
    .Update
  End With
End Sub

Public Sub ContactDelUser(uin As Long)
  ContRS.Filter = "UIN =" + Str$(uin)
  Set ContRSFilter = ContRS.OpenRecordset
  
  With ContRSFilter
    If .RecordCount > 0 Then
      .MoveFirst
      .Edit
      !bOnContact = False
    End If
    .Update
  End With
End Sub

Public Function ContactExist(uin As Long) As Boolean
  Dim TempStr As String
  On Error GoTo ErrorProc
  
  ContactExist = True
  TempStr = tvwContact.Nodes("o" & Trim$(Str$(uin))).Text
  Exit Function
  
ErrorProc:
  ContactExist = False
End Function

