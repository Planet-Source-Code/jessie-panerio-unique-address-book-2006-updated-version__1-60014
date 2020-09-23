Attribute VB_Name = "Module1"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Const HELPC = &H3&
Public Path As String
Public cnn As ADODB.Connection
Public adoview As ADODB.Recordset
Public findmode As Boolean
Public infomode As Boolean
Public modeval As Boolean
Public Mode As Boolean
Public strImgN As String
Public BImg() As Byte

Public Function lv_TimerCallBack(ByVal hWnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
     Dim tgtButton As lvButtons_H
     CopyMemory tgtButton, GetProp(hWnd, "lv_ClassID"), &H4
     Call tgtButton.TimerUpdate(GetProp(hWnd, "lv_TimerID"))
     CopyMemory tgtButton, 0&, &H4

End Function

Public Sub main()

    If Right(App.Path, 1) <> "\" Then
      Path = App.Path & "\"
    Else
      Path = App.Path
    End If
    If App.PrevInstance = True Then
      MsgBox "Address Book is already open.", vbOKOnly + vbInformation, "Address Book"
      End
    End If
    Call getconnected
    Call rs_view
    App.HelpFile = Path & "\JessiePanerio.HLP"
    Mode = True
    Load frmMain
    frmMain.Show
  
End Sub

Public Sub getconnected()

     Set cnn = New ADODB.Connection
     cnn.CursorLocation = adUseClient
     cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Path & "\AddressBook.mdb;Persist Security Info=False;Jet OLEDB:Database Password=panerio"
     cnn.Open
       
End Sub

Public Sub rs_view()

     Set adoview = New ADODB.Recordset
     adoview.Open "Select * from addressbook", cnn, adOpenStatic, adLockPessimistic
  
End Sub

Public Sub loaddataforviewing()
     
     On Error Resume Next
     With frmPersonalInfo
       If adoview.BOF = True Or adoview.EOF = True Then
         .lblrecno.Caption = " Record " & adoview.AbsolutePosition + 1 & " of " & adoview.MaxRecords
         .lblNickName1.Caption = ""
         .lblNickName2.Caption = ""
         .lblNickName3.Caption = ""
         Set .imgpic.Picture = Nothing
         Exit Sub
       Else
         Call LoadImage
         .lblrecno.Caption = " Record " & adoview.AbsolutePosition & " of " & adoview.recordcount
         .lblNickName1.Caption = adoview!nickname
         .lblNickName2.Caption = adoview!nickname
         .lblNickName3.Caption = adoview!nickname
         .lblLastName.Caption = adoview!lastname
         .lblFirstName.Caption = adoview!firstname
         .lblMiddleName.Caption = adoview!middlename
         .lblBirthDay.Caption = adoview!birthday
         .lblGender.Caption = adoview!gender
         .lblReligion.Caption = adoview!religion
         .lblCitizenship.Caption = adoview!citizenship
         .lblcivilstatus.Caption = adoview!civilstatus
         .lblContactHome.Caption = adoview!homeaddr
         .lblContactLandLine.Caption = adoview!homeno
         .lblContactMobile.Caption = adoview!homemobileno
         .lblOfficeName.Caption = adoview!officename
         .lblOfficeAddr.Caption = adoview!officeaddr
         .lblOfficePhone.Caption = adoview!officeno
         .lblEmail1.Caption = adoview!email1
         .lblEmail2.Caption = adoview!email2
         .lblEmail3.Caption = adoview!email3
       End If
     End With

End Sub

Public Sub loaddatatoedit()
 
     If adoview.BOF = True Or adoview.EOF = True Then
        Exit Sub
     Else
      With frmAddNew
       Call LoadImage
       .txtNickName1.Text = adoview!nickname
       .txtNickName.Text = adoview!nickname
       .txtLastName.Text = adoview!lastname
       .txtFirstName.Text = adoview!firstname
       .txtMiddleName.Text = adoview!middlename
       .MaskEdBoxBirthDay.Text = adoview!birthday
       .cmbGender.Text = adoview!gender
       .txtReligion.Text = adoview!religion
       .txtCitizenship.Text = adoview!citizenship
       .txtCivilStatus.Text = adoview!civilstatus
       .txtContactHome.Text = adoview!homeaddr
       .txtContactLandLine.Text = adoview!homeno
       .txtContactMobile.Text = adoview!homemobileno
       .txtOfficeName.Text = adoview!officename
       .txtOfficeAddr.Text = adoview!officeaddr
       .txtOfficePhone.Text = adoview!officeno
       .txtEmail1.Text = adoview!email1
       .txtEmail2.Text = adoview!email2
       .txtEmail3.Text = adoview!email3
       .txtPictureName = adoview!picfilename
      End With
     End If

End Sub

Public Sub WriteDataFromControls()
     
     With frmAddNew
        If .txtPictureName.Text = "" Then
           Call nopicture
           GoTo jessiepanerio:
        Else
         
jessiepanerio:
           Call Image
           adoview!picfilename = .txtPictureName
           adoview.Fields("picblob").AppendChunk BImg
           adoview!nickname = .txtNickName.Text
           adoview!lastname = .txtLastName.Text
           adoview!firstname = .txtFirstName.Text
           adoview!middlename = .txtMiddleName.Text
           adoview!birthday = .MaskEdBoxBirthDay.Text
           adoview!gender = .cmbGender.Text
           adoview!religion = .txtReligion.Text
           adoview!citizenship = .txtCitizenship.Text
           adoview!civilstatus = .txtCivilStatus.Text
           adoview!homeaddr = .txtContactHome.Text
           adoview!homeno = .txtContactLandLine.Text
           adoview!homemobileno = .txtContactMobile.Text
           adoview!officename = .txtOfficeName.Text
           adoview!officeaddr = .txtOfficeAddr.Text
           adoview!officeno = .txtOfficePhone.Text
           adoview!email1 = .txtEmail1.Text
           adoview!email2 = .txtEmail2.Text
           adoview!email3 = .txtEmail3.Text
        End If
     End With

End Sub

Public Sub WriteDataFromControlsEdit()
     
     With frmAddNew
       If .txtPictureName.Text = adoview!picfilename Then
          GoTo jessiepanerio:
       Else
          Call Image
          adoview!picfilename = .txtPictureName
          adoview.Fields("picblob").AppendChunk BImg

jessiepanerio:
          adoview!nickname = .txtNickName.Text
          adoview!lastname = .txtLastName.Text
          adoview!firstname = .txtFirstName.Text
          adoview!middlename = .txtMiddleName.Text
          adoview!birthday = .MaskEdBoxBirthDay.Text
          adoview!gender = .cmbGender.Text
          adoview!religion = .txtReligion.Text
          adoview!citizenship = .txtCitizenship.Text
          adoview!civilstatus = .txtCivilStatus.Text
          adoview!homeaddr = .txtContactHome.Text
          adoview!homeno = .txtContactLandLine.Text
          adoview!homemobileno = .txtContactMobile.Text
          adoview!officename = .txtOfficeName.Text
          adoview!officeaddr = .txtOfficeAddr.Text
          adoview!officeno = .txtOfficePhone.Text
          adoview!email1 = .txtEmail1.Text
          adoview!email2 = .txtEmail2.Text
          adoview!email3 = .txtEmail3.Text
       End If
     End With

End Sub

Public Sub clearlblcontrols()

     With frmPersonalInfo
       .lblNickName1.Caption = ""
       .lblNickName2.Caption = ""
       .lblNickName3.Caption = ""
       .lblLastName.Caption = ""
       .lblFirstName.Caption = ""
       .lblMiddleName.Caption = ""
       .lblBirthDay.Caption = ""
       .lblGender.Caption = ""
       .lblReligion.Caption = ""
       .lblCitizenship.Caption = ""
       .lblcivilstatus.Caption = ""
       .lblContactHome.Caption = ""
       .lblContactLandLine.Caption = ""
       .lblContactMobile.Caption = ""
       .lblOfficeName.Caption = ""
       .lblOfficeAddr.Caption = ""
       .lblOfficePhone.Caption = ""
       .lblEmail1.Caption = ""
       .lblEmail2.Caption = ""
       .lblEmail3.Caption = ""
       Set .imgpic.Picture = Nothing
     End With
    
End Sub

Public Sub clearcontrols()
   
     With frmAddNew
       .txtNickName.Text = ""
       .txtLastName.Text = ""
       .txtFirstName.Text = ""
       .txtMiddleName.Text = ""
       .MaskEdBoxBirthDay.Mask = ""
       .MaskEdBoxBirthDay.Text = ""
       .cmbGender.Text = ""
       .txtReligion.Text = ""
       .txtCitizenship.Text = ""
       .txtCivilStatus.Text = ""
       .txtContactHome.Text = ""
       .txtContactLandLine.Text = ""
       .txtContactMobile.Text = ""
       .txtOfficeName.Text = ""
       .txtOfficeAddr.Text = ""
       .txtOfficePhone.Text = ""
       .txtEmail1.Text = ""
       .txtEmail2.Text = ""
       .txtEmail3.Text = ""
       .txtPictureName.Text = ""
       Set .imgpic.Picture = Nothing
     End With
      
End Sub

Public Sub cmdCover(value1 As Boolean, value2 As Boolean, value3 As Boolean, value4 As Boolean, value5 As Boolean)
  
     With frmAddNew
       .lvButtons_H1Cover.Visible = value1 And .lvbutton1New.Enabled = value1
       .lvButtons_H2Cover.Visible = value2 And .lvbutton2Cancel.Enabled = value2
       .lvButtons_H3Cover.Visible = value3 And .lvbutton3Save.Enabled = value3
       .lvButtons_H4Cover.Visible = value4 And .lvbutton4Close.Enabled = value4
       .cmdAddPicture.Visible = value5
     End With
   
End Sub

Public Sub lvbutton(Value As Boolean)

     With frmAddNew
       .lvbutton1New.Enabled = Value
       .lvbutton2Cancel.Enabled = Value
       .lvbutton3Save.Enabled = Value
       .lvbutton4Close.Enabled = Value
     End With

End Sub
Public Sub editbutton(Value As Boolean)
    
     With frmAddNew
       .lvbuttonsEditCancel.Visible = Value
       .lvbuttonsEditSave.Visible = Value
       .lvbuttonsEditClose.Visible = Value
       .cmdChangePicture.Visible = Value
     End With
    
End Sub

Public Sub lockcontrols(Value As String)

     With frmAddNew
       .txtNickName.Locked = Value
       .txtLastName.Locked = Value
       .txtFirstName.Locked = Value
       .txtMiddleName.Locked = Value
       .txtCover.Visible = Value
       .cmbGender.Locked = Value
       .txtReligion.Locked = Value
       .txtCitizenship.Locked = Value
       .txtCivilStatus.Locked = Value
       .txtContactHome.Locked = Value
       .txtContactLandLine.Locked = Value
       .txtContactMobile.Locked = Value
       .txtOfficeName.Locked = Value
       .txtOfficeAddr.Locked = Value
       .txtOfficePhone.Locked = Value
       .txtEmail1.Locked = Value
       .txtEmail2.Locked = Value
       .txtEmail3.Locked = Value
     End With

End Sub

Public Sub Image()

     On Error Resume Next
     Dim IntNum As Integer
     IntNum = FreeFile
     Open strImgN For Binary As #IntNum
     ReDim BImg(FileLen(strImgN))
     Get #IntNum, , BImg
     Close #1

End Sub

Public Sub LoadImage()
    
     On Error Resume Next
     Dim ImgS As Long
     Dim OS As Long
     Dim TmpPic As String
     Const conCS = 100
     TmpPic = App.Path & "\tmpPic.bmp"
     If Len(Dir(TmpPic)) > 0 Then
       Kill TmpPic
     End If
     Dim F As Integer
     F = FreeFile
     Open App.Path & "\tmpPic.bmp" For Binary As #F
     ImgS = adoview.Fields("picblob").ActualSize
     Do While OS < ImgS
       BImg() = adoview _
       ("picblob").GetChunk(conCS)
       Put #F, , BImg
       OS = OS + conCS
     Loop
     Close #F
     If infomode = True And findmode = False Then
       frmPersonalInfo.imgpic.Picture = LoadPicture(App.Path & "\tmpPic.bmp")
       Kill App.Path & "\tmpPic.bmp"
     ElseIf findmode = True And infomode = False Then
       frmFind.imgpic.Picture = LoadPicture(App.Path & "\tmpPic.bmp")
       Kill App.Path & "\tmpPic.bmp"
     End If
      
End Sub

Public Sub nopicture()
 
     With frmAddNew
        .jessiepanerio.InitDir = App.Path & "\imgjessiepanerio"
        .jessiepanerio.FileName = App.Path & "\imgjessiepanerio\jessiepanerio.jpg"
            If .jessiepanerio.FileName <> "" Then
               strImgN = .jessiepanerio.FileName
               .txtPictureName.Text = "No Picture"
               .imgpic.Picture = LoadPicture(.jessiepanerio.FileName)
            End If
     End With

End Sub

Public Sub LoadDataIntoFile(DataName As Integer, FileName As String)
    
     Dim myArray() As Byte
     Dim myFile As Long
     If Dir(FileName) = "" Then
        myArray = LoadResData(DataName, "CUSTOM")
        myFile = FreeFile
        Open FileName For Binary Access Write As #myFile
        Put #myFile, , myArray
        Close #myFile
     End If
     
End Sub

Public Sub validate()
             
     Dim rs As New ADODB.Recordset
     rs.Open "Select * From addressbook Where nickname = '" & frmAddNew.txtNickName.Text & "'", cnn, adOpenStatic, adLockReadOnly
     If rs.recordcount < 1 Then
        modeval = False
        Exit Sub
     Else
        modeval = True
     End If
     Set rs = Nothing
     
End Sub
