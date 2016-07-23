VERSION 5.00
Begin VB.Form members 
   BackColor       =   &H00000000&
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "D-Store - Members"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   885
   ClientWidth     =   12225
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton confirm 
      Caption         =   "Confirm"
      Height          =   495
      Left            =   9480
      TabIndex        =   10
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox memail 
      Height          =   375
      Left            =   9960
      TabIndex        =   9
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox mcontact 
      Height          =   375
      Left            =   9960
      TabIndex        =   8
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox madd2 
      Height          =   375
      Left            =   9960
      TabIndex        =   6
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox madd1 
      Height          =   375
      Left            =   9960
      TabIndex        =   5
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox mfore 
      Height          =   375
      Left            =   9960
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox msur 
      Height          =   375
      Left            =   9960
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox MemberID 
      Enabled         =   0   'False
      Height          =   375
      Left            =   9960
      TabIndex        =   1
      Text            =   "ID"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.ComboBox mdistrict 
      Height          =   315
      Left            =   9960
      TabIndex        =   7
      Text            =   "======District======"
      Top             =   4080
      Width           =   1935
   End
   Begin VB.ComboBox mtitle 
      Height          =   315
      Left            =   9960
      TabIndex        =   2
      Text            =   "Title"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.ListBox heading 
      Height          =   240
      Left            =   240
      TabIndex        =   16
      Top             =   1080
      Width           =   8295
   End
   Begin VB.ListBox memberlist 
      Height          =   4740
      Left            =   240
      TabIndex        =   15
      Top             =   1320
      Width           =   8295
   End
   Begin VB.Frame menu 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      Begin VB.Label showall 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Show all Members"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8040
         TabIndex        =   14
         Top             =   165
         Width           =   2415
      End
      Begin VB.Label searchm 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Search Member"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5280
         TabIndex        =   13
         Top             =   165
         Width           =   2295
      End
      Begin VB.Label editm 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Edit details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2760
         TabIndex        =   12
         Top             =   165
         Width           =   2055
      End
      Begin VB.Label createm 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "New Member"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   165
         Width           =   2055
      End
   End
   Begin VB.Label lblMID 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      Caption         =   "Member ID"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8880
      TabIndex        =   26
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      Caption         =   "Title"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9360
      TabIndex        =   25
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblsurname 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      Caption         =   "Surname"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9000
      TabIndex        =   24
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblforename 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      Caption         =   "Forename"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9000
      TabIndex        =   23
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lbladd1 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      Caption         =   "Address 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9000
      TabIndex        =   22
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lbladd2 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      Caption         =   "Address 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9000
      TabIndex        =   21
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lbldistrict 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      Caption         =   "District"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9000
      TabIndex        =   20
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblmcontact 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      Caption         =   "Contact Number"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8640
      TabIndex        =   19
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label lblemail 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      Caption         =   "Email"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9240
      TabIndex        =   18
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label checkusero 
      Height          =   615
      Left            =   1920
      TabIndex        =   17
      Top             =   6960
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "members"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim useroption As Integer
Dim selectedanitem As Boolean
Dim DB As New ADODB.Connection

'Resets the form
Private Sub Form_Load()
Call resetform
Call showallm
Call createm_Click
heading.Clear
heading.additem "[ID]" & Chr(9) & "[Title]" & "  " & "[Surname]" & Chr(9) & "[Forename]" & Chr(9) & "[Phone Number]" & Chr(9) & "[Email]" & Chr(9) & Chr(9) & Chr(9) & "[Points]"
End Sub

'Resets the forecolor of members button in mainmenu into white
Private Sub Form_Unload(Cancel As Integer)
Mainmenu.fwfmembers.ForeColor = &HFFFFFF
End Sub

'Changes to create mode and resets the form
Private Sub createm_Click()
    Call resetform
    useroption = 1
    createm.ForeColor = &HFF&
    editm.ForeColor = &HFFFFFF
    searchm.ForeColor = &HFFFFFF
End Sub

'Changes to edit mode and resets the form
Private Sub editm_Click()
    Call resetform
    useroption = 2
    createm.ForeColor = &HFFFFFF
    editm.ForeColor = &HFF&
    searchm.ForeColor = &HFFFFFF
    selectedanitem = False
End Sub

'Changes to search mode and resets the form
Private Sub searchm_Click()
    Call resetform
    useroption = 3
    createm.ForeColor = &HFFFFFF
    editm.ForeColor = &HFFFFFF
    searchm.ForeColor = &HFF&
    MemberID.Text = ""
    MemberID.Enabled = True
    MemberID.SetFocus
End Sub

'Calls the function that lists all members details into the list
Private Sub showall_Click()
Call showallm
End Sub

'Fills all the details of the chosen member in edit mode into the fields
Private Sub memberlist_Click()
If useroption = 2 Then
    numberofmember = memberlist.ListCount - 1
    For recentitem = 0 To numberofmember
        If memberlist.Selected(recentitem) = True Then
            itemid = recentitem + 1
            recentitem = numberofmember
        End If
    Next recentitem
selectedanitem = True
Dim rsc As New ADODB.Recordset
    DB.Open Path
        rsc.Open "select * from members where [memberid] = " & itemid, Path, adOpenKeyset, adLockOptimistic
            MemberID.Text = rsc![MemberID]
            mtitle.Text = rsc![Title]
            msur.Text = rsc![surname]
            mfore.Text = rsc![forename]
            madd1.Text = rsc![address1]
            madd2.Text = rsc![address2]
            mdistrict.Text = rsc![address3]
            mcontact.Text = rsc![contact number]
            memail.Text = rsc![Email]
        rsc.Close
    DB.Close
End If
End Sub

'Runs different functions depending on which mode is chosen
Private Sub confirm_Click()
Select Case useroption
Case 1
    checkok = checkvalid()
    If checkok = True Then
        Call fillin
        MsgBox "Member added!"
        Call initialise
        Call initialtitle
        Call initialmdistrict
        Call showallm
    End If
Case 2
    checkok = checkvalid()
    If checkok = True Then
        Call applychange
        MsgBox "Changes applied!"
        Call initialise
        Call initialtitle
        Call initialmdistrict
        Call showallm
    End If
Case 3
    Dim checkcrit As Integer
        checkcrit = 0
        If MemberID.Text <> "" Then checkcrit = checkcrit + 1
        If mtitle.Text <> "" Then checkcrit = checkcrit + 1
        If msur.Text <> "" Then checkcrit = checkcrit + 1
        If mfore.Text <> "" Then checkcrit = checkcrit + 1
        If madd1.Text <> "" Then checkcrit = checkcrit + 1
        If madd2.Text <> "" Then checkcrit = checkcrit + 1
        If mdistrict.Text <> "" Then checkcrit = checkcrit + 1
        If mcontact.Text <> "" Then checkcrit = checkcrit + 1
        If memail.Text <> "" Then checkcrit = checkcrit + 1
        If checkcrit > 0 Then
            Dim rsc As New ADODB.Recordset
            DB.Open Path
                rsc.Open "Select * from members", Path, adOpenKeyset, adLockOptimistic
                    memberlist.Clear
                    For i = 1 To rsc.RecordCount
                        match = 0
                        pad = "                                                                                          "
                        If MemberID.Text <> "" And rsc![MemberID] = MemberID.Text Then match = match + 1
                            m1 = Left(rsc![MemberID] & pad, 6)
                        If mtitle.Text <> "" And rsc![Title] = mtitle.Text Then match = match + 1
                            title1 = Left(rsc![Title] & pad, 4)
                        If msur.Text <> "" Then
                            If rsc![surname] = changecase(msur.Text) Then match = match + 1
                        End If
                            surname1 = Left(rsc![surname] & pad, 15)
                        If mfore.Text <> "" And rsc![forename] = mfore.Text Then match = match + 1
                            forename1 = Left(rsc![forename] & pad, 20)
                        If madd1.Text <> "" And rsc![address1] = madd1.Text Then match = match + 1
                        If madd2.Text <> "" And rsc![address2] = madd2.Text Then match = match + 1
                        If mdistrict.Text <> "" And rsc![address3] = mdistrict.Text Then match = match + 1
                        If mcontact.Text <> "" And rsc![contact number] = mcontact.Text Then match = match + 1
                            contact1 = rsc![contact number]
                        If memail.Text <> "" And rsc![Email] = memail.Text Then match = match + 1
                            email1 = Left(rsc![Email] & pad, 30)
                        points1 = Left(rsc![Points] & pad, 5)
                        If match = checkcrit Then memberlist.additem m1 & Chr(9) & title1 & "    " & surname1 & Chr(9) & forename1 & Chr(9) & contact1 & Chr(9) & email1 & Chr(9) & points1
                        rsc.MoveNext
                    Next i
                rsc.Close
            DB.Close
        Else
            itemlist.Clear
        End If
End Select
End Sub

'==============================Procedures/Functions==============================
'Saves all new member's details into the database
Private Function fillin()
Dim rsc As New ADODB.Recordset
    DB.Open Path
        rsc.Open "Select * from Members", Path, adOpenKeyset, adLockOptimistic
        If rsc.RecordCount <> 0 Then rsc.MoveLast
        rsc.AddNew
            rsc![Title] = mtitle.Text
            rsc![surname] = changecase(msur.Text)
            rsc![forename] = changecase(mfore.Text)
            rsc![contact number] = mcontact.Text
            rsc![Email] = memail.Text
            rsc![address1] = madd1.Text
            rsc![address2] = madd2.Text
            rsc![address3] = mdistrict.Text
            rsc![date of register] = Date
        rsc.Update
    DB.Close
End Function

'Saves all changes of member details into the database
Private Function applychange()
    numberofmember = memberlist.ListCount - 1
    For recentitem = 0 To numberofmember
        If memberlist.Selected(recentitem) = True Then
            itemid = recentitem + 1
            recentitem = numberofmember
        End If
    Next recentitem
Dim rsc As New ADODB.Recordset
    DB.Open Path
        rsc.Open "Select * from members where [memberid] = " & itemid, Path, adOpenKeyset, adLockOptimistic
            rsc![Title] = mtitle.Text
            rsc![surname] = msur.Text
            rsc![forename] = mfore.Text
            rsc![contact number] = mcontact.Text
            rsc![Email] = memail.Text
            rsc![address1] = madd1.Text
            rsc![address2] = madd2.Text
            rsc![address3] = mdistrict.Text
        rsc.Update
        rsc.Close
    DB.Close
End Function

'Validations for the fields
Private Function checkvalid()
checkvalid = False
If mtitle.Text <> "" Then
    lbltitle.ForeColor = &HFFFFFF
    If msur.Text <> "" And Len(msur.Text) <= 15 Then
        lblsurname.ForeColor = &HFFFFFF
        If mfore.Text <> "" And Len(mfore.Text) <= 20 Then
            lblforename.ForeColor = &HFFFFFF
            If madd1.Text <> "" And Len(madd1.Text) <= 30 Then
                lbladd1.ForeColor = &HFFFFFF
                If madd2.Text <> "" And Len(madd2.Text) <= 30 Then
                    lbladd2.ForeColor = &HFFFFFF
                    If mdistrict.Text <> "" And mdistrict.Text <> "==Hong Kong Island==" And mdistrict.Text <> "=====Kowloon=====" And mdistrict.Text <> "==New Territories==" Then
                        lbldistrict.ForeColor = &HFFFFFF
                        If mcontact.Text <> "" Then
                            numberused = False
                            Dim rsc As New ADODB.Recordset
                            DB.Open Path
                                rsc.Open "Select * from members where [contact number] = '" & mcontact.Text & "'", Path, adOpenKeyset, adLockOptimistic
                                If rsc.RecordCount <> 0 Then
                                    numberused = True
                                    If useroption = 2 And rsc![MemberID] = MemberID.Text Then numberused = False
                                End If
                                rsc.Close
                            DB.Close
                            If numberused = False Then
                                If Len(mcontact.Text) = 8 Then
                                    If Left(mcontact.Text, 1) = 2 Or Left(mcontact.Text, 1) = 3 Or Left(mcontact.Text, 1) = 9 Or Left(mcontact.Text, 1) = 6 Then
                                        For checkno = 1 To 8
                                            If Asc(Mid(mcontact.Text, checkno, 1)) < 48 Or Asc(Mid(mcontact.Text, checkno, 1)) > 57 Then
                                                numberok = False
                                                checkno = 8
                                            Else
                                                numberok = True
                                            End If
                                        Next checkno
                                    End If
                                End If
                            End If
                            If numberok = True Then
                                lblmcontact.ForeColor = &HFFFFFF
                            emailused = False
                            DB.Open Path
                                rsc.Open "Select * from members where [email] = '" & memail.Text & "'", Path, adOpenKeyset, adLockOptimistic
                                If rsc.RecordCount <> 0 Then
                                    emailused = True
                                    If useroption = 2 And rsc![MemberID] = MemberID.Text Then emailused = False
                                End If
                                rsc.Close
                            DB.Close
                                If emailused = False Then
                                    emailvalid = False
                                    If memail.Text <> "" Then
                                        nof@ = InStr(memail, "@")
                                            If nof@ = 0 Then
                                                numbermemail = False
                                                checkno = Len(memail.Text)
                                            Else
                                                emailvalid = True
                                            End If
                                    End If
                                    If Len(memail.Text) > 30 Then emailvalid = False
                                    If emailvalid = True Then
                                        checkvalid = True
                                        lblemail.ForeColor = &HFFFFFF
                                    Else
                                        lblemail.ForeColor = &HFF&
                                        If memail.Text <> "" Then MsgBox "Invalid email!"
                                        If memail.Text = "" Then MsgBox "Field email can't be empty!"
                                    End If
                                Else
                                    lblemail.ForeColor = &HFF&
                                    MsgBox "Email address already used by another user!"
                                End If
                            Else
                                lblmcontact.ForeColor = &HFF&
                                MsgBox "Contact number already used by another user!"
                            End If
                        Else
                            lblmcontact.ForeColor = &HFF&
                            MsgBox "Invalid phone number!"
                        End If
                    Else
                        lbldistrict.ForeColor = &HFF&
                        If mdistrict.Text = "" Then
                            MsgBox "Field district can't be empty!"
                        Else
                            MsgBox "Please choose a valid district"
                        End If
                    End If
                Else
                    lbladd2.ForeColor = &HFF&
                    If madd2.Text = "" Then MsgBox "Field address2 can't be empty!"
                    If Len(madd2.Text) > 30 Then MsgBox "Field address2 is too long!"
                End If
            Else
                lbladd1.ForeColor = &HFF&
                If madd1.Text = "" Then MsgBox "Field address1 can't be empty!"
                If Len(madd1.Text) > 30 Then MsgBox "Field address1 is too long!"
            End If
        Else
            lblforename.ForeColor = &HFF&
            If mfore.Text = "" Then MsgBox "Field forename can't be empty!"
            If Len(mfore.Text) > 20 Then MsgBox "Field forename is too long!"
        End If
    Else
        lblsurname.ForeColor = &HFF&
        If msur.Text = "" Then MsgBox "Field surname can't be empty!"
        If Len(msur.Text) > 15 Then MsgBox "Field surname is too long!"
    End If
Else
    lbltitle.ForeColor = &HFF&
    If mtitle.Text = "" Then MsgBox "Field title can't be empty!"
    If Len(mtitle.Text) > 4 Then MsgBox "Field title is too long!"
End If
End Function

'Empties all the fields and look for the new member ID if new member is added
Private Function initialise()
Dim rsc As New ADODB.Recordset
MemberID.Enabled = False
    DB.Open Path
        rsc.Open "Select [MemberID] from MEMBERS", Path, adOpenKeyset, adLockOptimistic
            If rsc.RecordCount <> 0 Then
                rsc.MoveLast
                MemberID.Text = rsc![MemberID] + 1
            Else
                MemberID.Text = "1"
            End If
        rsc.Close
    DB.Close
    msur.Text = ""
    mfore.Text = ""
    madd1.Text = ""
    madd2.Text = ""
    mcontact.Text = ""
    memail.Text = ""
End Function

'Resets the combo box mtitle
Private Function initialtitle()
    mtitle.Clear
    mtitle.additem "Mr"
    mtitle.additem "Mrs"
    mtitle.additem "Miss"
    mtitle.additem "Ms"
End Function

'Resets the combo box mDistrict
Private Function initialmdistrict()
mdistrict.Clear
mdistrict.additem "==Hong Kong Island=="
mdistrict.additem "Central and Western"
mdistrict.additem "Eastern"
mdistrict.additem "Southern"
mdistrict.additem "Wan Chai"
mdistrict.additem "=====Kowloon====="
mdistrict.additem "Kowloon City"
mdistrict.additem "Kwun Tong"
mdistrict.additem "Sham Shui Po"
mdistrict.additem "Wong Tai Sin"
mdistrict.additem "Yau Tsim Mong"
mdistrict.additem "==New Territories=="
mdistrict.additem "Islands"
mdistrict.additem "Kwai Tsing"
mdistrict.additem "North"
mdistrict.additem "Sai Kung"
mdistrict.additem "Sha Tin"
mdistrict.additem "Tai Po"
mdistrict.additem "Tsuen Wan"
mdistrict.additem "Tuen Mun"
mdistrict.additem "Yuen Long"
End Function

'Fills details of all members into the list
Private Function showallm()
memberlist.Clear
Dim rsc As New ADODB.Recordset
DB.Open Path
    rsc.Open "Select * from members", Path, adOpenKeyset, adLockOptimistic
        For i = 1 To rsc.RecordCount
            pad = "                                                                                          "
            m1 = Left(rsc![MemberID] & pad, 6)
            title1 = Left(rsc![Title] & pad, 4)
            surname1 = Left(rsc![surname] & pad, 15)
            forename1 = Left(rsc![forename] & pad, 20)
            contact1 = rsc![contact number]
            email1 = Left(rsc![Email] & pad, 30)
            points1 = Left(rsc![Points] & pad, 5)
            memberlist.additem m1 & Chr(9) & title1 & "    " & surname1 & Chr(9) & forename1 & Chr(9) & contact1 & Chr(9) & email1 & Chr(9) & points1
            rsc.MoveNext
        Next i
    rsc.Close
DB.Close
End Function

'Changes the first letter of the word passed in to upper case
Private Function changecase(obj As String)
If obj <> "" And Asc(Left(obj, 1)) > 96 And Asc(Left(obj, 1)) < 123 Then
    changecase = Chr(Asc(Left(obj, 1)) - 32) & Right(obj, Len(obj) - 1)
Else
    changecase = obj
End If
End Function

'Calls the function that resets the form, combo box mDistrict and mtitle
Private Function resetform()
Call initialise
Call initialmdistrict
Call initialtitle
End Function
