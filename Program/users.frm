VERSION 5.00
Begin VB.Form users 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "D-Store - Users"
   ClientHeight    =   5580
   ClientLeft      =   1860
   ClientTop       =   1995
   ClientWidth     =   8865
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox password 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   6600
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox repassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   6600
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox fullname 
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.ComboBox userstatus 
      Height          =   315
      Left            =   6600
      TabIndex        =   5
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton confirm 
      Caption         =   "Confirm"
      Default         =   -1  'True
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   4800
      Width           =   1935
   End
   Begin VB.ListBox Title 
      Height          =   255
      ItemData        =   "users.frx":0000
      Left            =   240
      List            =   "users.frx":0002
      TabIndex        =   15
      Top             =   1440
      Width           =   5055
   End
   Begin VB.Frame menu 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   720
      TabIndex        =   12
      Top             =   240
      Width           =   7455
      Begin VB.Label deleteuser 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Delete user"
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
         Left            =   5640
         TabIndex        =   16
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label edituser 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Edit user"
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
         Left            =   3120
         TabIndex        =   14
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label adduser 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Add user"
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
         Left            =   480
         TabIndex        =   13
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.ListBox userlist 
      Height          =   2985
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   5055
   End
   Begin VB.TextBox username 
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblname 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   10
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lbluserstatus 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "User status"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   9
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblrepassword 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Re-enter"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblpassword 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Password"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblusername 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Username"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   0
      Top             =   1560
      Width           =   855
   End
End
Attribute VB_Name = "users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim useroption As Integer
Dim idofuser As Integer
Dim selectedanitem As Boolean

'Resets the form
Private Sub Form_Load()
Title.additem "[ID]" & "  [Username]" & "    [Password]" & "               [Fullname]" & "            [Status]"
Call adduser_Click
End Sub

'Resets the forecolor of user button in mainmenu into white
Private Sub Form_Unload(Cancel As Integer)
Mainmenu.fwfu.ForeColor = &HFFFFFF
End Sub

'Changes to add user mode and resets the form
Private Sub adduser_Click()
useroption = 1
username.Enabled = True
fullname.Enabled = True
adduser.ForeColor = &HFF&
edituser.ForeColor = &HFFFFFF
password.PasswordChar = "*"
repassword.PasswordChar = "*"
selectedanitem = False
Call resetform
End Sub

'Changes to edit user mode and resets the form
Private Sub edituser_Click()
useroption = 2
username.Enabled = False
fullname.Enabled = False
adduser.ForeColor = &HFFFFFF
edituser.ForeColor = &HFF&
selectedanitem = False
Call resetform
End Sub

'Calls function to delete user when in edit mode
Private Sub deleteuser_Click()
If useroption = 2 And selectedanitem = True Then Call checkdelete
End Sub

'Adds user or changes the details of user
Private Sub confirm_Click()
Lines = 1
Select Case useroption
Case 1
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim line As Integer
    Set ts = fso.OpenTextFile(App.Path & "\databases\users.dat")
    While Not ts.AtEndOfStream
        ts.ReadLine
        Lines = Lines + 1
    Wend
    ts.Close
    checkok = checkvalid()
    If checkok = True Then
        If checkusername = True Then
            Open (App.Path & "\databases\users.dat") For Append As #1
            writein = Lines & "," & username.Text & "," & password.Text & "," & fullname.Text & "," & UCase(userstatus.Text)
                Print #1, writein
            Close #1
            MsgBox "User created successfully!"
            Call resetform
        End If
    End If
Case 2
    If selectedanitem = True Then
        checkok = checkvalid()
        If checkok = True Then
            Dim userid() As Integer
            Dim usname() As String
            Dim pasword() As String
            Dim fulname() As String
            Dim usestatus() As String
            i = userlist.ListCount
            ReDim userid(i)
            ReDim usname(i)
            ReDim pasword(i)
            ReDim fulname(i)
            ReDim usestatus(i)
            i = 1
            If checkusername = True Then
            allowtochange = True
                Open (App.Path & "\databases\users.dat") For Input As #1
                    While Not EOF(1)
                        Input #1, field1, field2, field3, field4, field5
                        If field1 = idofuser Then
                            userid(i) = idofuser
                            usname(i) = username.Text
                            pasword(i) = password.Text
                            fulname(i) = fullname.Text
                            usestatus(i) = userstatus.Text
                            If fulname(i) = uname Then allowtochange = False
                        Else
                            userid(i) = field1
                            usname(i) = field2
                            pasword(i) = field3
                            fulname(i) = field4
                            usestatus(i) = field5
                        End If
                        i = i + 1
                    Wend
                Close #1
                If allowtochange = True Then
                    Open (App.Path & "\databases\users.dat") For Output As #1
                        For i = 1 To userlist.ListCount
                            writein = userid(i) & "," & usname(i) & "," & pasword(i) & "," & fulname(i) & "," & usestatus(i)
                            Print #1, writein
                        Next i
                    Close #1
                    MsgBox "Change applied!"
                Else
                    MsgBox "Not allow to change own account in this mode"
                End If
            End If
            Call resetform
        End If
    End If
    Call loadtable
    selectedanitem = False
End Select
End Sub

'Fills all details of that user when the chosen user is clicked in the list
Private Sub userlist_Click()
If useroption = 2 Then
    numberofuser = userlist.ListCount
    For RecentUser = 0 To numberofuser - 1
        If userlist.Selected(RecentUser) = True Then
            idofuser = RecentUser + 1
            RecentUser = numberofuser
        End If
    Next RecentUser
    Open (App.Path & "\databases\users.dat") For Input As #1
        While Not EOF(1) And userfound = False
            Input #1, field1, field2, field3, field4, field5
                If idofuser = field1 Then
                    username.Text = field2
                    password.Text = field3
                    password.PasswordChar = ""
                    repassword.PasswordChar = ""
                    fullname.Text = field4
                    userstatus.Text = field5
                End If
        Wend
    Close #1
selectedanitem = True
End If
End Sub
'============================Procedures/Functions===============================
'Checks the existence of a username
Private Function checkusername()
checkusername = True
If useroption = 1 Then
    Open (App.Path & "\databases\users.dat") For Input As #1
        While Not EOF(1)
            Input #1, field1, field2, field3, field4, field5
                If LCase(username.Text) = LCase(field2) Then
                    checkusername = False
                    MsgBox "Username already used!"
                    Close #1
                    Exit Function
                End If
        Wend
    Close #1
End If
End Function

'Validation rules of adding and changing user details
Private Function checkvalid()
checkvalid = False
    If username.Text <> "" And Len(username.Text) <= 10 Then
        lblusername.ForeColor = &HFFFFFF
        For checkno = 1 To Len(username.Text)
            If Asc(Mid(username.Text, checkno, 1)) < 48 Or Asc(Mid(username.Text, checkno, 1)) > 122 Or (Mid(username.Text, checkno, 1)) = ":" Or (Mid(username.Text, checkno, 1)) = ";" Or (Mid(username.Text, checkno, 1)) = "<" Or (Mid(username.Text, checkno, 1)) = "=" Or (Mid(username.Text, checkno, 1)) = ">" Or (Mid(username.Text, checkno, 1)) = "?" Or (Mid(username.Text, checkno, 1)) = "@" Or (Mid(username.Text, checkno, 1)) = "[" Or (Mid(username.Text, checkno, 1)) = "\" Or (Mid(username.Text, checkno, 1)) = "\" Or (Mid(username.Text, checkno, 1)) = "^" Or (Mid(username.Text, checkno, 1)) = "_" Or (Mid(username.Text, checkno, 1)) = "'" Then
                checkno = Len(username.Text)
                MsgBox "Can't have symbols for username!!"
                GoTo Quit
            End If
        Next checkno
        If password.Text <> "" And Len(password.Text) <= 15 Then
            lblpassword.ForeColor = &HFFFFFF
            For checkno = 1 To Len(password.Text)
                If Asc(Mid(password.Text, checkno, 1)) < 48 Or Asc(Mid(password.Text, checkno, 1)) > 122 Or (Mid(password.Text, checkno, 1)) = ":" Or (Mid(password.Text, checkno, 1)) = ";" Or (Mid(password.Text, checkno, 1)) = "<" Or (Mid(password.Text, checkno, 1)) = "=" Or (Mid(password.Text, checkno, 1)) = ">" Or (Mid(password.Text, checkno, 1)) = "?" Or (Mid(password.Text, checkno, 1)) = "@" Or (Mid(password.Text, checkno, 1)) = "[" Or (Mid(password.Text, checkno, 1)) = "\" Or (Mid(password.Text, checkno, 1)) = "\" Or (Mid(password.Text, checkno, 1)) = "^" Or (Mid(password.Text, checkno, 1)) = "_" Or (Mid(password.Text, checkno, 1)) = "'" Then
                    checkno = Len(password.Text)
                    MsgBox "Cannot have symbols for password!!"
                    GoTo Quit
                End If
            Next checkno
            If fullname.Text <> "" And Len(fullname.Text) <= 15 Then
                lblname.ForeColor = &HFFFFFF
                If password.Text = repassword.Text Then
                    If UCase(userstatus.Text) = "U" Or UCase(userstatus.Text) = "A" Or UCase(userstatus.Text) = "S" Or UCase(userstatus.Text) = "B" Then
                        lbluserstatus.ForeColor = &HFFFFFF
                        checkvalid = True
                    Else
                        lbluserstatus.ForeColor = &HFF&
                        If userstatus.Text = "" Then MsgBox "Field user status can't be empty!"
                    End If
                Else
                    lblrepassword.ForeColor = &HFF&
                    If repassword.Text = "" Then MsgBox "Please enter the password again to confirm"
                    If password.Text <> repassword.Text And repassword.Text <> "" Then MsgBox "Passwords not match!"
                End If
            Else
                lblname.ForeColor = &HFF&
                If Len(fullname.Text) > 15 Then MsgBox "The name is too long!"
                If fullname.Text = "" Then MsgBox "Field fullname can't be empty!"
            End If
        Else
            lblpassword.ForeColor = &HFF&
            If Len(password.Text) > 15 Then MsgBox "The password is too long!"
            If password.Text = "" Then MsgBox "Field password can't be empty!"
        End If
    Else
        lblusername.ForeColor = &HFF&
        If Len(username.Text) > 15 Then MsgBox "The username is too long!"
        If username.Text = "" Then MsgBox "Field username can't be empty!"
    End If
    
Quit:
Exit Function
End Function

'Confirmation of deleting a user and delete the chosen user if confirm
Private Function checkdelete()
Dim Msg, Style, Answer
    Msg = ("Are you sure you want to delete this user?")
    Style = vbYesNo + vbInformation
    Answer = MsgBox(Msg, Style, "Confirm")
If Answer = vbYes Then
        Dim userid() As Integer
        Dim usname() As String
        Dim pasword() As String
        Dim fulname() As String
        Dim usestatus() As String
        i = userlist.ListCount
        ReDim userid(i)
        ReDim usname(i)
        ReDim pasword(i)
        ReDim fulname(i)
        ReDim usestatus(i)
        i = 1
        Open (App.Path & "\databases\users.dat") For Input As #1
            While Not EOF(1)
                Input #1, field1, field2, field3, field4, field5
                If field1 = idofuser Then
                    If uname = field2 Then
                        MsgBox "You can't delete your own account!"
                        Close #1
                        Exit Function
                    Else
                        Reorder = 1
                        i = i - 1
                    End If
                Else
                    userid(i) = field1 - Reorder
                    usname(i) = field2
                    pasword(i) = field3
                    fulname(i) = field4
                    usestatus(i) = field5
                End If
                i = i + 1
            Wend
        Close #1
        Open (App.Path & "\databases\users.dat") For Output As #1
            For i = 1 To userlist.ListCount - 1
                writein = userid(i) & "," & usname(i) & "," & pasword(i) & "," & fulname(i) & "," & usestatus(i)
                Print #1, writein
            Next i
        Close #1
        Call resetform
        MsgBox "User deleted!"
End If
Call loadtable
End Function

'Calls functions that resets the form
Private Function resetform()
Call initialform
Call initialstatus
Call loadtable
Call resetwordc
End Function

'Empties the fields
Private Function initialform()
    username.Text = ""
    password.Text = ""
    repassword.Text = ""
    fullname.Text = ""
End Function

'Resets the combo box userstatus and adds options of user access rights
Private Function initialstatus()
    userstatus.Clear
    userstatus.additem "U"
    userstatus.additem "S"
    userstatus.additem "B"
    userstatus.additem "A"
End Function

'Loads the user details into the list
Private Function loadtable()
userlist.Clear
Open (App.Path & "\databases\users.dat") For Input As #1
    While Not EOF(1) And userfound = False
        Input #1, field1, field2, field3, field4, field5
        pad = "                                                                                          "
        username1 = Left(field2 & pad, 10)
        password1 = Left(field3 & pad, 20)
        fullname1 = Left(field4 & pad, 20)
        userlist.additem "[" & field1 & "]    " & username1 & Chr(9) & password1 & Chr(9) & fullname1 & Chr(9) & field5
    Wend
Close #1
End Function

'Resets the forcolor of the text next to the text box
Private Function resetwordc()
lblusername.ForeColor = &HFFFFFF
lblpassword.ForeColor = &HFFFFFF
lblrepassword.ForeColor = &HFFFFFF
lblname.ForeColor = &HFFFFFF
lbluserstatus.ForeColor = &HFFFFFF
End Function
