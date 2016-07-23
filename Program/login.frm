VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Login 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "D-Store - Login"
   ClientHeight    =   3315
   ClientLeft      =   2385
   ClientTop       =   4725
   ClientWidth     =   4365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timetologinagain 
      Left            =   3720
      Top             =   1680
   End
   Begin VB.Timer statusbartime 
      Interval        =   100
      Left            =   3720
      Top             =   2280
   End
   Begin VB.Frame frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   3615
      Begin VB.TextBox username 
         Height          =   405
         Left            =   960
         TabIndex        =   1
         Top             =   120
         Width           =   2415
      End
      Begin VB.TextBox password 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton loginconfirm 
         Caption         =   "&Login"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton exitorcancel 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label txtusername 
         BackColor       =   &H00000000&
         Caption         =   "Username"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
      Begin VB.Label txtpassword 
         BackColor       =   &H00000000&
         Caption         =   "Password"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Dstore 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "D-Store"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   8
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tries As Integer
Dim counttime As Integer
'=====================================Login==========================================
'Resets the form and ends the program if the same process is being run already
Private Sub Form_Load()
If App.PrevInstance Then End
Call resetform
End Sub

'Closes the form when press exit button and resets the form when cancel
Private Sub exitorcancel_click()
If exitorcancel.Enabled = True Then
    If exitorcancel.Caption = "&Exit" Then End
    If exitorcancel.Caption = "&Cancel" Then Call resetform
    username.SetFocus
End If
End Sub

'Updates the statusbar every second
Private Sub statusbartime_Timer()
StatusBar.Panels.Item(1).Width = "2000"
StatusBar.Panels.Item(1).Text = Now()
End Sub

'Calls the function that checks the fields are empty or not to enable the "Enter"
'button and changes the button "Exit" to "Cancel"
Private Sub username_Change()
Call enableloginconfirm
Call exitorclear
End Sub

'Calls the function that checks the fields are empty or not to enable the "Enter"
'button and changes the button "Exit" to "Cancel"
Private Sub password_Change()
Call enableloginconfirm
Call exitorclear
End Sub

'Checks the username and password against the User database and if correct unloads
'the login form and shows the mainmenu form, or calls the function that records the
'attempt of tries to login
Private Sub loginconfirm_Click()
Dim userfound As Boolean
Dim fso As New FileSystemObject
fileexist = fso.FileExists(App.Path & "\databases\users.dat")
userfound = False
    If fileexist = False Then
        MsgBox "No user records!"
        Call resetform
        username.SetFocus
    Else
        Open (App.Path & "\databases\users.dat") For Input As #1
            While Not EOF(1) And userfound = False
                Input #1, field1, field2, field3, field4, field5
                User = field2
                pass = field3
                uname = field4
                tempstatus = field5
                If LCase(username.Text) = LCase(User) Then
                        If LCase(password.Text) = LCase(pass) Then
                            status = tempstatus
                            Call resetform
                            userfound = True
                            Unload Me
                            Mainmenu.Show
                            logintime = Time()
                        Else
                            Call faillogin
                            Close #1
                            Exit Sub
                        End If
                    Else
                        If EOF(1) = True Then Call faillogin
                    End If
            Wend
        Close #1
    End If
End Sub

'Counts the time of access denied and when time is over the user will be allowed
'to login again
Private Sub timetologinagain_Timer()
If counttime = 0 Then
    timetologinagain.Interval = 0
    exitorcancel.Enabled = True
    tries = 0
    StatusBar.Panels.Remove (2)
    timetologinagain.Interval = 0
Else
    counttime = counttime - 1
    StatusBar.Panels.Item(2).Text = counttime & " seconds left"
End If
End Sub

'When press the enter will calls the button "Enter"
Private Sub password_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And loginconfirm.Enabled = True Then
        Call loginconfirm_Click
    End If
End Sub

'When press the enter will calls the button "Enter"
Private Sub username_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And loginconfirm.Enabled = True Then
        Call loginconfirm_Click
    End If
End Sub

'===========================Procedures/Functions===================================
'Resets the form and disables the "Enter" button
Private Function resetform()
username.Text = ""
password.Text = ""
Call enableloginconfirm
End Function

'Changes the text of Exit/Cancel button by checking are the fields username and
'password empty
Private Function exitorclear()
If username.Text <> "" Or password.Text <> "" Then
    exitorcancel.Caption = "&Cancel"
Else
    exitorcancel.Caption = "&Exit"
End If
End Function

'Changes the text of Exit/Cancel button by checking are the fields username and
'password empty
Private Function enableloginconfirm()
If username.Text = "" And password.Text = "" Then
    loginconfirm.Enabled = False
Else
    If tries < 4 Then loginconfirm.Enabled = True
End If
End Function

'Adds 1 to the variable tries
Private Function attempts()
tries = tries + 1
End Function

'Function that ban the access to the program for 5 minutes if too much tries
Private Function accessdenied()
If tries > 3 Then
    MsgBox "Access banned for 5 minutes"
    exitorcancel.Enabled = False
    counttime = 300
    timetologinagain.Interval = "1000"
    StatusBar.Panels.Add (2)
End If
If counttime = 0 Then
    timetologinagain.Interval = "0"
    tried = 0
End If
End Function

'Resets the form and calls up functions that checks the attempt of login
Private Function faillogin()
MsgBox "Incorrect username or password!"
uname = ""
status = ""
username.SetFocus
Call attempts
Call accessdenied
Call resetform
End Function
