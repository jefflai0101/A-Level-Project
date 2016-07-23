VERSION 5.00
Begin VB.Form changepassword 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "D-Store - Change Password"
   ClientHeight    =   2865
   ClientLeft      =   4515
   ClientTop       =   1470
   ClientWidth     =   4305
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton confirm 
      Caption         =   "Confirm"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox oldpass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox confirmnew 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox newpass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblpass2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "New password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblpass3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Confirm new"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblpass1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Old password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "changepassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Resets the form
Private Sub Form_Load()
Call resetform
End Sub

'Resets the forecolor of changepassword button in mainmenu into white
Private Sub Form_Unload(Cancel As Integer)
Mainmenu.fwfcp.ForeColor = &HFFFFFF
End Sub

'Actions when the confirm button was pressed
Private Sub confirm_Click()
If status = "A" Then Unload users
If checkvalid = True Then
    Dim userid() As Integer
    Dim usname() As String
    Dim pasword() As String
    Dim fulname() As String
    Dim usestatus() As String
    Lines = 0
    Dim fsys As New FileSystemObject
    Dim ts As TextStream
    Dim line As Integer
    Set ts = fsys.OpenTextFile(App.Path & "\databases\users.dat")
    While Not ts.AtEndOfStream
        ts.ReadLine
        Lines = Lines + 1
    Wend
    ts.Close
    ReDim userid(Lines)
    ReDim usname(Lines)
    ReDim pasword(Lines)
    ReDim fulname(Lines)
    ReDim usestatus(Lines)
    i = 1
        Open (App.Path & "\databases\users.dat") For Input As #1
            While Not EOF(1)
                Input #1, field1, field2, field3, field4, field5
                    If uname = field4 Then
                        userid(i) = field1
                        usname(i) = field2
                        pasword(i) = newpass.Text
                        fulname(i) = field4
                        usestatus(i) = field5
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
        Open (App.Path & "\databases\users.dat") For Output As #1
            For i = 1 To Lines
                writein = userid(i) & "," & usname(i) & "," & pasword(i) & "," & fulname(i) & "," & usestatus(i)
                Print #1, writein
            Next i
        Close #1
        MsgBox "Change applied!"
        MsgBox "Relogin required"
        Unload Me
        Login.Show
        Unload Mainmenu
End If
End Sub

'Validation rules of the changepassword process
Private Function checkvalid()
checkvalid = False
If oldpass.Text <> "" And newpass.Text <> "" And confirmnew.Text <> "" Then
    If newpass.Text = confirmnew.Text Then
        If oldpass.Text <> newpass.Text Then
            Open (App.Path & "\databases\users.dat") For Input As #1
                While Not EOF(1)
                    Input #1, field1, field2, field3, field4, field5
                        If field4 = uname Then
                            If oldpass.Text = field3 Then checkvalid = True
                        End If
                Wend
            Close #1
        Else
            MsgBox "The new password can't be the same as old one"
        End If
    Else
        MsgBox "New passwords not match!"
    End If
Else
    MsgBox "Can't have empty field(s)!"
End If
End Function

'Resets the form
Private Function resetform()
oldpass.Text = ""
newpass.Text = ""
confirmnew.Text = ""
End Function
