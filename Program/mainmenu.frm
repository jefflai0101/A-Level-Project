VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Mainmenu 
   BackColor       =   &H00000000&
   Caption         =   "D-Store - Main Menu"
   ClientHeight    =   11010
   ClientLeft      =   1110
   ClientTop       =   1305
   ClientWidth     =   15240
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11629.95
   ScaleMode       =   0  'User
   ScaleWidth      =   17339.19
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11055
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Width           =   13095
   End
   Begin VB.Timer mainmenutime 
      Interval        =   100
      Left            =   11760
      Top             =   120
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   10635
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label logout 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Log out"
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
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label fwforders 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Orders"
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
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label fwfmembers 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Members"
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
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label fwfbuac 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Backup/Archive"
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
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label fwfcp 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Change Password"
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
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label fwfu 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Users"
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
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label fwfreports 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Reports"
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
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label fwfstock 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stoc&k"
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
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label fwfsales 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Sales"
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
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label maintext 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "D-Store"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.Line separateline 
      BorderColor     =   &H00FFFF00&
      X1              =   2047.936
      X2              =   2047.936
      Y1              =   0
      Y2              =   7506.124
   End
End
Attribute VB_Name = "Mainmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'Calls functions that open different forms according to what button was pressed
Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case (KeyAscii)
Case 115
    Call fwfsales_Click
Case 107
    Call fwfstock_Click
Case 111
    Call fwforders_Click
Case 109
    Call fwfmembers_Click
Case 117
    Call fwfu_Click
Case 114
    Call fwfreports_Click
Case 98
    Call fwfbuac_Click
Case 99
    Call fwfcp_Click
Case 108
    Call logout_Click
End Select
End Sub

'Resets the form, points the path of the database and checks the year now and
'create new database for this year if there is no existence of this year's database
Private Sub Form_Load()
    relog = False
    Path = "provider = microsoft.jet.oledb.4.0; data source = " & App.Path & "\Databases\" & Format(Date, "yy") & ".mdb"
    Call checkdatabase
    StatusBar.Panels.Add (1)
    StatusBar.Panels.Add (2)
    StatusBar.Panels.Add (3)
    StatusBar.Panels.Item(1).AutoSize = sbrContents
    StatusBar.Panels.Item(2).AutoSize = sbrContents
    StatusBar.Panels.Item(3).AutoSize = sbrContents
    StatusBar.Panels.Item(4).AutoSize = sbrContents
    separateline.Y2 = Me.Height
    maintext.Visible = True
    Call initialbuttons
    If status = "U" Then Call userright
    If status = "S" Then Call stockright
    If status = "B" Then Call bothright
    Unload users
    Unload sales
    Unload stock
    Unload orders
    Unload members
    Unload reports
    Unload changepassword
    Unload backuparchive
End Sub

'Resets the buttons position and visibility for Admin User Right
Private Function initialbuttons()
    fwfsales.Visible = True
        fwfsales.Top = 1057.269
    fwfstock.Visible = True
        fwfstock.Top = 1691.63
    fwforders.Visible = True
        fwforders.Top = 2325.991
    fwfmembers.Visible = True
        fwfmembers.Top = 2960.352
    fwfu.Visible = True
        fwfu.Top = 3594.713
    fwfreports.Visible = True
        fwfreports.Top = 4229.074
    fwfbuac.Visible = True
        fwfbuac.Top = 4863.435
    fwfcp.Visible = True
        fwfcp.Top = 5497.796
    logout.Top = 6132.157
End Function

'Resets the buttons position and visibility for User User Right
Private Function userright()
    fwfcp.Top = 3594.713
    logout.Top = 4229.074
    fwfu.Visible = False
    fwfreports.Visible = False
    fwfbuac.Visible = False
End Function

'Resets the buttons position and visibility for Stockman User Right
Private Function stockright()
    fwfstock.Top = 1057.269
    fwfcp.Top = 1691.63
    logout.Top = 2325.991
    fwfsales.Visible = False
    fwforders.Visible = False
    fwfmembers.Visible = False
    fwfu.Visible = False
    fwfreports.Visible = False
    fwfbuac.Visible = False
End Function

'Resets the buttons position and visibility for User and Stockman User Right
Private Function bothright()
    fwfcp.Top = 3594.713
    logout.Top = 4229.074
    fwfu.Visible = False
    fwfreports.Visible = False
    fwfbuac.Visible = False
End Function

'Resets the size of the form
Private Sub Form_Resize()
    Call resizeform
End Sub

'Logs out when the cross button on the top right corner is pressed
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    Login.Show
End Sub

'Calls the action of unload mainmenu
Private Sub logout_Click()
    Call Form_Unload(1)
End Sub

'Updates the time, date, weekday and user's name every second
Private Sub mainmenutime_Timer()
    StatusBar.Panels.Item(1).Text = Format(Time, "HH:MM:SS AMPM")
    StatusBar.Panels.Item(2).Text = Day(Date) & "/" & MonthName(Month(Date)) & "/" & Year(Date)
    StatusBar.Panels.Item(3).Text = WeekdayName(Weekday(Date), , vbSunday)
    StatusBar.Panels.Item(4).Text = "Current user: " & uname
End Sub

'Shows form within the frame in the mainmenu form
Private Sub fwfu_Click()
    users.Show
    fwfu.ForeColor = &HFF&
    FormWithinForm Me.Frame, users
End Sub

'Shows form within the frame in the mainmenu form
Private Sub fwfreports_Click()
    reports.Show
    fwfreports.ForeColor = &HFF&
    FormWithinForm Me.Frame, reports
End Sub

'Shows form within the frame in the mainmenu form
Private Sub fwfsales_Click()
    sales.Show
    fwfsales.ForeColor = &HFF&
    FormWithinForm Me.Frame, sales
End Sub

'Shows form within the frame in the mainmenu form
Private Sub fwforders_Click()
    orders.Show
    fwforders.ForeColor = &HFF&
    FormWithinForm Me.Frame, orders
End Sub

'Shows form within the frame in the mainmenu form
Private Sub fwfmembers_Click()
    members.Show
    fwfmembers.ForeColor = &HFF&
    FormWithinForm Me.Frame, members
End Sub

'Shows form within the frame in the mainmenu form
Private Sub fwfstock_Click()
    stock.Show
    fwfstock.ForeColor = &HFF&
    FormWithinForm Me.Frame, stock
End Sub

'Shows form within the frame in the mainmenu form
Private Sub fwfbuac_Click()
    backuparchive.Show
    fwfbuac.ForeColor = &HFF&
    FormWithinForm Me.Frame, backuparchive
End Sub

'Shows form within the frame in the mainmenu form
Private Sub fwfcp_Click()
    changepassword.Show
    fwfcp.ForeColor = &HFF&
    FormWithinForm Me.Frame, changepassword
End Sub

'====================================Procedures/Functions========================================
Private Function checkdatabase()
Dim fso As New FileSystemObject
    'Checks the existence of this year's database and the default database
    thisyear = Format(Date, "yy") & ".mdb"
    fileexist = fso.FileExists(App.Path & "\databases\database.mdb")
    If fileexist = False Then
        MsgBox "Unable to locate the database, please contact your technician"
        End
    End If
    fileexist = fso.FileExists(App.Path & "\databases\" & thisyear)
    'If not exist (usually means new year) then create DB for this year by copying the default DB and name by year "yy"
    If fileexist = False Then
        fso.CopyFile App.Path & "\Databases\database.mdb", App.Path & "\Databases\" & thisyear, True
        lastyear = "0" & (Val(Format(Date, "yy")) - 1) & ".mdb"
        fileexist = fso.FileExists(App.Path & "\databases\" & lastyear)
        'Check does last year's database exist
        'If exist then copy the table "Orders", "Members" and "Stock" into this year's database
        If fileexist = True Then
            lypath = "provider = microsoft.jet.oledb.4.0; data source = " & App.Path & "\Databases\" & lastyear
            thisyearpath = App.Path & "\Databases\" & thisyear
            typath = "provider = microsoft.jet.oledb.4.0; data source = " & App.Path & "\Databases\" & thisyear
            Dim tablenames() As String
            ReDim tablenames(3)
            tablenames(1) = "Members"
            tablenames(2) = "Orders"
            tablenames(3) = "Stock"
            For i = 1 To 3
                Call deletetable(tablenames(i), typath)
                Call copytable(thisyearpath, tablenames(i), tablenames(i), lypath)
            Next i
        End If
    End If
End Function

'Reset the size of the form
Private Function resizeform()
If Me.Height > 2235 Then
    maintext.Left = (separateline.X1 + Me.Width - maintext.Width) / 2
    Frame.Left = separateline.X2 + 100
    Frame.Width = Me.Width
    Frame.Height = Me.Height - (StatusBar.Height + Frame.Top)
End If
End Function

'Function that make forms show within another form
Public Function FormWithinForm(Parent As Object, Child As Object)
On Error Resume Next
SetParent Child.hWnd, Parent.hWnd
FormWithinForm = (Err.Number = 0 And Err.LastDllError = 0)
End Function

'Copies table from one database into another
Public Sub copytable(ByVal dbpath$, ByVal targetTable$, ByVal NewTableName$, path1)
        DB.Open path1
            DB.Execute "SELECT " & targetTable & ".* INTO " & targetTable & " IN '" & dbpath & "' From " & NewTableName
        DB.Close
End Sub

'Deletes a certain table from a database
Public Sub deletetable(ByVal targetTable$, path1)
DB.Open path1
    DB.Execute "Drop table " & targetTable
DB.Close
End Sub
