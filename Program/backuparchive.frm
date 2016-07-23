VERSION 5.00
Begin VB.Form backuparchive 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Backup and Archive"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3915
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer refreshdrivecount 
      Interval        =   1000
      Left            =   3240
      Top             =   2400
   End
   Begin VB.CommandButton confirm 
      Caption         =   "Confirm"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
   End
   Begin VB.ComboBox dselect 
      Height          =   315
      Left            =   480
      TabIndex        =   5
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Frame options 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   2895
      Begin VB.OptionButton archive 
         BackColor       =   &H00000000&
         Caption         =   "Archive"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton backup 
         BackColor       =   &H00000000&
         Caption         =   "Backup"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.ComboBox cdbyear 
      Height          =   300
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label dbyear 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "backuparchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numberoffound As Integer
'Resets the form
Private Sub Form_Load()
Call checkfiles
Call checkdrive
backup.Value = False
archive.Value = False
End Sub

'Calls up the function that checks the existence of databases in the database folder
Private Sub cdbyear_Change()
Call checkfiles
End Sub

'Calls the function that checks the existence of the USB drive every 10 seconds
Private Sub refreshdrivecount_Timer()
Dim fso As New FileSystemObject
Dim odrives As Drives
Set odrives = fso.Drives
drcount = odrives.Count
If drcount > numberoffound Then
    numberoffound = drcount
    Call checkdrive
    MsgBox "Drive found"
End If
If drcount < numberoffound Then Call checkdrive
End Sub

'Reconfirm when User tries to backup or archive database
Private Sub confirm_Click()
If backup.Value <> False Or archive.Value <> False And cdbyear.Text <> "" Then
    If backup.Value = True Then baword = "Are you sure you want to backup this file?"
    If archive.Value = True Then baword = "Are you sure you want to archive this file?"
    If dselect.Text = "Drive not found" Then
        MsgBox "Can't locate usb drive!"
    Else
        Dim Msg, Style, Answer
                Msg = baword
                Style = vbYesNo + vbInformation
                Answer = MsgBox(Msg, Style, "Confirm")
                If Answer = vbYes Then
                    If backup.Value = True Then Call backupfile(cdbyear)
                    If archive.Value = True Then Call archivefile(cdbyear)
                End If
    End If
End If
End Sub

'Resets the forecolor of backup/archive button in mainmenu into white
Private Sub Form_Unload(Cancel As Integer)
Mainmenu.fwfbuac.ForeColor = &HFFFFFF
End Sub
'==============================Procedures/Functions================================
'Backup database onto the USB drive
Private Function backupfile(obj As Object)
Dim fso As New FileSystemObject
Dim thisfile As File
Set thisfile = fso.GetFile(App.Path & "\databases\" & obj & ".mdb")
    If fso.FolderExists(dselect.Text & "databases\") = False Then fso.CreateFolder (dselect.Text & "databases\")
    thisfile.Copy dselect.Text & "databases\"
MsgBox "Backup complete"
Call Form_Load
End Function

'Archives database onto the USB drive
Private Function archivefile(obj As Object)
Dim fso As New FileSystemObject
Dim thisfile As File
Set thisfile = fso.GetFile(App.Path & "\databases\" & obj & ".mdb")
allowtoarchive = True
If Left(cdbyear.Text, 2) = Format(Date, "yy") Then
    MsgBox "Not allow to archive the recent year's record!"
    allowtoarchive = False
End If
If allowtoarchive = True Then
    If fso.FolderExists(dselect.Text & "databases\") = False Then fso.CreateFolder (dselect.Text & "databases\")
    thisfile.Copy dselect.Text & "databases\", True
    thisfile.Delete True
    MsgBox "Archive complete"
End If
Call Form_Load
End Function

'Check all the databases' name in the folder Databases
Private Function checkfiles()
Dim fso As New FileSystemObject
Dim ofolder As Folder
Dim oCurrentFile As File
Dim oFileColl As Files
Dim folderpath As String

folderpath = App.Path & "\databases"
Set ofolder = fso.GetFolder(folderpath)
Set oFileColl = ofolder.Files
If oFileColl.Count > 2 Then
cdbyear.Clear
With cdbyear
For Each oCurrentFile In oFileColl
If oCurrentFile.Name <> "users.dat" Then
    If oCurrentFile.Name <> "Database.mdb" Then
        If Right(oCurrentFile.Name, 4) <> ".txt" Then
            If Right(oCurrentFile.Name, 4) <> ".ldb" Then
                .AddItem Left(oCurrentFile.Name, 2)
            End If
        End If
    End If
End If
Next
End With
End If
End Function

'Checks the existence of the USB drive
Private Function checkdrive()
Dim fso As New FileSystemObject
Dim oDrive As Drive
Dim odrives As Drives
Dim oFile As File
Set odrives = fso.Drives
dselect.Enabled = False
dselect.Text = "Drive not found"
Dim drivesfound As Integer
drivesfound = 0
numberoffound = odrives.Count
For Each oDrive In odrives
    Select Case oDrive.DriveType
        Case 1
            dselect.AddItem oDrive.DriveLetter & ":"
            If oDrive.IsReady = True Then
                For Each oFile In oDrive.RootFolder.Files
                    If oFile.Name = "uniquedriveidentifier.dat" And drivesfound = 0 Then
                        dselect.Text = oDrive.DriveLetter & ":\"
                        drivesfound = 1
                    ElseIf drivesfound > 1 Then
                        dselect.AddItem oDrive.DriveLetter & ":\"
                        drivesfound = drivesfound + 1
                        dselect.Enabled = True
                    End If
                Next
            End If
    End Select
Next oDrive
End Function

