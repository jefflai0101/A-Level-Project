VERSION 5.00
Begin VB.Form reports 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "D-Store - Reports"
   ClientHeight    =   6435
   ClientLeft      =   6660
   ClientTop       =   1845
   ClientWidth     =   5925
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbyear 
      Height          =   315
      Left            =   960
      TabIndex        =   8
      Text            =   "year"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton printstocklist 
      Caption         =   "Stock List"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton printresult 
      Caption         =   "Print"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame results 
      BackColor       =   &H00000000&
      Height          =   4335
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   4935
      Begin VB.Label tts 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label gp 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label cos 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label sales 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   13
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lbltsm 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Total Sales made"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label lblgp 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Gross Profit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label tpprice 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Cost of Sales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label tprice 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Total Sales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblresultmonth 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "MONTH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.CommandButton calculate 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox cbmonth 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Text            =   "month"
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label year 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblmonth 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As New ADODB.Connection

'Resets the form and calls function that calculates the statistics of this month
Private Sub Form_Load()
Dim cyear As String
Dim monthnumber As String
Call checkyear
cbmonth.Clear
sales.Caption = ""
cos.Caption = ""
gp.Caption = ""
tts.Caption = ""
cyear = Format(Date, "yy")
monthnumber = Format(Date, "mm")
Call calculations(cyear, monthnumber)
End Sub

'Reset the combo box cbmonth
Private Sub cbyear_Click()
Call resetmonth
End Sub

'Calls the function that calculates the statistics of this month
Private Sub calculate_Click()
Dim cyear As String
Dim monthnumber As String
If cbyear.Text <> "" Then cyear = cbyear.Text
If cbmonth.Text <> "" Then monthnumber = "0" & lookforvalue(cbmonth.Text)
If cbyear.Text <> "" And cbmonth.Text <> "" Then Call calculations(cyear, monthnumber)
End Sub

'Deletes the notepad files and reset the colour of report button in the
'mainmenu form into white
Private Sub Form_Unload(Cancel As Integer)
Dim fso As New FileSystemObject
If fso.FileExists(App.Path & "\databases\" & lblresultmonth.Caption & ".txt") = True Then fso.DeleteFile (App.Path & "\databases\" & lblresultmonth.Caption & ".txt")
If fso.FileExists(App.Path & "\databases\Stock List.txt") = True Then fso.DeleteFile (App.Path & "\databases\Stock List.txt")
Mainmenu.fwfreports.ForeColor = &HFFFFFF
End Sub

'Prints out the statistics of the month that is recently displayed
Private Sub printresult_Click()
Dim fso As New FileSystemObject
Dim oFile As File
pathoftext = App.Path & "\databases\" & lblresultmonth.Caption & ".txt"
If fso.FileExists(pathoftext) = True Then fso.DeleteFile (pathoftext)
Call ctxt
Shell ("notepad.exe /p " & pathoftext)
Set oFile = fso.GetFile(pathoftext)
oFile.Delete True
End Sub

'Creates a report of stock using notepad, print out and deletes the file
Private Sub printstocklist_Click()
Dim fso As New FileSystemObject
Dim rsstock As New ADODB.Recordset
    pathoftext = App.Path & "\databases\Stock List.txt"
    fso.CreateTextFile (pathoftext)
        Open (pathoftext) For Output As #1
            pad = "                         "
            pad1 = "================================================================================"
            printemptyline (2)
            Print #1, pad1
            printemptyline (2)
            Print #1, "  [Stock ID]  [Brand]                         [Size] [Cost] [Price][Quantity]"
            printemptyline (2)
            DB.Open Path
            rsstock.Open "Select * from [Stock] where [quantity] > 0", Path, adOpenKeyset, adLockOptimistic
                With rsstock
                For i = 1 To .RecordCount
                    Print #1, "  " & Left(![stockid] & pad, 14) & Left(![brand] & pad, 30) & Chr(9) & Left(![Size] & pad, 2) & "   $" & Left(![pprice] & pad, 5) & "  " & "$" & Left(![price] & pad, 5) & "    " & ![quantity]
                    printemptyline (1)
                    .MoveNext
                Next i
                End With
            rsstock.Close
            DB.Close
            printemptyline (1)
            Print #1, pad1
        Close #1
Shell ("notepad.exe /p " & pathoftext)
fso.DeleteFile (pathoftext)
End Sub

'Prints a empty line into a notepad file
Private Function printemptyline(o)
For i = 1 To o
    Print #1, pad
Next i
End Function

'Creates a report of statistics using notepad
Private Function ctxt()
Dim rssale As New ADODB.Recordset
Dim fso As New FileSystemObject
Dim tstream As TextStream
Dim oFile As File
pathtodb = "provider = microsoft.jet.oledb.4.0; data source = " & App.Path & "\databases\" & Mid(lblresultmonth.Caption, 3, 2) & ".mdb"
pathoftext = App.Path & "\databases\" & lblresultmonth.Caption & ".txt"
If lblresultmonth.Caption <> "MONTH" Then
    fso.CreateTextFile (pathoftext)
    Open (pathoftext) For Output As #1
        pad = "                         "
        pad1 = "================================================================================"
        printemptyline (4)
        Print #1, pad & Left("Sales" & pad, 27) & sales.Caption
        printemptyline (2)
        Print #1, pad & Left("Cost of Sales" & pad, 27) & cos.Caption
        printemptyline (2)
        Print #1, pad & Left("Gross Profit" & pad, 27) & gp.Caption
        printemptyline (2)
        Print #1, pad & Left("Total Sales made" & pad, 27) & tts.Caption
        printemptyline (2)
        Print #1, pad1
        printemptyline (2)
        Print #1, Left(pad & pad, 40) & "Items" & pad & pad
        printemptyline (2)
        Print #1, "    [Sale ID]    [Stock ID]       [Cost]      [Price]   [Quantity][Sale/Order]"
        printemptyline (2)
        DB.Open pathtodb
        target = "0" & lookforvalue(Mid(lblresultmonth.Caption, 6, Len(lblresultmonth.Caption) - 5))
        rssale.Open "select * from [sales] where Month(date) = " & target, pathtodb, adOpenKeyset, adLockOptimistic
            With rssale
            For i = 1 To .RecordCount
                If ![fromorders] = True Then msge = "Order"
                If ![fromorders] = False Then msge = "Sale"
                Print #1, "        " & Left(![saleid] & pad, 8) & Left(![stockid] & pad, 16) & "  " & "$" & Left(![pprice] & pad, 12) & "$" & Left(![price] & pad, 12) & Left(![quantity] & pad, 9) & msge
                printemptyline (1)
                .MoveNext
            Next i
            End With
        rssale.Close
        DB.Close
        printemptyline (1)
        Print #1, pad1
    Close #1
End If
End Function

'Calculates the statistics of for the chosen month
Private Function calculations(cyear As String, monthnumber As String)
Dim pathtotarget As String
Dim rsdb As New ADODB.Recordset
pathtt = "provider = microsoft.jet.oledb.4.0; data source = " & App.Path & "\databases\" & cyear & ".mdb"
    DB.Open pathtt
        rsdb.Open "select * from [sales] where Month(date) = " & monthnumber, pathtt, adOpenKeyset, adLockOptimistic
        If rsdb.RecordCount = 0 Then
            MsgBox "No records found in this month"
            Call checkyear
            cbmonth.Clear
            rsdb.Close
            DB.Close
            Exit Function
        End If
            For i = 1 To rsdb.RecordCount
                If rsdb![quantity] > 1 Then
                    tempprice = Val(tempprice) + (rsdb![price] * rsdb![quantity])
                    temppprice = Val(temppprice) + (rsdb![pprice] * rsdb![quantity])
                Else
                    tempprice = Val(tempprice) + rsdb![price]
                    temppprice = Val(temppprice) + rsdb![pprice]
                End If
                rsdb.MoveNext
            Next i
        sales.Caption = "$" & tempprice
        cos.Caption = "$" & temppprice
        gp.Caption = "$" & Val(tempprice) - Val(temppprice)
        If Val(tempprice) - Val(temppprice) < 0 Then gp.Caption = "($" & Val(tempprice) - Val(temppprice) & ")"
        tts.Caption = rsdb.RecordCount & " times"
        lblresultmonth.Caption = Format(Date, "YYYY") & " " & lookforvalue(Val(monthnumber))
        rsdb.Close
    DB.Close
Call checkyear
cbmonth.Clear
Dim fso As New FileSystemObject
If fso.FileExists(App.Path & "\" & lblresultmonth.Caption & ".txt") = True Then fso.DeleteFile (App.Path & "\" & lblresultmonth.Caption & ".txt")
Call ctxt
End Function

'Shows all valid options of month choices in the combo box cbmonth
Private Function resetmonth()
Dim i As Integer
cbmonth.Clear
If cbyear.Text <> "" Then
    If cbyear.Text = Format(Date, "yy") Then sameyear = True
If sameyear = True Then
    For i = 1 To Val(Format(Date, "mm"))
        cbmonth.additem lookforvalue(i)
    Next i
Else
    For i = 1 To 12
        cbmonth.additem lookforvalue(i)
    Next i
End If
End If
End Function

'look for name or the number of a month with passing in value
Private Function lookforvalue(target)
If target = "January" Then lookforvalue = 1
If target = "Feburary" Then lookforvalue = 2
If target = "March" Then lookforvalue = 3
If target = "April" Then lookforvalue = 4
If target = "May" Then lookforvalue = 5
If target = "June" Then lookforvalue = 6
If target = "July" Then lookforvalue = 7
If target = "August" Then lookforvalue = 8
If target = "September" Then lookforvalue = 9
If target = "October" Then lookforvalue = 10
If target = "November" Then lookforvalue = 11
If target = "December" Then lookforvalue = 12
If target = 1 Then lookforvalue = "January"
If target = 2 Then lookforvalue = "Feburary"
If target = 3 Then lookforvalue = "March"
If target = 4 Then lookforvalue = "April"
If target = 5 Then lookforvalue = "May"
If target = 6 Then lookforvalue = "June"
If target = 7 Then lookforvalue = "July"
If target = 8 Then lookforvalue = "August"
If target = 9 Then lookforvalue = "September"
If target = 10 Then lookforvalue = "October"
If target = 11 Then lookforvalue = "November"
If target = 12 Then lookforvalue = "December"
End Function

'Shows only .mdb files that are named for year
Private Function checkyear()
Dim fso As New FileSystemObject
Dim ofolder As Folder
Dim oFile As File
Dim oFiles As Files
Dim folderpath As String

folderpath = App.Path & "\databases"
Set ofolder = fso.GetFolder(folderpath)
Set oFiles = ofolder.Files
If oFiles.Count > 2 Then
cbyear.Clear
With cbyear
For Each oFile In oFiles
If oFile.Name <> "users.dat" Then
    If oFile.Name <> "Database.mdb" Then
        If Right(oFile.Name, 4) <> ".txt" Then
            If Right(oFile.Name, 4) <> ".ldb" Then
                .additem Left(oFile.Name, 2)
            End If
        End If
    End If
End If
Next
End With
End If
End Function
