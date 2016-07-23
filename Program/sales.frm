VERSION 5.00
Begin VB.Form sales 
   BackColor       =   &H00000000&
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "D-Store - Sales"
   ClientHeight    =   6015
   ClientLeft      =   630
   ClientTop       =   1800
   ClientWidth     =   10590
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox heading 
      Height          =   240
      Left            =   240
      TabIndex        =   18
      Top             =   1080
      Width           =   7455
   End
   Begin VB.CommandButton removeitem 
      Caption         =   "Remove item"
      Height          =   495
      Left            =   9120
      TabIndex        =   5
      Top             =   285
      Width           =   1215
   End
   Begin VB.Frame customerdetails 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   8040
      TabIndex        =   10
      Top             =   960
      Width           =   2415
      Begin VB.CommandButton candis 
         Caption         =   "Cancel Discount"
         Height          =   495
         Left            =   360
         TabIndex        =   21
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton ok 
         Caption         =   "OK"
         Height          =   375
         Left            =   1800
         TabIndex        =   19
         Top             =   1920
         Width           =   495
      End
      Begin VB.ComboBox pointsuse 
         Height          =   315
         Left            =   360
         TabIndex        =   14
         Top             =   1920
         Width           =   1335
      End
      Begin VB.ComboBox memberid 
         Height          =   315
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label tprice 
         BackColor       =   &H00000000&
         Caption         =   "total price"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label lblbefore 
         BackColor       =   &H00000000&
         Caption         =   "1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblafter 
         BackColor       =   &H00000000&
         Caption         =   "1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label lblmbptas 
         BackColor       =   &H00000000&
         Caption         =   "Points Gained"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label lblmbpt 
         BackColor       =   &H00000000&
         Caption         =   "Member points"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblmemberid 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Member ID"
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
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton additem 
      Caption         =   "Add item"
      Height          =   495
      Left            =   7680
      TabIndex        =   4
      Top             =   285
      Width           =   1215
   End
   Begin VB.CommandButton confirm 
      Caption         =   "Confirm Sales"
      Height          =   495
      Left            =   8400
      TabIndex        =   9
      Top             =   5400
      Width           =   1695
   End
   Begin VB.ComboBox discount 
      Height          =   315
      Left            =   6720
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.ComboBox quantity 
      Height          =   315
      Left            =   4320
      TabIndex        =   2
      Text            =   "quantity"
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox stockid 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.ListBox saleslist 
      Height          =   3660
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   7455
   End
   Begin VB.Label lbldiscount 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      Caption         =   "Discount"
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
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblquantity 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      Caption         =   "Quantity"
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
      Left            =   3000
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblstockid 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      Caption         =   "Stock ID"
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
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Calls the function that resets the form
Private Sub Form_Load()
Call resetform
heading.additem "    [Stock ID]" & Chr(9) & "[Brand]" & Chr(9) & Chr(9) & "         [Quantity]" & Chr(9) & "[Price]" & Chr(9) & "[Discount]"
End Sub

'Resets the forecolor of sales button in mainmenu into white
Private Sub Form_Unload(Cancel As Integer)
Mainmenu.fwfsales.ForeColor = &HFFFFFF
End Sub

'Adds the item into the sale list, calculates the points after the sale
Private Sub additem_Click()
Dim price As Integer
If stockid.Text <> "" And quantity.Text <> "" Then
    Dim rsstock As New ADODB.Recordset
    DB.Open Path
        rsstock.Open "Select * from stock where [stockid] = '" & stockid.Text & "'", Path, adOpenKeyset, adLockOptimistic
            If rsstock.RecordCount = 0 Then
                MsgBox "Invalid stock ID"
            Else
                stockid1 = rsstock![stockid]
                price = rsstock![price]
                brand1 = rsstock![brand]
            End If
        rsstock.Close
    DB.Close
    Dim discountrate As Integer
    If discount.Text <> "" Then
        If quantity > 1 Then
            MsgBox "Discount only allow for 1 item!"
            Exit Sub
        Else
            discountrate = Val(Left(discount.Text, 2))
            price = pricecalculation(price, discountrate)
        End If
    End If
    
    pad = "                                                                              "
    brand1 = Left(brand1 & pad, 30)
    quantity1 = Left(quantity.Text & pad, 2)
        For i = 0 To saleslist.ListCount - 1
            If UCase(stockid.Text) = Left(saleslist.List(i), 11) Then
                If Left(Right(saleslist.List(i), 3), 2) = discountrate Then
                    salequantity = Mid(saleslist.List(i), 44, 2)
                    If Right(salequantity, 1) = Chr(9) Then salequantity = Left(salequantity, 1)
                    quantity1 = Val(quantity.Text) + Val(salequantity)
                    saleslist.Selected(i) = True
                    tprice.Caption = "$" & Val(Mid(tprice.Caption, 2, Len(tprice.Caption))) - salep(i)
                    saleslist.removeitem (i)
                    i = saleslist.ListCount - 1
                End If
            End If
        Next i
    saleslist.additem stockid1 & Chr(9) & brand1 & Chr(9) & quantity1 & Chr(9) & "$" & price & Chr(9) & discount.Text
    tprice.Caption = "$" & Val(Mid(tprice.Caption, 2, Len(tprice.Caption))) + (price * quantity1)
    lblafter.Caption = Val(lblafter.Caption) + Int((price * quantity1))
    Call resetquantity
    Call resetdiscount
    stockid.Text = ""
    stockid.SetFocus
    confirm.Enabled = True
End If
End Sub

'Removes item from the list, recalculates the points after the sale
Private Sub removeitem_Click()
If saleslist.ListCount <> 0 Then
    For i = 0 To saleslist.ListCount - 1
        If saleslist.Selected(i) = True Then
            saleprice = salep(i)
            salequantity = Val(Mid(saleslist.List(i), 44, 2))
            If Right(salequantity, 1) = Chr(9) Then salequantity = Left(salequantity, 1)
            m = i
            discountrate = finddiscount(i)
            If Right(saleslist.List(i), 1) = "*" Then lblafter.Caption = Val(lblafter.Caption) + ((100 - discountrate) * 100)
            lblafter.Caption = Val(lblafter.Caption) - Val(saleprice * salequantity)
            tprice.Caption = "$" & (Val(Right(tprice.Caption, Len(tprice.Caption) - 1)) - Val(saleprice * salequantity))
        End If
    Next i
    saleslist.removeitem (m)
End If
If saleslist.ListCount = 0 Then confirm.Enabled = False
End Sub

'Cancels the discount of an item, recalculates the points after the sale
Private Sub candis_Click()
countforper = InStr(saleslist.List(saleslist.ListIndex), "%")
If countforper <> 0 Then
        For i = 0 To saleslist.ListCount - 1
                If saleslist.Selected(i) = True Then
                    saleprice = salep(i)
                    discountrate = finddiscount(i)
                    newprice = Int(pricerevcal(Val(saleprice), Val(discountrate)))
                    If Right(newprice, 1) = 1 Then newprice = newprice - 1
                    If Right(newprice, 1) = 9 Then newprice = newprice + 1
                    tempstockid = Left(saleslist.List(i), 44)
                    editeditem = tempstockid & Chr(9) & "$" & newprice & Chr(9)
                    lblafter.Caption = Val(lblafter) - saleprice + newprice
                    If Right(saleslist.List(i), 1) = "*" Then lblafter.Caption = Val(lblafter) + ((100 - discountrate) * 100)
                    saleslist.removeitem (i)
                End If
        Next i
        saleslist.additem editeditem
        tprice.Caption = "$" & Val(Right(tprice.Caption, Len(tprice.Caption) - 1)) - saleprice + newprice
End If
End Sub

'Discounts a item with member points, recalculates the points after the sale
Private Sub ok_Click()
Dim saleprice As String
Dim pointdiscount As Integer
allowtodiscount = True
If pointsuse.Text <> "" Then
    For i = 0 To saleslist.ListCount - 1
        If saleslist.Selected(i) Then
            If Right(saleslist.List(i), 1) = "%" Or Right(saleslist.List(i), 1) = "*" Then
                allowtodiscount = False
            End If
        End If
        salequantity = Trim(Mid(saleslist.List(i), 44, 2))
    Next i
    If allowtodiscount = True And salequantity = 1 Then
        pointstotakeaway = Left(pointsuse.Text, InStr(pointsuse.Text, "=") - 1)
        pointdiscount = Left(Right(pointsuse.Text, 3), 2)
        For i = 0 To saleslist.ListCount - 1
                If saleslist.Selected(i) = True Then
                    saleprice = salep(i)
                    newprice = pricecalculation(Val(saleprice), pointdiscount)
                    tempstockid = Left(saleslist.List(i), 44)
                    editeditem = tempstockid & Chr(9) & "$" & newprice & Chr(9) & pointdiscount & "%*"
                    saleslist.removeitem (i)
                End If
        Next i
        saleslist.additem editeditem
        lblafter.Caption = Val(lblafter.Caption) - Val(saleprice) + newprice - pointstotakeaway
        tprice.Caption = "$" & Val(Right(tprice.Caption, Len(tprice.Caption) - 1)) - saleprice + newprice
    Else
        If allowtodiscount = False Then MsgBox "No discount is allowed on discounted items"
        If salequantity <> 1 Then MsgBox "Only one item can be discounted"
    End If
End If
Call saleslist_Click
End Sub

'Checks that only enables the combo box pointsuse and command button OK and candis
'when there is only one item on the list is chosen
Private Sub saleslist_Click()
For i = 0 To saleslist.ListCount - 1
    If saleslist.Selected(i) = True Then m = m + 1
Next i
Call calculatepointstouse
If m = 1 Then
    pointsuse.Enabled = True
    ok.Enabled = True
Else
    pointsuse.Enabled = False
    ok.Enabled = False
End If
candis.Enabled = True
End Sub

'Adds all the details of items of this sales into the table Sales
Private Sub confirm_Click()
Dim rssales As New ADODB.Recordset
Dim rsstock As New ADODB.Recordset
Dim rsc As New ADODB.Recordset
DB.Open Path
    rssales.Open "Sales", Path, adOpenKeyset, adLockOptimistic
    With rssales
        If .RecordCount = 0 Then
            saleid = 1
        Else
            .MoveLast
            saleid = rssales![saleid] + 1
        End If
    End With
    rssales.Close
DB.Close
timeofsale = Time
For i = 0 To saleslist.ListCount - 1
salequantity = Mid(saleslist.List(i), 44, 2)
If Right(salequantity, 1) = Chr(9) Or Right(salequantity, 1) = "" Then salequantity = Left(salequantity, 1)
rsstock.Open "select [quantity] from [Stock] where [stockid] = '" & Left(saleslist.List(i), 11) & "'", Path, adOpenKeyset, adLockOptimistic
    If rsstock![quantity] - salequantity < 0 Then
        MsgBox "Error occured, the process will be stopped now"
        Call resetform
        rsstock.Close
        Exit Sub
    End If
rsstock.Close
saleprice = salep(i)
    DB.Open Path
        rssales.Open "Sales", Path, adOpenKeyset, adLockOptimistic
        With rssales
            .AddNew
            ![saleid] = saleid
            tempstockid = Left(saleslist.List(i), 11)
            ![stockid] = tempstockid
            ![quantity] = salequantity
            rsstock.Open "select * from stock where stockid = '" & tempstockid & "'", Path, adOpenKeyset, adLockOptimistic
                ![pprice] = rsstock![pprice]
                ![price] = saleprice
                rsstock![quantity] = rsstock![quantity] - salequantity
                rsstock.Update
            rsstock.Close
            ![Time] = timeofsale
            ![Date] = Date
            If memberid.Text <> "" Then
                ![member] = True
                ![memberid] = Val(Left(memberid.Text, InStr(memberid.Text, ")") + 1))
            Else
                ![memberid] = 0
            End If
            ![fromorders] = False
            ![staff] = uname
        End With
        rssales.Update
        rssales.Close
    DB.Close
Next i
If memberid.Text <> "" Then
    DB.Open Path
        rsc.Open "select * from members where [memberid] = " & Val(Left(memberid.Text, InStr(memberid.Text, ")") + 1)), Path, adOpenKeyset, adLockOptimistic
            rsc![Points] = Int(Val(lblbefore.Caption)) + Int(Val(lblafter.Caption))
            rsc.Update
        rsc.Close
    DB.Close
End If
MsgBox "Sale made!"
stock.refreshofstock.Interval = "10"
stock.refreshofstock.Interval = "0"
Call printoutreceipt(timeofsale, Date, saleid)
Call resetform
End Sub

'Checks is the Stock ID valid and looks for the quantity when stockID is valid
Private Sub stockid_Change()
Dim quantity1 As Integer
If Len(stockid.Text) < 11 Then
    Call resetquantity
ElseIf Len(stockid.Text) = 11 Then
    tempstockid = UCase(stockid.Text)
    Dim rsstock As New ADODB.Recordset
    DB.Open Path
        rsstock.Open "Select * from stock where [stockid] = '" & tempstockid & "'", Path, adOpenKeyset, adLockOptimistic
            If rsstock.RecordCount = 0 Then
                MsgBox "Invalid stock ID"
                rsstock.Close
                DB.Close
                Exit Sub
            Else
                quantity1 = rsstock![quantity]
            End If
        rsstock.Close
    DB.Close

    Dim stocktotakeaway As Integer
    stocktotakeaway = checknewquantity(tempstockid)
    Call addquantity(quantity1, stocktotakeaway)
ElseIf Len(stockid.Text) > 11 Then
    stockid.Text = ""
End If
End Sub

'Enables the combo box discount when combo box quantity is not empty
Private Sub quantity_click()
If quantity.Text <> "" Then discount.Enabled = True
End Sub

'Looks for the member point for the chosen user
Private Sub memberid_Click()
searchcritria = Val(Left(memberid.Text, InStr(memberid.Text, ")") + 1))
Dim rsc As New ADODB.Recordset
    DB.Open Path
        rsc.Open "Select [points] from members where [memberid] = " & searchcritria, Path, adOpenKeyset, adLockOptimistic
            lblbefore.Caption = rsc![Points]
        rsc.Close
    DB.Close
    Call calculatepointstouse
memberid.Enabled = False
End Sub

'Calls the function that fills the discount rate available for using member points
Private Sub pointsuse_Change()
Call calculatepointstouse
End Sub

'Disables the combo boxes quantity, discount
Private Sub quantity_Change()
Call resetquantity
discount.Clear
Call resetdiscount
Call stockid_Change
End Sub

'Calls the function that resets the combo box Member ID
Private Sub discount_Change()
Call resetdiscount
End Sub

'Calls the function that resets the combo box Member ID
Private Sub memberid_Change()
Call resetmemberid
End Sub
'==========================Procedures/Functions===============================
'Adds the discount rate for using member points
Private Function calculatepointstouse()
pointsuse.Clear
If lblafter.Caption = "" Or lblafter.Caption = 0 Then
    Points = Val(lblbefore.Caption)
Else
    If Val(lblafter) > 0 Then Points = Val(lblbefore.Caption) - Val(lblafter.Caption)
    If Val(lblafter) < 0 Then Points = Val(lblbefore.Caption) + Val(lblafter.Caption)
End If
times = Int(Points / 500)
pointstouse = 500
discount1 = 95
For i = 1 To times
    pointsuse.additem pointstouse & "=>" & discount1 & "%"
    pointstouse = pointstouse + 500
    discount1 = discount1 - 5
    If discount1 = 45 Then i = times
Next i
End Function

'Resets the combo box MemberID
Private Function resetmemberid()
memberid.Clear
Dim rsc As New ADODB.Recordset
    DB.Open Path
        rsc.Open "select * from members order by [memberid]", Path, adOpenKeyset, adLockOptimistic
            For i = 1 To rsc.RecordCount
                memberid.additem rsc![memberid] & ")" & rsc![Title] & " " & rsc![surname]
                rsc.MoveNext
            Next i
        rsc.Close
    DB.Close
End Function

'
Private Function checknewquantity(stockid)
checknewquantity = 0
itemfound = False
For i = 0 To saleslist.ListCount
    stockidinlist = Left(saleslist.List(i), 11)
    If stockid = stockidinlist Then
        currentitem = Mid(saleslist.List(i), 44, 2)
        If Right(currentitem, 1) = " " Then currentitem = Left(currentitem, 1)
        itemfound = True
        checknewquantity = checknewquantity + currentitem
    End If
Next i
If itemfound = False Then checknewquantity = 0
End Function

'Recalculates the original price
Private Function pricecalculation(anything As Integer, discountrate As Integer)
    pricecalculation = (anything * Val(Left(discountrate, 2))) / 100
End Function

'Calculates the discount
Private Function pricerevcal(anything As Integer, discountrate As Integer)
    pricerevcal = (anything / Val(Left(discountrate, 2))) * 100
End Function

'Calculates the net quantity left including the quantity in the database as well as
'the quantity already added in this sale
Private Function addquantity(quantity1 As Integer, stocktotakeaway As Integer)
    quantity.Clear
    quantity1 = quantity1 - stocktotakeaway
    If quantity1 > 0 Then
        quantity.Enabled = True
        For i = 1 To quantity1
            quantity.additem i
        Next i
    Else
        MsgBox stockid.Text & " has no stock"
    End If
End Function

'Disables the combo box quantity
Private Function resetquantity()
quantity.Enabled = False
quantity.Clear
discount.Enabled = False
End Function

'Creates a sales receipt, prints out and deletes the file
Private Function printoutreceipt(tos, d, sid)
Dim fso As New FileSystemObject
Dim rssale As New ADODB.Recordset
Dim rsstock As New ADODB.Recordset
pathtotext = App.Path & "\databases\D-Store Receipt.txt"
fso.CreateTextFile (pathtotext)
pad = "                                                           "
pad1 = "================================================================================"
Open (pathtotext) For Output As #1
    printemptyline (2)
    Print #1, pad1
    printemptyline (1)
    Print #1, "         [Date]:" & d & " " & tos & "            [Staff]:" & uname
    printemptyline (1)
    Print #1, "         [Sale ID]:" & sid & "         [Total Price]:" & tprice.Caption
    printemptyline (1)
    If memberid.Text = "" Then
        Print #1, "         [Member ID]:N/A" & "     [Points Gained]:" & Val(lblafter.Caption) & "        [Points]:0"
    Else
        Print #1, "         [Member ID]:" & Val(Left(memberid.Text, InStr(memberid.Text, ")") + 1)) & "     [Points Gained]:" & Val(lblafter.Caption) & "        [Points]:" & (Val(lblbefore.Caption) + Val(lblafter.Caption))
    End If
    
    printemptyline (1)
    Print #1, pad1
    printemptyline (2)
    Print #1, "         [Stock ID]          [Price Each]   [Quantity]    [Discount]"
    printemptyline (1)
DB.Open Path
    rssale.Open "select * from [sales] where [saleid] = " & sid, Path, adOpenKeyset, adLockOptimistic
    With rssale
       For i = 1 To rssale.RecordCount
            rsstock.Open "select [price] from [stock] where [stockid] = '" & ![stockid] & "'", Path, adOpenKeyset, adLockOptimistic
                If Val(![price]) / Val(rsstock![price]) <> 1 Then dic = Val(![price]) / Val(rsstock![price]) * 100 & "%"
                If Val(![price]) / Val(rsstock![price]) = 1 Then dic = ""
            rsstock.Close
            Print #1, "         " & ![stockid] & Chr(9) & Chr(9) & "$" & Left(![price], 6) & Chr(9) & Chr(9) & ![quantity] & Chr(9) & "      " & dic
            printemptyline (1)
            .MoveNext
        Next i
    .Close
    End With
DB.Close
    printemptyline (1)
    Print #1, pad1
Close #1
Shell ("notepad.exe /p " & pathtotext)
fso.DeleteFile (pathtotext)
End Function

'Prints an empty line into a notepad file
Private Function printemptyline(o)
For i = 1 To o
    Print #1, pad
Next i
End Function

'Resets the combo box discount and adds options of discount rates
Private Function resetdiscount()
discount.Clear
discount.Enabled = False
discount.additem "95%"
discount.additem "90%"
discount.additem "85%"
discount.additem "80%"
discount.additem "75%"
discount.additem "70%"
discount.additem "65%"
End Function

'Function that empties the field
Private Function resetform()
Call resetmemberid
Call resetquantity
Call resetdiscount
saleslist.Clear
lblbefore.Caption = ""
lblafter.Caption = ""
stockid.Text = ""
candis.Enabled = False
confirm.Enabled = False
pointsuse.Enabled = False
ok.Enabled = False
End Function

'Gets the price of the chosen item
Private Function salep(i)
salep = Right(Left(saleslist.List(i), InStr(saleslist.List(i), "$") + 5), 5)
If Left(salep, 1) = "$" Then salep = Mid(salep, 2, Len(salep) - 1)
jlength = Len(salep)
For j = 0 To jlength - 1
    Resize = jlength - 1 - j
    If Mid(salep, jlength - j, 1) = Chr(9) Then salep = Left(salep, Resize)
Next j
End Function

'Gets the discount rate of the chosen item
Private Function finddiscount(i)
finddiscount = Trim(Right(saleslist.List(i), 4))
If Left(finddiscount, 1) = Chr(9) Then finddiscount = Right(finddiscount, Len(finddiscount) - 1)
If Right(finddiscount, 1) = "*" Then finddiscount = Left(finddiscount, Len(finddiscount) - 1)
If Right(finddiscount, 1) = "%" Then finddiscount = Left(finddiscount, Len(finddiscount) - 1)
End Function
