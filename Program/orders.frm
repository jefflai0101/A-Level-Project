VERSION 5.00
Begin VB.Form orders 
   BackColor       =   &H00000000&
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "D-Store - Orders"
   ClientHeight    =   6255
   ClientLeft      =   1260
   ClientTop       =   1395
   ClientWidth     =   10125
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox heading 
      Height          =   240
      Left            =   3960
      TabIndex        =   19
      Top             =   1320
      Width           =   5895
   End
   Begin VB.ListBox itemlist 
      Height          =   4260
      Left            =   3960
      Style           =   1  '項目包含核取方塊
      TabIndex        =   13
      Top             =   1560
      Width           =   5895
   End
   Begin VB.Frame morder 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   3495
      Begin VB.CommandButton removeitem 
         Caption         =   "Remove item"
         Height          =   495
         Left            =   1800
         TabIndex        =   18
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton additem 
         Caption         =   "Add item"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox ordern 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   1560
         Width           =   1695
      End
      Begin VB.ComboBox quantity 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   2775
         Width           =   1815
      End
      Begin VB.TextBox stockid 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   2160
         Width           =   1575
      End
      Begin VB.ComboBox memberid 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label tprice 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label lblordern 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Order number"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label tp 
         BackColor       =   &H00000000&
         Caption         =   "Total price"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label orderdetails 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Order details"
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
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblquantity 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Quantity"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2775
         Width           =   855
      End
      Begin VB.Label lblstockid 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Stock ID"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblmid 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Member ID"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.ListBox orderlist 
      Height          =   4200
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Frame menu 
      BackColor       =   &H00000000&
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.Label cancelo 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Cancel Orders"
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
         Left            =   7200
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label confirmo 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Confirm Orders"
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
         Left            =   3840
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label makeo 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Make Orders"
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
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.ListBox heading2 
      Height          =   240
      Left            =   240
      TabIndex        =   20
      Top             =   1320
      Width           =   3495
   End
End
Attribute VB_Name = "orders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As New ADODB.Connection

'Resets the form
Private Sub Form_Load()
Call resetform
heading.additem "    [Stock ID]" & Chr(9) & "[Brand]" & Chr(9) & Chr(9) & "         [Quantity]" & Chr(9) & "[Price]"
heading2.additem "    [OID]" & Chr(9) & "[MID]" & Chr(9) & "[Member Name]" & Chr(9) & "[Price]"
End Sub

'Resets the forecolor of orders button in mainmenu into white
Private Sub Form_Unload(Cancel As Integer)
Mainmenu.fwforders.ForeColor = &HFFFFFF
End Sub

'Runs shows the table that allows user to add orders and prints out receipt
'or adds the order into database
Private Sub makeo_Click()
If makeo.Caption = "Make Orders" Then
    orderlist.Visible = False
    orderlist.Clear
    morder.Visible = True
    itemlist.Clear
    makeo.Caption = "Add Order"
    makeo.ForeColor = &HFF&
    Call initialmember
    Call initialorderid
    Call resetquantity
    stockid.Text = ""
ElseIf makeo.Caption = "Add Order" Then
    If itemlist.ListCount > 0 Then
        Dim Msg, Style, Answer
        Msg = ("Are you sure you want to add this order?")
        Style = vbYesNo + vbInformation
        Answer = MsgBox(Msg, Style, "Confirm")
        If Answer = vbYes Then
            Call addorder
        Else
            MsgBox "Order not completed"
        End If
    End If
    Call resetform
End If
End Sub

'Confirms the order and prints out receipt
Private Sub confirmo_Click()
Dim ordernumber As Integer
ordernumber = getonumber()
Dim Msg, Style, Answer
    Msg = "Are you sure you want to confirm order #" & ordernumber & "?"
    Style = vbYesNo + vbInformation
    Answer = MsgBox(Msg, Style, "Confirm")
If Answer = vbYes Then
    Call transfertosale(ordernumber)
    Call deleteorder(ordernumber)
    MsgBox "Order confirmed"
    Call loadtable
    itemlist.Clear
End If
End Sub

'Cancels the order made
Private Sub cancelo_Click()
Dim ordernumber As Integer
Dim rsorder As New ADODB.Recordset
Dim rsstock As New ADODB.Recordset
ordernumber = getonumber()
If ordernumber > 0 Then
    Dim Msg, Style, Answer
        Msg = "Are you sure you want to cancel order #" & ordernumber & "?"
        Style = vbYesNo + vbInformation
        Answer = MsgBox(Msg, Style, "Confirm")
    If Answer = vbYes Then
        DB.Open Path
            rsorder.Open "select * from [orders] where [ordernumber] = " & ordernumber, Path, adOpenKeyset, adLockOptimistic
                tempq = rsorder![quantity]
                tempsid = rsorder![stockid]
            rsorder.Close
            rsstock.Open "select * from [stock] where [stockid] = '" & tempsid & "'", Path, adOpenKeyset, adLockOptimistic
                rsstock![quantity] = Val(rsstock![quantity]) + Val(tempq)
                rsstock.Update
            rsstock.Close
        DB.Close
        Call deleteorder(ordernumber)
        MsgBox "Order cancelled"
        Call loadtable
        itemlist.Clear
    End If
End If
End Sub

'Adds item onto the list of a single order
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
    pad = "                                                                              "
    brand1 = Left(brand1 & pad, 30)
    quantity1 = Left(quantity.Text & pad, 2)
        For i = 0 To itemlist.ListCount - 1
            If stockid.Text = Left(itemlist.List(i), 11) Then
                orderquantity = Mid(itemlist.List(i), 44, 2)
                If Right(orderquantity, 1) = Chr(9) Then orderquantity = Left(orderquantity, 1)
                quantity1 = Val(quantity.Text) + Val(orderquantity)
                itemlist.Selected(i) = True
                Call removeitem_Click
            End If
        Next i
    itemlist.additem stockid1 & Chr(9) & brand1 & Chr(9) & quantity1 & Chr(9) & "$" & price
    tprice.Caption = "$" & (Val(Right(tprice.Caption, Len(tprice.Caption) - 1)) + price * quantity.Text)
    Call resetquantity
    stockid.Text = ""
    stockid.SetFocus
End If
End Sub

'Shows all items of that chosen order
Private Sub orderlist_Click()
orderid = Left(orderlist.List(orderlist.ListIndex), 5)
chr9empty = False
For i = 1 To 5
    If chr9empty = False Then
        If Mid(orderid, i, 1) = Chr(9) Or Mid(orderid, i, 1) = " " Then
            chr9empty = True
        Else
            tempoid = tempoid & Mid(orderid, i, 1)
        End If
    End If
Next i
itemlist.Clear
Dim rsorder As New ADODB.Recordset
Dim rsstock As New ADODB.Recordset
    DB.Open Path
    rsorder.Open "select * from orders where [ordernumber] = " & tempoid, Path, adOpenKeyset, adLockOptimistic
        With rsorder
        For i = 1 To .RecordCount
            rsstock.Open "select [brand] from stock where [stockid] = '" & ![stockid] & "'", Path, adOpenKeyset, adLockOptimistic
            pad = "                                                                              "
            tempstockid = Left(rsstock![brand] & pad, 30)
            itemlist.additem ![stockid] & Chr(9) & tempstockid & Chr(9) & ![quantity] & Chr(9) & "$" & ![price]
            .MoveNext
            rsstock.Close
        Next i
        End With
    rsorder.Close
    DB.Close
End Sub

'Removes selected item from the list of order items
Private Sub removeitem_Click()
If itemlist.ListCount <> 0 Then
    For i = 0 To itemlist.ListCount - 1
        If itemlist.Selected(i) = True Then
            saleprice = Right(Left(itemlist.List(i), InStr(itemlist.List(i), "$") + 5), 5)
                If Left(saleprice, 1) = "$" Then saleprice = Mid(saleprice, 2, Len(saleprice) - 1)
                    salequantity = Mid(itemlist.List(i), 44, 2)
                    If Right(salequantity, 1) = Chr(9) Then salequantity = Left(salequantity, 1)
                    chr9found = False
                    For j = 1 To Len(saleprice)
                        If Mid(saleprice, j, 1) <> Chr(9) And chr9found = False Then
                            tempprice = tempprice & Mid(saleprice, j, 1)
                        Else
                            chr9found = True
                        End If
                    Next j
                        tprice.Caption = "$" & (Val(Right(tprice.Caption, Len(tprice.Caption) - 1)) - tempprice * salequantity)
        End If
        If itemlist.Selected(i) = True Then itemlist.removeitem (i)
    Next i
End If
End Sub

'Resets the member combo box when attempt to change
Private Sub memberid_Change()
initialmember
End Sub

'Shows the form members if create member is chosen
Private Sub memberid_Click()
If MemberID.Text = "Create member" Then
    members.Show
    Call Mainmenu.FormWithinForm(Mainmenu.Frame, members)
    Call initialmember
End If
End Sub

'Resets the combo box quantity and calls the change of button
Private Sub quantity_Change()
Call resetquantity
Call stockid_Change
End Sub

'Checks the quantity of that stock when the text box is changed
Private Sub stockid_Change()
Dim quantity1 As Integer
If Len(stockid.Text) < 11 Then
    Call resetquantity
ElseIf Len(stockid.Text) = 11 Then
    Dim rsstock As New ADODB.Recordset
    DB.Open Path
        rsstock.Open "Select * from stock where [stockid] = '" & UCase(stockid.Text) & "'", Path, adOpenKeyset, adLockOptimistic
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
    stocktotakeaway = checknewquantity(stockid)
    Call addquantity(quantity1, stocktotakeaway)
ElseIf Len(stockid.Text) > 11 Then
    stockid.Text = ""
End If
End Sub

'===============================Procedures/Functions=================================
'Checks the quantity of a chosen stock against the database
Private Function checknewquantity(stockid As Object)
checknewquantity = 0
itemfound = False
For i = 0 To itemlist.ListCount
    stockidinlist = Left(itemlist.List(i), 11)
    If stockid = stockidinlist Then
        checknewquantity = Mid(itemlist.List(i), 44, 2)
        If Right(checknewquantity, 1) = Chr(9) Then checknewquantity = Left(checknewquantity, 1)
        itemfound = True
    End If
Next i
If itemfound = False Then checknewquantity = 0
End Function

'Disables and clears the quantity combo box
Private Function resetquantity()
quantity.Enabled = False
quantity.Clear
End Function

'Calculates the net quantity left including the quantity in the database as well as
'the quantity already added in this order
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

'Looks up the value of the recent order id
Private Function initialorderid()
Dim rsorder As New ADODB.Recordset
    DB.Open Path
        rsorder.Open "Select * from orders", Path, adOpenKeyset, adLockOptimistic
            If rsorder.RecordCount <> 0 Then
                rsorder.MoveLast
                ordern.Text = rsorder![ordernumber] + 1
            Else
                ordern.Text = "1"
            End If
        rsorder.Close
    DB.Close
End Function

'Adds members into the combo box MemberID and the option "Create member"
Private Function initialmember()
MemberID.Clear
Dim rsc As New ADODB.Recordset
    DB.Open Path
        rsc.Open "select * from members order by [memberid]", Path, adOpenKeyset, adLockOptimistic
            For i = 1 To rsc.RecordCount
                MemberID.additem rsc![MemberID] & ")" & rsc![Title] & " " & rsc![surname]
                rsc.MoveNext
            Next i
        rsc.Close
    DB.Close
MemberID.additem "Create member"
End Function

'Adds order details of items into the database
Private Function addorder()
Dim rsstock As New ADODB.Recordset
If MemberID.Text <> "" Then
    For i = 0 To itemlist.ListCount - 1
        If Right(Mid(itemlist.List(i), 44, 2), 1) = " " Or Right(Mid(itemlist.List(i), 44, 2), 1) = Chr(9) Then salequantity = Left(salequantity, 1)
        idofitem = Left(itemlist.List(i), 11)
        saleprice = Right(Left(itemlist.List(i), InStr(itemlist.List(i), "$") + 5), 5)
            If Left(saleprice, 1) = Chr(9) Then saleprice = Mid(saleprice, 2, Len(saleprice) - 1)
            If Left(saleprice, 1) = "$" Then saleprice = Mid(saleprice, 2, Len(saleprice) - 1)
                salequantity = Mid(itemlist.List(i), 44, 2)
                If Right(salequantity, 1) = Chr(9) Then salequantity = Left(salequantity, 1)
                    chr9found = False
                    For j = 1 To Len(saleprice)
                        If Mid(saleprice, j, 1) <> Chr(9) And chr9found = False Then
                            tempprice = tempprice & Mid(saleprice, j, 1)
                        Else
                            chr9found = True
                        End If
                    Next j
        rsstock.Open "select [quantity] from [Stock] where [stockid] = '" & idofitem & "'", Path, adOpenKeyset, adLockOptimistic
            If rsstock![quantity] - salequantity < 0 Then
                MsgBox "Error occured, the process will be stopped now"
                Call resetform
                rsstock.Close
                Exit Function
            End If
        rsstock.Close
        Dim rsorder As New ADODB.Recordset
        DB.Open Path
            rsstock.Open "select * from [Stock] where [stockid] = '" & idofitem & "'", Path, adOpenKeyset, adLockOptimistic
                rsstock![quantity] = rsstock![quantity] - salequantity
                rsstock.Update
            rsorder.Open "[Orders]", Path, adOpenKeyset, adLockOptimistic
                With rsorder
                .AddNew
                ![ordernumber] = ordern.Text
                ![MemberID] = Val(Left(MemberID.Text, InStr(MemberID.Text, ")") + 1))
                ![stockid] = idofitem
                ![pprice] = rsstock![pprice]
                ![price] = saleprice
                ![quantity] = salequantity
                ![Time] = Time()
                ![Date] = Date
                ![staff] = uname
                .Update
                End With
            rsorder.Close
            rsstock.Close
        DB.Close
    Next i
    stock.refreshofstock.Interval = "10"
    stock.refreshofstock.Interval = "0"
    MsgBox "Order added"
    Call printorderreceipt(Time, Date, Val(ordern.Text))
Else
    MsgBox "Order not completed because MemebrID is missing"
    Exit Function
End If
End Function

'Loads all the orders into the list
Private Function loadtable()
Dim rsorder As New ADODB.Recordset
Dim rsc As New ADODB.Recordset
orderlist.Clear
    rsorder.Open "[Orders]", Path, adOpenKeyset, adLockOptimistic
        If rsorder.RecordCount <> 0 Then
            getid = rsorder.RecordCount
        Else
            getid = 1
        End If
    rsorder.Close
For i = 1 To getid
    totalprice = 0
    rsorder.Open "Select * from [Orders] where [ordernumber] = " & i, Path, adOpenKeyset, adLockOptimistic
        If rsorder.RecordCount <> 0 Then
                rsorder.MoveFirst
                For j = 1 To rsorder.RecordCount
                    totalprice = Val(totalprice) + Val(rsorder![price])
                    rsorder.MoveNext
                Next j
            rsorder.MoveFirst
        End If
    If rsorder.RecordCount > 0 Then
        rsc.Open "select * from [Members] where [memberid] = " & rsorder![MemberID], Path, adOpenKeyset, adLockOptimistic
            With rsc
                pad = "                                                                                          "
                mname = Left(![Title] & " " & ![surname] & pad, 19)
                orderlist.additem i & Chr(9) & rsc![MemberID] & Chr(9) & mname & Chr(9) & "$" & totalprice
            End With
        rsc.Close
    End If
    rsorder.Close
Next i
End Function

'Transfers all order details of a chosen, confirmed order from the table "Orders"
'to table "Sales" in the database
Private Function transfertosale(ordernumber As Integer)
Dim rsorder As New ADODB.Recordset
Dim rssale As New ADODB.Recordset
confirmtime = Time
DB.Open Path
    rssale.Open "[sales]", Path, adOpenKeyset, adLockOptimistic
        rssale.MoveLast
        tempsid = rssale![saleid] + 1
    rssale.Close
    rsorder.Open "select * from [orders] where [ordernumber] = " & ordernumber, Path, adOpenKeyset, adLockOptimistic
        For i = 1 To rsorder.RecordCount
            rssale.Open "[sales]", Path, adOpenKeyset, adLockOptimistic
            With rssale
                .AddNew
                ![saleid] = tempsid
                ![stockid] = rsorder![stockid]
                ![pprice] = rsorder![pprice]
                ![price] = rsorder![price]
                ![quantity] = rsorder![quantity]
                ![Time] = confirmtime
                ![Date] = Date
                ![member] = True
                ![MemberID] = rsorder![MemberID]
                ![fromorders] = True
                ![staff] = uname
                .Update
            .Close
            End With
            rsorder.MoveNext
        Next i
    rsorder.Close
DB.Close
Call printotsreceipt(confirmtime, Date, tempsid)
End Function

'Deletes all details of the chosen cancelled order
Private Function deleteorder(ordernumber As Integer)
Dim rsorder As New ADODB.Recordset
Dim rsstock As New ADODB.Recordset
    DB.Open Path
        rsorder.Open "delete * from [orders] where [ordernumber] = " & ordernumber, Path, adOpenKeyset, adLockOptimistic
    DB.Close
Call loadtable
End Function

'Looks the order ID of a chosen order
Private Function getonumber()
ordernumber = Left(orderlist.List(orderlist.ListIndex), 4)
    invalidchrfound = False
    endswhen = Len(ordernumber)
    For i = 1 To endswhen
        If Mid(ordernumber, i, 1) = Chr(9) Then invalidchrfound = True
        If Mid(ordernumber, i, 1) = " " Then invalidchrfound = True
        If invalidchrfound = False Then temp = temp & Mid(ordernumber, i, 1)
    Next i
    getonumber = Val(temp)
End Function

'Print empty line into a notepad
Private Function printemptyline(o)
For i = 1 To o
    Print #1, pad
Next i
End Function

'Creates a receipt of confirmed order using notepad, prints out and deletes the file
Private Function printotsreceipt(tos, d, sid)
Dim fso As New FileSystemObject
Dim rssale As New ADODB.Recordset
pathtotext = App.Path & "\databases\Confirmed Order.txt"
fso.CreateTextFile (pathtotext)
pad = "                                                           "
pad1 = "================================================================================"
Open (pathtotext) For Output As #1
    printemptyline (2)
    Print #1, pad1
    printemptyline (1)
    Print #1, "         [Date]:" & d & " " & tos & "            [Staff]:" & uname
    printemptyline (1)
While meid = ""
DB.Open Path
    rssale.Open "select * from [sales] where [SaleID] = " & sid, Path, adOpenKeyset, adLockOptimistic
    With rssale
       For i = 1 To rssale.RecordCount
            meid = rssale![MemberID]
            tempprice = tempprice + Val(rssale![price])
            rssale.MoveNext
       Next i
    .Close
    End With
DB.Close
Wend
    Print #1, "     [Sale ID]:" & sid & "         [Total Price]: $" & tempprice & "     [Member ID]:" & meid
    printemptyline (1)
    Print #1, pad1
    printemptyline (2)
    Print #1, "         [Stock ID]          [Price Each]   [Quantity]"
    printemptyline (1)
DB.Open Path
    rssale.Open "select * from [sales] where [saleid] = " & sid, Path, adOpenKeyset, adLockOptimistic
    With rssale
       For i = 1 To rssale.RecordCount
            Print #1, "         " & ![stockid] & Chr(9) & Chr(9) & "$" & Left(![price], 6) & Chr(9) & Chr(9) & ![quantity]
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
Dim rsc As New ADODB.Recordset
DB.Open Path
rsc.Open "Select * from [members] where [memberid] = " & meid, Path, adOpenKeyset, adLockOptimistic
    rsc![Points] = Val(rsc![Points]) + (tempprice / 2)
    rsc.Update
rsc.Close
DB.Close
End Function

'Creates a report of order made using notepad, print out and deletes the file
Private Function printorderreceipt(tos, d, oid)
Dim fso As New FileSystemObject
Dim rsorder As New ADODB.Recordset
pathtotext = App.Path & "\databases\Order Receipets.txt"
fso.CreateTextFile (pathtotext)
pad = "                                                           "
pad1 = "================================================================================"
Open (pathtotext) For Output As #1
    printemptyline (2)
    Print #1, pad1
    printemptyline (1)
    Print #1, "         [Date]:" & d & " " & tos & "            [Staff]:" & uname
    printemptyline (1)
While meid = ""
DB.Open Path
    rsorder.Open "select * from [orders] where [ordernumber] = " & oid, Path, adOpenKeyset, adLockOptimistic
    With rsorder
       For i = 1 To rsorder.RecordCount
             meid = rsorder![MemberID]
            tempprice = tempprice + Val(rsorder![price])
            rsorder.MoveNext
       Next i
    .Close
    End With
DB.Close
Wend
    Print #1, "     [Order ID]:" & oid & "         [Total Price]: $" & tempprice & "     [Member ID]:" & meid
    printemptyline (1)
    Print #1, pad1
    printemptyline (2)
    Print #1, "         [Stock ID]          [Price Each]   [Quantity]"
    printemptyline (1)
DB.Open Path
    rsorder.Open "select * from [orders] where [ordernumber] = " & oid, Path, adOpenKeyset, adLockOptimistic
    With rsorder
       For i = 1 To rsorder.RecordCount
            Print #1, "         " & ![stockid] & Chr(9) & Chr(9) & "$" & Left(![price], 6) & Chr(9) & Chr(9) & ![quantity]
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

'Empties the fields and hides the order table
Private Function resetform()
orderlist.Visible = True
morder.Visible = False
makeo.Caption = "Make Orders"
Call loadtable
makeo.ForeColor = &HFFFFFF
tprice.Caption = "$0"
itemlist.Clear
stockid.Text = ""
End Function
