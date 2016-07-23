VERSION 5.00
Begin VB.Form stock 
   BackColor       =   &H00000000&
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "D-Store - Stock Manage"
   ClientHeight    =   7200
   ClientLeft      =   240
   ClientTop       =   1275
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   13095
   ShowInTaskbar   =   0   'False
   Begin VB.Timer refreshofstock 
      Left            =   240
      Top             =   240
   End
   Begin VB.ListBox itemlist 
      Height          =   4740
      Left            =   240
      TabIndex        =   22
      Top             =   1680
      Width           =   8775
   End
   Begin VB.ListBox Titles 
      Height          =   240
      Left            =   240
      TabIndex        =   21
      Top             =   1440
      Width           =   8775
   End
   Begin VB.Frame itemdetails 
      BackColor       =   &H00000000&
      Caption         =   "Item details"
      ForeColor       =   &H00FFFFFF&
      Height          =   5415
      Left            =   9240
      TabIndex        =   4
      Top             =   1320
      Width           =   3615
      Begin VB.ComboBox brand 
         Height          =   315
         Left            =   1440
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton confirm 
         Caption         =   "Confirm"
         Height          =   495
         Left            =   1080
         TabIndex        =   20
         Top             =   4440
         Width           =   1695
      End
      Begin VB.TextBox quantity 
         Height          =   375
         Left            =   1440
         TabIndex        =   19
         Top             =   3600
         Width           =   1935
      End
      Begin VB.TextBox pos 
         Height          =   375
         Left            =   1440
         TabIndex        =   18
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox pp 
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox descrip 
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox combosize 
         Height          =   315
         Left            =   1440
         TabIndex        =   14
         Text            =   "Choose a size"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox stockid 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label quan1 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Quantity"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label pos1 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Price of Sale"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label pprice1 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Purchase Price"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label descrip1 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Description"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label size1 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Size"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label brand1 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Brand"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label stockid1 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Stock ID"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame menu 
      BackColor       =   &H00000000&
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   11415
      Begin VB.Label sallitem 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Show All Items"
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
         Left            =   9120
         TabIndex        =   23
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label sitem 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Search Items"
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
         Left            =   6960
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label ditem 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Delete Items"
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
         Left            =   4680
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label eitem 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Edit Items"
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
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label aitem 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Add Items"
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
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim useroption As Integer
Dim selectedanitem As Boolean
Dim DB As New ADODB.Connection
Dim rsstock As New ADODB.Recordset

'Resets the form, and only disables the aitem, eitem and ditem button when the
'user's access right is "User"
Private Sub Form_Load()
pad = "                                                                                          "
Titles.additem "[Stock ID]" & Chr(9) & Left("[Brand]" & pad, 30) & Chr(9) & "[Size]   " & Left("[Descriptions]" & pad, 35) & Chr(9) & "[PPrice]" & Chr(9) & "[Price]" & Chr(9) & "[Quantity]"
Call aitem_Click
If status = "U" Then
    aitem.Enabled = False
    eitem.Enabled = False
    ditem.Enabled = False
    Call sitem_Click
End If
End Sub

'Resets the forecolor of changepassword button in mainmenu into white
Private Sub Form_Unload(Cancel As Integer)
Mainmenu.fwfstock.ForeColor = &HFFFFFF
End Sub

'Shows all items in the table stock
Private Sub refreshofstock_Timer()
Call sallitem_Click
End Sub

'Changes to add item mode and resets the form
Private Sub aitem_Click()
useroption = 1
aitem.ForeColor = &HFF&
eitem.ForeColor = &HFFFFFF
sitem.ForeColor = &HFFFFFF
Call resetform
End Sub

'Changes to edit item mode and resets the form
Private Sub eitem_Click()
useroption = 2
selectedanitem = False
eitem.ForeColor = &HFF&
aitem.ForeColor = &HFFFFFF
sitem.ForeColor = &HFFFFFF
End Sub

'Changes to search item mode and resets the form
Private Sub sitem_Click()
    aitem.ForeColor = &HFFFFFF
    eitem.ForeColor = &HFFFFFF
    sitem.ForeColor = &HFF&
    useroption = 3
End Sub

'Deletes a item from the table Stock and updates the list
Private Sub ditem_Click()
If selectedanitem = True And useroption = 2 Then
    Dim Msg, Style, Answer
    Msg = ("Are you sure you want to delete this item?")
    Style = vbYesNo + vbInformation
    Answer = MsgBox(Msg, Style, "Confirm")
    If Answer = vbYes Then
        numberofitem = itemlist.ListCount - 1
        itemid = getitemid
        DB.Open Path
            rsstock.Open "delete * from stock where [stockid] = '" & itemid & "'", Path, adOpenKeyset, adLockOptimistic
        DB.Close
        MsgBox "Item deleted!"
    End If
    Call resetform
End If
End Sub

'Shows all items in the table Stock
Private Sub sallitem_Click()
selectedanitem = False
itemlist.Clear
DB.Open Path
    rsstock.Open "Select * from Stock", Path, adOpenKeyset, adLockOptimistic
        For i = 1 To rsstock.RecordCount
            pad = "                                                                                          "
            brand2 = Left(rsstock![brand] & pad, 30)
            size2 = Left(rsstock![Size] & pad, 2)
            desc1 = Left(rsstock![descriptions] & pad, 35)
            pp1 = "$" & Left(rsstock![pprice] & pad, 5)
            p1 = "$" & Left(rsstock![price] & pad, 5)
            qu1 = Left(rsstock![quantity] & pad, 2)
            itemlist.additem rsstock![stockid] & Chr(9) & brand2 & Chr(9) & size2 & Chr(9) & desc1 & Chr(9) & pp1 & Chr(9) & p1 & Chr(9) & qu1
            rsstock.MoveNext
        Next i
    rsstock.Close
DB.Close
End Sub

'Fills all details of the chosen item into the field when in edit mode
Private Sub itemlist_Click()
If useroption = 2 Then
    numberofitem = itemlist.ListCount - 1
    For recentitem = 0 To numberofitem
        If itemlist.Selected(recentitem) = True Then
            itemid = Left(itemlist.List(recentitem), 11)
        End If
    Next recentitem
selectedanitem = True
    DB.Open Path
        rsstock.Open "select * from Stock where [stockid] = '" & itemid & "'", Path, adOpenKeyset, adLockOptimistic
            stockid.Text = rsstock![stockid]
            brand.Text = rsstock![brand]
            combosize.Text = rsstock![Size]
            descrip.Text = rsstock![descriptions]
            pp.Text = rsstock![pprice]
            pos.Text = rsstock![price]
            quantity.Text = rsstock![quantity]
        rsstock.Close
    DB.Close
End If
End Sub

'Adds new item, saves the changes or search for item with criteria
Private Sub confirm_Click()
Select Case useroption
Case 1
    idok = checkid
    If idok = True Then
        If checkvalid = True Then
            Call fillin
            MsgBox "Item added!"
            Call aitem_Click
            Call resetform
        End If
    End If
Case 2
    If selectedanitem = True Then
        If checkvalid = True Then
            Call applychange
            Call eitem_Click
            Call resetform
            MsgBox "Change applied!"
        End If
    End If
Case 3
    Dim checkcrit As Integer
    checkcrit = 0
    If stockid.Text <> "" Then checkcrit = checkcrit + 1
    If brand.Text <> "" Then checkcrit = checkcrit + 1
    If combosize.Text <> "" Then checkcrit = checkcrit + 1
    If descrip.Text <> "" Then checkcrit = checkcrit + 1
    If pp.Text <> "" Then checkcrit = checkcrit + 1
    If pos.Text <> "" Then checkcrit = checkcrit + 1
    If quantity.Text <> "" Then checkcrit = checkcrit + 1
    If checkcrit > 0 Then
        DB.Open Path
            rsstock.Open "Select * from stock", Path, adOpenKeyset, adLockOptimistic
                itemlist.Clear
                With rsstock
                For i = 1 To rsstock.RecordCount
                    match = 0
                    pad = "                                                                                          "
                    If stockid.Text <> "" And ![stockid] = UCase(stockid.Text) Then match = match + 1
                        stid = ![stockid]
                    If brand.Text <> "" And UCase(![brand]) = UCase(brand.Text) Then match = match + 1
                        brand2 = Left(![brand] & pad, 30)
                    If combosize.Text <> "" And UCase(![Size]) = UCase(combosize.Text) Then match = match + 1
                        size2 = Left(![Size] & pad, 2)
                    If descrip.Text <> "" And UCase(![descriptions]) = UCase(descrip.Text) Then match = match + 1
                        desc1 = Left(![descriptions] & pad, 35)
                    If pp.Text <> "" And ![pprice] = pp.Text Then match = match + 1
                        pp1 = "$" & Left(![pprice] & pad, 5)
                    If pos.Text <> "" And ![price] = pos.Text Then match = match + 1
                        p1 = "$" & Left(![price] & pad, 5)
                    If quantity.Text <> "" And ![quantity] = quantity.Text Then match = match + 1
                        qu1 = Left(![quantity] & pad, 2)
                    If match = checkcrit Then itemlist.additem stid & Chr(9) & brand2 & Chr(9) & size2 & "         " & desc1 & Chr(9) & pp1 & Chr(9) & p1 & Chr(9) & qu1
                    rsstock.MoveNext
                Next i
                End With
            rsstock.Close
        DB.Close
    Else
        itemlist.Clear
        Call resetform
    End If
End Select
End Sub

'============================Procedures/Functions==============================
'Adds all details of an item into the table Stock
Private Function fillin()
    DB.Open Path
        rsstock.Open "Select * from stock", Path, adOpenKeyset, adLockOptimistic
        rsstock.AddNew
            rsstock![stockid] = UCase(stockid.Text)
            rsstock![brand] = brand.Text
            rsstock![Size] = combosize.Text
            rsstock![descriptions] = descrip.Text
            rsstock![pprice] = pp.Text
            rsstock![price] = pos.Text
            rsstock![quantity] = quantity.Text
        rsstock.Update
        rsstock.Close
    DB.Close
End Function

'Saves all changes of an item's details into the database
Private Function applychange()
    itemid = getitemid
    DB.Open Path
        rsstock.Open "Select * from stock where [stockid] = '" & itemid & "'", Path, adOpenKeyset, adLockOptimistic
            rsstock![stockid] = UCase(stockid.Text)
            rsstock![brand] = brand.Text
            rsstock![Size] = combosize.Text
            rsstock![descriptions] = descrip.Text
            rsstock![pprice] = pp.Text
            rsstock![price] = pos.Text
            rsstock![quantity] = quantity.Text
        rsstock.Update
        rsstock.Close
    DB.Close
End Function

'Checks the existence of an stock ID against the database
Private Function checkid()
DB.Open Path
    rsstock.Open "select [stockid] from stock where [stockid] = '" & stockid.Text & "'", Path, adOpenKeyset, adLockOptimistic
        If rsstock.RecordCount = 0 Then
            checkid = True
        Else
            MsgBox "Stock ID already used"
            checkid = False
        End If
    rsstock.Close
DB.Close
End Function

'Validation rules for adding or changing items
Private Function checkvalid()
checkvalid = False
If Len(stockid.Text) = 11 Then
    stockid1.ForeColor = &HFFFFFF
    If brand.Text <> "" And Len(brand.Text) <= 30 Then
        brand1.ForeColor = &HFFFFFF
        If combosize.Text = "S" Or combosize.Text = "M" Or combosize.Text = "L" Or combosize.Text = "XL" Then
            size1.ForeColor = &HFFFFFF
            If descrip.Text <> "" And Len(descrip.Text) <= 35 Then
                descrip1.ForeColor = &HFFFFFF
                If pp.Text <> "" Then
                    For checkno = 1 To Len(pp.Text)
                        If Asc(Mid(pp.Text, checkno, 1)) < 48 Or Asc(Mid(pp.Text, checkno, 1)) > 57 Then
                            ppok = False
                            checkno = Len(pp.Text)
                        Else
                            ppok = True
                        End If
                    Next checkno
                    If ppok = True Then
                        If Val(pp.Text) > 0 And Val(pp.Text) <= 50000 Then
                            pprice1.ForeColor = &HFFFFFF
                            If pos.Text <> "" Then
                                For checkno = 1 To Len(pos.Text)
                                    If Asc(Mid(pos.Text, checkno, 1)) < 48 Or Asc(Mid(pos.Text, checkno, 1)) > 57 Then
                                        posok = False
                                        checkno = Len(pos.Text)
                                    Else
                                        posok = True
                                    End If
                                Next checkno
                                If posok = True Then
                                    If Val(pos.Text) > 0 And Val(pos.Text) <= 50000 Then
                                        pos1.ForeColor = &HFFFFFF
                                            If quantity.Text <> "" Then
                                                For checkno = 1 To Len(quantity.Text)
                                                    If Asc(Mid(quantity.Text, checkno, 1)) < 48 Or Asc(Mid(quantity.Text, checkno, 1)) > 57 Then
                                                        quok = False
                                                        checkno = Len(pos.Text)
                                                    Else
                                                        quok = True
                                                    End If
                                                Next checkno
                                                If quok = True Then
                                                    If quantity.Text <> "" And Val(quantity.Text) >= 0 And Val(quantity.Text) <= 1000 Then
                                                        quan1.ForeColor = &HFFFFFF
                                                        checkvalid = True
                                                    Else
                                                        MsgBox "Invalid stock amount!"
                                                        quan1.ForeColor = &HFF&
                                                    End If
                                                Else
                                                    quan1.ForeColor = &HFF&
                                                    MsgBox "Characters are not allowed in this field!"
                                                End If
                                            Else
                                                MsgBox "Field quantity can't be empty"
                                            End If
                                    Else
                                        pos1.ForeColor = &HFF&
                                        MsgBox "Invalid price"
                                    End If
                                Else
                                    pos1.ForeColor = &HFF&
                                    MsgBox "Characters are not allowed in this field!"
                                End If
                            Else
                                pos1.ForeColor = &HFF&
                                MsgBox "Field Price of Sale can't be empty"
                            End If
                        Else
                            pprice1.ForeColor = &HFF&
                            MsgBox "Invalid price"
                        End If
                    Else
                        pprice1.ForeColor = &HFF&
                        MsgBox "Characters are not allowed in this field!"
                    End If
                Else
                    pprice1.ForeColor = &HFF&
                    MsgBox "Field Price of Sale can't be empty"
                End If
            Else
                descrip1.ForeColor = &HFF&
                If descrip.Text = "" Then MsgBox "Field Descriptions can't be empty"
                If Len(descrip.Text) > 35 Then MsgBox "Field descriptions is too long"
            End If
        Else
            size1.ForeColor = &HFF&
            If combosize.Text <> "" Then MsgBox "Invalid size!"
            If combosize.Text = "" Then MsgBox "Field Size can't be empty"
        End If
    Else
        brand1.ForeColor = &HFF&
        If brand.Text = "" Then MsgBox "Field brand can't be empty"
        If Len(brand.Text) > 30 Then MsgBox "Field brand is too long"
    End If
Else
    stockid1.ForeColor = &HFF&
    If stockid.Text = "" Then MsgBox "Field StockID can't be empty"
    If Len(stockid.Text) <> 11 And stockid.Text <> "" Then MsgBox "Invalid StockID length"
End If
End Function

'Gets the chosne item's stock ID
Private Function getitemid()
numberofitem = itemlist.ListCount - 1
For recentitem = 0 To numberofitem
    If itemlist.Selected(recentitem) = True Then
        getitemid = Left(itemlist.List(recentitem), 11)
    End If
Next recentitem
End Function

'Calls functions that resets the form
Private Function resetform()
Call initialise
Call initialsize
Call initialbrand
Call sallitem_Click
End Function

'Empties the fields
Private Function initialise()
stockid.Text = ""
descrip.Text = ""
pp.Text = ""
pos.Text = ""
quantity.Text = ""
End Function

'Resets the combo box combosize and add options of sizes
Private Function initialsize()
combosize.Clear
combosize.additem "S"
combosize.additem "M"
combosize.additem "L"
combosize.additem "XL"
End Function

'Reset the combo box brand and add options of brands
Private Function initialbrand()
brand.Clear
    DB.Open Path
        rsstock.Open "Select [brand] from Stock order by [brand]", Path, adOpenKeyset, adLockOptimistic
            If rsstock.RecordCount > 0 Then
                brand.additem rsstock![brand]
                temp = rsstock![brand]
                For i = 1 To rsstock.RecordCount
                    If rsstock![brand] <> temp Then
                        brand.additem rsstock![brand]
                        temp = rsstock![brand]
                    End If
                    rsstock.MoveNext
                Next i
            End If
        rsstock.Close
    DB.Close
End Function
