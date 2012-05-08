VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form state 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "图书查询"
   ClientHeight    =   9600
   ClientLeft      =   2490
   ClientTop       =   1380
   ClientWidth     =   14070
   Icon            =   "state.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   14070
   Begin VB.CommandButton Command5 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   11
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "查询"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   1680
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "state.frx":000C
      Left            =   2400
      List            =   "state.frx":0022
      TabIndex        =   6
      Text            =   "书名"
      Top             =   720
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      Caption         =   "检索"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   6015
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   5535
      End
      Begin VB.Label Label1 
         Caption         =   "检索类型"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "借出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5415
      Left            =   600
      TabIndex        =   0
      Top             =   3120
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   9551
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "检索内容"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   360
      TabIndex        =   2
      Top             =   2640
      Width           =   13215
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   6120
         Width           =   10335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "操作"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   6960
      TabIndex        =   3
      Top             =   240
      Width           =   6615
      Begin VB.CommandButton Command4 
         Caption         =   "增加图书"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   10
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "归还"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   2775
      End
   End
End
Attribute VB_Name = "state"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Public xlApp As Excel.Application
Public rowc As Integer '得到列数
Public XlBook As Excel.Workbook
Public xlSheet As Excel.Worksheet
Public key As String

Private Sub Command1_Click()
If Me.ListView1.SelectedItem Is Nothing Then
MsgBox "未能选择对象", vbExclamation, "未能选择对象"
Else
If xlSheet.Cells(state.ListView1.SelectedItem, 4) <> 0 Then
MsgBox "该书已被借阅,不能借出", vbExclamation, "已被借阅"
Else
borrow.Show
Me.Visible = False
End If
End If
End Sub

Private Sub Command2_Click()
If Me.ListView1.SelectedItem Is Nothing Then
MsgBox "未能选择对象", vbExclamation, "未能选择对象"
Else
If xlSheet.Cells(state.ListView1.SelectedItem, 4) = 0 Then
MsgBox "该书未被借阅,不能归还", vbExclamation, "未被借阅"
Else
re.Show
Me.Visible = False
End If
End If
End Sub

Private Sub Command3_Click()
rowc = xlSheet.[A65536].End(xlUp).Row

i = 0
Me.AutoRedraw = False
  ListView1.Visible = False
  ListView1.ListItems.Clear
  ListView1.Visible = True
  Me.AutoRedraw = True

key = Me.Text1
Select Case Combo1.Text
Case "编号":
For k = 2 To rowc
'xlSheet.Cells(k, 8) = InStr(xlSheet.Cells(k, 2), Me.Frame1.Caption)
If InStr(UCase(xlSheet.Cells(k, 1)), UCase(key)) > 0 Then
add xlSheet.Cells(k, 1), xlSheet.Cells(k, 2), xlSheet.Cells(k, 3), xlSheet.Cells(k, 4), xlSheet.Cells(k, 5), xlSheet.Cells(k, 6), xlSheet.Cells(k, 7), Val(k)
End If
Next k
Case "书名":
For k = 2 To rowc
'xlSheet.Cells(k, 8) = InStr(xlSheet.Cells(k, 2), Me.Frame1.Caption)
If InStr(UCase(xlSheet.Cells(k, 2)), UCase(key)) > 0 Then
add xlSheet.Cells(k, 1), xlSheet.Cells(k, 2), xlSheet.Cells(k, 3), xlSheet.Cells(k, 4), xlSheet.Cells(k, 5), xlSheet.Cells(k, 6), xlSheet.Cells(k, 7), Val(k)
End If
Next k
Case "借阅日期":
For k = 2 To rowc
'xlSheet.Cells(k, 8) = InStr(xlSheet.Cells(k, 2), Me.Frame1.Caption)
If InStr(UCase(xlSheet.Cells(k, 5)), UCase(key)) > 0 Then
add xlSheet.Cells(k, 1), xlSheet.Cells(k, 2), xlSheet.Cells(k, 3), xlSheet.Cells(k, 4), xlSheet.Cells(k, 5), xlSheet.Cells(k, 6), xlSheet.Cells(k, 7), Val(k)
End If
Next k
Case "借阅人姓名":
For k = 2 To rowc
'xlSheet.Cells(k, 8) = InStr(xlSheet.Cells(k, 2), Me.Frame1.Caption)
If InStr(UCase(xlSheet.Cells(k, 6)), UCase(key)) > 0 Then
add xlSheet.Cells(k, 1), xlSheet.Cells(k, 2), xlSheet.Cells(k, 3), xlSheet.Cells(k, 4), xlSheet.Cells(k, 5), xlSheet.Cells(k, 6), xlSheet.Cells(k, 7), Val(k)
End If
Next k
Case "联系方式":
For k = 2 To rowc
'xlSheet.Cells(k, 8) = InStr(xlSheet.Cells(k, 2), Me.Frame1.Caption)
If InStr(UCase(xlSheet.Cells(k, 7)), UCase(key)) > 0 Then
add xlSheet.Cells(k, 1), xlSheet.Cells(k, 2), xlSheet.Cells(k, 3), xlSheet.Cells(k, 4), xlSheet.Cells(k, 5), xlSheet.Cells(k, 6), xlSheet.Cells(k, 7), Val(k)
End If
Next k
Case "系列名称":
For k = 2 To rowc
'xlSheet.Cells(k, 8) = InStr(xlSheet.Cells(k, 2), Me.Frame1.Caption)
If InStr(UCase(xlSheet.Cells(k, 3)), UCase(key)) > 0 Then
add xlSheet.Cells(k, 1), xlSheet.Cells(k, 2), xlSheet.Cells(k, 3), xlSheet.Cells(k, 4), xlSheet.Cells(k, 5), xlSheet.Cells(k, 6), xlSheet.Cells(k, 7), Val(k)
End If
Next k
Case Else:
For k = 2 To rowc
'xlSheet.Cells(k, 8) = InStr(xlSheet.Cells(k, 2), Me.Frame1.Caption)
If InStr(UCase(xlSheet.Cells(k, 2)), UCase(key)) > 0 Then
add xlSheet.Cells(k, 1), xlSheet.Cells(k, 2), xlSheet.Cells(k, 3), xlSheet.Cells(k, 4), xlSheet.Cells(k, 5), xlSheet.Cells(k, 6), xlSheet.Cells(k, 7), Val(k)
End If
Next k
End Select
Me.Label2 = "有关" & key & "的搜索结果共有" & Me.ListView1.ListItems.Count & "项"
End Sub

Private Sub Command4_Click()
addb.Show
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()

i = 0
ListView1.View = lvwReport
  ListView1.HideSelection = False
ListView1.FullRowSelect = True
ListView1.ColumnHeaders.add , "line", "编录行号"
ListView1.ColumnHeaders.add , "Number", "编号"
ListView1.ColumnHeaders.add , "Name", "书名"
ListView1.ColumnHeaders.add , "xl", "系列名称"
ListView1.ColumnHeaders.add , "state", "借阅状态"
ListView1.ColumnHeaders.add , "date", "借阅日期"
ListView1.ColumnHeaders.add , "bname", "借阅人姓名"
ListView1.ColumnHeaders.add , "contact", "联系方式"
Set xlApp = CreateObject("Excel.Application") '创建Excel对象
Set XlBook = xlApp.Workbooks.Open(data.file) '打开已经存在的Excel工作
Set xlSheet = XlBook.Worksheets(1)



End Sub
Private Function add(number As String, name As String, xl As String, state As Integer, bdate As String, bname As String, contact As String, line As String)
 i = i + 1
 Dim itmx As ListItem
  '添加column1的名称。
    Set itmx = ListView1.ListItems.add(i, , line)
    '使用SubItemIndex将SubItem与正确的ColumnHeader关联。使用关键字("Sex")指定正确的ColumnHeader。
    itmx.SubItems(ListView1.ColumnHeaders("Name").SubItemIndex) = name
    
    '使用ColumnHeader关键字将SubItems字符串与
'正确的ColumnHeader关联。
If state = 1 Then
itmx.SubItems(ListView1.ColumnHeaders("state").SubItemIndex) = "已借阅"
Else
itmx.SubItems(ListView1.ColumnHeaders("state").SubItemIndex) = "未借阅"
End If
itmx.SubItems(ListView1.ColumnHeaders("Number").SubItemIndex) = number
itmx.SubItems(ListView1.ColumnHeaders("xl").SubItemIndex) = xl
itmx.SubItems(ListView1.ColumnHeaders("date").SubItemIndex) = bdate
itmx.SubItems(ListView1.ColumnHeaders("contact").SubItemIndex) = contact
itmx.SubItems(ListView1.ColumnHeaders("bname").SubItemIndex) = bname
End Function
 


Private Sub Form_Unload(Cancel As Integer)
XlBook.Close '关闭工作簿
Set XlBook = Nothing '从内存中清除
xlApp.Quit '关闭excel
Set xlApp = Nothing '从内存中清除
End Sub



Private Sub ListView1_DblClick()
If xlSheet.Cells(state.ListView1.SelectedItem, 4) = 0 Then
borrow.Show
Me.Visible = False
Else
re.Show
Me.Visible = False
End If
End Sub
