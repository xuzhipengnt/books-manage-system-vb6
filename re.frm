VERSION 5.00
Begin VB.Form re 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "确认归还"
   ClientHeight    =   4470
   ClientLeft      =   4755
   ClientTop       =   3480
   ClientWidth     =   5880
   Icon            =   "re.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   5880
   Begin VB.CommandButton Command1 
      Caption         =   "取消归还"
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
      Left            =   3360
      TabIndex        =   1
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton re 
      Caption         =   "确认归还"
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
      Left            =   600
      TabIndex        =   0
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label contact 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label bdate 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label bbname 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label number 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label bname 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "联系方式:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "借阅人:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label dasf 
      BackStyle       =   0  'Transparent
      Caption         =   "借阅日期:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "书名:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "确信要归还这本书吗？"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "代号:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "re"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Index = state.ListView1.SelectedItem
With state.xlSheet
Me.bname = .Cells(state.ListView1.SelectedItem, 6)
Me.number = .Cells(state.ListView1.SelectedItem, 1)
Me.bbname = .Cells(state.ListView1.SelectedItem, 2)
Me.bdate = .Cells(state.ListView1.SelectedItem, 5)
Me.contact = .Cells(state.ListView1.SelectedItem, 7)
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
state.Visible = True
End Sub

Private Sub re_Click()
With state.xlSheet
.Cells(state.ListView1.SelectedItem, 4) = 0
.Cells(state.ListView1.SelectedItem, 5) = ""
.Cells(state.ListView1.SelectedItem, 6) = ""
.Cells(state.ListView1.SelectedItem, 7) = ""
End With
state.XlBook.Save
Dim itmx As ListItem
Set itmx = state.ListView1.SelectedItem
itmx.SubItems(4) = "未借阅"  'text1.text 第几列 text2.text 内容
itmx.SubItems(5) = ""
itmx.SubItems(6) = ""
itmx.SubItems(7) = ""
Set itmx = Nothing
Unload Me
End Sub
