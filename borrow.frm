VERSION 5.00
Begin VB.Form borrow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "书籍借阅"
   ClientHeight    =   6735
   ClientLeft      =   6345
   ClientTop       =   2625
   ClientWidth     =   6525
   Icon            =   "borrow.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   6525
   Begin VB.CommandButton Command2 
      Caption         =   "取消借阅"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   14
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确认借阅"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   13
      Top             =   5520
      Width           =   1935
   End
   Begin VB.TextBox contact 
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
      Left            =   1920
      TabIndex        =   12
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   "借阅人信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   5775
      Begin VB.TextBox bbname 
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
         Left            =   1680
         TabIndex        =   11
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label7 
         Caption         =   "联系方式："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   10
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "姓    名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "日    期："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   8
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label bdate 
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
         Left            =   1800
         TabIndex        =   7
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         TabIndex        =   6
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "书籍信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      Begin VB.Label number 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         TabIndex        =   4
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label bname 
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
         Left            =   1680
         TabIndex        =   3
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "编   号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "书   名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
   End
End
Attribute VB_Name = "borrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Me.bbname <> "" And Me.contact <> "" Then
With state.xlSheet
.Cells(state.ListView1.SelectedItem, 4) = 1
.Cells(state.ListView1.SelectedItem, 6) = bbname
.Cells(state.ListView1.SelectedItem, 5) = bdate
.Cells(state.ListView1.SelectedItem, 7) = contact
End With
state.XlBook.Save
Dim itmx As ListItem
Set itmx = state.ListView1.SelectedItem
itmx.SubItems(4) = "已借阅"  'text1.text 第几列 text2.text 内容
itmx.SubItems(6) = bbname
itmx.SubItems(5) = bdate
itmx.SubItems(7) = contact
Set itmx = Nothing
Unload Me
Else
MsgBox "联系人信息不全", vbExclamation, "错误"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
With state.xlSheet
Me.bname = .Cells(state.ListView1.SelectedItem, 2)
Me.number = .Cells(state.ListView1.SelectedItem, 1)
End With
Me.bdate = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
state.Visible = True
End Sub
