VERSION 5.00
Begin VB.Form addb 
   Caption         =   "增加图书"
   ClientHeight    =   4710
   ClientLeft      =   6720
   ClientTop       =   3375
   ClientWidth     =   6225
   Icon            =   "add.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4710
   ScaleWidth      =   6225
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
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
      Left            =   3360
      TabIndex        =   9
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确认增加"
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
      Left            =   720
      TabIndex        =   8
      Top             =   3720
      Width           =   1695
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
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5775
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
         ItemData        =   "add.frx":000C
         Left            =   1440
         List            =   "add.frx":002B
         TabIndex        =   5
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox number 
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
         Left            =   1440
         TabIndex        =   4
         Top             =   2040
         Width           =   3135
      End
      Begin VB.TextBox bname 
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
         Left            =   1440
         TabIndex        =   3
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "系   列："
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
         Left            =   480
         TabIndex        =   7
         Top             =   1320
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
         Height          =   495
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Width           =   1455
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
         TabIndex        =   1
         Top             =   2040
         Width           =   1455
      End
   End
   Begin VB.Label Label3 
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
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
End
Attribute VB_Name = "addb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
r = state.xlSheet.[A65536].End(xlUp).Row + 1
If Me.bname <> "" And Me.number <> "" And Me.Combo1.Text <> "" Then
With state.xlSheet
.Cells(r, 1) = number
.Cells(r, 2) = "《" & bname & "》"
.Cells(r, 3) = Me.Combo1
.Cells(r, 4) = 0
End With
state.XlBook.Save
MsgBox "已成功增加" & "《" & bname & "》", vbInformation, "成功"
Unload Me
Else
MsgBox "书籍信息不全", vbExclamation, "错误"
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

