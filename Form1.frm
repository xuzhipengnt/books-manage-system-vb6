VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6255
   ClientLeft      =   5085
   ClientTop       =   3435
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2160
      Top             =   2760
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "图书版        "
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2055
      Left            =   4200
      TabIndex        =   1
      Top             =   3720
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "个人书籍杂志借阅管理系统"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2055
      Left            =   3600
      TabIndex        =   0
      Top             =   1440
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   5775
      Left            =   0
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public key
Dim tt As Integer
Private Sub Form_Load()
data.file = App.Path & "\data.xls"
If App.PrevInstance Then
Me.Timer1.Enabled = False
    MsgBox "这个程序已经打开，不能再次运行", vbExclamation, "重复运行错误"
    AppActivate App.Title
  End
End If
If IsFileOpen(data.file) Then
Me.Timer1.Enabled = False
 MsgBox "data.xls文档已经被打开,请退出该文档再进行操作" & Dir(data.file), vbExclamation, "程序将退出"
 End
End If

tt = 3
End Sub

  Public Function IsFileOpen(sFile As Variant) As Boolean
          IsFileOpen = False
          Dim openFile     As New FileSystemObject, targetFileName           As String
          If Not openFile.FileExists(sFile) Then
                  MsgBox "文件不存在！"
                  Exit Function
          End If
            
          targetFileName = "c:\temp"
          On Error GoTo ErrOpen
          openFile.MoveFile sFile, targetFileName
          openFile.MoveFile targetFileName, sFile
          Debug.Print targetFileName
          Exit Function
ErrOpen:
          IsFileOpen = True
  End Function
  


Private Sub Label2_Click()

End Sub

Private Sub Timer1_Timer()
tt = tt - 1
If tt < 0 Then
state.Show
Unload Me
End If

End Sub
