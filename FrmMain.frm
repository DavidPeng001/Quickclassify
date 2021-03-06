VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "桌面A计划"
   ClientHeight    =   4920
   ClientLeft      =   9030
   ClientTop       =   3795
   ClientWidth     =   6465
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6465
   Begin VB.CommandButton CmdSetting 
      Caption         =   "自定义整理"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4320
      TabIndex        =   6
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "一键整理"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2400
      TabIndex        =   5
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton CmdSearch 
      Caption         =   "搜索"
      Height          =   600
      Left            =   240
      TabIndex        =   3
      Top             =   4080
      Width           =   1815
   End
   Begin VB.ListBox List 
      Height          =   2790
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   6015
   End
   Begin VB.TextBox TxtDir 
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton CmdOpen 
      Caption         =   "选择文件夹"
      Height          =   360
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label LblTips 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请提前备份你的重要文件，并关闭该目录下正在运行的程序或文档。"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   5400
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdSearch_Click()

    If Trim$(TxtDir) = "" Then
        MsgBox "请选择要整理的文件夹", vbInformation, "提示"
        Exit Sub
    End If
    'Trim$ can remove space in two sides of the string
    If Dir(TxtDir, vbDirectory) = "" Then
        MsgBox "路径不合法，请重新选择", vbInformation, "提示"
        TxtDir = ""
        Exit Sub
    End If
    

    Call ListFile
    
End Sub

Private Sub ListFile()
    Dim strPath As String
    
    CmdSearch.Enabled = False
    CmdOpen.Enabled = False
    LblTips.Visible = True
    
    strPath = TxtDir & IIf(Right$(TxtDir, 1) <> "\", "\,", "")

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder_a = fso.GetFolder(TxtDir) '路径设为你的文件夹A路径

    List.Clear
 
    For Each f In folder_a.Files
            List.AddItem f & "   "
            '"" & f.DateCreated
        
    Next
    
    LblTips = "搜索完毕，结果如下："
    CmdSearch.Enabled = True
    CmdOpen.Enabled = True
    CmdStart.Enabled = True
    CmdSetting.Enabled = True
    Exit Sub
ErrHandler:                             '用户按“取消”按钮。
    Exit Sub
End Sub

Private Sub CmdOpen_Click()
    TxtDir = GetDirectory
End Sub

Private Sub CmdSetting_Click()
    FrmMain.Hide
    FrmSetting.Show
End Sub

Private Sub CmdStart_Click()
    IsMusic = True
    IsAudio = True
    IsFile = True
    IsPicture = True
    x = MsgBox("确定进行文件整理吗？", vbQuestion + vbOKCancel, "提示")
    
    If x = 1 Then
        Call Classify
        CmdStart.Enabled = False
        CmdSetting.Enabled = False
    End If
End Sub
