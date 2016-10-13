VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ļ���һ������"
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
      Caption         =   "�Զ�������"
      Height          =   615
      Left            =   4200
      TabIndex        =   6
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "һ������"
      Height          =   615
      Left            =   2160
      TabIndex        =   5
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton CmdSearch 
      Caption         =   "����"
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
      Caption         =   "ѡ���ļ���"
      Height          =   360
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label LblTips 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ǰ���������Ҫ�ļ������رո�Ŀ¼���������еĳ�����ĵ���"
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
        MsgBox "��ѡ��Ҫ������ļ���", vbInformation, "��ʾ"
        Exit Sub
    End If
    'Trim$ can remove space in two sides of the string
    If Dir(TxtDir, vbDirectory) = "" Then
        MsgBox "·�����Ϸ���������ѡ��", vbInformation, "��ʾ"
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
    Set folder_a = fso.GetFolder(TxtDir) '·����Ϊ����ļ���A·��

    List.Clear
 
    For Each f In folder_a.Files
            List.AddItem f & "   "
            '"" & f.DateCreated
        
    Next
    
    LblTips = "������ϣ�������£�"
    CmdSearch.Enabled = True
    CmdOpen.Enabled = True
    Exit Sub
ErrHandler:                             '�û�����ȡ������ť��
    Exit Sub
End Sub

Private Sub CmdOpen_Click()
    TxtDir = GetDirectory
End Sub

Private Sub CmdStart_Click()
    IsMusic = True
    IsAudio = True
    IsFile = True
    
    'To do:remove space
    
    Dim strFormat As String
    Dim intLen As Integer
    Dim intFolder As Integer
    Dim strFile As String
    TxtDir = Trim$(TxtDir)
    intFolder = Len(TxtDir.Text)
    
    For n = 0 To List.ListCount Step 1
        List.List(n) = Trim$(List.List(n))
        strFormat = Right(List.List(n), 3)
        intLen = Len(List.List(n))
        strFile = Right(List.List(n), intLen - intFolder - 1)
        If (strFormat = "mp3" Or strFormat = "mav") And IsMusic = True Then
                        
           ' Name List.List(n) As (TxtDir + "\Music\" + strFile)
        ElseIf (strFormat = "txt") And IsMusic = True Then
            
            Name List.List(n) As (TxtDir + "\File\" + strFile)
            
        End If
    Next n

End Sub
