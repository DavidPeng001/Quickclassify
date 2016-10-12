VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ļ���һ������"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   375
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
   ScaleHeight     =   5415
   ScaleWidth      =   6465
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton CmdOk 
      Caption         =   "����"
      Height          =   600
      Left            =   240
      TabIndex        =   3
      Top             =   4080
      Width           =   1590
   End
   Begin VB.ListBox List 
      Height          =   2595
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


Private Sub CmdOk_Click()

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
    
    CmdOk.Enabled = False
    CmdOpen.Enabled = False
    LblTips.Visible = True
    
    strPath = TxtDir & IIf(Right$(TxtDir, 1) <> "\", "\,", "")

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder_a = fso.GetFolder(TxtDir) '·����Ϊ����ļ���A·��

    List.Clear
 
    For Each f In folder_a.Files
            List.AddItem f & "     " & f.DateCreated
        
    Next
    
    LblTips = "������ϣ�������£�"
    CmdOk.Enabled = True
    CmdOpen.Enabled = True
    Exit Sub
ErrHandler:                             '�û�����ȡ������ť��
    Exit Sub
End Sub

Private Sub CmdOpen_Click()
    TxtDir = GetDirectory
End Sub

