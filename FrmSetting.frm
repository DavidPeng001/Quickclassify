VERSION 5.00
Begin VB.Form FrmSetting 
   Caption         =   "�Զ�������"
   ClientHeight    =   2565
   ClientLeft      =   8775
   ClientTop       =   4350
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2565
   ScaleWidth      =   3945
   Begin VB.CommandButton CmdNext 
      Caption         =   "����"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.CheckBox ChkPicture 
      Caption         =   "ͼƬ"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CheckBox ChkFile 
      Caption         =   "�ĵ�"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CheckBox ChkAudiio 
      Caption         =   "��Ƶ"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CheckBox ChkMusic 
      Caption         =   "����"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton CmdBack 
      Caption         =   "����"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label LblTip 
      Caption         =   "��ѡ�������������ļ�����"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "FrmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBack_Click()
FrmMain.Show
Unload Me
End Sub

Private Sub CmdNext_Click()
    IsMusic = True
    IsAudio = True
    IsFile = True
    IsPicture = True
    
    If ChkMusic = 1 Then
        IsMusic = False
    End If
    If ChkAudio = 1 Then
        IsAudio = False
    End If
    If ChkFile = 1 Then
        IsFile = False
    End If
    If ChkPicture = 1 Then
        IsPicture = False
    End If

    x = MsgBox("ȷ�������ļ�������", vbQuestion + vbOKCancel, "��ʾ")
    
    If x = 1 Then
        Call Classify
        FrmMain.CmdStart.Enabled = False
        FrmMain.CmdSetting.Enabled = False
    End If
    
    FrmMain.Show
    Unload Me
End Sub
