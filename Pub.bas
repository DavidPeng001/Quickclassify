Attribute VB_Name = "Pub"
Public IsMusic As Boolean
Public IsAudio As Boolean
Public IsFile As Boolean
Public IsPicture As Boolean


Declare Function SHGetPathFromIDList _
        Lib "shell32.dll" _
        Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
                                      ByVal pszPath As String) As Long
' SHGetPathFromIDListA can get a file path in a disk

Declare Function SHBrowseForFolder _
        Lib "shell32.dll" _
        Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
' SHGetPathFromIDListA can get a folder path in a disk

Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Function GetDirectory(Optional Msg) As String
    Dim bInfo As BROWSEINFO
    Dim path  As String
    Dim r     As Long, x As Long, pos As Integer
    ' Root folder = Desktop
    bInfo.pidlRoot = 0&

    ' Title in the dialog
    If IsMissing(Msg) Then
        bInfo.lpszTitle = "��ѡ����Ҫ������ļ���"
    Else
        bInfo.lpszTitle = Msg
    End If
    ' Type of directory to return
    bInfo.ulFlags = &H1
    ' Display the dialog
    x = SHBrowseForFolder(bInfo)
    'SHBrowseForFolder can call BrowseForFolder Dialog
    
    ' Parse the result
    path = Space$(512)
    ' return some spase
    r = SHGetPathFromIDList(ByVal x, ByVal path)

    If r Then
        pos = InStr(path, Chr$(0))
        GetDirectory = Left(path, pos - 1)
        'the function instr can find and return the spot of the 2nd string in 1st string
        'the function LEFT can return some charcter from the left of string
        
    Else
        GetDirectory = ""
    End If
End Function


Sub Classify()

    Dim strFormat As String
    Dim intLen As Integer
    Dim intFolder As Integer
    Dim strFile As String
    TxtDir = Trim$(FrmMain.TxtDir)
    intFolder = Len(FrmMain.TxtDir)
    
    For n = 0 To FrmMain.List.ListCount - 1 Step 1
    
        FrmMain.List.List(n) = Trim$(FrmMain.List.List(n))
        
        strFormat = Right(FrmMain.List.List(n), 3)
        intLen = Len(FrmMain.List.List(n))
        strFile = Right(FrmMain.List.List(n), intLen - intFolder - 1)
        
        If (strFormat = "mp3" Or strFormat = "mav" Or strFormat = "acc" Or strFormat = "lac" Or strFormat = "wma" Or strFormat = "m4a") And IsMusic = True Then
            If Dir(FrmMain.TxtDir + "\����\", vbDirectory) = "" Then
                 MkDir (FrmMain.TxtDir + "\����\")
            End If
            
            Name FrmMain.List.List(n) As (FrmMain.TxtDir + "\����\" + strFile)
            
        ElseIf (strFormat = "txt" Or strFormat = "doc" Or strFormat = "ocx" Or strFormat = "wps" Or strFormat = "ppt" Or strFormat = "pps" Or strFormat = "ptx" Or strFormat = "xls") And IsMusic = True Then
            If Dir(FrmMain.TxtDir + "\�ĵ�\", vbDirectory) = "" Then
                 MkDir (FrmMain.TxtDir + "\�ĵ�\")
            End If
            
            Name FrmMain.List.List(n) As (FrmMain.TxtDir + "\�ĵ�\" + strFile)
            
        ElseIf (strFormat = "mp4" Or strFormat = "mkv" Or strFormat = "mvb" Or strFormat = "flv" Or strFormat = "mpg" Or strFormat = "mov" Or strFormat = "mob") And IsMusic = True Then
            If Dir(FrmMain.TxtDir + "\��Ƶ\", vbDirectory) = "" Then
                 MkDir (FrmMain.TxtDir + "\��Ƶ\")
            End If
            
            Name FrmMain.List.List(n) As (FrmMain.TxtDir + "\��Ƶ\" + strFile)
            
        ElseIf (strFormat = "jpg" Or strFormat = "png" Or strFormat = "gif" Or strFormat = "bmp" Or strFormat = "ico") And IsMusic = True Then
            If Dir(FrmMain.TxtDir + "\��Ƭ\", vbDirectory) = "" Then
                 MkDir (FrmMain.TxtDir + "\��Ƭ\")
            End If
            
            Name FrmMain.List.List(n) As (FrmMain.TxtDir + "\��Ƭ\" + strFile)
            
        End If
    Next n
    
    MsgBox "���������", vbInformation, "��ʾ"

End Sub


