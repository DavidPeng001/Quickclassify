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
        bInfo.lpszTitle = "请选择你要整理的文件夹"
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

