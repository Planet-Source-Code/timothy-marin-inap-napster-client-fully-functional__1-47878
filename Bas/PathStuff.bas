Attribute VB_Name = "PathStuff"
      Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" _
  (ByVal hWndOwner As Long, ByVal nFolder As Long, _
  pidl As ITEMIDLIST) As Long

Declare Function SHGetPathFromIDList Lib "Shell32.dll" _
  Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
  ByVal pszPath As String) As Long
  
Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

      Private Const BIF_RETURNONLYFSDIRS = 1
      Private Const BIF_DONTGOBELOWDOMAIN = 2
      Private Const MAX_PATH = 260

      Private Declare Function SHBrowseForFolder Lib "shell32" _
                                        (lpbi As BrowseInfo) As Long
                                        
      Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                        (ByVal lpString1 As String, ByVal _
                                        lpString2 As String) As Long
Public Type SH_ITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SH_ITEMID
End Type

      Private Type BrowseInfo
         hWndOwner      As Long
         pIDLRoot       As Long
         pszDisplayName As Long
         lpszTitle      As Long
         ulFlags        As Long
         lpfnCallback   As Long
         lParam         As Long
         iImage         As Long
      End Type
      
Public Function FileExist(strFile As String) As Boolean
    If PathFileExists(strFile) = 1 Then
        FileExist = True
    ElseIf PathFileExists(strFile) = 0 Then
        FileExist = False
    End If
End Function

Public Function fGetSpecialFolder(CSIDL As Long) As String
    Dim sPath As String
    Dim IDL As ITEMIDLIST
    '
    ' Retrieve info about system folders such as the
    ' "Desktop" folder. Info is stored in the IDL structure.
    '
    fGetSpecialFolder = ""
    If SHGetSpecialFolderLocation(frmMain.hWnd, CSIDL, IDL) = 0 Then
        '
        ' Get the path from the ID list, and return the folder.
        '
        sPath = Space$(MAX_PATH)
        If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath) Then
            fGetSpecialFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1) & "\"
        End If
    End If
End Function

Public Function SelectFolder(coosh As Long)
      'Opens a Treeview control that displays the directories in a computer

         Dim lpIDList As Long
         Dim sBuffer As String
         Dim szTitle As String
         Dim tBrowseInfo As BrowseInfo

         szTitle = "Select Folder"
         With tBrowseInfo
            .hWndOwner = coosh
            .lpszTitle = lstrcat(szTitle, "")
            .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
         End With

         lpIDList = SHBrowseForFolder(tBrowseInfo)

         If (lpIDList) Then
            sBuffer = Space(MAX_PATH)
            SHGetPathFromIDList lpIDList, sBuffer
            sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
            SelectFolder = sBuffer
         End If
         
      End Function



