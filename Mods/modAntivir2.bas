Attribute VB_Name = "modAntivir2"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit

Private Function OpenKey(lhKey As Long, SubKey As String, ulOptions As Long) As Long
Dim lhKeyOpen As Long
Dim lResult As Long

lhKeyOpen = 0
lResult = RegOpenKeyEx(lhKey, SubKey, 0, ulOptions, lhKeyOpen)

If lResult <> ERROR_SUCCESS Then
OpenKey = 0
Else
OpenKey = lhKeyOpen
End If
End Function

Private Function CreateKey(lhKey As Long, SubKey As String, NewSubKey As String) As Boolean
Dim lhKeyOpen As Long
Dim lhKeyNew As Long
Dim lDisposition As Long
Dim lResult As Long
Dim Security As SECURITY_ATTRIBUTES

lhKeyOpen = OpenKey(lhKey, SubKey, KEY_CREATE_SUB_KEY)
lResult = RegCreateKeyEx(lhKeyOpen, NewSubKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, Security, lhKeyNew, lDisposition)

If lResult = ERROR_SUCCESS Then
CreateKey = True
RegCloseKey (lhKeyNew)
Else
CreateKey = False
End If

RegCloseKey (lhKeyOpen)
End Function

Private Function SetValue(lhKey As Long, SubKey As String, sValue As String) As Boolean
Dim lhKeyOpen As Long
Dim lResult As Long
Dim lTyp As Long
Dim lByte As Long

lByte = Len(sValue)
lTyp = REG_SZ
lhKeyOpen = OpenKey(lhKey, SubKey, KEY_SET_VALUE)
lResult = RegSetValue(lhKey, SubKey, lTyp, sValue, lByte)

If lResult <> ERROR_SUCCESS Then
SetValue = False
Else
SetValue = True
RegCloseKey (lhKeyOpen)
End If
End Function

Public Function RegisterFile(sFileExt As String, sFileDescr As String, sAppID As String, sOpenCmd As String, sIconFile As String) As Boolean
Dim hKey As Long
Dim bSuccess As Boolean
Dim bSuccess2 As Boolean
    
bSuccess = False
hKey = HKEY_LOCAL_MACHINE
  
If CreateKey(hKey, REG_PRIMARY_KEY, sFileExt) Then
 If SetValue(hKey, REG_PRIMARY_KEY & sFileExt, sAppID) Then
  If CreateKey(hKey, REG_PRIMARY_KEY, sAppID) Then
   If SetValue(hKey, REG_PRIMARY_KEY & sAppID, sFileDescr) Then
    If CreateKey(hKey, REG_PRIMARY_KEY & sAppID, _
        REG_SHELL_KEY & REG_SHELL_OPEN_KEY & _
        REG_SHELL_OPEN_COMMAND_KEY) Then
        bSuccess = SetValue(hKey, REG_PRIMARY_KEY & _
        sAppID & "\" & REG_SHELL_KEY & _
        REG_SHELL_OPEN_KEY & _
        REG_SHELL_OPEN_COMMAND_KEY, sOpenCmd)
     If CreateKey(hKey, REG_PRIMARY_KEY & sAppID, _
        REG_ICON_KEY) Then
        bSuccess2 = SetValue(hKey, REG_PRIMARY_KEY & _
        sAppID & "\" & REG_ICON_KEY, sIconFile)
     End If
    End If
   End If
  End If
 End If
End If

RegisterFile = (bSuccess = bSuccess2)
End Function
Public Function CalcCRC(strfilename As String) As String

  Dim cCRC32  As New cCRC32
  Dim lCRC32  As Long
  Dim cStream As New cBinaryFileStream

    cStream.file = strfilename
    lCRC32 = cCRC32.GetFileCrc32(cStream)
    CalcCRC = Hex$(lCRC32)

End Function

Public Function LongFilename( _
    ByRef FilePath As String) As String
  Const INVALID_VALUE = -1
  Dim hFind As Long
  Dim WFD As WIN32_FIND_DATA
  
  hFind = FindFirstFileA(FilePath, WFD)
  If hFind <> INVALID_VALUE Then
    LongFilename = LeftB$(WFD.cFileName, _
        InStrB(WFD.cFileName, vbNullChar))
    FindClose hFind
  End If
End Function

Public Sub CheckExe()

    On Error GoTo Ignore
    If GetSetting(AV.AVname, "Settings", "CRC", CalcCRC(App.Path & "\" & App.EXEName & ".exe")) <> CalcCRC(App.Path & "\" & App.EXEName & ".exe") Then
        MsgBox "Program heavily damaged!!!", vbCritical + vbOKOnly, "Error"
        End
    End If
    SaveSetting AV.AVname, "Settungs", "CRC", CalcCRC(App.Path & "\" & App.EXEName & ".exe")
Ignore:

End Sub

Public Sub Checkfolder(Optional ByVal StrFolder As String)

  Dim Result As Variant
  Dim c      As Collection

    On Error Resume Next
    If StrFolder = vbNullString Then
        Set Result = SH.BrowseForFolder(frmMain.hWnd, "Select the folder you want to have scanned", 1)
    End If
    With Result.Items.Item
        FullPathSearch .Path, c, , , , True
    End With 'RESULT.ITEMS.ITEM
    On Error GoTo 0

End Sub

Private Function FindFiles(ByVal Path As String, _
                           ByRef Files As Collection, _
                           Optional ByVal Pattern As String = "*.*", _
                           Optional ByVal Attributes As VbFileAttribute = vbNormal, _
                           Optional ByVal Recursive As Boolean = True) As Long

  Const vbErr_PathNotFound As Long = 76
  Const INVALID_VALUE      As Long = -1
  Dim FileAttr             As Long
  Dim FileName             As String
  Dim hFind                As Long
  Dim WFD                  As WIN32_FIND_DATA

    If Mid$(Path, Len(Path) - 1, 1) <> "\" Then
        Path = Path & "\"
    End If
    If Files Is Nothing Then
        Set Files = New Collection
    End If
    Pattern = LCase$(Pattern)
    hFind = FindFirstFileA(Path & "*", WFD)
    If hFind = INVALID_VALUE Then
        err.Raise vbErr_PathNotFound
    End If
    Do
        FileName = LeftB$(WFD.cFileName, InStrB(WFD.cFileName, vbNullChar))
        FileAttr = GetFileAttributesA(Path & FileName)
        If FileAttr And vbDirectory Then
            If Recursive Then
                If FileAttr <> INVALID_VALUE Then
                    If FileName <> "." Then
                        If FileName <> ".." Then
                            FindFiles = FindFiles + FindFiles(Path & FileName, Files, Pattern, Attributes)
                        End If
                    End If
                End If
            End If
         Else 'NOT FILEATTR...
            If (FileAttr And Attributes) = Attributes Then
                If LCase$(FileName) Like Pattern Then
                    FindFiles = FindFiles + 1
                    Files.Add Path & FileName
                End If
            End If
        End If
    Loop While FindNextFileA(hFind, WFD)
    FindClose hFind

End Function

Public Sub FullPathSearch(ByRef Path As String, _
                          ByRef Files As Collection, _
                          Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, _
                          Optional ByVal Pattern As String = "*.*", _
                          Optional ByVal Attributes As VbFileAttribute = vbNormal, _
                          Optional ByVal Recursive As Boolean = False)

    
  Dim Candidates As Collection
  Dim file       As Variant

    Running = True
    If Files Is Nothing Then
        Set Files = New Collection
    End If
    FindFiles Path, Candidates, Pattern, Attributes, Recursive
    For Each file In Candidates
        If CheckFile(file) Then
            Exit Sub
        End If
        If Running = False Then
            Exit Sub
        End If
    Next file

End Sub
Public Sub Log(strLog As String)
Dim ff As Integer
ff = FreeFile
On Error Resume Next
MkDir App.Path & "\Logs"
Open App.Path & "\Logs\" & Replace(Date, "/", "_") & ".txt" For Append As #ff
Print #ff, "[" & Time & "] " & strLog
Close #ff
End Sub
