Attribute VB_Name = "Module1"
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public MyPassword As String
Public UsePassword As Boolean
Public ExitExtract As Boolean
Public UsePath As String
Public OneFileOnly As Boolean

Function GetPicName(FileTitle As String) As String

    Dim Ext As String
    Ext = LCase(FileTitle)
    'EXT = Right(EXT, 3)
    Ext = GetExtension(Ext)
    Select Case Ext
        Case "bmp", "jpg", "psd", "psp"
            GetPicName = "bmp"
        Case "gif"
            GetPicName = "gif"
        Case "exe"
            GetPicName = "exe"
        Case "hlp"
            GetPicName = "hlp"
        Case "txt"
            GetPicName = "txt"
        Case "frm"
            GetPicName = "frm"
        Case "zip"
            GetPicName = "zfile"
        Case "rar"
            GetPicName = "rar"
        Case "doc", "dot"
            GetPicName = "doc"
        Case "xls"
            GetPicName = "xls"
         Case "htm", "html", "shtml", "dhtml", "php3"
            GetPicName = "htm"
         Case "wav", "mp3", "mp2"
            GetPicName = "wav"
         Case "reg"
            GetPicName = "reg"
         Case "ttf"
            GetPicName = "ttf"
          Case "log"
            GetPicName = "log"
          Case "ini", "inf"
            GetPicName = "ini"
          Case "bas"
            GetPicName = "bas"
          Case "mdb"
            GetPicName = "mdb"
          Case "dll", "ocx"
            GetPicName = "dll"
          Case "vbp"
            GetPicName = "vbp"
          Case "cls"
            GetPicName = "cls"
          Case "ctl"
            GetPicName = "ctl"
          Case "pdf"
            GetPicName = "pdf"
          Case "psp"
            GetPicName = "psd"
        Case Else
            GetPicName = "win"
    End Select
    
End Function

Function GetExtension(mFile As String) As String
    GetExtension = Right(mFile, Len(mFile) - InStrRev(mFile, "."))
End Function


Function GetType(FileTitle As String) As String

    Dim Ext As String
    Ext = LCase(FileTitle)
    'EXT = Right(EXT, 3)
    Ext = GetExtension(Ext)
    Select Case Ext
        Case "bmp", "jpg", "psd", "psp"
            GetType = "Windows Bitmap"
        Case "gif"
            GetType = "Picture"
        Case "exe"
            GetType = "Application"
        Case "hlp"
            GetType = "Help File"
        Case "txt"
            GetType = "Text File"
        Case "frm"
            GetType = "Visual Basic Form"
        Case "zip", "rar"
            GetType = "Compressed File"
        Case "doc", "dot"
            GetType = "Microsoft Word Document"
        Case "xls"
            GetType = "Microsoft Excel Sheet"
         Case "htm", "html", "shtml", "dhtml", "php3"
            GetType = "HTML File"
         Case "wav", "mp3", "mp2"
            GetType = "Sound File"
         Case "reg"
            GetType = "Registry File"
         Case "ttf"
            GetType = "True Type Font"
          Case "log"
            GetType = "Log File"
          Case "ini", "inf"
            GetType = "Configuration File"
          Case "bas"
            GetType = "Visual Basic Module"
          Case "mdb"
            GetType = "Microsoft Access Database"
          Case "dll", "ocx"
            GetType = "Application Extension"
          Case "vbp"
            GetType = "Visual Basic Project File"
          Case "cls"
            GetType = "Visual Basic Class Module"
          Case "ctl"
            GetType = "Visual Basic Control"
          Case "pdf"
            GetType = "Adobe Acrobat Portable File"
          Case "psp"
            GetType = "Adobe Photoshop Document"
        Case Else
            GetType = UCase(Ext) & " File"
    End Select
    
End Function

'Function ValidateDir(ByVal aDir As String) As Boolean
'    Const ValidChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ01234567890\.()+=-_"
'    Dim i As Integer, Root As String, Char As String
'    For i = 1 To Len(aDir)
'        Char = Mid(Ext, i, 1)
'        If InStr(ValidChars, Char) = 0 Then
'            ValidateDir = False
'            Exit Function
'        End If
'    Next i
'    ValidateDir = True
'End Function


Public Function MakeSureDirExists(TheDirectory As String) As Boolean
    Dim sDirTest As String
    Dim i As Long
    Dim sPath As String
    Dim iCounter As Integer
    Dim sTempDir As String
    Dim NewDirs As New Collection
    sPath = TheDirectory
    
    On Error GoTo XF
    
    If Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If
    
    iCounter = 1
    
    Do Until i = Len(sPath)
        iCounter = InStr(iCounter, sPath, "\")
        If Len(sTempDir) = Len(sPath) Then Exit Do
        If iCounter <> 0 Then
            sTempDir = Left(sPath, iCounter)
            sDirTest = Dir(sTempDir, vbDirectory)
            If sDirTest = "" And Right(sTempDir, 2) <> ":\" Then
                MkDir sTempDir
                NewDirs.Add sTempDir
            End If
        End If

        iCounter = iCounter + 1
        i = i + 1
    Loop
    MakeSureDirExists = True
    Set NewDirs = Nothing
    Exit Function
XF:
    For i = 1 To NewDirs.Count
        If Dir(NewDirs(i)) <> "" Then
            RmDir (NewDirs(i))
        End If
    Next i
    Set NewDirs = Nothing
    MakeSureDirExists = False
End Function

Public Function GetSysDir() As String
    Dim strSyspath As String
    strSyspath = String(145, Chr(0))
    strSyspath = Left(strSyspath, GetSystemDirectory(strSyspath, 145))
    GetSysDir = strSyspath
End Function

Public Function ExtractLibs(intResNr As Integer, strPath As String) As Boolean
    Dim intFileNumber As Integer
    Dim bLibBuffer() As Byte
    '101 = Unzip32.dll
    '102 = zip32.dll
    On Error GoTo Errhandler
    bLibBuffer = LoadResData(intResNr, "DLL")
    intFileNumber = FreeFile
    Open strPath For Binary Access Write As #intFileNumber
        Put #intFileNumber, , bLibBuffer
    Close #intFileNumber
    On Error GoTo 0
    ExtractLibs = True
    Exit Function
Errhandler:
    ExtractLibs = False
End Function
