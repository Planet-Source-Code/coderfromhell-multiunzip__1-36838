Attribute VB_Name = "ModUnzip"
Public WhatDoIDo As Integer
Public BTNCaption As Integer
Public ShouldILoad As Boolean
Public IhaveDoneIt As Boolean
Public BadFileNames As New Collection
Public Test As Boolean
Public PasswordNeeded As Boolean

'###########################
'Public Enum EXT
'    ExtractAllFiles_D = 0
'    ExtractSelectedFiles_D = 1
'    CompressAllFiles_D = 2
'    CompressSelectedFiles_D = 3
'End Enum
'
'Public Enum BTNCAP
'    Compress = 0
'    Decompress = 1
'End Enum

'###########################

Function UnzipZip(ZipPath As String, ZipFile As String, ExtractTo As String, HonorDirs As Boolean) As Boolean

On Error GoTo vbErrorHandler

If Right(ZipPath, 1) <> "\" Then
    ZipPath = ZipPath & "\"
End If

If ZipFile = "" Then
    UnzipZip = True
    Exit Function
End If
'
' Unzip the ZIPTEST.ZIP file to the Windows Temp Directory
'
    Dim oUnZip As CGUnzipFiles
    
    Set oUnZip = New CGUnzipFiles
    
    With oUnZip
'
' What Zip File ?
'
        .ZipFileName = ZipPath & ZipFile
'
' Where are we zipping to ?
'
        .ExtractDir = ExtractTo
'
' Keep Directory Structure of Zip ?
'
        .HonorDirectories = HonorDirs
        
        .TestZip = False
'
' Unzip and Display any errors as required

'
        If .Unzip <> 0 Then
            'MsgBox .GetLastMessage '<<<<<<<<<<<<<<
            IhaveDoneIt = False
            BadFileNames.Add ZipFile
        End If
    End With
    
    Set oUnZip = Nothing
    'MsgBox "\ZIPTEST.ZIP Extracted Successfully to " & "G:\other stuff\"
    UnzipZip = True
    
    Exit Function

vbErrorHandler:
    MsgBox Err.Number & " " & "Form1::cmdUnZip_Click" & " " & Err.Description
    UnzipZip = False
End Function

