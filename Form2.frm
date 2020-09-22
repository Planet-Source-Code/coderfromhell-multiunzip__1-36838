VERSION 5.00
Begin VB.Form frmExtract 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extract"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   7500
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3720
      Width           =   7560
      Begin VB.Label Label1 
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   60
         Width           =   7455
      End
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      Top             =   3360
      Width           =   3255
   End
   Begin VB.CommandButton cmdNewFolder 
      Caption         =   "New &Folder"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   3280
      Width           =   1215
   End
   Begin VB.Frame framePassword 
      Caption         =   "Password:"
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   2655
      Begin VB.CheckBox Check2 
         Caption         =   "Use this password for each archive."
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkMask 
         Caption         =   "Mask Password"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "§"
         TabIndex        =   8
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Extract Options:"
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton optFiles 
         Caption         =   "Extract Selected Files."
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   2415
      End
      Begin VB.OptionButton optFiles 
         Caption         =   "Extract All Files."
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.CheckBox chkStruct 
         Caption         =   "Keep Directory Structure."
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   2415
      End
   End
   Begin Project1.DirTree Dir1 
      Height          =   3135
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5530
      DefaultSelectedBold=   -1  'True
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "&Extract"
      Default         =   -1  'True
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FilesPath As String
'Public OneFileOnly As Boolean

Private Sub Check2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    QwickInfo 6
End Sub

Private Sub chkMask_Click()
    If chkMask.Value = 0 Then
        txtPassword.PasswordChar = ""
    Else
        txtPassword.PasswordChar = "§"
    End If
End Sub

Private Sub chkMask_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    QwickInfo 5
End Sub

Private Sub chkStruct_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    QwickInfo 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    QwickInfo 10
End Sub

Private Sub cmdExtract_Click()
    If MakeSureDirExists(txtPath.Text) Then
        Me.Hide
        MyPassword = txtPassword.Text
        UsePassword = IIf(Check2.Value = 1, True, False)
        If optFiles(0).Value = True Then
            ExtractAllFiles
        ElseIf optFiles(1).Value = True Then
            ExtractSelectedFiles
        End If
        Form1.Dir1.Refresh
        Form1.Dir1.Path = UsePath
        Unload Me
    Else
        MsgBox "Please enter a valid path!", vbOKOnly, "Invalid Path!"
        txtPath.SelStart = 0
        txtPath.SelLength = Len(txtPath.Text)
        Dir1.SetFocus
    End If
End Sub

Private Sub cmdExtract_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    QwickInfo 9
End Sub

Private Sub cmdNewFolder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    QwickInfo 11
End Sub

Private Sub Dir1_MouseMove(x As Single, y As Single)
    QwickInfo 7
End Sub

Private Sub Dir1_PathChange(NewPath As String)
    txtPath.Text = NewPath
End Sub

Private Sub Form_Load()
    Dir1.GoList
    Dir1.Path = Module1.UsePath
    If Right(txtPath, 1) <> "\" Then txtPath = txtPath & "\"
    If Module1.OneFileOnly Then
        Dim Stg As String, i As Integer
        i = Len(txtPath)
        Stg = Form1.LV1.SelectedItem.Text
        txtPath.Text = txtPath.Text & Left(Stg, Len(Stg) - 4)
        txtPath.SelStart = i
        txtPath.SelLength = Len(Stg)
        optFiles(1).Value = True
    Else
        txtPath.SelStart = Len(txtPath)
    End If
    Me.TestZip
    Show 1, Form1
End Sub

Public Sub TestZip()
    Test = True
    UnzipZip UsePath, Form1.LV1.SelectedItem.Text, "C:\Temp\", False
    Dim ctl As Control
    If PasswordNeeded = True Then
        framePassword.Enabled = True
        For Each ctl In Me.Controls
            If ctl.Container Is framePassword Then
                ctl.Enabled = True
            End If
        Next
    Else
        For Each ctl In Me.Controls
            If ctl.Container Is framePassword Then
                ctl.Enabled = False
            End If
        Next
        framePassword.Enabled = False
    End If
End Sub

''''''''''''''''
'''''''''''''''''
''''''''''''''''''
'''''''''''''''''''
''''''''''''''''''''    Extract Process --------------->>>>>
'''''''''''''''''''
''''''''''''''''''
'''''''''''''''''
''''''''''''''''












Private Sub ExtractAllFiles()

    Dim Completed As Boolean
    Dim i As Integer
    Dim ABC As String
    Dim oo As Long
    Dim ff As Long
        
    Let ABC = txtPath.Text
        
    If Right(ABC, 1) <> "\" Then ABC = ABC & "\"
    
    Load frmProgress
    frmProgress.FileCount = Form1.ZipCount
    frmProgress.lblFile.Caption = ""
    frmProgress.Show
    DoEvents
    
    For i = 1 To Form1.LV1.ListItems.Count
        If Form1.LV1.ListItems(i).Tag = "zipfile" Then
            If ExitExtract = True Then ExitExtract = False: Exit Sub
            frmProgress.FileName = Form1.LV1.ListItems(i)
            Test = False
            Completed = UnzipZip(UsePath, Form1.LV1.ListItems(i), ABC, IIf(chkStruct.Value = 1, True, False))
            frmProgress.UpdateProgress
            'Picture2.Height = Picture2.Height + Valu2Use
            'Picture2.Refresh
        End If
    Next i
    Unload frmProgress
    If Completed = True Then
        If IhaveDoneIt Then
            MsgBox "Extraction Successfull"
        Else
            'This is where you open a form with list of bad files
            If BadFileNames.Count = 0 Then
                MsgBox "Extraction Successfull"
                Exit Sub
            End If
            Load frmBadFiles
            For ff = 1 To BadFileNames.Count
                frmBadFiles.Text1.Text = frmBadFiles.Text1.Text & BadFileNames(ff) & vbCrLf
            Next ff
            frmBadFiles.Show 1
            Set BadFileNames = Nothing
            Set BadFileNames = New Collection
        End If
    Else
        MsgBox "Error, Can't complete extraction."
    End If
End Sub

Private Sub ExtractSelectedFiles()

    Dim Completed As Boolean
    Dim i As Integer
    Dim ABC As String
    Dim oo As Long
    Dim ff As Long
    Dim cCount As Long
    
    Let ABC = txtPath.Text
        
    If Right(ABC, 1) <> "\" Then ABC = ABC & "\"
    
    Load frmProgress
    cCount = 0
    For i = 1 To Form1.LV1.ListItems.Count
        If Form1.LV1.ListItems(i).Selected And Form1.LV1.ListItems(i).Tag = "zipfile" Then
            cCount = cCount + 1
        End If
    Next i
    frmProgress.FileCount = cCount
    frmProgress.lblFile.Caption = ""
    frmProgress.Show , Form1
    DoEvents
    For i = 1 To Form1.LV1.ListItems.Count
        If Form1.LV1.ListItems(i).Selected Then
            If Form1.LV1.ListItems(i).Tag = "zipfile" Then
                If ExitExtract = True Then ExitExtract = False: Exit Sub
                frmProgress.FileName = Form1.LV1.ListItems(i)
                Test = False
                Completed = UnzipZip(UsePath, Form1.LV1.ListItems(i), ABC, IIf(chkStruct.Value = 1, True, False))
                frmProgress.UpdateProgress
                'Picture2.Height = Picture2.Height + Valu2Use
                'Picture2.Refresh
            End If
        End If
    Next i
    Unload frmProgress
    If Completed = True Then
        If IhaveDoneIt Then
            MsgBox "Extraction Successfull"
        Else
            'This is where you open a form with list of bad files
            If BadFileNames.Count = 0 Then
                MsgBox "Extraction Successfull"
                Exit Sub
            End If
            Load frmBadFiles
            For ff = 1 To BadFileNames.Count
                frmBadFiles.Text1.Text = frmBadFiles.Text1.Text & BadFileNames(ff) & vbCrLf
            Next ff
            frmBadFiles.Show 1
            Set BadFileNames = Nothing
            Set BadFileNames = New Collection
        End If
    Else
        MsgBox "Error, Can't complete extraction."
    End If
End Sub


Private Sub QwickInfo(ID As Integer)
    Dim Info As String
    Select Case ID
        Case 0: Info = ""
        Case 1: Info = "This extracts the files into their subfolders."
        Case 2: Info = "Extracts all the files in the active folder."
        Case 3: Info = "Extracts all the selected files in the active folder."
        Case 4: Info = "Type a password here if all the archives require the same password."
        Case 5: Info = "Hides/Shows the password"
        Case 6: Info = "If this is checked this password will be used on each archive."
        Case 7: Info = "Select a folder to extract the archives contents to."
        Case 8: Info = "You can type your own path here and folders are automatically created."
        Case 9: Info = "Click here to start extracting."
        Case 10: Info = "Click here to cancel."
        Case 11: Info = "click here to create a new folder."
    End Select
    Label1.Caption = Info
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    QwickInfo 0
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    QwickInfo 0
End Sub

Private Sub framePassword_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    QwickInfo 0
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    QwickInfo 0
End Sub

Private Sub optFiles_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case Index
        Case 0: QwickInfo 2
        Case 1: QwickInfo 3
    End Select
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    QwickInfo 0
End Sub

Private Sub txtPassword_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    QwickInfo 4
End Sub

Private Sub txtPath_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    QwickInfo 8
End Sub
