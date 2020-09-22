VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Multi Unzip"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   9030
      TabIndex        =   7
      Top             =   810
      Width           =   9030
      Begin VB.TextBox Text1 
         Height          =   285
         HideSelection   =   0   'False
         Left            =   720
         TabIndex        =   11
         ToolTipText     =   "Type the address here!"
         Top             =   60
         Width           =   7575
      End
      Begin VB.ComboBox Text1Drop 
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   1200
         Style           =   1  'Simple Combo
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   0
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.CommandButton Command1 
         Caption         =   "GO"
         Height          =   285
         Left            =   8400
         TabIndex        =   9
         ToolTipText     =   "Click here to goto the address!"
         Top             =   60
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A&ddress"
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   100
         Width           =   855
      End
   End
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   6120
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComctlLib.ImageList imglst 
      Left            =   6840
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   30
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0442
            Key             =   "win"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0556
            Key             =   "exe"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":066A
            Key             =   "hlp"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0782
            Key             =   "txt"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0896
            Key             =   "bmp"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":09AE
            Key             =   "frm"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0E02
            Key             =   "nfo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1256
            Key             =   "dll"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1572
            Key             =   "ttf"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":188E
            Key             =   "chm"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1BAA
            Key             =   "htm"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1EC6
            Key             =   "wav"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":21E2
            Key             =   "mdb"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":24FE
            Key             =   "log"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":281A
            Key             =   "bat"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2B36
            Key             =   "reg"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2E52
            Key             =   "vbp"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":316E
            Key             =   "bas"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":348A
            Key             =   "cls"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":37A6
            Key             =   "ctl"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3AC2
            Key             =   "rar"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3DDE
            Key             =   "zip"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":40FA
            Key             =   "doc"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4416
            Key             =   "xls"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4732
            Key             =   "ini"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4A4E
            Key             =   "pdf"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4D6A
            Key             =   "psd"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5086
            Key             =   "gif"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":53A2
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":57F6
            Key             =   "zfile"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   4560
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox picSep 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   8640
      ScaleHeight     =   5895
      ScaleWidth      =   135
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   135
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   2295
      Left            =   4440
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imglst"
      SmallIcons      =   "imglst"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   7832
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Size (Bytes)"
         Object.Width           =   2540
      EndProperty
   End
   Begin Project1.DirTree Dir1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   7646
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   1429
      ButtonWidth     =   1693
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgtoolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   2100
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "    Extract    "
            Key             =   "Extract"
            Description     =   "Extract"
            Object.ToolTipText     =   "Click here to extract the file(s)"
            ImageIndex      =   1
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList imgtoolbar 
         Left            =   5520
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":5C4A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":5F66
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   2055
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Image imgSep 
      Height          =   5895
      Left            =   3840
      MouseIcon       =   "Form1.frx":6282
      MousePointer    =   99  'Custom
      Top             =   1080
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   -120
      Picture         =   "Form1.frx":63D4
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   2265
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cX As Single
Private Const MaxLeft = 1000
Private Const MaxRight = 1000
Private CurItem As ListItem

'Layout
Private mLeft As Single
Private mTop As Single
Private mWidth As Single
Private mHeight As Single

Public ZipCount As Long
Private UserChange As Boolean

Public RealLen As Byte
Public AutoInput As Boolean
Dim LoadingPath As Boolean


Private Sub Command1_Click()
    Dim ss As Integer
    If Dir(Text1.Text, vbDirectory) <> "" Then
        LoadingPath = True
        ss = Text1.SelStart
        Dir1.Path = Text1.Text
        LoadingPath = False
        If Len(Text1.Text) >= ss Then Text1.SelStart = ss
        Text1_Change
    End If
End Sub

Private Sub Dir1_Click()
On Error GoTo CheckError
    Dim DirReturn As String, Pth As String, i As Integer, pic As String, LI As ListItem
    LV1.ListItems.Clear
    List1.Clear
    List2.Clear
    ZipCount = 0
    Pth = Dir1.Path
    If Right(Pth, 1) <> "\" Then Pth = Pth & "\"
    DirReturn = Dir(Pth & "*.*", vbDirectory + vbNormal)
    Do Until DirReturn = ""
        If Not DirReturn = "." And Not DirReturn = ".." Then
            If (GetAttr(Pth & DirReturn) And vbDirectory) = vbDirectory Then
                List1.AddItem DirReturn
            ElseIf (GetAttr(Pth & DirReturn) And vbNormal) = vbNormal Then ' And LCase(Right(DirReturn, 4)) = ".zip" Then
                List2.AddItem DirReturn
            End If
        End If
        DirReturn = Dir
    Loop
    For i = 0 To List1.ListCount - 1
        Set LI = LV1.ListItems.Add(, Pth & List1.List(i), List1.List(i), "folder", "folder")
        LI.SubItems(1) = " "
        LI.Tag = "folder"
    Next i
    For i = 0 To List2.ListCount - 1
        pic = GetPicName(List2.List(i))
        If LCase(Right(List2.List(i), 4)) = ".zip" Then
            Set LI = LV1.ListItems.Add(, Pth & List2.List(i), List2.List(i), pic, pic)
        End If
        LI.SubItems(1) = FileLen(Pth & List2.List(i))
'        If LCase(Right(LI, 4)) <> ".zip" Then
'            LI.Ghosted = True
'            LI.ForeColor = &H404000   '&H808080
'        Else
            ZipCount = ZipCount + 1
            LI.Tag = "zipfile"
'        End If
    Next i
CheckError:
    If Err.Number = 52 Then
        Dim TmpMsg As VbMsgBoxResult
        TmpMsg = MsgBox("This device is unavailable at the moment. Would you like to retry?", vbRetryCancel, "Device Unavailable")
        If TmpMsg = vbRetry Then
            Dir1_Click
        End If
    End If
End Sub

Private Sub Dir1_PathChange(NewPath As String)
    UserChange = False
    Text1.Text = NewPath
    UserChange = True
End Sub

Private Sub Form_Load()
    Dim SysDir As String, SysFile As String
    SysDir = Module1.GetSysDir()
    SysFile = "\unzip32.dll"
    If Dir(SysDir & SysFile) = "" Then ExtractLibs 101, SysDir & SysFile
    SysFile = "\zip32.dll"
    If Dir(SysDir & SysFile) = "" Then ExtractLibs 102, SysDir & SysFile
    
    Dim cmd As String
    Picture1.PaintPicture Image2.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
    Me.Width = GetSetting("MultiUnzip", "Layout", "Width", 8000)
    Me.Height = GetSetting("MultiUnzip", "Layout", "Height", 7000)
    Me.Left = GetSetting("MultiUnzip", "Layout", "Left", 1000)
    Me.Top = GetSetting("MultiUnzip", "Layout", "Top", 1000)
    mLeft = Left
    mTop = Top
    mWidth = Width
    mHeight = Height
    Me.WindowState = GetSetting("MultiUnzip", "Layout", "WinState", 0)
    imgSep.Left = GetSetting("MultiUnzip", "Layout", "Sep", 3800)
    Form_Resize
    Dir1.GoList
    cmd = Command$
'    cmd = "H:\Darrens work#\VB\Downloaded\Game\wacarat.zip"
    Show
    DoEvents
    If Trim(cmd) = "" Then
        LoadingPath = True
        Dir1.Path = GetSetting("MultiUnzip", "Layout", "Dir", "C:\My Documents\")
        LoadingPath = False
        Text1_Change
    Else
        If FileExists(cmd) And LCase(Right(cmd, 4)) = ".zip" Then
            LoadMyFile cmd
        Else
            LoadingPath = True
            Dir1.Path = GetSetting("MultiUnzip", "Layout", "Dir", "C:\My Documents\")
            LoadingPath = False
            Text1_Change
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        imgSep.Top = Picture2.Top + Picture2.Height
        imgSep.Height = ScaleHeight - Picture2.Top + Picture2.Height
        imgSep.Width = 60
        Dir1.Move 0, Picture2.Top + Picture2.Height, imgSep.Left, ScaleHeight - (Picture2.Top + Picture2.Height)
        LV1.Move imgSep.Left + imgSep.Width, Picture2.Top + Picture2.Height
        LV1.Width = ScaleWidth - LV1.Left
        LV1.Height = ScaleHeight - (Picture2.Top + Picture2.Height)
        If Me.WindowState <> vbMaximized Then
            mLeft = Left
            mTop = Top
            mWidth = Width
            mHeight = Height
        End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "MultiUnzip", "Layout", "Width", mWidth
    SaveSetting "MultiUnzip", "Layout", "Height", mHeight
    SaveSetting "MultiUnzip", "Layout", "Left", mLeft
    SaveSetting "MultiUnzip", "Layout", "Top", mTop
    SaveSetting "MultiUnzip", "Layout", "WinState", Me.WindowState
    SaveSetting "MultiUnzip", "Layout", "Sep", imgSep.Left
    SaveSetting "MultiUnzip", "Layout", "Dir", Dir1.Path
    
    Set CurItem = Nothing
End Sub

Private Sub imgSep_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        cX = x
        picSep.Move imgSep.Left, imgSep.Top, imgSep.Width, imgSep.Height
        picSep.Visible = True
    End If
End Sub

Private Sub imgSep_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Dim TmpX As Single
        TmpX = imgSep.Left + x - cX
        If TmpX > MaxLeft And TmpX < (ScaleWidth - MaxRight) Then
            picSep.Move TmpX
        ElseIf TmpX < MaxLeft Then
            TmpX = MaxLeft + cX
            picSep.Left = TmpX
        ElseIf TmpX > (ScaleWidth - MaxRight) Then
            TmpX = (ScaleWidth - MaxRight) + cX
            picSep.Left = TmpX
        End If
    End If
End Sub

Private Sub imgSep_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        picSep.Visible = False
        imgSep.Left = picSep.Left
        Form_Resize
    End If
End Sub

Private Sub LV1_DblClick()
On Error GoTo XS
    If LV1.SelectedItem = CurItem Then
        If CurItem.Tag = "folder" Then
            Dir1.Path = CurItem.Key & "\"
        ElseIf LCase(Right(CurItem.Text, 4)) = ".zip" Then
            UsePath = Dir1.Path
            OneFileOnly = True
            Load frmExtract
            'frmExtract.FilesPath = Dir1.Path
            'frmExtract.OneFileOnly = True
            'frmExtract.Dir1.Path = Dir1.Path
            'frmExtract.txtPath = frmExtract.txtPath & Left(LV1.SelectedItem.Text, Len(LV1.SelectedItem.Text) - 4)
'            frmExtract.optFiles(1).Value = True
'            frmExtract.TestZip
'            frmExtract.Show 1, Me
        End If
    End If
XS:
End Sub

Private Sub LV1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set CurItem = Item
End Sub

Private Sub LV1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then Dir1.Refresh
End Sub

Private Sub LV1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call LV1_DblClick
    End If
End Sub

Private Sub Picture2_Resize()
Text1.Width = Picture2.Width - (Text1.Left + 100 + Command1.Width)
Command1.Left = Text1.Left + Text1.Width + 60
End Sub

Private Sub Text1_Change()
On Local Error GoTo XS
    If Right(Text1.Text, 1) = "\" And Not LoadingPath Then
        Text1Drop.Clear
        If Dir(Text1.Text, vbDirectory) <> "" Then
            Dim DirReturn As String, Path As String
            Path = Text1.Text
            DirReturn = Dir(Path & "*.*", vbNormal + vbDirectory + vbSystem + vbHidden + vbReadOnly)
            Do Until DirReturn = ""
                DoEvents
                If DirReturn <> "." And DirReturn <> ".." Then
                    Text1Drop.AddItem Path & DirReturn
                End If
                DirReturn = Dir()
            Loop
        End If
    ElseIf Not LoadingPath Then
        AutoCompleteChange Text1, Text1Drop
    End If
XS:
End Sub

Private Sub Text1_GotFocus()
    If Not Text1.Text = "" Then
        'Text1.SelStart = 0
        Text1.SelStart = Len(Text1.Text)
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode >= 65 And KeyCode <= 90) Or (KeyCode >= 48 And KeyCode <= 57) Or (KeyCode >= 96 And KeyCode <= 105) Then
        'It's a letter or number
'        AutoCompleteKeyPress Text1, KeyCode
    End If
    AutoCompleteKeyPress Text1, KeyCode
End Sub

Sub AutoCompleteChange(UserBox As Object, Comb As ComboBox)
    Dim i As Integer
    If UserBox.Text <> "" And AutoInput = False Then
        RealLen = Len(UserBox.Text)
        Do
            If LCase(UserBox.Text) = LCase(Comb.List(i)) Then
                Exit Sub
            ElseIf LCase(UserBox.Text) = LCase(Left(Comb.List(i), RealLen)) Then
                AutoInput = True
                UserBox.Text = Comb.List(i)
                UserBox.SelStart = RealLen
                UserBox.SelLength = Len(UserBox.Text) - RealLen
            End If
            i = i + 1
        Loop Until i = Comb.ListCount
    Else
        AutoInput = False
    End If
End Sub

Sub AutoCompleteKeyPress(UserBox As Object, KeyCode As Integer)
    If KeyCode = 8 Then
        If RealLen > 0 And UserBox.SelLength > 0 Then
            UserBox.SelStart = RealLen - 1
            UserBox.SelLength = Len(UserBox.Text) - RealLen + 1
        End If
    ElseIf KeyCode = 46 Then
        If UserBox.SelLength <> 0 Then
            UserBox.Text = Left(UserBox.Text, RealLen)
            AutoInput = True
        End If
    ElseIf KeyCode = 13 Then
        Command1_Click
    End If
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
'    AutoCompleteKeyPress Text1, KeyAscii
End Sub

'Private Sub Text1_Change()
'    If UserChange Then
'        CheckAddress
'        AutoCompleteChange Text1
'    End If
'End Sub
'
'Private Sub Text1_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyDelete Then UserChange = False: Exit Sub
'    UserChange = True
'    AutoCompleteKeyPress Text1, KeyAscii
'    If KeyAscii = 32 And Text1.SelText <> "" Then
'        Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
'        Text1.SelStart = Len(Text1.Text)
'        KeyAscii = 0
'    End If
'End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Extract"
            If ZipCount = 0 Then Exit Sub
            Module1.UsePath = Dir1.Path
            Module1.OneFileOnly = False
            On Error GoTo XS
            Load frmExtract
'            frmExtract.FilesPath = Dir1.Path
'            frmExtract.OneFileOnly = False
'            frmExtract.Dir1.Path = Dir1.Path
'            frmExtract.TestZip
'            frmExtract.Show 1, Me
    End Select
XS:
End Sub

'Sub AutoCompleteChange(Comb As ComboBox)
'    Dim i As Integer
'    If Comb.Text <> "" And AutoInput = False Then
'        RealLen = Len(Comb.Text)
'        i = 0
'        Do Until i = Comb.ListCount
'            If LCase(Comb.Text) = LCase(Comb.List(i)) Then
'                Exit Sub
'            ElseIf LCase(Comb.Text) = LCase(Left(Comb.List(i), RealLen)) Then
'                AutoInput = True
'                Comb.Text = Comb.List(i)
'                Comb.SelStart = RealLen
'                Comb.SelLength = Len(Comb.Text) - RealLen
'            End If
'            i = i + 1
'        Loop
'    Else
'        AutoInput = False
'    End If
'End Sub
'
'Sub AutoCompleteKeyPress(Comb As ComboBox, KeyCode As Integer)
'    If KeyCode = 8 Then
'        If RealLen > 0 And Comb.SelLength > 0 Then
'            Comb.SelStart = RealLen - 1
'            Comb.SelLength = Len(Comb.Text) - RealLen + 1
'        End If
'    ElseIf KeyCode = 46 Then
'        If Comb.SelLength <> 0 Then
'            Comb.Text = Left(Comb.Text, RealLen)
'            AutoInput = True
'        End If
'    End If
'End Sub

'Private Sub CheckAddress()
'    Dim stg As String, fle As String
'    UserChange = False
'    On Error GoTo XS
'    stg = GetPathFromFileName(Text1.Text)
'    If Dir(stg & "*.*", vbDirectory) <> "" Then
'        For i = Text1.ListCount - 1 To 0 Step -1
'            Text1.RemoveItem i
'        Next i
'        fle = Dir(stg & "*.*", vbDirectory)
'        Do Until fle = ""
'            If fle <> "." And fle <> ".." Then
'                If (GetAttr(stg & fle) And vbDirectory) = vbDirectory Then
'                    If Right(stg, 1) <> "\" Then
'                        Text1.AddItem "\" & stg & fle
'                    Else
'                        Text1.AddItem stg & fle
'                    End If
'                End If
'            End If
'            fle = Dir
'        Loop
'    Else
'        Text1.Clear
'    End If
'XS:
'    UserChange = True
'End Sub

Private Function GetPathFromFileName(FileName As String) As String
    If Right(FileName, 1) = "\" Then GetPathFromFileName = FileName: Exit Function
    If Trim(FileName) = "" Then Exit Function
    If InStr(1, FileName, "\") = 0 Then GetPathFromFileName = FileName: Exit Function
    Dim Posa As Long
    Posa = InStrRev(FileName, "\")
    GetPathFromFileName = Left(FileName, Posa)
End Function

Private Function GetFileFromPath(ByVal Pth As String) As String
    Dim Posa As Integer
    Posa = InStrRev(Pth, "\")
    GetFileFromPath = Mid(Pth, Posa + 1)
End Function

Private Function FileExists(FileName As String) As Boolean
    If Dir(FileName) <> "" Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

Private Sub LoadMyFile(ByVal FileName As String)
    Dim Stg As String, i As Long, fle As String
    LoadingPath = True
    Stg = GetPathFromFileName(FileName)
    Dir1.Path = Stg
    fle = GetFileFromPath(FileName)
    For i = 1 To LV1.ListItems.Count
        If LCase(LV1.ListItems(i).Text) = LCase(fle) Then
            LV1.ListItems(i).Selected = True
            LV1.ListItems(i).EnsureVisible
            Set CurItem = LV1.ListItems(i)
            Call LV1_DblClick
            Exit For
        End If
    Next i
    LoadingPath = False
End Sub
