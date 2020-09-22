VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl DirTree 
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   ScaleHeight     =   5220
   ScaleWidth      =   6495
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   3480
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.ImageList img 
      Left            =   3360
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DirTree.ctx":0000
            Key             =   "unknown"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DirTree.ctx":0944
            Key             =   "fixed"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DirTree.ctx":30F8
            Key             =   "ram"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DirTree.ctx":58AC
            Key             =   "remove"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DirTree.ctx":8060
            Key             =   "cd"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DirTree.ctx":A814
            Key             =   "folder1"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DirTree.ctx":CFC8
            Key             =   "open1"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DirTree.ctx":F77C
            Key             =   "remote"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DirTree.ctx":F8D8
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DirTree.ctx":F9EC
            Key             =   "open"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView DirTree 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   3836
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img"
      Appearance      =   1
   End
End
Attribute VB_Name = "DirTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private nNode As Node
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Public Event Click()
Public Event PathChange(NewPath As String)
Public Event MouseMove(x As Single, y As Single)
Private m_Path As String
Private Const DefaultPath = "C:\"
Private mDefaultSelectedBold As Boolean
Dim EndPath As String

Private Sub DisplayDir(ByVal Pth, ByVal Parent)
Dim j As Integer
    On Error Resume Next
    If Right(Pth, 1) <> "\" Then Pth = Pth & "\"
    tmp = Dir(Pth, vbDirectory)
    Do Until tmp = ""
        If tmp <> "." And tmp <> ".." Then
            If GetAttr(Pth & tmp) And vbDirectory Then
                'I use ListBox with property Sorted=True to
                'alphabetize directories. Easy eh? ;-)
                List1.AddItem StrConv(tmp, vbProperCase)
                'StrConv function convert for example
                '"WINDOWS" to "Windows"
            End If
        End If
        tmp = Dir
    Loop
    'Add sorted directory names to TreeView
    For j = 1 To List1.ListCount
        Set nNode = DirTree.Nodes.Add(Parent, tvwChild, Pth & List1.List(j - 1), List1.List(j - 1), "folder", "open")
    Next j
    List1.Clear
End Sub

Private Sub GoToPath(ByVal Path As String)
On Error GoTo XS
    Dim Paths() As String, i As Integer, MyRoot As String, CurPath As String, DI As Node, TmpStg As String
    If Left(Path, 1) <> "\" Then Path = "\" & Path
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    Paths = Split(Path, "\")
    For i = 0 To UBound(Paths)
        If Trim(Paths(i)) <> "" Then
            If Right(Paths(i), 1) <> "\" Then Paths(i) = Paths(i) & "\"
            If MyRoot = "" Then MyRoot = Paths(i): Exit For
        End If
    Next i
    If MyRoot = "" Then GoTo XS
    For i = 0 To UBound(Paths)
        If Trim(Paths(i)) <> "" Then
            CurPath = CurPath & StrConv(Paths(i), vbProperCase)
            Set DI = DirTree.Nodes(CurPath)
            If i = UBound(Paths) - 1 Then
                NodeClicked DI
                DI.Selected = True
                DI.Expanded = True
                If mDefaultSelectedBold Then DI.Bold = True
                GoTo XS
            End If
            DI.Expanded = True
            If Right(CurPath, 1) <> "\" Then CurPath = CurPath & "\"
        End If
    Next i
XS:
End Sub

Private Sub LoadTreeView()
    DirTree.Nodes.Clear
    Dim DriveNum As String
    Dim DriveType As Long
    DriveNum = 64
    On Error Resume Next
    Do
        DriveNum = DriveNum + 1
        DriveType = GetDriveType(Chr$(DriveNum) & ":\")
        If DriveNum > 90 Then Exit Do
        Select Case DriveType
'            Case 0: Set nNode = DirTree.Nodes.Add(, , "root" & DriveNum, StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "unknown")
'                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
'            Case 2: Set nNode = DirTree.Nodes.Add(, , "root" & DriveNum, "(" & Chr$(DriveNum) & ":)", "remove")
'            Case 3: Set nNode = DirTree.Nodes.Add(, , "root" & DriveNum, StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "fixed")
'                    If DriveNum = 67 Then nNode.Expanded = True: nNode.Selected = True
'                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
'            Case 4: Set nNode = DirTree.Nodes.Add(, , "root" & DriveNum, StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "remote")
'                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
'            Case 5: Set nNode = DirTree.Nodes.Add(, , "root" & DriveNum, StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "cd")
'                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
'            Case 6: Set nNode = DirTree.Nodes.Add(, , "root" & DriveNum, StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "ram")
'                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
            Case 0: Set nNode = DirTree.Nodes.Add(, , Chr(DriveNum) & ":\", StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "unknown")
                    DisplayDir Mid(DirTree.Nodes(Chr(DriveNum) & ":\").Text, Len(DirTree.Nodes(Chr(DriveNum) & ":\").Text) - 2, 2), Chr(DriveNum) & ":\"
            Case 2: Set nNode = DirTree.Nodes.Add(, , Chr(DriveNum) & ":\", "(" & Chr$(DriveNum) & ":)", "remove")
            Case 3: Set nNode = DirTree.Nodes.Add(, , Chr(DriveNum) & ":\", StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "fixed")
'                    If DriveNum = 67 Then nNode.Expanded = True: nNode.Selected = True
                    DisplayDir Mid(DirTree.Nodes(Chr(DriveNum) & ":\").Text, Len(DirTree.Nodes(Chr(DriveNum) & ":\").Text) - 2, 2), Chr(DriveNum) & ":\"
            Case 4: Set nNode = DirTree.Nodes.Add(, , Chr(DriveNum) & ":\", StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "remote")
                    DisplayDir Mid(DirTree.Nodes(Chr(DriveNum) & ":\").Text, Len(DirTree.Nodes(Chr(DriveNum) & ":\").Text) - 2, 2), Chr(DriveNum) & ":\"
            Case 5: Set nNode = DirTree.Nodes.Add(, , Chr(DriveNum) & ":\", StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "cd")
                    DisplayDir Mid(DirTree.Nodes(Chr(DriveNum) & ":\").Text, Len(DirTree.Nodes(Chr(DriveNum) & ":\").Text) - 2, 2), Chr(DriveNum) & ":\"
            Case 6: Set nNode = DirTree.Nodes.Add(, , Chr(DriveNum) & ":\", StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "ram")
                    DisplayDir Mid(DirTree.Nodes(Chr(DriveNum) & ":\").Text, Len(DirTree.Nodes(Chr(DriveNum) & ":\").Text) - 2, 2), Chr(DriveNum) & ":\"
        End Select
    Loop
End Sub


Private Sub DirTree_Expand(ByVal Node As MSComctlLib.Node)
Dim j As Integer
    DirTree.Refresh
    LockWindowUpdate DirTree.hwnd
    For j = Node.Child.FirstSibling.Index To Node.Child.LastSibling.Index
        NodeClicked DirTree.Nodes(j)
    Next j
    NodeClicked Node
    Node.Selected = True
    DirTree.Refresh
    LockWindowUpdate 0
End Sub


Private Sub DirTree_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then Me.Refresh
End Sub


Private Sub DirTree_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(x, y)
End Sub

Private Sub DirTree_NodeClick(ByVal Node As MSComctlLib.Node)
    EndPath = ExtractPath(Node.FullPath)
    NodeClicked Node
End Sub

Private Function ExtractPath(ByVal Pth As String) As String
    Dim Stg As String, lft As String, rht As String
    Stg = Mid(Pth, InStr(1, Pth, "(") + 1)
    lft = Left(Stg, InStr(1, Stg, ")") - 1)
    rht = Mid(Stg, InStr(1, Stg, ")") + 1)
    Stg = lft & rht
    If Right(Stg, 1) <> "\" Then Stg = Stg & "\"
    ExtractPath = Stg
End Function

Private Sub NodeClicked(Node As MSComctlLib.Node)
    Dim Path As String
    If Left(Node.Key, 4) = "root" Then
        On Error Resume Next
        If Node.Children > 0 Then GoTo Skok
        DisplayDir Mid(Node.Text, Len(Node.Text) - 2, 2), Node.Key
    End If
    Path = Mid(Node.FullPath, InStr(1, Node.FullPath, ":") - 1, 2) & Mid(Node.FullPath, InStr(1, Node.FullPath, ":") + 2)
    If Node.Children > 0 Then GoTo Skok
    DisplayDir Path, Node.Index

Skok:
    Path = Mid(Node.FullPath, InStr(1, Node.FullPath, ":") - 1, 2) & Mid(Node.FullPath, InStr(1, Node.FullPath, ":") + 2)
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    m_Path = Path

If LCase(Path) = LCase(EndPath) Then RaiseEvent Click
RaiseEvent PathChange(m_Path)
End Sub

Private Sub UserControl_InitProperties()
m_Path = "C:\"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mDefaultSelectedBold = PropBag.ReadProperty("DefaultSelectedBold", False)
End Sub

Private Sub UserControl_Resize()
DirTree.Width = UserControl.Width
DirTree.Height = UserControl.Height
End Sub

Public Sub GoList()
LoadTreeView
End Sub

Public Property Get Path() As String
    Path = m_Path
End Property

Public Property Let Path(New_Path As String)
    EndPath = New_Path
    If Right(EndPath, 1) <> "\" Then EndPath = EndPath & "\"
    GoToPath New_Path
End Property

Public Sub GoUpALevel()
    Dim Temp As String
    Temp = RemoveRightSection(m_Path)
    Me.Path = Temp
End Sub

Private Function RemoveRightSection(txt As String) As String
Dim strga As String
If txt = "" Then Exit Function
If InStr(1, txt, "\") = 0 Then Exit Function
strga = txt
If Right(strga, 1) = "\" Then strga = Left(strga, Len(strga) - 1)
Do Until Right(strga, 1) = "\"
    strga = Left(strga, Len(strga) - 1)
Loop
RemoveRightSection = strga
End Function

Public Property Get DefaultSelectedBold() As Boolean
    DefaultSelectedBold = mDefaultSelectedBold
End Property

Public Property Let DefaultSelectedBold(New_Val As Boolean)
    mDefaultSelectedBold = New_Val
    PropertyChanged "DefaultSelectedBold"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DefaultSelectedBold", mDefaultSelectedBold, False)
End Sub

Public Sub Refresh()
     DirTree.Nodes.Clear
     LoadTreeView
     GoToPath EndPath
End Sub
