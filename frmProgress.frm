VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extracting..."
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   ControlBox      =   0   'False
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   1748
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00EC7E44&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00EC7E44&
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   4155
      TabIndex        =   3
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label lblFile 
      Caption         =   "?"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   960
      UseMnemonic     =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "File:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Multi Unzip is extracting the selected files."
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   200
      Picture         =   "frmProgress.frx":030A
      Top             =   200
      Width           =   480
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mFileCount As Single
Private mProgJump As Single
Private ProgWid As Single

Public Property Get FileCount() As Single
    FileCount = mFileCount
End Property

Public Property Let FileCount(New_Val As Single)
    mFileCount = New_Val
    mProgJump = picProgress.ScaleWidth / New_Val
    ProgWid = 0
End Property

'----------

Private Sub cmdCancel_Click()
    ExitExtract = True
    Unload Me
End Sub

Public Property Let FileName(ByVal NewFile As String)
    If Len(NewFile) > 35 Then
        NewFile = Left(NewFile, 32) & "..."
    End If
    lblFile.Caption = NewFile
    lblFile.Refresh
End Property

Public Sub UpdateProgress()
    ProgWid = ProgWid + mProgJump
    picProgress.Line (0, 0)-(ProgWid, picProgress.ScaleHeight), vbHighlight, BF
    picProgress.Refresh
End Sub

