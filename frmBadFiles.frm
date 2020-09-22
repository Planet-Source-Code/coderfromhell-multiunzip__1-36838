VERSION 5.00
Begin VB.Form frmBadFiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bad Archives"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "frmBadFiles.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2168
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   840
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   $"frmBadFiles.frx":0442
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmBadFiles.frx":04DF
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmBadFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
