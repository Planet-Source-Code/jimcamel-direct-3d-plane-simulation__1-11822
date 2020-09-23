VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Pick your plane"
   ClientHeight    =   2496
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3744
   LinkTopic       =   "Form2"
   ScaleHeight     =   2496
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Select Plane"
      Height          =   372
      Left            =   1800
      TabIndex        =   3
      Top             =   2040
      Width           =   1932
   End
   Begin VB.FileListBox File1 
      Height          =   1992
      Left            =   1800
      Pattern         =   "*.x"
      TabIndex        =   2
      Top             =   0
      Width           =   1932
   End
   Begin VB.DirListBox Dir1 
      Height          =   2016
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1692
   End
   Begin VB.DriveListBox Drive1 
      Height          =   288
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1692
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
planename = File1.Path & "\" & File1.FileName
Load (Form1)
Unload (Form2)
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
End Sub
