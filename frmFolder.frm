VERSION 5.00
Begin VB.Form frmFolder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add folder"
   ClientHeight    =   3105
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   4575
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4575
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Select a folder to wipe (all its contents will be permanently lost)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub OKButton_Click()
Dim i As Long
Dim FilePath As String
FilePath = File1.Path
If Right(FilePath, 1) <> "\" Then
    FilePath = FilePath & "\"
End If
For i = 0 To (File1.ListCount - 1)
    File1.ListIndex = i
    frmMain.lstFiles.AddItem FilePath & File1.FileName
Next
Unload Me
End Sub
