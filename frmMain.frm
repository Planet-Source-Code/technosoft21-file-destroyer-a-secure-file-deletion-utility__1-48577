VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "File Destroyer"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   3360
      Width           =   1335
   End
   Begin VB.ComboBox cboStrength 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   1560
      List            =   "frmMain.frx":0019
      TabIndex        =   8
      Text            =   "8"
      Top             =   2880
      Width           =   735
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   2760
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdShred 
      Caption         =   "Destroy them!"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   5895
   End
   Begin VB.CommandButton cmdAddFolder 
      Caption         =   "Add folder"
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdRemoveFile 
      Caption         =   "Remove file"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "Add file"
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.ListBox lstFiles 
      Height          =   1815
      ItemData        =   "frmMain.frx":0036
      Left            =   120
      List            =   "frmMain.frx":0038
      TabIndex        =   2
      Top             =   960
      Width           =   5895
   End
   Begin VB.Label lblBlue 
      BackColor       =   &H80000002&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Left            =   120
      Top             =   3840
      Width           =   7335
   End
   Begin VB.Label lblGrey 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   7335
   End
   Begin VB.Label lblStrength 
      Caption         =   "Normal"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Shredder strength: "
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Files to shred:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "FILE DESTROYER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CancelPressed As Boolean
Public Progress As Integer

Private Sub cboStrength_Change()
Dim Strength As Integer
Strength = CInt(cboStrength.Text)

If Strength < 4 Then
    lblStrength.Caption = "Very quick"
ElseIf Strength < 6 Then
    lblStrength.Caption = "Quick"
ElseIf Strength < 9 Then
    lblStrength.Caption = "Normal"
ElseIf Strength < 12 Then
    lblStrength.Caption = "Strong"
ElseIf Strength < 19 Then
    lblStrength.Caption = "Paranoid"
ElseIf Strength < 28 Then
    lblStrength.Caption = "Very paranoid"
Else
    lblStrength.Caption = "Maximum destruction"
End If
End Sub

Private Sub cboStrength_Click()
cboStrength_Change
End Sub

Private Sub cmdAddFile_Click()
On Error GoTo Quit
With dlgFile
    .DialogTitle = "Select file to destroy"
    .Filter = "All files (*.*)|*.*"
    .CancelError = True
    .ShowOpen
    lstFiles.AddItem .FileName
End With
Quit:
End Sub

Private Sub cmdAddFolder_Click()
frmFolder.Show vbModal, Me
End Sub

Private Sub cmdCancel_Click()
CancelPressed = True
End Sub

Private Sub cmdRemoveFile_Click()
On Error Resume Next
Dim OldIndex As Integer
OldIndex = lstFiles.ListIndex - 1
If OldIndex < 0 Then OldIndex = 0
lstFiles.RemoveItem lstFiles.ListIndex
lstFiles.SetFocus
lstFiles.ListIndex = OldIndex
End Sub

Private Sub cmdShred_Click()
Dim Confirm As VbMsgBoxResult
Dim i As Integer
Confirm = MsgBox("Are you absolutely sure you want to permanently delete all these files?", vbExclamation + vbYesNo)
If Confirm <> vbYes Then
    Exit Sub
Else
    cmdCancel.Enabled = True
    cmdShred.Enabled = False
    For i = 0 To (lstFiles.ListCount - 1)
        Progress = 0
        ShredFile lstFiles.List(0), CInt(cboStrength.Text)
        lstFiles.RemoveItem 0
        DoEvents
        If CancelPressed = True Then
            lblBlue.Visible = False
            cmdCancel.Enabled = False
            cmdShred.Enabled = True
            DoEvents
            Exit Sub
        End If
    Next
    lstFiles.Clear
    MsgBox "The files have been destroyed permanently!", vbInformation + vbOKOnly
    lblBlue.Visible = False
    cmdCancel.Enabled = False
    cmdShred.Enabled = True
End If
End Sub

Private Sub Form_Load()
CancelPressed = False
Progress = 0
End Sub

Sub UpdateProgress()
If Progress > 100 Then Progress = 100
lblBlue.Visible = True
lblBlue.Width = lblGrey.Width * (Progress / 100)
End Sub
