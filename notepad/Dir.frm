VERSION 5.00
Begin VB.Form Dir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dir"
   ClientHeight    =   4455
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3120
      TabIndex        =   5
      Top             =   1680
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SELECT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   3480
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   765
      Left            =   3120
      TabIndex        =   3
      Top             =   2160
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Select Location"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Location Of File"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "Dir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
notepad.filelocation = Text1.Text
notepad.fonts
notepad.Enabled = True
notepad.Visible = True
End Sub

Private Sub Dir1_Change()
Text1.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
If Drive1.Drive <> "" Then
Dir1.Path = Drive1.Drive
End If
End Sub



Private Sub Text1_Click()
If Text1.Text <> "" Then
Dir1.Path = Text1.Text
Drive1.Drive = Dir1.Path
End If
End Sub
