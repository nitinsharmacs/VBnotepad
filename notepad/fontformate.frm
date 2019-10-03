VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Font"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10530
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdColor 
      Left            =   5760
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "color"
      Height          =   975
      Left            =   600
      TabIndex        =   16
      Top             =   7200
      Width           =   4215
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4200
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image custom 
         BorderStyle     =   1  'Fixed Single
         Height          =   480
         Left            =   3120
         Picture         =   "fontformate.frx":0000
         Stretch         =   -1  'True
         ToolTipText     =   "Custom"
         Top             =   360
         Width           =   585
      End
      Begin VB.Label greentextlabel 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "News706 BT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   2280
         TabIndex        =   19
         Top             =   360
         Width           =   495
      End
      Begin VB.Label redtextlabel 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "News706 BT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1560
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.Label bluetextlabel 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "News706 BT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   840
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton ok 
      BackColor       =   &H0080C0FF&
      Caption         =   "OK"
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cancel 
      BackColor       =   &H008080FF&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "effects"
      Height          =   1935
      Left            =   600
      TabIndex        =   10
      Top             =   4920
      Width           =   4215
      Begin VB.CheckBox Check2 
         Caption         =   "underline"
         Height          =   300
         Left            =   480
         TabIndex        =   12
         Top             =   1320
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Caption         =   "strikeout"
         Height          =   420
         Left            =   480
         TabIndex        =   11
         Top             =   600
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "sample"
      Height          =   2415
      Left            =   5520
      TabIndex        =   9
      Top             =   4680
      Width           =   4335
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "A B C"
         Height          =   1335
         Left            =   960
         TabIndex        =   13
         Top             =   1200
         Width           =   2535
      End
   End
   Begin VB.TextBox fontsizeshow 
      Height          =   420
      Left            =   8760
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox fontstyleshow 
      Height          =   420
      Left            =   4920
      TabIndex        =   7
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox fontshow 
      Height          =   420
      Left            =   600
      TabIndex        =   6
      Top             =   1080
      Width           =   3375
   End
   Begin VB.ListBox sizelist 
      Height          =   2460
      Left            =   8760
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.ListBox fontstylelist 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   4920
      TabIndex        =   2
      Top             =   1560
      Width           =   2895
   End
   Begin VB.ListBox fontlist 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Size"
      Height          =   375
      Left            =   8760
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Font style"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Font"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bluetextlabel_Click()
Label4.ForeColor = bluetextlabel.ForeColor
End Sub

Private Sub cancel_Click()
Form1.Visible = False
notepad.Enabled = True
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
Label4.FontStrikethru = True
Else
Label4.FontStrikethru = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Label4.FontUnderline = True
Else
Label4.FontUnderline = False
End If
End Sub



Private Sub custom_Click()
cdColor.ShowColor

Label4.ForeColor = cdColor.Color

End Sub

Private Sub fontlist_Click()
fontshow.Text = fontlist.Text
Label4.Font = fontlist.Text
End Sub

Private Sub fontshow_Change()
If fontshow.Text <> "" Then
Label4.Font = fontshow.Text
End If
End Sub

Private Sub fontsizeshow_Change()
If fontsizeshow.Text > 0 Then
Label4.FontSize = fontsizeshow.Text
End If
End Sub



Private Sub fontstylelist_Click()
fontstyleshow.Text = fontstylelist.Text
If fontstyleshow.Text = "regular" Then
Label4.FontBold = False
Label4.FontItalic = False
ElseIf fontstyleshow.Text = "bold" Then
Label4.FontBold = True
Label4.FontItalic = False
ElseIf fontstyleshow.Text = "italic" Then
Label4.FontItalic = True
Label4.FontBold = False
ElseIf fontstyleshow.Text = "bold italic" Then
Label4.FontItalic = True
Label4.FontBold = True
End If
End Sub

Private Sub Form_Load()
fontshow.Text = notepad.Text1.Font
fontsizeshow.Text = notepad.Text1.FontSize
If notepad.Text1.FontBold = True And notepad.Text1.FontItalic = False Then
fontstyleshow.Text = "bold"
ElseIf notepad.Text1.FontItalic = True And notepad.Text1.FontBold = False Then
fontstyleshow.Text = "italic"
ElseIf notepad.Text1.FontBold = True And notepad.Text1.FontItalic = True Then
fontstyleshow.Text = "bold italic"
Else
fontstyleshow.Text = "regular"
End If
End Sub

Private Sub greentextlabel_Click()
Label4.ForeColor = greentextlabel.ForeColor
End Sub

Private Sub ok_Click()
If fontshow.Text <> "" Then
notepad.Text1.Font = fontshow.Text
End If
If fontsizeshow.Text <> "" Then
notepad.Text1.FontSize = fontsizeshow.Text
End If
If fontstyleshow.Text = "regular" Then
notepad.Text1.FontBold = False
notepad.Text1.FontItalic = False
ElseIf fontstyleshow.Text = "bold" Then
notepad.Text1.FontBold = True
notepad.Text1.FontItalic = False
ElseIf fontstyleshow.Text = "italic" Then
notepad.Text1.FontItalic = True
notepad.Text1.FontBold = False
ElseIf fontstyleshow.Text = "bold italic" Then
notepad.Text1.FontItalic = True
notepad.Text1.FontBold = True
End If

If Check1.Value = 1 Then
notepad.Text1.FontStrikethru = True
Else
notepad.Text1.FontStrikethru = False
End If
If Check2.Value = 1 Then
notepad.Text1.FontUnderline = True
Else
notepad.Text1.FontUnderline = False
End If
Form1.Visible = False
notepad.Text1.ForeColor = Label4.ForeColor

notepad.Enabled = True
End Sub

Private Sub redtextlabel_Click()
Label4.ForeColor = redtextlabel.ForeColor
End Sub

Private Sub sizelist_Click()
fontsizeshow.Text = sizelist.Text
Label4.FontSize = sizelist.Text
End Sub
