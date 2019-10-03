VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form notepad 
   Caption         =   "notepad"
   ClientHeight    =   14355
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   24060
   Enabled         =   0   'False
   Icon            =   "notepad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   14355
   ScaleWidth      =   24060
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog cd 
      Left            =   19560
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   14295
      HideSelection   =   0   'False
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   24015
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Index           =   0
      WindowList      =   -1  'True
      Begin VB.Menu filenew 
         Caption         =   "New                           "
         Shortcut        =   ^N
      End
      Begin VB.Menu filesave 
         Caption         =   "Save                          "
         Shortcut        =   ^S
      End
      Begin VB.Menu filesaveas 
         Caption         =   "Save As                          "
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu fileopen 
         Caption         =   "Open                         "
         Shortcut        =   ^O
      End
      Begin VB.Menu fileprint 
         Caption         =   "Print                           "
         Shortcut        =   ^P
      End
      Begin VB.Menu fileexit 
         Caption         =   "Exit                            "
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu formate 
      Caption         =   "Formate"
      Index           =   1
      Begin VB.Menu formatefont 
         Caption         =   "Font                      "
         Shortcut        =   +{F12}
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Index           =   2
      Begin VB.Menu editFind 
         Caption         =   "Find                         "
         Shortcut        =   ^F
      End
      Begin VB.Menu editwraptext 
         Caption         =   "Wrap Text               "
         Begin VB.Menu editwraptextwrap 
            Caption         =   "Wrap          "
         End
         Begin VB.Menu editwrapunwrap 
            Caption         =   "Unwrap"
         End
      End
      Begin VB.Menu editselectall 
         Caption         =   "Select All                    "
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "notepad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim linetext, openflag, newflag, newsaveflag As String
Dim filenameroot As Variant
Dim counterForFileLabel, characterslength, newfilesize As Integer
Public filelocation As String


Private Sub editFind_Click()
Find.Visible = True
notepad.ZOrder 1
Find.ZOrder 0
End Sub

Private Sub editselectall_Click()
If Text1.Text <> "" Then   ' {
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End If      '}
End Sub

Private Sub File_Click(Index As Integer)
If Index = 1 Then

End If
End Sub

Private Sub fileexit_Click()
End
End Sub

Private Sub filenew_Click()
newbox
End Sub

Private Sub fileopen_Click()
Dim an As Variant
If Text1.Text <> "" Then            ';;;;

If newflag = "yes" And newsaveflag = "yes" Then            '===

If newfilesize <> Len(Text1.Text) Then
an = MsgBox("               Want to save Changes ?              ", vbYesNoCancel + vbQuestion, "Save Changes ?")
If an = vbYes Then
realsavebox
ElseIf an = vbNo Then
openbox
End If

Else
openbox
End If

ElseIf openflag = "open" And characterslength <> Len(Text1.Text) Then
an = MsgBox("              Want to save changes ?              ", vbYesNoCancel + vbQuestion, "Save Changes ?")

If an = vbYes Then          '[[[[
realsavebox
openbox
ElseIf an = vbNo Then
openbox
End If                      '[[[[

ElseIf openflag = "open" And characterslength = Len(Text1.Text) Then
openbox

ElseIf openflag <> "open" Then

an = MsgBox("              Save File ?             ", vbYesNoCancel + vbQuestion, "Save ?")

If an = vbYes Then          '---
realsavebox
openbox
ElseIf an = vbNo Then
openbox
End If                      '---

End If                  '===


Else
openbox
End If              ';;;;
End Sub



Private Sub fileprint_Click()



cd.ShowFont
End Sub

Private Sub filesave_Click()
realsavebox
End Sub



Private Sub filesaveas_Click()
savebox
End Sub

Private Sub Form_Load()

If Dir.Text1.Text <> "" Then
filelocation = Dir.Text1.Text & "\fonts.txt"
MsgBox filelocation
Open filelocation For Input As #1
Dim a, b As Integer
Dim fontdata As String
b = 2
Do
Input #1, fontdata
Form1.fontlist.AddItem fontdata
Loop Until EOF(1)
End If


For a = 1 To 15
Form1.sizelist.AddItem a + a
Next
Close #1

Form1.fontstylelist.AddItem "regular"
Form1.fontstylelist.AddItem "bold"
Form1.fontstylelist.AddItem "italic"
Form1.fontstylelist.AddItem "bold italic"

' initialise for counterForFileLabel
counterForFileLabel = 1


If Find.Enabled = False Then
notepad.Enabled = True
End If

openflag = "no"
newflag = "yes"
newsaveflag = "no"
notepad.Caption = "No Title"


Dir.Visible = True
End Sub

Private Sub formatefont_Click()
Form1.Visible = True
Form1.SetFocus
Form1.ZOrder (0)
notepad.ZOrder (1)
notepad.Enabled = False
End Sub


' sub procedure for openbox
Sub openbox()

cd.ShowOpen
If cd.FileName <> "" Then   '---
filenameroot = cd.FileName
Text1.Text = ""
Open cd.FileName For Input As #1
Do
Input #1, linetext
Text1.Text = Text1.Text & linetext
Loop Until EOF(1)
notepad.Caption = cd.FileName
Close #1
characterslength = Len(Text1.Text)
openflag = "open"

End If

End Sub

'sub procedure for saveAsbox

Sub savebox()

cd.ShowSave
If cd.FileName <> "" Then
Open cd.FileName For Output As #1
Print #1, Text1.Text
Close #1

notepad.Caption = cd.FileName
If newflag = "yes" Then
newfilesize = Len(Text1.Text)
newsaveflag = "yes"
End If

End If

End Sub


'sub procedure for new

Sub newbox()
Dim ask As Variant
If openflag = "open" Then
If Len(Text1.Text) <> characterslength Then
ask = MsgBox("              Want to Save Changes ?              ", vbYesNoCancel + vbQuestion, "Save Changes")

If ask = vbYes Then
realsavebox
Text1.Text = ""
notepad.Caption = "No Title"
newflag = "yes"
ElseIf ask = vbNo Then
Text1.Text = ""
newflag = "yes"
notepad.Caption = "No Title"
End If

Else
Text1.Text = ""
notepad.Caption = "No Title"
newflag = "yes"
End If
openflag = "close"
Else

If Text1.Text <> "" Then
ask = MsgBox("             Save File ?             ", vbYesNoCancel + vbQuestion, "Save ?")

If ask = vbYes Then
savebox
Text1.Text = ""
notepad.Caption = "No Title"
newflag = "yes"
ElseIf ask = vbNo Then
Text1.Text = ""
notepad.Caption = "No Title"
newflag = "yes"
End If

End If

End If
openflag = "close"
End Sub

' sub procedure for save
Sub realsavebox()
If openflag = "open" Then
Open filenameroot For Output As #2
Print #2, Text1.Text
Close #2

ElseIf openflag <> "open" Then
savebox
End If
End Sub


' sub procedure of fonts

Public Sub fonts()

If Dir.Text1.Text <> "" Then
filelocation = Dir.Text1.Text & "\fonts.txt"
MsgBox filelocation
Open filelocation For Input As #1
Dim a, b As Integer
Dim fontdata As String
b = 2
Do
Input #1, fontdata
Form1.fontlist.AddItem fontdata
Loop Until EOF(1)
End If


End Sub

