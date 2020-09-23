VERSION 5.00
Begin VB.Form Findshow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find and Replace"
   ClientHeight    =   1635
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6165
   Icon            =   "Find.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1425
      TabIndex        =   7
      Top             =   600
      Width           =   2130
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   240
      Width           =   2115
   End
   Begin VB.CommandButton FindButton 
      Caption         =   "Find"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3720
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton FindNextButton 
      Caption         =   "Find Again"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   1305
   End
   Begin VB.CommandButton ReplaceButton 
      Caption         =   "Replace"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3720
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton ReplaceAllButton 
      Caption         =   "Replace All"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4800
      TabIndex        =   2
      Top             =   600
      Width           =   1275
   End
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4800
      TabIndex        =   1
      Top             =   1080
      Width           =   1245
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Case sensitive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   30
      TabIndex        =   0
      Top             =   1155
      Width           =   2040
   End
   Begin VB.Label Label1 
      Caption         =   "Find What"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   15
      TabIndex        =   9
      Top             =   240
      Width           =   1320
   End
   Begin VB.Label Label2 
      Caption         =   "Replace with"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   1290
   End
End
Attribute VB_Name = "Findshow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const HWND_TOPMOST = -1
Const SWP_SHOWWINDOW = &H40



Dim Position As Integer

Private Sub FindButton_Click()
Dim compare As Integer

Position = 0
If Check1.Value = 1 Then
    compare = vbBinaryCompare
Else
    compare = vbTextCompare
End If
Position = InStr(Position + 1, Mainform.Text1.Text, Text1.Text, compare)
If Position > 0 Then
    ReplaceButton.Enabled = True
    ReplaceAllButton.Enabled = True
    Mainform.Text1.SelStart = Position - 1
    Mainform.Text1.SelLength = Len(Text1.Text)
    Mainform.SetFocus
Else
    MsgBox "Word not found"
    ReplaceButton.Enabled = False
    ReplaceAllButton.Enabled = False
End If
End Sub

Private Sub FindNextButton_Click()
Dim compare As Integer

If Check1.Value = 1 Then
    compare = vbBinaryCompare
Else
    compare = vbTextCompare
End If
Position = InStr(Position + 1, Mainform.Text1.Text, Text1.Text, compare)
If Position > 0 Then
    Mainform.Text1.SelStart = Position - 1
    Mainform.Text1.SelLength = Len(Text1.Text)
    Mainform.SetFocus
Else
    MsgBox "String not found"
    ReplaceButton.Enabled = False
    ReplaceAllButton.Enabled = False
End If

End Sub

Private Sub Command5_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Dim ret As Long
Dim retvalue As Integer
Load Findshow
    retvalue = SetWindowPos(Me.hwnd, HWND_TOPMOST, 300, 100, _
               450, 125, SWP_SHOWWINDOW)
    
End Sub

Private Sub ReplaceButton_Click()
Dim compare As Integer

    Mainform.Text1.SelText = Text2.Text
    If Check1.Value = 1 Then
        compare = vbBinaryCompare
    Else
        compare = vbTextCompare
    End If
    Position = InStr(Position + 1, Mainform.Text1.Text, Text1.Text, compare)
    If Position > 0 Then
        Mainform.Text1.SelStart = Position - 1
        Mainform.Text1.SelLength = Len(Text1.Text)
        Mainform.SetFocus
    Else
        MsgBox "String not found"
        ReplaceButton.Enabled = False
        ReplaceAllButton.Enabled = False
    End If
    
End Sub

Private Sub ReplaceAllButton_Click()
Dim compare As Integer

    Mainform.Text1.SelText = Text2.Text
    If Check1.Value = 1 Then
        compare = vbBinaryCompare
    Else
        compare = vbTextCompare
    End If
    Position = InStr(Position + 1, Mainform.Text1.Text, Text1.Text, compare)
    While Position > 0
        Mainform.Text1.SelStart = Position - 1
        Mainform.Text1.SelLength = Len(Text1.Text)
        Mainform.Text1.SelText = Text2.Text
        Position = Position + Len(Text2.Text)
        Position = InStr(Position + 1, Mainform.Text1.Text, Text1.Text)
    Wend
        ReplaceButton.Enabled = False
        ReplaceAllButton.Enabled = False
        MsgBox "Done replacing"
End Sub


