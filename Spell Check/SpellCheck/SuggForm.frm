VERSION 5.00
Begin VB.Form SuggestionsForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Word's Spelling Suggestions"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   Icon            =   "SuggForm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "R e p l a c e"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   1845
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C l o s e"
      BeginProperty Font 
         Name            =   "Modern"
         Size            =   11.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2160
      TabIndex        =   2
      Top             =   3720
      Width           =   1845
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   3045
      TabIndex        =   1
      Top             =   570
      Width           =   2865
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   570
      Width           =   2865
   End
   Begin VB.Label Label2 
      Caption         =   "Suggested"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3060
      TabIndex        =   4
      Top             =   240
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Word Errors"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   240
      Width           =   1320
   End
End
Attribute VB_Name = "SuggestionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo CloseError
    AppWord.ActiveDocument.Close False
    If NewInstance Then
        AppWord.Quit
    End If
    Me.Hide
    Exit Sub
CloseError:
    MsgBox "Couldn't close document!"
    Me.Hide
End Sub

Private Sub Command2_Click()
    If List1.ListIndex = -1 Or List2.ListIndex = -1 Then
        MsgBox "Please select a word and an alternate spelling"
        Exit Sub
    End If
    Mainform.Text1.Text = Replace(Mainform.Text1.Text, List1.Text, List2.Text)
End Sub

Private Sub List1_Click()
    Screen.MousePointer = vbHourglass
    Set CorrectionsCollection = _
        AppWord.GetSpellingSuggestions(List1.Text)
    List2.Clear
    For iSuggWord = 1 To CorrectionsCollection.Count
        List2.AddItem CorrectionsCollection.Item(iSuggWord)
    Next
    Screen.MousePointer = vbDefault
End Sub
