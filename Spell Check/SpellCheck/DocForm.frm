VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Mainform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SpellCheck Project- Requires MSWORD"
   ClientHeight    =   5760
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9495
   Icon            =   "DocForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5730
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "DocForm.frx":08CA
      Top             =   0
      Width           =   9480
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu Save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu Split0 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Begin VB.Menu Fandr 
         Caption         =   "&Find and Replace"
         Shortcut        =   ^F
      End
      Begin VB.Menu Selall 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu Spell 
         Caption         =   "&Spell Check"
      End
   End
   Begin VB.Menu Opt 
      Caption         =   "Options"
      Begin VB.Menu TextAuto 
         Caption         =   "Text Auto Filler"
      End
   End
   Begin VB.Menu Hel 
      Caption         =   "Help"
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This example will show you how to import MSWORD spell Check
'to your vb Applications.
'References: MSWORD.OLB is used
'You need to include this file while distributing this application
'=====================================
'VOTE FOR ME
'=====================================
'This code may be small yet powerful. Who wishes to install
'a 5 MB Dictionary with their application. Isn't this a simple
'Replacement.
Dim NewInstance As Boolean

Private Sub About_Click()
MsgBox "Hi there, don't forget to VOTE FOR ME OK" & vbCrLf & "Also download the other codes i have submitted" & vbCrLf & vbCrLf & "Email: sriharish@msn.com"

End Sub

Private Sub Fandr_Click()
Findshow.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Open_Click()
Dim readfile As String
CommonDialog1.Filter = "Text File |*.txt|"
CommonDialog1.ShowOpen
If Len(CommonDialog1.FileTitle) > 0 Then
Close #1
Open CommonDialog1.FileName For Input As #1
Text1.Text = ""
Do Until EOF(1)
Line Input #1, readfile
Text1.Text = Text1.Text + vbCrLf + readfile
Loop
Close #1
End If
End Sub

Private Sub Save_Click()
If Text1.Text = "" Then
MsgBox "Please type something. Anythign will do"
Exit Sub
End If
CommonDialog1.Filter = "Text File|*.txt|"
CommonDialog1.ShowSave
If Len(CommonDialog1.FileTitle) > 1 Then
Close #1
Open CommonDialog1.FileName & ".txt" For Output As #1
Print #1, Text1.Text
Close #1
End If
End Sub

Private Sub Selall_Click()
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Spell_Click()
Dim DRange As Range
Wait.Show
Wait.Label1.Caption = "Importing Dictionary... "
    On Error Resume Next
    Set AppWord = GetObject(, "Word.Application")
    If AppWord Is Nothing Then
        Set AppWord = CreateObject("Word.Application")
        If AppWord Is Nothing Then
            MsgBox "Could not import dictionary.MS Word may not be installed in your PC use the file MSWORD.OLB provided with the source code"
            End
        Else
            NewInstance = True
        End If
    Else
        NewInstance = False
    End If
On Error GoTo ErrorHandler
    AppWord.Documents.Add
    Wait.Show
    Wait.Label1.Caption = "checking words..."
    Set DRange = AppWord.ActiveDocument.Range
    DRange.InsertAfter Text1.Text
    Set SpellCollection = DRange.SpellingErrors
    If SpellCollection.Count > 0 Then
        SuggestionsForm.List1.Clear
        SuggestionsForm.List2.Clear
        For iWord = 1 To SpellCollection.Count
            SuggestionsForm!List1.AddItem SpellCollection.Item(iWord)
            If SuggestionsForm!List1.List(SuggestionsForm!List1.NewIndex) = SuggestionsForm!List1.List(SuggestionsForm!List1.NewIndex + 1) Then
                SuggestionsForm!List1.RemoveItem SuggestionsForm!List1.NewIndex
            End If
        Next
    End If
       SuggestionsForm.Show
    Unload Wait
    Exit Sub
    
ErrorHandler:
    MsgBox "The following error occured during the document's spelling" & vbCrLf & Err.Description
End Sub
Private Sub TextAuto_Click()
MsgBox "You can download Text Autofiller Plugin for this application. Read Important.txt for more", vbInformation
End Sub
