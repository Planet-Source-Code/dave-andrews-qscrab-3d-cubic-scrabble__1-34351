VERSION 5.00
Begin VB.Form frmAna 
   Caption         =   "ANAGRAMS"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3405
   Icon            =   "frmAna.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   3405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.ListBox List1 
      Height          =   1980
      IntegralHeight  =   0   'False
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frmAna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const LB_FINDSTRING = &H18F
Const LB_FINDSTRINGEXACT = &H1A2
Const CB_FINDSTRING = &H14C
Const CB_FINDSTRINGEXACT = &H158
Dim AllStop As Boolean
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Function InList(sStringToFind As String, lstListBox As ListBox) As Boolean
    InList = SendMessageByString(lstListBox.hwnd, LB_FINDSTRING, -1, sStringToFind) >= 0
End Function

Function InCombo(sStringToFind, cbCombo As ComboBox) As Boolean
    InCombo = SendMessageByString(cbCombo.hwnd, CB_FINDSTRING, -1, sStringToFind) >= 0
End Function


Sub GetAnagrams()
Dim i As Long
List1.Clear
CheckGram Text1.Text
For i = 0 To UBound(AllWords)
    DoEvents
    If AllStop Then Exit Sub
    If IsWord(AllWords(i)) Then
        If Not InList(AllWords(i), List1) Then
            List1.AddItem AllWords(i)
        End If
    End If
Next i
End Sub




Sub CheckMSWord(lWord As String)
Dim objMSWord As New Word.Application
Dim sugList As SpellingSuggestions
Dim sug As SpellingSuggestion
DoEvents
'Set objMSWord = New Word.Application
objMSWord.WordBasic.FileNew 'open a doc
objMSWord.Visible = False 'hide the doc
If objMSWord.CheckSpelling(lWord) Then
    MsgBox " True"
End If
objMSWord.Quit
Set objMSWord = Nothing
End Sub

Private Sub cmdClear_Click()
AllStop = True
Text1.Text = ""
List1.Clear
End Sub

Private Sub Form_Load()
'initMSWord
InitWindow Me
List1.Height = Me.ScaleHeight - List1.Top
LoadDict
End Sub

Private Sub Form_Resize()
On Error Resume Next
List1.Height = Me.ScaleHeight - List1.Top
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set objMSWord = Nothing
End
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    AllStop = False
    GetAnagrams
End If
End Sub

