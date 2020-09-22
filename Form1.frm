VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   387
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   629
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3735
      Left            =   480
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   8
      Top             =   1440
      Width           =   3735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "export From File"
      Height          =   615
      Left            =   7320
      TabIndex        =   7
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Left            =   3360
      Top             =   2760
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Export"
      Height          =   495
      Left            =   7200
      TabIndex        =   6
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Remove Duplicates"
      Height          =   495
      Left            =   7200
      TabIndex        =   5
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "WSWords"
      Height          =   495
      Left            =   7200
      TabIndex        =   4
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "UK"
      Height          =   495
      Left            =   7200
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "King James"
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Capitalize"
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HTML"
      Height          =   495
      Left            =   7200
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RMC As New clsDX73D
Dim RMC2 As New clsDX73D
Private Sub Command1_Click()
Exit Sub
Dim MyDB As Database
Dim MyRS As Recordset
Set MyDB = DBEngine.Workspaces(0).OpenDatabase("QScrab.mdb")
Set MyRS = MyDB.OpenRecordset("Words")
Dim FF As Integer
Dim i As Integer
Dim Pth As String
Dim fName As String
Dim Temp As String
Dim Word As String
fName = Dir(App.Path & "\v003\", vbNormal)

Do While fName <> ""
    FF = FreeFile
    Open App.Path & "\v003\" & fName For Input As #FF
    Do While Not EOF(FF)
        Line Input #FF, Temp
        For i = 1 To Len(Temp)
            If Mid(Temp, i, 4) = "</B>" Then
                Word = Mid(Temp, 7, i - 7)
                Me.Caption = fName & " - " & Word
                DoEvents
                Exit For
            End If
        Next i
        MyRS.AddNew
        MyRS("Word") = UCase(Left(Word, 50))
        MyRS.Update
    Loop
    Close #FF
    fName = Dir(, vbNormal)
Loop
Close #FF
Set MyRS = Nothing
Set MyDB = Nothing
End Sub

Private Sub Command2_Click()
Exit Sub
Dim MyDB As Database
Dim MyRS As Recordset
Set MyDB = DBEngine.Workspaces(0).OpenDatabase("QScrab.mdb")
Set MyRS = MyDB.OpenRecordset("Words")
Do While Not MyRS.EOF
    MyRS.Edit
    MyRS("Word") = UCase(MyRS("Word"))
    Me.Caption = UCase(MyRS("Word"))
    DoEvents
    MyRS.Update
    MyRS.MoveNext
Loop
Set MyRS = Nothing
Set MyDB = Nothing
End Sub


Private Sub Command3_Click()
Exit Sub
Dim MyDB As Database
Dim MyRS As Recordset
Set MyDB = DBEngine.Workspaces(0).OpenDatabase("QScrab.mdb")
Set MyRS = MyDB.OpenRecordset("Words", dbOpenDynaset)
Dim FF As Integer
Dim i As Integer
Dim Pth As String
Dim fName As String
Dim Temp As String
Dim Word As String
fName = App.Path & "\KJWORDS.TXT"
FF = FreeFile
Open fName For Input As #FF
Do While Not EOF(FF)
    Line Input #FF, Word
    Me.Caption = Word
    DoEvents
    MyRS.AddNew
    MyRS("Word") = UCase(Left(Word, 50))
    If Asc(Left(Word, 1)) < 96 Then
        MyRS("Proper") = True
    End If
    MyRS.Update
Loop
Close #FF
Set MyRS = Nothing
Set MyDB = Nothing
End Sub


Private Sub Command4_Click()
'Exit Sub
Dim MyDB As Database
Dim MyRS As Recordset
Set MyDB = DBEngine.Workspaces(0).OpenDatabase("QScrab.mdb")
Set MyRS = MyDB.OpenRecordset("Words", dbOpenDynaset)
Dim FF As Integer
Dim i As Integer
Dim Pth As String
Dim fName As String
Dim Temp As String
Dim Word As String
fName = App.Path & "\UKACD17.TXT"
FF = FreeFile
Open fName For Input As #FF
Do While Not EOF(FF)
    Line Input #FF, Word
    Me.Caption = Word
    DoEvents
    If Asc(Left(Word, 1)) > 96 Then
        MyRS.FindFirst "Word = '" & UCase(Word) & "'"
        If Not MyRS.NoMatch Then
            MyRS.Edit
            MyRS("Proper") = False
            MyRS("Ok") = True
            MyRS.Update
        End If
    End If
Loop
Close #FF
Set MyRS = Nothing
Set MyDB = Nothing
End Sub


Private Sub Command5_Click()
Exit Sub
Dim MyDB As Database
Dim MyRS As Recordset
Set MyDB = DBEngine.Workspaces(0).OpenDatabase("QScrab.mdb")
Set MyRS = MyDB.OpenRecordset("Words", dbOpenDynaset)
Dim FF As Integer
Dim i As Integer
Dim Pth As String
Dim fName As String
Dim Temp As String
Dim Word As String
fName = App.Path & "\WSWORDS1.TXT"
FF = FreeFile
Open fName For Input As #FF
Do While Not EOF(FF)
    Line Input #FF, Word
    Me.Caption = Word
    DoEvents
    MyRS.AddNew
    MyRS("Word") = UCase(Left(Word, 50))
    If Asc(Left(Word, 1)) < 96 Then
        MyRS("Proper") = True
    End If
    MyRS.Update
Loop
Close #FF
Set MyRS = Nothing
Set MyDB = Nothing
End Sub


Private Sub Command6_Click()
Exit Sub
Dim MyDB As Database
Dim MyRS As Recordset
Set MyDB = DBEngine.Workspaces(0).OpenDatabase("QScrab.mdb")
Set MyRS = MyDB.OpenRecordset("Select * From Words Order By Word", dbOpenDynaset)
Dim Temp As String
Do While Not MyRS.EOF
    If MyRS("Word") = Temp Then
        MyRS.Edit
        MyRS("Duplicate") = True
        MyRS.Update
    End If
    Temp = MyRS("Word")
    Me.Caption = Temp
    DoEvents
    MyRS.MoveNext
Loop
Set MyRS = Nothing
Set MyDB = Nothing
End Sub


Private Sub Command7_Click()
Dim MyDB As Database
Dim MyRS As Recordset
Dim FF As Integer
Dim Temp As String
FF = FreeFile
Set MyDB = DBEngine.Workspaces(0).OpenDatabase("QScrab.mdb")
Set MyRS = MyDB.OpenRecordset("Select Word From Words Where OK = True Order By Word")
Do While Not MyRS.EOF
    If Left(Temp, 1) <> Left(MyRS("Word"), 1) Then
        Close #FF
        Open "Dictionary\" & Left(MyRS("Word"), 1) & ".dic" For Output As #FF
    End If
    Temp = MyRS("Word")
    Print #FF, Temp
    Me.Caption = Temp
    MyRS.MoveNext
    DoEvents
Loop
Set MyRS = Nothing
Set MyDB = Nothing
Close #FF
End Sub


Private Sub Command8_Click()
Dim MyDB As Database
Dim MyRS As Recordset
Dim FF As Integer
Dim F2 As Integer
Dim Temp As String
Dim Temp2 As String
fName = App.Path & "\UKACD17.TXT"
F2 = FreeFile
Open fName For Input As #F2
Do While Not EOF(F2)
    Input #F2, Temp
    If Asc(Left(Temp, 1)) > 96 And Asc(Left(Temp, 1)) < 123 Then
        Temp = UCase(Temp)
        If Left(Temp, 1) <> Left(Temp2, 1) Then
            Close #FF
            FF = FreeFile
            Open "Dictionary\" & Left(Temp, 1) & ".dic" For Append As #FF
        End If
        Print #FF, Temp
        Me.Caption = Temp
        DoEvents
        Temp2 = Temp
    End If
Loop
Close #FF
Close #F2
End Sub


Private Sub Form_Load()
RMC.InitDx Me
RMC.mFrO.AddVisual RMC.CreateBoxMesh(4, 4, 4)
RMC.mFrO.SetPosition Nothing, 0, 0, 10
RMC.Resize Me

RMC2.InitDx Picture1
RMC2.mFrO.AddVisual RMC2.CreateBoxMesh(4, 4, 4)
RMC2.mFrO.SetPosition Nothing, 0, 0, 10
RMC2.Resize Me
Timer1.Interval = 1
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RMC.MouseDown X, Y
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RMC.MouseMove X, Y
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RMC.MouseUp
End Sub

Private Sub Form_Resize()
RMC.Resize Me
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RMC2.MouseDown X, Y
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RMC2.MouseMove X, Y
End Sub


Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RMC2.MouseUp
End Sub


Private Sub Timer1_Timer()
RMC.Update
RMC2.Update
End Sub


