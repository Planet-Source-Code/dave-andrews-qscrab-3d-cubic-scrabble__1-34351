Attribute VB_Name = "modQScrab"
Option Explicit
Option Base 1
Public Const pi = 3.141592654
Type QPos
    Side As Integer
    Row As Integer
    Col As Integer
End Type
Type Tile
    Char As String * 1
    Value As Integer
    Count As Integer
    Blank As Boolean
    Chosen As Boolean
    Played As Boolean
    Pos As Integer
End Type
Type QCoord
    DxID As Integer
    Up As QPos
    Down As QPos
    Left As QPos
    Right As QPos
    Piece As Tile
    Type As Integer
    Status As Integer
End Type
Type QWord
    Word As String
    Types As String
End Type
Public QCube() As QCoord
'Type Player
'    Rack(7) As Piece
'End Type
'Public Letters() As Piece
'Public Players() As Player
Public NoTile As Tile
Public Letters(27) As Tile
Public dRow As Boolean
Dim WordCount As Integer

Function Bonus(tRack() As Tile) As Integer
Dim i As Integer
For i = 1 To UBound(tRack)
    If tRack(i).Played = False Then
        Bonus = 0
        Exit Function
    End If
Next i
Bonus = 50
End Function

Sub DefineLetters()
Dim i As Integer
Dim FF As Integer
Dim Char As String * 1
Dim Count As Integer
Dim Value As Integer
FF = FreeFile
i = 1
Open App.Path & "\Letter.Set" For Input As #FF
Do While Not EOF(FF)
    Input #FF, Char, Count, Value
    Letters(i).Char = Char
    Letters(i).Value = Value
    Letters(i).Count = Count
    Letters(i).Played = False
    Letters(i).Chosen = False
    i = i + 1
Loop
Close #FF
End Sub




Function GetPoints(Char As String) As Integer
Select Case Char
    Case "A", "E", "I", "L", "N", "O", "R", "S", "T", "U"
        GetPoints = 1
    Case "D", "G"
        GetPoints = 2
    Case "B", "C", "M", "P"
        GetPoints = 3
    Case "F", "H", "V", "W", "Y"
        GetPoints = 4
    Case "K"
        GetPoints = 5
    Case "J", "X"
        GetPoints = 8
    Case "Q", "Z"
        GetPoints = 10
End Select
End Function
Function CheckWord(Word As String) As Boolean
Dim FF As Integer
Dim temp As String
Word = UCase(Word)
If Word = "" Then
    CheckWord = False
    Exit Function
End If
FF = FreeFile
Open "Dictionary\" & Left(Word, 1) & ".dic" For Input As #1
Do While Not EOF(FF)
    Input #FF, temp
    If temp = Word Then
        Close #FF
        CheckWord = True
        Exit Function
    End If
Loop
Close #FF
CheckWord = False
End Function

Sub ClearQCube()
Dim i As Integer
Dim j As Integer
Dim k As Integer
ReDim QCube(6, 7, 7) As QCoord
NoTile.Blank = False
NoTile.Char = ""
NoTile.Chosen = False
NoTile.Count = 0
NoTile.Played = False
NoTile.Pos = 0
NoTile.Value = 0
For i = 1 To 6
    For j = 1 To 7
        For k = 1 To 7
            QCube(i, j, k).Status = 0
            QCube(i, j, k).Type = GetModVal(i, j, k)
            QCube(i, j, k).Piece = NoTile
            QCube(i, j, k).Up = GetUp(i, j, k)
            QCube(i, j, k).Down = GetDown(i, j, k)
            QCube(i, j, k).Left = GetLeft(i, j, k)
            QCube(i, j, k).Right = GetRight(i, j, k)
        Next k
    Next j
Next i
DefineLetters
End Sub


Function GetModVal(Side As Integer, Col As Integer, Row As Integer) As Integer
Dim PosWord As String * 2
PosWord = Col & Row
'0 = normal
'1 = Double Letter
'2 = Tripple Letter
'3 = Double Word
'4 = tripple Word
Select Case Side
    Case 1, 3, 5
        Select Case PosWord
            Case "14", "31", "47", "75": GetModVal = 1
            Case "17": GetModVal = 4
            Case "22", "62": GetModVal = 2
            Case "26", "35", "44", "53", "71": GetModVal = 3
            Case Else: GetModVal = 0
        End Select
    Case 2, 4, 6
        Select Case PosWord
            Case "74", "57", "41", "13": GetModVal = 1
            Case "71": GetModVal = 4
            Case "66", "26": GetModVal = 2
            Case "62", "53", "44", "35", "17": GetModVal = 3
            Case Else: GetModVal = 0
        End Select
End Select
End Function
Function GetLeft(Side As Integer, Col As Integer, Row As Integer) As QPos
Select Case Side
    Case 1, 3, 5
        Select Case Col
            Case 1
                GetLeft.Side = 0
                GetLeft.Col = 0
                GetLeft.Row = 0
            Case 2, 3, 4, 5, 6, 7
                GetLeft.Side = Side
                GetLeft.Row = Row
                GetLeft.Col = Col - 1
        End Select
    Case 2, 4, 6
        Select Case Col
            Case 1
                GetLeft.Side = Side - 1
                GetLeft.Row = Row
                GetLeft.Col = 7
            Case 2, 3, 4, 5, 6, 7
                GetLeft.Side = Side
                GetLeft.Row = Row
                GetLeft.Col = Col - 1
        End Select
End Select

End Function

Function GetRight(Side As Integer, Col As Integer, Row As Integer) As QPos
Select Case Side
    Case 1, 3, 5
        Select Case Col
            Case 1, 2, 3, 4, 5, 6
                GetRight.Side = Side
                GetRight.Row = Row
                GetRight.Col = Col + 1
            Case 7
                GetRight.Side = Side + 1
                GetRight.Row = Row
                GetRight.Col = 1
        End Select
    Case 2, 4, 6
        Select Case Col
            Case 1, 2, 3, 4, 5, 6
                GetRight.Side = Side
                GetRight.Row = Row
                GetRight.Col = Col + 1
            Case 7
                GetRight.Side = 0
                GetRight.Col = 0
                GetRight.Row = 0
        End Select
End Select
End Function
Function GetUp(Side As Integer, Col As Integer, Row As Integer) As QPos
Select Case Side
    Case 1, 3, 5
        Select Case Row
            Case 1, 2, 3, 4, 5, 6
                GetUp.Side = Side
                GetUp.Row = Row + 1
                GetUp.Col = Col
            Case 7
                GetUp.Side = 0
                GetUp.Col = 0
                GetUp.Row = 0
        End Select
    Case 2, 4, 6
        Select Case Row
            Case 1, 2, 3, 4, 5, 6
                GetUp.Side = Side
                GetUp.Row = Row + 1
                GetUp.Col = Col
            Case 7
                If Side = 6 Then
                    GetUp.Side = 1
                Else
                    GetUp.Side = Side + 1
                End If
                GetUp.Row = 1
                GetUp.Col = Col
        End Select
End Select
End Function
Function GetDown(Side As Integer, Col As Integer, Row As Integer) As QPos
Select Case Side
    Case 1, 3, 5
        Select Case Row
            Case 1
                If Side = 1 Then
                    GetDown.Side = 6
                Else
                    GetDown.Side = Side - 1
                End If
                GetDown.Row = 7
                GetDown.Col = Col
            Case 2, 3, 4, 5, 6, 7
                GetDown.Side = Side
                GetDown.Row = Row - 1
                GetDown.Col = Col
        End Select
    Case 2, 4, 6
        Select Case Row
            Case 1
                GetDown.Side = 0
                GetDown.Col = 0
                GetDown.Row = 0
            Case 2, 3, 4, 5, 6, 7
                GetDown.Side = Side
                GetDown.Row = Row - 1
                GetDown.Col = Col
        End Select
End Select
End Function
Function GoodWordPlacement(OnRow As Boolean, Rack() As Tile) As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim FoundStart As Boolean
Dim FoundEnd As Boolean
Dim PartOfWord As Boolean
Dim Word As String
Dim WStart As QPos
Dim NextPos As QPos
Dim cPos As QPos
Dim TCount As Integer
For i = 1 To 7
    If Rack(i).Played Then TCount = TCount + 1
Next i
For i = 1 To 6
    For j = 1 To 7
        For k = 1 To 7
            If QCube(i, j, k).Status = 1 Then
                WStart.Side = i
                WStart.Col = j
                WStart.Row = k
                GoTo FindFirst
            End If
        Next k
    Next j
Next i
FindFirst:
NextPos = WStart
If OnRow Or TCount = 1 Then  'go to the left until you reach 0
    Do While QCube(NextPos.Side, NextPos.Col, NextPos.Row).Left.Col > 0
        If QCube(NextPos.Side, NextPos.Col, NextPos.Row).Status = 0 Then Exit Do
        NextPos = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Left
    Loop
    Do While QCube(NextPos.Side, NextPos.Col, NextPos.Row).Right.Col > 0
        If Not FoundStart Then
            If QCube(NextPos.Side, NextPos.Col, NextPos.Row).Piece.Played Then
                FoundStart = True
            End If
        End If
        If FoundStart And Not FoundEnd Then
            If QCube(NextPos.Side, NextPos.Col, NextPos.Row).Piece.Played Then
                Word = Word & QCube(NextPos.Side, NextPos.Col, NextPos.Row).Piece.Char
                cPos = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Up
                If cPos.Side > 0 Then If QCube(cPos.Side, cPos.Col, cPos.Row).Status = 2 Then PartOfWord = True
                cPos = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Down
                If cPos.Side > 0 Then If QCube(cPos.Side, cPos.Col, cPos.Row).Status = 2 Then PartOfWord = True
                cPos = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Left
                If cPos.Side > 0 Then If QCube(cPos.Side, cPos.Col, cPos.Row).Status = 2 Then PartOfWord = True
                cPos = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Right
                If cPos.Side > 0 Then If QCube(cPos.Side, cPos.Col, cPos.Row).Status = 2 Then PartOfWord = True
            Else
                FoundEnd = True
            End If
        ElseIf FoundStart And FoundEnd Then
            If QCube(NextPos.Side, NextPos.Col, NextPos.Row).Status = 1 Then
                Word = ""
            End If
        End If
        NextPos = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Right
    Loop
End If
If Len(Word) > 1 And TCount = 1 Then
    GoTo SkipCol
ElseIf TCount = 1 Then
    Word = ""
    NextPos = WStart
    FoundStart = False
    FoundEnd = False
End If
If Not OnRow Or TCount = 1 Then 'Go Up until you reach 0
    Do While QCube(NextPos.Side, NextPos.Col, NextPos.Row).Up.Row > 0
        If QCube(NextPos.Side, NextPos.Col, NextPos.Row).Status = 0 Then Exit Do
        NextPos = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Up
    Loop
    Do While QCube(NextPos.Side, NextPos.Col, NextPos.Row).Down.Row > 0
        If Not FoundStart Then
            If QCube(NextPos.Side, NextPos.Col, NextPos.Row).Piece.Played Then
                FoundStart = True
            End If
        End If
        If FoundStart And Not FoundEnd Then
            If QCube(NextPos.Side, NextPos.Col, NextPos.Row).Piece.Played Then
                Word = Word & QCube(NextPos.Side, NextPos.Col, NextPos.Row).Piece.Char
                cPos = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Up
                If cPos.Side > 0 Then If QCube(cPos.Side, cPos.Col, cPos.Row).Status = 2 Then PartOfWord = True
                cPos = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Down
                If cPos.Side > 0 Then If QCube(cPos.Side, cPos.Col, cPos.Row).Status = 2 Then PartOfWord = True
                cPos = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Left
                If cPos.Side > 0 Then If QCube(cPos.Side, cPos.Col, cPos.Row).Status = 2 Then PartOfWord = True
                cPos = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Right
                If cPos.Side > 0 Then If QCube(cPos.Side, cPos.Col, cPos.Row).Status = 2 Then PartOfWord = True
            Else
                FoundEnd = True
            End If
        ElseIf FoundStart And FoundEnd Then
            If QCube(NextPos.Side, NextPos.Col, NextPos.Row).Status = 1 Then
                Word = ""
            End If
        End If
        NextPos = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Down
    Loop
End If
SkipCol:
If WordCount > 1 And Not PartOfWord Then Word = ""
GoodWordPlacement = Word
End Function

Function CheckNewWords(ByRef wList() As QWord) As Boolean
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim wCount As Integer
Dim NextPos As QPos
wCount = 1
ReDim wList(wCount) As QWord
For i = 1 To 6
    For j = 1 To 7
        For k = 1 To 7
            If QCube(i, j, k).Status = 1 Then
                NextPos.Side = i
                NextPos.Col = j
                NextPos.Row = k
                FindWords NextPos, wList()
            End If
        Next k
    Next j
Next i
RemoveWordDups wList()

For i = 1 To UBound(wList)
    If Not CheckWord(wList(i).Word) Then
        CheckNewWords = False
        Exit Function
    End If
Next i
WordCount = WordCount + UBound(wList)
CheckNewWords = True
End Function
Sub FindWords(StartPos As QPos, ByRef wList() As QWord)
Dim NextPos As QPos
Dim Rlet As String
Dim Llet As String
Dim Ulet As String
Dim DLet As String
Dim SLet As String
Dim Rtyp As String
Dim Ltyp As String
Dim Utyp As String
Dim Dtyp As String
Dim Styp As String
Dim i As Integer
SLet = QCube(StartPos.Side, StartPos.Col, StartPos.Row).Piece.Char
Styp = QCube(StartPos.Side, StartPos.Col, StartPos.Row).Type
Dim ColWord As String
Dim RowWord As String
'Right
NextPos = QCube(StartPos.Side, StartPos.Col, StartPos.Row).Right
Do While NextPos.Side > 0
    If QCube(NextPos.Side, NextPos.Col, NextPos.Row).Status > 0 Then
        Rlet = Rlet & QCube(NextPos.Side, NextPos.Col, NextPos.Row).Piece.Char
        If QCube(NextPos.Side, NextPos.Col, NextPos.Row).Piece.Blank Then
            Rtyp = Rtyp & "9"
        Else
            Rtyp = Rtyp & QCube(NextPos.Side, NextPos.Col, NextPos.Row).Type
        End If
    Else
        Exit Do
    End If
    NextPos = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Right
Loop
'left
NextPos = QCube(StartPos.Side, StartPos.Col, StartPos.Row).Left
Do While NextPos.Side > 0
    If QCube(NextPos.Side, NextPos.Col, NextPos.Row).Status > 0 Then
        Llet = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Piece.Char & Llet
        If QCube(NextPos.Side, NextPos.Col, NextPos.Row).Piece.Blank Then
            Ltyp = "9" & Ltyp
        Else
            Ltyp = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Type & Ltyp
        End If
    Else
        Exit Do
    End If
    NextPos = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Left
Loop
'Up
NextPos = QCube(StartPos.Side, StartPos.Col, StartPos.Row).Up
Do While NextPos.Side > 0
    If QCube(NextPos.Side, NextPos.Col, NextPos.Row).Status > 0 Then
        Ulet = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Piece.Char & Ulet
        If QCube(NextPos.Side, NextPos.Col, NextPos.Row).Piece.Blank Then
            Utyp = "9" & Utyp
        Else
            Utyp = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Type & Utyp
        End If
    Else
        Exit Do
    End If
    NextPos = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Up
Loop
'Up
NextPos = QCube(StartPos.Side, StartPos.Col, StartPos.Row).Down
Do While NextPos.Side > 0
    If QCube(NextPos.Side, NextPos.Col, NextPos.Row).Status > 0 Then
        DLet = DLet & QCube(NextPos.Side, NextPos.Col, NextPos.Row).Piece.Char
        If QCube(NextPos.Side, NextPos.Col, NextPos.Row).Piece.Blank Then
            Dtyp = Dtyp & "9"
        Else
            Dtyp = Dtyp & QCube(NextPos.Side, NextPos.Col, NextPos.Row).Type
        End If
    Else
        Exit Do
    End If
    NextPos = QCube(NextPos.Side, NextPos.Col, NextPos.Row).Down
Loop
ColWord = Ulet & SLet & DLet
RowWord = Llet & SLet & Rlet
i = UBound(wList) + 1
If Ulet <> "" Or DLet <> "" Then
    ReDim Preserve wList(i) As QWord
    wList(i).Word = ColWord
    wList(i).Types = Utyp & Styp & Dtyp
End If
i = UBound(wList) + 1
If Rlet <> "" Or Llet <> "" Then
    ReDim Preserve wList(i) As QWord
    wList(i).Word = RowWord
    wList(i).Types = Ltyp & Styp & Rtyp
End If
End Sub

Sub RemoveWordDups(ByRef wList() As QWord)
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim Found As Boolean
k = 1
Dim tList(10) As QWord
For i = 2 To UBound(wList)
    Found = False
    For j = i + 1 To UBound(wList)
        If wList(i).Word = wList(j).Word Then Found = True
    Next j
    If Not Found Then
        tList(k) = wList(i)
        k = k + 1
    End If
Next i
If k = 1 Then Exit Sub
k = k - 1
ReDim wList(k) As QWord
For i = 1 To k
    wList(i) = tList(i)
Next i
End Sub

Function ScoreWords(wList() As QWord) As Integer
'0 = normal
'1 = Double Letter
'2 = Tripple Letter
'3 = Double Word
'4 = tripple Word
Dim tScore As Integer
Dim i As Integer
Dim j As Integer
Dim WordMod As Integer
For i = 1 To UBound(wList)
    WordMod = 1
    tScore = 0
    For j = 1 To Len(wList(i).Word)
        Select Case Mid(wList(i).Types, j, 1)
            Case 0
                tScore = tScore + GetPoints(Mid(wList(i).Word, j, 1))
            Case 1
                tScore = tScore + (2 * GetPoints(Mid(wList(i).Word, j, 1)))
            Case 2
                tScore = tScore + (3 * GetPoints(Mid(wList(i).Word, j, 1)))
            Case 3
                tScore = tScore + GetPoints(Mid(wList(i).Word, j, 1))
                WordMod = WordMod * 2
            Case 4
                tScore = tScore + GetPoints(Mid(wList(i).Word, j, 1))
                WordMod = WordMod * 3
        End Select
    Next j
    ScoreWords = ScoreWords + (WordMod * tScore)
Next i
End Function
Function TileSetOk(Side As Integer, Col As Integer, Row As Integer) As Boolean
Dim OnRow As Boolean
Dim OnCol As Boolean
Dim TCount As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
'First Check Rows
If QCube(Side, Col, Row).Status <> 0 Then
    TileSetOk = False
    Exit Function
End If
For i = 1 To 6
    For j = 1 To 7
        For k = 1 To 7
            If QCube(i, j, k).Piece.Chosen = True Or QCube(i, j, k).Status = 1 Then
                If j = Col And k <> Row Then OnCol = True
                If k = Row And j <> Col Then OnRow = True
                TCount = TCount + 1
            End If
        Next k
    Next j
Next i
If TCount = 1 Then
    TileSetOk = True
ElseIf OnCol And OnRow Then
    TileSetOk = False
ElseIf OnCol Or OnRow Then
    TileSetOk = True
    If TCount = 2 Then
        If OnCol Then dRow = False Else dRow = True
    ElseIf dRow <> OnRow Then
        TileSetOk = False
    End If
Else
    TileSetOk = False
End If

End Function


