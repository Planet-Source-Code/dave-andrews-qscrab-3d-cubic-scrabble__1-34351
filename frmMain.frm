VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "QScrab"
   ClientHeight    =   6510
   ClientLeft      =   195
   ClientTop       =   405
   ClientWidth     =   7725
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   515
   Begin VB.PictureBox picRight 
      Align           =   4  'Align Right
      BackColor       =   &H00000000&
      Height          =   5415
      Left            =   6555
      ScaleHeight     =   357
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   2
      Top             =   0
      Width           =   1170
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtScore 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         TabIndex        =   8
         Text            =   "0"
         Top             =   0
         Width           =   1095
      End
      Begin VB.PictureBox picMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   0
         ScaleHeight     =   239
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   3
         Top             =   1680
         Width           =   1095
         Begin VB.CommandButton Command2 
            Caption         =   "Submit Word"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   0
            TabIndex        =   18
            Top             =   2640
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   17
            Top             =   2160
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Exchange Letters"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   19
            Top             =   1680
            Width           =   1095
         End
         Begin VB.CommandButton cmdGoSide 
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   720
            TabIndex        =   10
            Top             =   1320
            Width           =   375
         End
         Begin VB.CommandButton cmdGoSide 
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   360
            TabIndex        =   11
            Top             =   1320
            Width           =   375
         End
         Begin VB.CommandButton cmdGoSide 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   0
            TabIndex        =   12
            Top             =   1320
            Width           =   375
         End
         Begin VB.CommandButton cmdGoSide 
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   720
            TabIndex        =   13
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton cmdGoSide 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   360
            TabIndex        =   14
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton cmdGoSide 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   15
            Top             =   960
            Width           =   375
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000FFFF&
            X1              =   0
            X2              =   72
            Y1              =   43
            Y2              =   43
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "SIDE"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "3X Word"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Index           =   0
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "3X Letter"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   135
            Index           =   1
            Left            =   0
            TabIndex        =   6
            Top             =   150
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "2X Word"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0FF&
            Height          =   135
            Index           =   2
            Left            =   0
            TabIndex        =   5
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "2X Letter"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   135
            Index           =   3
            Left            =   0
            TabIndex        =   4
            Top             =   450
            Width           =   1095
         End
      End
   End
   Begin VB.PictureBox picRack 
      Align           =   2  'Align Bottom
      BackColor       =   &H00000000&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   511
      TabIndex        =   0
      Top             =   5415
      Width           =   7725
   End
   Begin VB.PictureBox picMain 
      Align           =   3  'Align Left
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      ScaleHeight     =   361
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   529
      TabIndex        =   1
      Top             =   0
      Width           =   7935
   End
   Begin VB.Timer Timer1 
      Left            =   10200
      Top             =   6480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim QFrame As Direct3DRMFrame3
Dim Qmesh As Direct3DRMMeshBuilder3
Dim QTex As Direct3DRMTexture3
Dim QSurf As DirectDrawSurface4
Dim QSurfDesc As DDSURFACEDESC2
Dim QDD As DirectDraw7
Dim PFrame As Direct3DRMFrame3
Dim PMesh As Direct3DRMMeshBuilder3
Dim RFrame(7) As Direct3DRMFrame3
Dim Rack(7) As Tile
Dim SelTile As Tile
Dim NewWords() As QWord
Dim RMC As New clsDX73D
Dim rmcRack As New clsDX73D
Sub BadMove(Status As String)
Dim i As Integer
For i = 1 To 7
    Rack(i).Chosen = False
    Rack(i).Played = False
Next i
ResetRack
RemoveTempWord
Me.Caption = Status
End Sub


Sub ClearRack()
On Error Resume Next
Dim i As Integer
For i = 1 To 7
    RFrame(i).DeleteVisual RFrame(i).GetVisual(0)
Next i
End Sub


Sub FOrceRack(Word As String)
Dim NewRack As String
Dim OldRack As String
BadMove "Exchanging Letters"
Dim i As Integer
For i = 1 To 7
    OldRack = OldRack & Rack(i).Char
Next i

NewRack = Word
ClearRack
SetRack NewRack
ReturnLetters OldRack
End Sub

Function GetPieceElement(pName As String) As Integer
Dim i As Integer
For i = 0 To QFrame.GetChildren.GetSize - 1
    If QFrame.GetChildren.GetElement(i).GetVisual(0).GetName = pName Then
        GetPieceElement = i
        'Exit Function
    End If
Next i
End Function

Private Function QSpace() As Direct3DRMMeshBuilder3
Set QSpace = RMC.mDrm.CreateMeshBuilder
Dim f As Direct3DRMFace2
Set f = RMC.mDrm.CreateFace
f.AddVertex 0.25, 0.25, -0.25
f.AddVertex 3.75, 0.25, -0.25
f.AddVertex 3.75, 3.75, -0.25
f.AddVertex 0.25, 3.75, -0.25
QSpace.AddFace f
Set f = Nothing
End Function

Sub RefreshRack()
Dim i As Integer
Dim Blank As Boolean
For i = 1 To 7
    'Set RFrame(i) = RMC.mDrm.CreateFrame(rmcRack.mFrO)
    Set Qmesh = RMC.CreateBoxMesh(3.8, 3.8, 1)
    If Rack(i).Char = " " Then Blank = True Else Blank = False
    Qmesh.SetTexture rmcRack.mDrm.LoadTexture("pine.bmp")
    SetTileChar Qmesh, Rack(i).Char, Blank
    RFrame(i).DeleteVisual RFrame(i).GetVisual(0)
    RFrame(i).AddVisual Qmesh
    Qmesh.SetName i
    RFrame(i).SetPosition Nothing, (i - 4) * 5, 0, 35
Next i
End Sub

Sub RemoveTempWord()
Dim i As Integer
Dim j As Integer
Dim k As Integer
For i = 1 To 6
    For j = 1 To 7
        For k = 1 To 7
            If QCube(i, j, k).Status = 1 Then
                QCube(i, j, k).Status = 0
                QFrame.DeleteChild QFrame.GetChildren.GetElement(GetPieceElement("Piece" & i & j & k))
            End If
        Next k
    Next j
Next i
RMC.Update
End Sub

Sub ReturnLetters(OldRack As String)
Dim i As Integer
For i = 1 To Len(OldRack)
    Select Case Mid(OldRack, i, 1)
        Case " "
            Letters(27).Count = Letters(27).Count + 1
        Case Else
            Letters(Asc(Mid(OldRack, i, 1)) - 64).Count = Letters(Asc(Mid(OldRack, i, 1)) - 64).Count + 1
    End Select
Next i
End Sub

Sub SetWord()
Dim i As Integer
Dim j As Integer
Dim k As Integer
For i = 1 To 6
    For j = 1 To 7
        For k = 1 To 7
            If QCube(i, j, k).Status = 1 Then
                QCube(i, j, k).Status = 2
                QCube(i, j, k).Type = 0
            End If
        Next k
    Next j
Next i
End Sub
Function GetElement(Side As Integer, Col As Integer, Row As Integer)
GetElement = ((Side - 1) * 49) + ((Col - 1) * 7) + (Row - 1)
End Function
Sub BuildCube()
Dim CFrame As Direct3DRMFrame3
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim SName As String
Dim R(1) As RECT
Set Qmesh = RMC.mDrm.CreateMeshBuilder
Set QFrame = RMC.mDrm.CreateFrame(RMC.mFrO)
SName = "Space4.x"
RMC.mFrO.SetSceneBackgroundRGB 1, 1, 1
'GoTo Skip:
'Side 1
For j = 0 To 6
    For k = 0 To 6
        Set Qmesh = RMC.mDrm.CreateMeshBuilder
        'Qmesh.AddMeshBuilder QSpace, D3DRMADDMESHBUILDER_FLATTENSUBMESHES
        Qmesh.LoadFromFile SName, 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        Qmesh.Translate j * 4, k * 4, 0
        'QMesh.SetColorRGB 1, 0, 0
        ''Qmesh.SetColorRGB 1, 0.5, 0
        SetModifiers Qmesh, 1, j + 1, k + 1
        Qmesh.SetName "Space" & "1" & j + 1 & k + 1
        Set CFrame = RMC.mDrm.CreateFrame(Nothing)
        CFrame.AddVisual Qmesh
        CFrame.SetPosition Nothing, -14, -14, -14
        QFrame.AddChild CFrame
        Set CFrame = Nothing
    Next k
Next j
'Side 2
For j = 0 To 6
    For k = 0 To 6
        Set Qmesh = RMC.mDrm.CreateMeshBuilder
        Qmesh.LoadFromFile SName, 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        Qmesh.Translate j * 4, k * 4, 0
        Qmesh.SetName "Space" & "2" & j + 1 & k + 1
        'QMesh.SetColorRGB 1, 0.5, 0
        ''Qmesh.SetColorRGB 1, 0.5, 0
        SetModifiers Qmesh, 2, j + 1, k + 1
        Set CFrame = RMC.mDrm.CreateFrame(Nothing)
        CFrame.AddVisual Qmesh
        CFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, -90 * pi / 180
        'CFrame.AddRotation D3DRMCOMBINE_AFTER, 0, 0, 1, 90 * pi / 180
        'CFrame.AddRotation D3DRMCOMBINE_AFTER, 0, 0, 0, -90 * pi / 180
        CFrame.SetPosition Nothing, 14, -14, -14
        QFrame.AddChild CFrame
        Set CFrame = Nothing
    Next k
Next j
'Side 3
For j = 0 To 6
    For k = 0 To 6
        Set Qmesh = RMC.mDrm.CreateMeshBuilder
        Qmesh.LoadFromFile SName, 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        Qmesh.Translate j * 4, k * 4, 0
        Qmesh.SetName "Space" & "3" & j + 1 & k + 1
        'QMesh.SetColorRGB 1, 1, 0
        ''Qmesh.SetColorRGB 1, 0.5, 0
        SetModifiers Qmesh, 3, j + 1, k + 1
        Set CFrame = RMC.mDrm.CreateFrame(Nothing)
        CFrame.AddVisual Qmesh
        CFrame.AddRotation D3DRMCOMBINE_REPLACE, 1, 0, 0, 90 * pi / 180
        CFrame.AddRotation D3DRMCOMBINE_AFTER, 0, 1, 0, -90 * pi / 180
        CFrame.SetPosition Nothing, 14, 14, -14
        QFrame.AddChild CFrame
        Set CFrame = Nothing
    Next k
Next j
'Side 4
For j = 0 To 6
    For k = 0 To 6
        Set Qmesh = RMC.mDrm.CreateMeshBuilder
        Qmesh.LoadFromFile SName, 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        Qmesh.Translate j * 4, k * 4, 0
        Qmesh.SetName "Space" & "4" & j + 1 & k + 1
        'QMesh.SetColorRGB 0, 1, 0
        ''Qmesh.SetColorRGB 1, 0.5, 0
        SetModifiers Qmesh, 4, j + 1, k + 1
        Set CFrame = RMC.mDrm.CreateFrame(Nothing)
        CFrame.AddVisual Qmesh
        CFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, 180 * pi / 180
        CFrame.AddRotation D3DRMCOMBINE_AFTER, 0, 0, 1, 90 * pi / 180
        CFrame.SetPosition Nothing, 14, 14, 14
        QFrame.AddChild CFrame
        Set CFrame = Nothing
    Next k
Next j
'Side 5
For j = 0 To 6
    For k = 0 To 6
        Set Qmesh = RMC.mDrm.CreateMeshBuilder
        Qmesh.LoadFromFile SName, 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        Qmesh.Translate j * 4, k * 4, 0
        Qmesh.SetName "Space" & "5" & j + 1 & k + 1
        'QMesh.SetColorRGB 0, 0, 1
        ''Qmesh.SetColorRGB 1, 0.5, 0
        SetModifiers Qmesh, 5, j + 1, k + 1
        Set CFrame = RMC.mDrm.CreateFrame(Nothing)
        CFrame.AddVisual Qmesh
        CFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, 90 * pi / 180
        CFrame.AddRotation D3DRMCOMBINE_AFTER, 1, 0, 0, -90 * pi / 180
        CFrame.SetPosition Nothing, -14, 14, 14
        QFrame.AddChild CFrame
        Set CFrame = Nothing
    Next k
Next j
'Side 6
For j = 0 To 6
    For k = 0 To 6
        Set Qmesh = RMC.mDrm.CreateMeshBuilder
        Qmesh.LoadFromFile SName, 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        Qmesh.Translate j * 4, k * 4, 0
        Qmesh.SetName "Space" & "6" & j + 1 & k + 1
        'QMesh.SetColorRGB 1, 0, 1
        ''Qmesh.SetColorRGB 1, 0.5, 0
        SetModifiers Qmesh, 6, j + 1, k + 1
        Set CFrame = RMC.mDrm.CreateFrame(Nothing)
        CFrame.AddVisual Qmesh
        CFrame.AddRotation D3DRMCOMBINE_REPLACE, 1, 0, 0, -90 * pi / 180
        'CFrame.AddRotation D3DRMCOMBINE_AFTER, 0, 1, 0, -90 * pi / 180
        CFrame.SetPosition Nothing, -14, -14, 14
        QFrame.AddChild CFrame
        Set CFrame = Nothing
    Next k
Next j
Skip:

'-----Inner box to reject see-through cracks
Set Qmesh = RMC.CreateBoxMesh(26, 26, 26)
Set CFrame = RMC.mDrm.CreateFrame(Nothing)
Qmesh.SetColorRGB 1, 0.5, 0
CFrame.AddVisual Qmesh

QFrame.AddChild CFrame
Set CFrame = Nothing
'--------------------
'-----------Edge-Borders------------------
Set Qmesh = RMC.mDrm.CreateMeshBuilder
Qmesh.LoadFromFile "Edges.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
Set CFrame = RMC.mDrm.CreateFrame(Nothing)
'Qmesh.SetTexture RMC.mDRM.LoadTexture("Rosewood.bmp")
CFrame.AddVisual Qmesh
CFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, 90 * pi / 180
QFrame.AddChild CFrame
'----------------------
QFrame.SetPosition Nothing, 0, 0, 55

Set RMC.mFrO = QFrame
'RotToSide 1
'Timer1.Interval = 1
Set CFrame = Nothing
End Sub
Function ChooseLetters(lCount As Integer) As String
Randomize
Dim i As Integer
Dim j As Integer
For i = 1 To 26
    j = j + Letters(i).Count
Next i
If j < lCount Then lCount = j
For i = 1 To lCount
    j = 0
    Do While j < 1 Or j > 27
        j = Abs(CInt(Rnd * 36) - 9)
    Loop
    Do While Letters(j).Count = 0
        j = CInt(Rnd * 26) + 1
    Loop
    Letters(j).Count = Letters(j).Count - 1
    ChooseLetters = ChooseLetters & Letters(j).Char
Next i
End Function
Function GetCNum(Side As Integer, Row As Integer, Col As Integer) As Long
GetCNum = (Side - 1) * 49
GetCNum = GetCNum + ((Row - 1) * 7)
GetCNum = GetCNum + Col
GetCNum = GetCNum - 1
End Function



Sub PickNewLetters()
Dim i As Integer
Dim j As Integer
Dim NewRack As String
Dim RLetters As String
For i = 1 To 7
    If Rack(i).Played Then
        j = j + 1
    Else
        If Rack(i).Blank Then Rack(i).Char = " "
        If Rack(i).Char <> "$" Then RLetters = RLetters & Rack(i).Char
    End If
Next i
ClearRack
NewRack = RLetters & ChooseLetters(j)
If NewRack = "" Then
    Me.Caption = "No More Letters"
    Exit Sub
End If
SetRack NewRack
ResetRack
End Sub

Sub ResetRack()
Dim i As Integer
For i = 1 To 7
    RFrame(i).SetRotation Nothing, 0, 0, 0, 0
    If Rack(i).Played = False Then
        RFrame(i).AddRotation D3DRMCOMBINE_REPLACE, 0, 0, 0, 0
    Else
        RFrame(i).AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, 180 * pi / 180
    End If
    RFrame(i).SetPosition Nothing, (i - 4) * 5, 0, 35
    Rack(i).Chosen = False
Next i
End Sub

Sub RotToSide(Index As Integer)
Select Case Index
    Case 1
        RMC.mFrO.AddRotation D3DRMCOMBINE_REPLACE, 0, 0, 0, 0
    Case 2
        RMC.mFrO.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, 90 * pi / 180
    Case 3
        RMC.mFrO.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, 90 * pi / 180
        RMC.mFrO.AddRotation D3DRMCOMBINE_AFTER, 1, 0, 0, -90 * pi / 180
    Case 4
        RMC.mFrO.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, 90 * pi / 180
        RMC.mFrO.AddRotation D3DRMCOMBINE_AFTER, 1, 0, 0, -90 * pi / 180
        RMC.mFrO.AddRotation D3DRMCOMBINE_AFTER, 0, 1, 0, 90 * pi / 180
    Case 5
        RMC.mFrO.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, 90 * pi / 180
        RMC.mFrO.AddRotation D3DRMCOMBINE_AFTER, 1, 0, 0, -90 * pi / 180
        RMC.mFrO.AddRotation D3DRMCOMBINE_AFTER, 0, 1, 0, 90 * pi / 180
        RMC.mFrO.AddRotation D3DRMCOMBINE_AFTER, 1, 0, 0, -90 * pi / 180
    Case 6
        RMC.mFrO.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, 90 * pi / 180
        RMC.mFrO.AddRotation D3DRMCOMBINE_AFTER, 1, 0, 0, -90 * pi / 180
        RMC.mFrO.AddRotation D3DRMCOMBINE_AFTER, 0, 1, 0, 90 * pi / 180
        RMC.mFrO.AddRotation D3DRMCOMBINE_AFTER, 1, 0, 0, -90 * pi / 180
        RMC.mFrO.AddRotation D3DRMCOMBINE_AFTER, 0, 1, 0, 90 * pi / 180
End Select
RMC.mFrO.SetPosition Nothing, 0, 0, 50
RMC.Update
End Sub

Sub SetModifiers(ByRef MyMesh As Direct3DRMMeshBuilder3, Side As Integer, Col As Integer, Row As Integer)
Dim PosWord As String * 2
PosWord = Col & Row
GoTo Solid
Select Case Side
    Case 1, 3, 5
        Select Case PosWord
            Case "14", "31", "47", "75": MyMesh.SetTexture RMC.mDrm.LoadTexture("ltblue_Pine.bmp")
            Case "17": MyMesh.SetTexture RMC.mDrm.LoadTexture("red_Pine.bmp")
            Case "22", "62", "66": MyMesh.SetTexture RMC.mDrm.LoadTexture("blue_Pine.bmp")
            Case "26", "35", "44", "53", "71": MyMesh.SetTexture RMC.mDrm.LoadTexture("pink_Pine.bmp")
            Case Else: MyMesh.SetTexture RMC.mDrm.LoadTexture("Pine.bmp")
        End Select
    Case 2, 4, 6
        Select Case PosWord
            Case "74", "57", "41", "13": MyMesh.SetTexture RMC.mDrm.LoadTexture("ltblue_Pine.bmp")
            Case "71": MyMesh.SetTexture RMC.mDrm.LoadTexture("red_Pine.bmp")
            Case "66", "26", "22": MyMesh.SetTexture RMC.mDrm.LoadTexture("blue_Pine.bmp")
            Case "62", "53", "44", "35", "17": MyMesh.SetTexture RMC.mDrm.LoadTexture("pink_Pine.bmp")
            Case Else: MyMesh.SetTexture RMC.mDrm.LoadTexture("Pine.bmp")
        End Select
End Select
MyMesh.SetQuality D3DRMRENDER_PHONG
Exit Sub
'-------SOLID--------------
Solid:
Select Case Side
    Case 1, 3, 5
        Select Case PosWord
            Case "14", "31", "47", "75": MyMesh.SetColorRGB 0.75, 1, 1
            Case "17": MyMesh.SetColorRGB 1, 0, 0
            Case "22", "62", "66": MyMesh.SetColorRGB 0, 0, 1
            Case "26", "35", "44", "53", "71": MyMesh.SetColorRGB 1, 0.75, 0.75
            Case Else: MyMesh.SetColorRGB 1, 0.5, 0
        End Select
    Case 2, 4, 6
        Select Case PosWord
            Case "74", "57", "41", "13": MyMesh.SetColorRGB 0.75, 1, 1
            Case "71": MyMesh.SetColorRGB 1, 0, 0
            Case "66", "26", "22": MyMesh.SetColorRGB 0, 0, 1
            Case "62", "53", "44", "35", "17": MyMesh.SetColorRGB 1, 0.75, 0.75
            Case Else: MyMesh.SetColorRGB 1, 0.5, 0
        End Select
End Select

End Sub

Sub SetRack(Tiles As String)
Dim i As Integer
Dim Blank As Boolean
For i = 1 To 7
    Rack(i) = NoTile
    Rack(i).Char = "$"
Next i
For i = 1 To 7
    'Set RFrame(i) = RMC.mDrm.CreateFrame(rmcRack.mFrO)
    Set Qmesh = RMC.CreateBoxMesh(3.8, 3.8, 1)
    If Mid(Tiles, i, 1) = " " Then Blank = True Else Blank = False
    Qmesh.SetTexture rmcRack.mDrm.LoadTexture("pine.bmp")
    SetTileChar Qmesh, Mid(Tiles, i, 1), Blank
    RFrame(i).AddVisual Qmesh
    Qmesh.SetName i
    RFrame(i).SetPosition Nothing, (i - 4) * 5, 0, 35
    Rack(i).Char = Mid(Tiles, i, 1)
    Rack(i).Blank = Blank
    Rack(i).Chosen = False
    Rack(i).Played = False
    Rack(i).Pos = i
Next i

SelTile.Chosen = False
End Sub
Sub SetTile(Side As Integer, Col As Integer, Row As Integer, Char As String, Blank As Boolean)
Dim CFrame As Direct3DRMFrame3
Set CFrame = QFrame.GetChildren.GetElement(GetCNum(Side, Row, Col)).CloneObject
Set Qmesh = RMC.CreateBoxMesh(3.8, 3.8, 1)
Qmesh.Translate (Col - 0.5) * 4, (Row - 0.5) * 4, 0
Qmesh.SetName "Piece" & Side & Col & Row
Qmesh.SetTexture RMC.mDrm.LoadTexture("Pine.bmp")
SetTileChar Qmesh, Char, Blank
CFrame.AddVisual Qmesh
QFrame.AddChild CFrame
RMC.Update
End Sub

Sub GetPiece()
Set PMesh = RMC.mDrm.CreateMeshBuilder
Set PFrame = RMC.mDrm.CreateFrame(RMC.mFrO)
PMesh.LoadFromFile "Piece.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
PMesh.SetTexture RMC.mDrm.LoadTexture("Rosewood.jpg")
PFrame.AddVisual PMesh
PFrame.SetPosition QFrame, 0, 0, -20
RMC.Update
End Sub
Sub SetTileChar(ByRef TileMesh As Direct3DRMMeshBuilder3, Char As String, Blank As Boolean)
Set QTex = RMC.CreateUpdateableTexture(120, 120, "")
Set QSurf = QTex.GetSurface(0)
Dim R(1) As RECT
Me.FontSize = 90
QSurf.SetFont Me.Font
QSurf.setDrawWidth 1
QSurf.SetForeColor CLng(&H4A9AC6) 'vbBlack
QSurf.SetFillColor CLng(&H4A9AC6)  'CLng(&H102478)
QSurf.SetFillStyle 0
QSurf.SetFontTransparency True
If Blank Then
    QSurf.DrawCircle 60, 60, 60
    QSurf.SetForeColor RGB(0.75, 0.25, 0.25)
    QSurf.DrawText 25, -10, Char, False
Else
    QSurf.DrawBox 0, 0, 120, 120
    QSurf.SetFillColor CLng(&H4A9AC6) 'vbBlack
    'QSurf.DrawBox 85, 85, 120, 120
    QSurf.SetForeColor RGB(0.75, 0.25, 0.25)
    QSurf.DrawText 25, -10, Char, False
    Me.FontSize = 24
    QSurf.SetFont Me.Font
    QSurf.SetForeColor vbYellow
    QSurf.DrawText 85, 85, GetPoints(Char), False
End If
QTex.Changed D3DRMTEXTURE_CHANGEDPIXELS, 0, R()
TileMesh.SetTextureCoordinates 1, 0, 0
TileMesh.SetTextureCoordinates 2, 1, 0
TileMesh.SetTextureCoordinates 3, 1, 1
TileMesh.SetTextureCoordinates 0, 0, 1
TileMesh.GetFace(0).SetTexture QTex
Set QTex = Nothing
Set QSurf = Nothing
'TileMesh.GetFace(1).SetColorRGB 0.25, 0.5, 0.75
'TileMesh.GetFace(2).SetColorRGB 0.25, 0.5, 0.75
'TileMesh.GetFace(3).SetColorRGB 0.25, 0.5, 0.75
'TileMesh.GetFace(4).SetColorRGB 0.25, 0.5, 0.75
'TileMesh.GetFace(5).SetColorRGB 0.25, 0.5, 0.75
'-----End texture---------
End Sub




Sub SlideTiles(PosA As Integer, PosB As Integer)

Dim i As Integer
Dim A As Integer
Dim B As Integer
If PosA < PosB Then
    A = PosA
    B = PosB
Else
    A = PosB
    B = PosA
End If
SwapRackTiles A, B
For i = A + 1 To B - 1
    SwapRackTiles i, B
Next i
End Sub

Sub SwapRackTiles(PosA As Integer, PosB As Integer)
Dim TMP As Tile
TMP = Rack(PosA)
Rack(PosA) = Rack(PosB)
Rack(PosB) = TMP
Rack(PosA).Pos = PosA
Rack(PosB).Pos = PosB
End Sub

Private Sub cmdGoSide_Click(Index As Integer)
RotToSide Index + 1
End Sub







Private Sub Command1_Click()
Dim NewRack As String
Dim OldRack As String
BadMove "Exchanging Letters"
Dim i As Integer
For i = 1 To 7
    OldRack = OldRack & Rack(i).Char
Next i
NewRack = ChooseLetters(7)
ClearRack
SetRack NewRack
ReturnLetters OldRack
End Sub

Private Sub Command2_Click()
Dim cWord As String
Dim i As Integer
Dim PTS As Integer
Dim pCount As Integer
For i = 1 To 7
    If Rack(i).Played Then pCount = pCount + 1
Next i
If pCount = 0 Then Exit Sub
cWord = GoodWordPlacement(dRow, Rack())
If cWord = "" Then
    BadMove "Illegal Placement"
ElseIf CheckNewWords(NewWords()) = False Then
    BadMove "Illegal Word"
Else
    SetWord
    For i = 1 To UBound(NewWords)
        List1.AddItem NewWords(i).Word
    Next i
    List1.Selected(List1.ListCount - 1) = True
    PTS = ScoreWords(NewWords) + Bonus(Rack())
    txtScore = txtScore + PTS
    Me.Caption = "Good Play, " & PTS & " Points"
    PickNewLetters
End If

End Sub

Private Sub Command3_Click()
BadMove "Clear Letters"
End Sub



Private Sub Form_DblClick()
Me.WindowState = vbMinimized
End Sub

Private Sub Form_Initialize()
RMC.InitDx picMain
rmcRack.InitDx picRack
RMC.Resize picMain
rmcRack.Resize picRack
Timer1.Interval = 1
ClearQCube
BuildCube
Dim i As Integer
For i = 1 To 7
    Set RFrame(i) = RMC.mDrm.CreateFrame(rmcRack.mFrO)
Next i
SetRack ChooseLetters(7)
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
StartMove Me
End Sub


Private Sub Form_Resize()
On Local Error Resume Next
picMain.Width = Me.ScaleWidth - picRight.ScaleWidth

End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'CHEATER!!!!
'If Shift = 3 Then
'    FOrceRack UCase(InputBox("Enter Rack"))
'    Exit Sub
'End If
On Error Resume Next
RMC.MouseDown X, Y
Randomize
Dim PickArray As Direct3DRMPickArray
Dim Desc As D3DRMPICKDESC
Dim SName As String
Dim Side As Integer
Dim Col As Integer
Dim Row As Integer
SName = RMC.Pick(X, Y)
If Left(SName, 5) = "Space" Then
    SName = Right(SName, 3)
    Side = Mid(SName, 1, 1)
    Col = Mid(SName, 2, 1)
    Row = Mid(SName, 3, 1)
    If SelTile.Chosen = True Then
        QCube(Side, Col, Row).Piece = SelTile
        If TileSetOk(Side, Col, Row) Then
            SetTile Side, Col, Row, SelTile.Char, SelTile.Blank
            Rack(SelTile.Pos).Played = True
            SelTile.Played = True
            ResetRack
            Set rmcRack.mFrO = Nothing
            SelTile.Chosen = False
            QCube(Side, Col, Row).Piece = SelTile
            QCube(Side, Col, Row).Status = 1 'testing
        Else
            Me.Caption = "Bad Move"
            QCube(Side, Col, Row).Piece = NoTile
        End If
    End If
End If

End Sub





Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RMC.MouseMove X, Y
End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RMC.MouseUp
End Sub


Private Sub picMain_Resize()
RMC.Resize picMain
End Sub

Private Sub picRack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'rmcRack.MouseMove X, Y
End Sub

Private Sub picRack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'rmcRack.MouseUp
End Sub


Private Sub picRack_Resize()
rmcRack.Resize picRack
End Sub

Private Sub picRight_Resize()
On Error Resume Next
List1.Height = picRight.ScaleHeight - picMenu.ScaleHeight - txtScore.Height
picMenu.Top = List1.Top + List1.Height
End Sub


Private Sub Timer1_Timer()
rmcRack.Update
RMC.Update
DoEvents
End Sub


Private Sub picRack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ExitMe
Dim rPos As Integer
Dim i As Integer
Dim R As String
rPos = CInt(rmcRack.Pick(X, Y))
If SelTile.Chosen Then
    If rPos <> SelTile.Pos Then
        SlideTiles SelTile.Pos, rPos
        RefreshRack
        ResetRack
        SelTile.Chosen = False
        Exit Sub
    Else
        ResetRack
        SelTile.Chosen = False
        Exit Sub
    End If
End If
ResetRack
'rmcRack.MouseDown X, Y
If Rack(rPos).Played = False Then
    If Rack(rPos).Blank Then
        R = InputBox("What would you like this to be?", "You Have A BLANK")
        R = UCase(R)
        If R = "" Then Exit Sub
        Rack(rPos).Char = R
    End If
    'RFrame(rPos).SetRotation Nothing, 0, 1, 1, 0.05
    Set rmcRack.mFrO = RFrame(rPos)
    rmcRack.mFrO.SetRotation Nothing, 0, 1, 1, 0.15
    'rmcRack.mFrO.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 0.5
    Rack(rPos).Chosen = True
    SelTile = Rack(rPos)
End If
ExitMe:
End Sub


