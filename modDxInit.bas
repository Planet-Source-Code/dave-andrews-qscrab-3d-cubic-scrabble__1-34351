Attribute VB_Name = "modDxInit"
Option Explicit

Public mdownX As Single
Public mDownY As Single
Public mMouseDown As Boolean

Public Type dxPTM
    dX As Single
    dY As Single
    Distance As Single
End Type

Public mDX7 As DirectX7
Public mDrm As Direct3DRM3
Public mDrw As DirectDraw7
Public mFrs As Direct3DRMFrame3
Public mFrC As Direct3DRMFrame3
Public mFrO As Direct3DRMFrame3
Public mFrL As Direct3DRMFrame3
Public mDev As Direct3DRMDevice3
Public mVpt As Direct3DRMViewport2
Public DxL1 As Direct3DRMLight
Public DxL2 As Direct3DRMLight
Public DXClipper As DirectDrawClipper
Public DxMeshB As Direct3DRMMeshBuilder3
Sub InitDx(Canvas As Object)
Set mDX7 = New DirectX7
Set mDrm = mDX7.Direct3DRMCreate
Set mDrw = mDX7.DirectDrawCreate("")
'-Set Up Camera Frames And Lights-
Set mFrs = mDrm.CreateFrame(Nothing)
Set mFrC = mDrm.CreateFrame(mFrs)
Set mFrO = mDrm.CreateFrame(mFrs)
Set mFrL = mDrm.CreateFrame(mFrs)
Set DxL1 = mDrm.CreateLightRGB(D3DRMLIGHT_DIRECTIONAL, 0.8, 0.8, 0.8)
Set DxL2 = mDrm.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.8, 0.8, 0.8)
mFrL.AddLight DxL1
mFrL.AddLight DxL2
'-------Create Display--------
Set mVpt = Nothing
Set mDev = Nothing
Set DXClipper = mDrw.CreateClipper(0)
Canvas.ScaleMode = vbPixels
DXClipper.SetHWnd Canvas.hWnd
Set mDev = mDrm.CreateDeviceFromClipper(DXClipper, "", 500, 500)
Set mVpt = mDrm.CreateViewport(mDev, mFrC, 0, 0, 500, 500)
End Sub
Sub LoadMesh()
Set DxMeshB = mDrm.CreateMeshBuilder()
mDrm.SetSearchPath App.Path
DxMeshB.LoadFromFile "Piece.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
DxMeshB.SetTexture mDrm.LoadTexture("Pine.jpg")
mFrO.AddVisual DxMeshB
mFrO.SetPosition Nothing, 0, 0, 10
End Sub


Sub PointToMouse(PTM As dxPTM, X As Single, Y As Single)
Dim sX As Single
Dim sY As Single
With PTM
    .dX = mdownX - X
    .dY = mDownY - Y
    sX = (.dX * .dX)
    sY = (.dY * .dY)
    .Distance = Sqr(sX + sY)
End With
End Sub

Sub Rotate(X As Single, Y As Single)
Dim PTM As dxPTM
Dim Theta As Single
PointToMouse PTM, X, Y
Theta = PTM.Distance / 1000
mFrO.SetRotation Nothing, PTM.dY, PTM.dX, 0, Theta
End Sub

