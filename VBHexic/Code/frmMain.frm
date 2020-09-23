VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "VBHexic"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   413
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LastBulletPos As Point

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Set the mouse position
    MousePos.X = X * (413 / Me.ScaleWidth)
    MousePos.Y = Y * (435 / Me.ScaleHeight)
    
    'Update the bullet position
    BulletPos = Game_Bullet_GetPos
    
    'Check if the position changed
    If BulletPos.X <> LastBulletPos.X Or BulletPos.Y <> LastBulletPos.Y Then
        Engine_Sfx_Play SfxClick
        LastBulletPos = BulletPos
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Check if a rotate is already in progress
    If RotateDir = 0 Then
        If Not Game_HexAreMoving Then
            
            'Rotate
            If Button = vbLeftButton Then
                RotateCount = 0
                RotateDir = 1
                RotateHex1 = SelectedHex1
                RotateHex2 = SelectedHex2
                RotateHex3 = SelectedHex3
            ElseIf Button = vbRightButton Then
                RotateCount = 0
                RotateDir = 2
                RotateHex1 = SelectedHex1
                RotateHex2 = SelectedHex2
                RotateHex3 = SelectedHex3
            End If
    
        End If
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Unload the engine
    EndGameLoop = True

End Sub
