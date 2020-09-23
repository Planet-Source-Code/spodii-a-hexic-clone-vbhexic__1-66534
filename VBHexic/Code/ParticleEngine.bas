Attribute VB_Name = "ParticleEngine"
Option Explicit

Private Type Effect
    X As Single 'X location of effect
    Y As Single 'Y location of effect
    Gfx As Byte 'Particle texture used
    R As Single
    G As Single
    B As Single
    Used As Boolean         'If the effect is in use
    EffectNum As Byte       'What number of effect that is used
    Modifier As Integer
    FloatSize As Long
    Direction As Integer
    Particles() As Particle 'Information on each particle
    Progression As Single
    PartVertex() As TLVERTEX    'Used for point render particles
    PreviousFrame As Long
    ParticleCount As Integer    'Number of particles total
    ParticlesLeft As Integer    'Number of particles left - Only for non-repetitive effects
End Type
Public Effect() As Effect   'List of all the active effects

'Constants with the order number for each effect
Public Const Glitter_Num As Byte = 1    'Explosion of glitter

Private Function Effect_FToDW(f As Single) As Long
Dim Buf As D3DXBuffer
Dim TempVal As Long

    'Cant say what this does since this is straight from Almar's code
    Set Buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData Buf, 0, 4, 1, f
    D3DX.BufferGetData Buf, 0, 4, 1, TempVal
    Effect_FToDW = TempVal

End Function

Public Sub Effect_Glitter_Begin(ByVal X As Single, ByVal Y As Single, ByVal Particles As Integer, ByVal Gfx As Byte, ByVal R As Single, ByVal G As Single, ByVal B As Single)
Dim EffectIndex As Byte
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = Glitter_Num         'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).R = R
    Effect(EffectIndex).G = G
    Effect(EffectIndex).B = B

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Glitter_Reset EffectIndex, LoopC
    Next LoopC

    'Set the initial time
    Effect(EffectIndex).PreviousFrame = CurrTime

End Sub

Private Sub Effect_Glitter_Reset(ByVal EffectIndex As Byte, ByVal Index As Long)
Dim TempDegree As Single

    'Set the temporary degree
    TempDegree = Rnd * 360

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X, Effect(EffectIndex).Y, Sin(TempDegree * DegreeToRadian) * (Rnd * 3.5 + 3.5), Cos(TempDegree * DegreeToRadian) * (Rnd * 3.5 + 3.5), 0, 1.5
    Effect(EffectIndex).Particles(Index).ResetColor Effect(EffectIndex).R, Effect(EffectIndex).G, Effect(EffectIndex).B, 1 + (Rnd * 0.2), 0.06 + (Rnd * 0.06)

End Sub

Private Sub Effect_Glitter_Update(ByVal EffectIndex As Byte)
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (CurrTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = CurrTime

    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used = True Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is out of the screen
            If Effect(EffectIndex).Particles(LoopC).sngY - 32 > frmMain.ScaleHeight Then Effect(EffectIndex).Particles(LoopC).sngA = 0

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Disable the particle
                Effect(EffectIndex).Particles(LoopC).Used = False

                'Subtract from the total particle count
                Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                'Check if the effect is out of particles
                If Effect(EffectIndex).ParticlesLeft <= 0 Then Effect(EffectIndex).Used = False

            Else

                'Set The Particle Information On The Particle Vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Sub Effect_Kill(ByVal EffectIndex As Byte, Optional ByVal KillAll As Boolean = False)
Dim LoopC As Long

    'Check if to kill all effects
    If KillAll = True Then

        'Loop through every effect
        For LoopC = 1 To NumEffects

            'Stop the effect
            Effect(LoopC).Used = False

        Next
    Else

        'Stop the selected effect
        Effect(EffectIndex).Used = False
    End If

End Sub

Private Function Effect_NextOpenSlot() As Byte
Dim EffectIndex As Byte

    'Find the next open effect slot
    Do
        EffectIndex = EffectIndex + 1   'Check the next slot
        If EffectIndex > NumEffects Then Exit Function 'Dont go over maximum amount
    Loop While Effect(EffectIndex).Used = True    'Check next if effect is in use

    'Return the next open slot
    Effect_NextOpenSlot = EffectIndex

End Function

Public Sub Effect_Render(ByVal EffectIndex As Byte)

    'Set the render state to point blitting
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Effect(EffectIndex).FloatSize
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    'Set the texture
    D3DDevice.SetTexture 0, TexParticle(Effect(EffectIndex).Gfx)

    'Draw all the particles at once
    D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, Effect(EffectIndex).ParticleCount, Effect(EffectIndex).PartVertex(0), Len(Effect(EffectIndex).PartVertex(0))

    'Reset the render state back to normal
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

End Sub

Sub Effect_UpdateAll()
Dim LoopC As Long

    'Update every effect in use
    For LoopC = 1 To NumEffects

        'Make sure the effect is in use
        If Effect(LoopC).Used = True Then

            'Find out which effect is selected, then update it
            If Effect(LoopC).EffectNum = Glitter_Num Then Effect_Glitter_Update LoopC

            'Render the effect
            Effect_Render LoopC

        End If

    Next

End Sub
