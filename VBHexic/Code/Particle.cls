VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Particle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Values Stored For Each Particle
Private mvarused As Boolean
Private mvarsngX As Single
Private mvarsngY As Single
Private mvarsngXSpeed As Single
Private mvarsngYSpeed As Single
Private mvarsngXAccel As Single
Private mvarsngYAccel As Single
Private mvarsngR As Single
Private mvarsngG As Single
Private mvarsngB As Single
Private mvarsngA As Single
Private mvarsngAlphaDecay As Single

Public Sub ResetColor(sngRed As Single, sngGreen As Single, sngBlue As Single, sngAlpha As Single, sngDecay As Single)

'Change The Particle Color Values

    sngR = sngRed
    sngG = sngGreen
    sngB = sngBlue
    sngA = sngAlpha
    sngAlphaDecay = sngDecay

End Sub

Public Sub ResetIt(X As Single, Y As Single, XSpeed As Single, YSpeed As Single, XAcc As Single, YAcc As Single)

'Reset The Particle's Variables

    sngX = X
    sngY = Y
    sngXSpeed = XSpeed
    sngYSpeed = YSpeed
    sngXAccel = XAcc
    sngYAccel = YAcc

End Sub

Public Property Get sngA() As Single

'Release The Particle Alpha Value

    sngA = mvarsngA

End Property

Public Property Let sngA(ByVal vData As Single)

'Set The Particle Alpha Value

    mvarsngA = vData

End Property

Public Property Let sngAlphaDecay(ByVal vData As Single)

'Set The Alpha Decay Value

    mvarsngAlphaDecay = vData

End Property

Public Property Get sngAlphaDecay() As Single

'Release The Particle Alpha Decay Value

    sngAlphaDecay = mvarsngAlphaDecay

End Property

Public Property Get sngB() As Single

'Release The Particle Blue Color

    sngB = mvarsngB

End Property

Public Property Let sngB(ByVal vData As Single)

'Set The Particle Blue Color

    mvarsngB = vData

End Property

Public Property Let sngG(ByVal vData As Single)

'Set The Particle Green Color

    mvarsngG = vData

End Property

Public Property Get sngG() As Single

'Release The Particle Green Color

    sngG = mvarsngG

End Property

Public Property Get sngR() As Single

'Release The Particle Red Color

    sngR = mvarsngR

End Property

Public Property Let sngR(ByVal vData As Single)

'Set The Particle Red Color

    mvarsngR = vData

End Property

Public Property Get sngX() As Single

'Release The X Position

    sngX = mvarsngX

End Property

Public Property Let sngX(ByVal vData As Single)

'Set The X Position

    mvarsngX = vData

End Property

Public Property Get sngXAccel() As Single

'Release The Particle X Acceleration

    sngXAccel = mvarsngXAccel

End Property

Public Property Let sngXAccel(ByVal vData As Single)

'Set The Particle X Acceleration

    mvarsngXAccel = vData

End Property

Public Property Get sngXSpeed() As Single

'Release The Particle X Speed

    sngXSpeed = mvarsngXSpeed

End Property

Public Property Let sngXSpeed(ByVal vData As Single)

'Set The Particle X Speed

    mvarsngXSpeed = vData

End Property

Public Property Get sngY() As Single

'Release The Y Position

    sngY = mvarsngY

End Property

Public Property Let sngY(ByVal vData As Single)

'Set The Y Position

    mvarsngY = vData

End Property

Public Property Let sngYAccel(ByVal vData As Single)

'Set The Particle Y Acceleration

    mvarsngYAccel = vData

End Property

Public Property Get sngYAccel() As Single

'Release The Particle Y Acceleration

    sngYAccel = mvarsngYAccel

End Property

Public Property Get sngYSpeed() As Single

'Release The Particle Y Speed

    sngYSpeed = mvarsngYSpeed

End Property

Public Property Let sngYSpeed(ByVal vData As Single)

'Set The Particle Y Speed

    mvarsngYSpeed = vData

End Property

Public Sub UpdateParticle(sngTime As Single)

'Update The Particle's Variables

    sngX = sngX + sngXSpeed * sngTime
    sngY = sngY + sngYSpeed * sngTime

    sngXSpeed = sngXSpeed + sngXAccel * sngTime
    sngYSpeed = sngYSpeed + sngYAccel * sngTime

    sngA = sngA - sngAlphaDecay * sngTime

End Sub

Public Property Let Used(ByVal vData As Boolean)

'Set The Used Value

    mvarused = vData

End Property

Public Property Get Used() As Boolean

'Release The Used Value

    Used = mvarused

End Property
