Attribute VB_Name = "Engine"
Option Explicit

'DirectX 8 Objects
Private DX As DirectX8
Private DS As DirectSound8
Private D3D As Direct3D8
Private DSEnum As DirectSoundEnum8
Public D3DDevice As Direct3DDevice8
Public D3DX As D3DX8

'Describes a transformable lit vertex
Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    Rhw As Single
    Color As Long
    Specular As Long
    Tu As Single
    Tv As Single
End Type

'List of textures loaded from bitmaps - each texture holds one .bmp file
Private TexHexagon As Direct3DTexture8      'Hexagon
Private TexBullet As Direct3DTexture8       'Bullet between selected hexagons
Private TexNumber(0 To 9) As Direct3DTexture8   'Numbers 0 to 9
Private TexStar As Direct3DTexture8         'Hexagon star
Private Const NumParticles As Byte = 1          'Number of particle textures
Public TexParticle(1 To NumParticles) As Direct3DTexture8   'Particle textures

'List of sound buffers - each buffer holds one .wav file
Private SfxRotate As DirectSoundSecondaryBuffer8
Public SfxClick As DirectSoundSecondaryBuffer8
Private SfxPop As DirectSoundSecondaryBuffer8

'Different vertexes - they are stored individually for ease of use and slight speed
Private PointVertex(0 To 3) As TLVERTEX 'Points vertex
Private HexVertex(0 To 3) As TLVERTEX   'Hexagon vertex
Private BulVertex(0 To 3) As TLVERTEX   'Bullet vertex
Private StarVertex(0 To 3) As TLVERTEX  'Star vertex

'Normal hexagon position table - get the position of a hexagon by it's 2D array
Private HexPoint(1 To HexFieldWidth, 1 To HexFieldHeight) As Point

'Center hexagon position table - get the center of a hexagon by it's 2D array
Private HexCenter(1 To HexFieldWidth, 1 To HexFieldHeight) As Point

'Linelist - used for debugging
Private LineListA(0 To 1) As TLVERTEX
Private LineListB(0 To 1) As TLVERTEX
Private LineListC(0 To 1) As TLVERTEX

'Frames Per Second calculation variables
Private FPSLastSecond As Long
Private LastCheck As Long
Private Elapsed As Long
Private FPS As Long
Private FPSCounter As Long
Public CurrTime As Long

'If value = 1, then do collision check routine
Private CheckForCollision As Byte

'Multiply a degree by this value to get the radian value: RadianVal = DegreeVal * (Pi / 180)
Public Const DegreeToRadian As Single = 0.0174532925

'Multiply a radian by this value to get the degree value: DegreeVal = RadianVal * (180 / Pi)
Public Const RadianToDegree As Single = 57.2957795

'Describes the return from a texture init
Private Type D3DXIMAGE_INFO_A
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type

'Rotation count (to make the pieces constantly rotate until a full cycle is made or a combo is made)
Private RotateDelay As Long
Public RotateCount As Byte
Public RotateDir As Byte    '1 for clockwise, 2 for counter-clockwise
Public RotateHex1 As Point
Public RotateHex2 As Point
Public RotateHex3 As Point

'Background information
Private BackLastChangeCount As Long
Private BackRotateSpeed As Single
Private BackRotation As Single
Private BackXOffset As Single
Private BackYOffset As Single
Private BackXSpeed As Single
Private BackYSpeed As Single

'Used to get the number of computer ticks in miliseconds
'Use timeGetTime since CurrTime does not display elapsed time too well (low resolution)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

Public Sub Draw_Background()

'************************************************************
'Draw the background to the game - called in Draw_Stage
'************************************************************
Dim X As Single
Dim Y As Single

    'Check if a new speed change is needed
    If CurrTime - 2500 > BackLastChangeCount Then   'Change every 2.5 seconds
        BackXSpeed = (0.05 + Rnd * 0.2)
        BackYSpeed = -(0.05 + Rnd * 0.1)
        BackRotateSpeed = 0.1 + (Rnd * 0.1)
        BackLastChangeCount = CurrTime
    End If

    'Move the offsets
    BackXOffset = BackXOffset + BackXSpeed * Elapsed
    BackYOffset = BackYOffset + BackYSpeed * Elapsed

    'Check if to lower the offset values
    If BackXOffset > HexWidth + 6 Then BackXOffset = BackXOffset Mod (HexWidth + 6)
    If BackYOffset < HexHeight + 6 Then BackYOffset = BackYOffset Mod (HexHeight + 6)

    'Rotate
    BackRotation = BackRotation + BackRotateSpeed * Elapsed
    If BackRotation > 360 Then BackRotation = BackRotation Mod 360

    'Set the texture
    D3DDevice.SetTexture 0, TexHexagon

    'Loop through each hexagon to be on the background
    For X = -(HexWidth * 2) To frmMain.ScaleWidth + HexWidth Step HexWidth + 6
        For Y = -(HexHeight * 2) To frmMain.ScaleHeight + HexHeight Step HexHeight + 6

            'Draw the hexagon
            Engine_Render_Rectangle X + BackXOffset, Y + BackYOffset, HexWidth, HexHeight, BackRotation, HexVertex, , 1677721855, 1677721855, 1677721855, 1677721855

        Next Y
    Next X

End Sub

Public Sub Draw_Stage()

'************************************************************
'Draw the game screen and do many calculations in the process
'************************************************************

Static BulletDegree As Single       'Primary bullet degree
Dim LineList(0 To 1) As TLVERTEX
Dim DrawSelectedHex As Boolean
Dim TempColor As Long
Dim X As Single
Dim Y As Single

    'Set the linelist
    LineList(0).Color = D3DColorARGB(255, 200, 200, 0)
    LineList(0).Rhw = 1
    LineList(0).Specular = 0
    LineList(0).Z = 0
    LineList(1) = LineList(0)

    'Clear the device with black
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, -1, 1#, 0

    'Start the device rendering
    D3DDevice.BeginScene

    '***** Draw the background *****
    Draw_Background

    '***** Draw the hexagons *****

    'Loop through all hexagons in the hexagon field
    For X = 1 To HexFieldWidth
        For Y = 1 To HexFieldHeight

            'Check if a hexagon exists at the point
            If HexPiece(X, Y).Color <> 0 Then

                'Update the hexagon coordinates if needed
                If Engine_MoveToPoint(HexPiece(X, Y).X, HexPiece(X, Y).Y, HexPiece(X, Y).TargetX, HexPiece(X, Y).TargetY, HexPieceMoveSpeed * Elapsed, HexPiece(X, Y).MoveRad) = 1 Then
                    CheckForCollision = 1
                    Engine_Sfx_Play SfxClick
                End If

                'Reset selected hex value
                DrawSelectedHex = False

                'Check if the hexagon is selected
                If X = SelectedHex1.X Then
                    If Y = SelectedHex1.Y Then DrawSelectedHex = True
                End If
                If X = SelectedHex2.X Then
                    If Y = SelectedHex2.Y Then DrawSelectedHex = True
                End If
                If X = SelectedHex3.X Then
                    If Y = SelectedHex3.Y Then DrawSelectedHex = True
                End If

                'Do automatic rotation
                If HexPiece(X, Y).Rotate = True Then
                    HexPiece(X, Y).Degree = HexPiece(X, Y).Degree + Elapsed * HexPieceRotationSpeed
                    If HexPiece(X, Y).Degree >= 360 Then
                        HexPiece(X, Y).Rotate = False
                        HexPiece(X, Y).Degree = 0
                    End If
                End If

                'Increase/decrease magnification value
                If DrawSelectedHex = False Then
                    HexPiece(X, Y).Magnification = HexPiece(X, Y).Magnification - (Elapsed * MagnificationRate)
                    If HexPiece(X, Y).Magnification < 0 Then HexPiece(X, Y).Magnification = 0
                Else
                    HexPiece(X, Y).Magnification = HexPiece(X, Y).Magnification + (Elapsed * MagnificationRate)
                    If HexPiece(X, Y).Magnification > MagnificationMax Then HexPiece(X, Y).Magnification = MagnificationMax
                End If
                
                'Update shrink value
                If HexPiece(X, Y).IsShrink Then
                    HexPiece(X, Y).Shrink = HexPiece(X, Y).Shrink - (Elapsed * ShrinkSpeed)
                    HexPiece(X, Y).Magnification = HexPiece(X, Y).Shrink
                    If HexPiece(X, Y).Shrink <= -30 Then Game_HexPiece_Remove X, Y
                End If

                'Draw the hexagon if not selected
                If Not DrawSelectedHex Then
                    Engine_Render_Rectangle HexPiece(X, Y).X - (HexPiece(X, Y).Magnification * 0.5), HexPiece(X, Y).Y - (HexPiece(X, Y).Magnification * 0.5) - Y, HexWidth + HexPiece(X, Y).Magnification, HexHeight + HexPiece(X, Y).Magnification, HexPiece(X, Y).Degree, HexVertex, TexHexagon, HexColor(HexPiece(X, Y).Color), HexColor(HexPiece(X, Y).Color), HexColor(HexPiece(X, Y).Color), HexColor(HexPiece(X, Y).Color)
                    If HexPiece(X, Y).Star = 1 Then Engine_Render_Rectangle HexPiece(X, Y).X - (HexPiece(X, Y).Magnification * 0.5), HexPiece(X, Y).Y - (HexPiece(X, Y).Magnification * 0.5) - Y, HexWidth + HexPiece(X, Y).Magnification, HexHeight + HexPiece(X, Y).Magnification, HexPiece(X, Y).Degree, HexVertex, TexStar, D3DColorARGB(255, 255, 255, 0), D3DColorARGB(255, 255, 255, 0), D3DColorARGB(255, 255, 255, 0), D3DColorARGB(255, 255, 255, 0)
                End If
                
            End If

        Next Y
    Next X
    
    'Now that all unselected have been drawn, draw selected so the magnification doesn't draw under unselected tiles
    If SelectedHex1.X > 0 And SelectedHex1.Y > 0 Then
        If HexPiece(SelectedHex1.X, SelectedHex1.Y).Color <> 0 Then
            Engine_Render_Rectangle HexPiece(SelectedHex1.X, SelectedHex1.Y).X - (HexPiece(SelectedHex1.X, SelectedHex1.Y).Magnification * 0.5), HexPiece(SelectedHex1.X, SelectedHex1.Y).Y - (HexPiece(SelectedHex1.X, SelectedHex1.Y).Magnification * 0.5) - SelectedHex1.Y, HexWidth + HexPiece(SelectedHex1.X, SelectedHex1.Y).Magnification, HexHeight + HexPiece(SelectedHex1.X, SelectedHex1.Y).Magnification, HexPiece(SelectedHex1.X, SelectedHex1.Y).Degree, HexVertex, TexHexagon, HexColor(HexPiece(SelectedHex1.X, SelectedHex1.Y).Color), HexColor(HexPiece(SelectedHex1.X, SelectedHex1.Y).Color), HexColor(HexPiece(SelectedHex1.X, SelectedHex1.Y).Color), HexColor(HexPiece(SelectedHex1.X, SelectedHex1.Y).Color)
            If HexPiece(SelectedHex1.X, SelectedHex1.Y).Star = 1 Then Engine_Render_Rectangle HexPiece(SelectedHex1.X, SelectedHex1.Y).X - (HexPiece(SelectedHex1.X, SelectedHex1.Y).Magnification * 0.5), HexPiece(SelectedHex1.X, SelectedHex1.Y).Y + (HexPiece(SelectedHex1.X, SelectedHex1.Y).Magnification * 0.5) - Y, HexWidth + HexPiece(SelectedHex1.X, SelectedHex1.Y).Magnification, HexHeight + HexPiece(SelectedHex1.X, SelectedHex1.Y).Magnification, HexPiece(SelectedHex1.X, SelectedHex1.Y).Degree, HexVertex, TexStar, D3DColorARGB(255, 255, 255, 0), D3DColorARGB(255, 255, 255, 0), D3DColorARGB(255, 255, 255, 0), D3DColorARGB(255, 255, 255, 0)
        End If
    End If
    If SelectedHex2.X > 0 And SelectedHex2.Y > 0 Then
        If HexPiece(SelectedHex2.X, SelectedHex2.Y).Color <> 0 Then
            Engine_Render_Rectangle HexPiece(SelectedHex2.X, SelectedHex2.Y).X - (HexPiece(SelectedHex2.X, SelectedHex2.Y).Magnification * 0.5), HexPiece(SelectedHex2.X, SelectedHex2.Y).Y - (HexPiece(SelectedHex2.X, SelectedHex2.Y).Magnification * 0.5) - SelectedHex2.Y, HexWidth + HexPiece(SelectedHex2.X, SelectedHex2.Y).Magnification, HexHeight + HexPiece(SelectedHex2.X, SelectedHex2.Y).Magnification, HexPiece(SelectedHex2.X, SelectedHex2.Y).Degree, HexVertex, TexHexagon, HexColor(HexPiece(SelectedHex2.X, SelectedHex2.Y).Color), HexColor(HexPiece(SelectedHex2.X, SelectedHex2.Y).Color), HexColor(HexPiece(SelectedHex2.X, SelectedHex2.Y).Color), HexColor(HexPiece(SelectedHex2.X, SelectedHex2.Y).Color)
            If HexPiece(SelectedHex2.X, SelectedHex2.Y).Star = 1 Then Engine_Render_Rectangle HexPiece(SelectedHex2.X, SelectedHex2.Y).X - (HexPiece(SelectedHex1.X, SelectedHex1.Y).Magnification * 0.5), HexPiece(SelectedHex2.X, SelectedHex2.Y).Y + (HexPiece(SelectedHex1.X, SelectedHex1.Y).Magnification * 0.5) - Y, HexWidth + HexPiece(SelectedHex2.X, SelectedHex2.Y).Magnification, HexHeight + HexPiece(SelectedHex2.X, SelectedHex2.Y).Magnification, HexPiece(SelectedHex2.X, SelectedHex2.Y).Degree, HexVertex, TexStar, D3DColorARGB(255, 255, 255, 0), D3DColorARGB(255, 255, 255, 0), D3DColorARGB(255, 255, 255, 0), D3DColorARGB(255, 255, 255, 0)
        End If
    End If
    If SelectedHex3.X > 0 And SelectedHex3.Y > 0 Then
        If HexPiece(SelectedHex3.X, SelectedHex3.Y).Color <> 0 Then
            Engine_Render_Rectangle HexPiece(SelectedHex3.X, SelectedHex3.Y).X - (HexPiece(SelectedHex3.X, SelectedHex3.Y).Magnification * 0.5), HexPiece(SelectedHex3.X, SelectedHex3.Y).Y - (HexPiece(SelectedHex3.X, SelectedHex3.Y).Magnification * 0.5) - SelectedHex3.Y, HexWidth + HexPiece(SelectedHex3.X, SelectedHex3.Y).Magnification, HexHeight + HexPiece(SelectedHex3.X, SelectedHex3.Y).Magnification, HexPiece(SelectedHex3.X, SelectedHex3.Y).Degree, HexVertex, TexHexagon, HexColor(HexPiece(SelectedHex3.X, SelectedHex3.Y).Color), HexColor(HexPiece(SelectedHex3.X, SelectedHex3.Y).Color), HexColor(HexPiece(SelectedHex3.X, SelectedHex3.Y).Color), HexColor(HexPiece(SelectedHex3.X, SelectedHex3.Y).Color)
            If HexPiece(SelectedHex3.X, SelectedHex3.Y).Star = 1 Then Engine_Render_Rectangle HexPiece(SelectedHex3.X, SelectedHex3.Y).X - (HexPiece(SelectedHex1.X, SelectedHex1.Y).Magnification * 0.5), HexPiece(SelectedHex3.X, SelectedHex3.Y).Y + (HexPiece(SelectedHex1.X, SelectedHex1.Y).Magnification * 0.5) - Y, HexWidth + HexPiece(SelectedHex3.X, SelectedHex3.Y).Magnification, HexHeight + HexPiece(SelectedHex3.X, SelectedHex3.Y).Magnification, HexPiece(SelectedHex3.X, SelectedHex3.Y).Degree, HexVertex, TexStar, D3DColorARGB(255, 255, 255, 0), D3DColorARGB(255, 255, 255, 0), D3DColorARGB(255, 255, 255, 0), D3DColorARGB(255, 255, 255, 0)
        End If
    End If

    '***** Draw the selected hexagons indicators *****
    'Make sure that there are 3 selected hexagons
    If SelectedHex1.X <> 0 Then
        If SelectedHex1.Y <> 0 Then
            If SelectedHex2.X <> 0 Then
                If SelectedHex2.Y <> 0 Then
                    If SelectedHex3.X <> 0 Then
                        If SelectedHex3.Y <> 0 Then

                            'Update the Bullet Degree
                            BulletDegree = BulletDegree + Elapsed * 0.26
                            If BulletDegree > 360 Then BulletDegree = BulletDegree Mod 360

                            'Draw the bullet between the 3 selected hexagons
                            Engine_Render_Rectangle BulletPos.X - (BulWidth * 0.5), BulletPos.Y - (BulHeight * 0.5), BulWidth, BulHeight, BulletDegree, BulVertex(), TexBullet, -754974976, -754974976, -754974976, -754974976
            
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    '***** Draw the point numbers *****
    'Loop through all the points
    For X = 1 To NumPointDisplays
        If PointDisplay(X).Used Then
            
            'Update the variables
            PointDisplay(X).Y = PointDisplay(X).Y - Elapsed * 0.1
            PointDisplay(X).Color.A = PointDisplay(X).Color.A - Elapsed * 0.1
            TempColor = D3DColorARGB(PointDisplay(X).Color.A, PointDisplay(X).Color.R, PointDisplay(X).Color.G, PointDisplay(X).Color.B)
            
            'Check if the number is out of the screen
            If PointDisplay(X).Y < -16 Then PointDisplay(X).Used = 0
            If PointDisplay(X).Color.A < 0 Then PointDisplay(X).Used = 0
            
            'Draw the points
            If PointDisplay(X).Used Then Engine_Render_Rectangle PointDisplay(X).X - (PointDisplay(X).Magnifier * 0.5), PointDisplay(X).Y - (PointDisplay(X).Magnifier * 0.5), 16 + PointDisplay(X).Magnifier, 16 + PointDisplay(X).Magnifier, 0, PointVertex(), TexNumber(PointDisplay(X).Number), TempColor, TempColor, TempColor, TempColor

        End If
    Next X
    
    '***** Draw the particle effects *****
    Effect_UpdateAll

    'DEBUG - Draw the selected hexagon triangle indicator
    If DEBUGMODE_ShowSelectionTriangles Then
        D3DDevice.SetTexture 0, Nothing
        D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 1, LineListA(0), Len(LineListA(0))
        D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 1, LineListB(0), Len(LineListB(0))
        D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 1, LineListC(0), Len(LineListC(0))
    End If

    'End the device rendering
    D3DDevice.EndScene

    'Display the textures drawn to the device
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

End Sub

Public Sub Engine_DeInit()

'************************************************************
'Unload the game engine and all the variables
'************************************************************
Dim Frm As Form
Dim i As Long

    'Unload the textures and devices and such
    If Not DX Is Nothing Then Set DX = Nothing
    If Not D3DX Is Nothing Then Set D3DX = Nothing
    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
    If Not TexHexagon Is Nothing Then Set TexHexagon = Nothing
    If Not TexBullet Is Nothing Then Set TexBullet = Nothing
    If Not TexStar Is Nothing Then Set TexStar = Nothing
    For i = 0 To 9
        If Not TexNumber(i) Is Nothing Then Set TexNumber(i) = Nothing
    Next i
    For i = 1 To NumParticles
        If Not TexParticle(i) Is Nothing Then Set TexParticle(i) = Nothing
    Next i

    'Unload all forms
    For Each Frm In VB.Forms
        Unload Frm
    Next

    'End it all off
    End

End Sub

Private Function Engine_GetAngle(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Single

'************************************************************
'Get the angle between two points
'************************************************************

Dim intB As Integer
Dim intC As Integer

'Get the absolute value of the differences

    intB = Abs(X2 - X1)
    intC = Abs(Y2 - Y1)

    'Dont divide by 0 - apply soh-cah-toa instead
    If intB <> 0 Then Engine_GetAngle = Atn(intC / intB) * RadianToDegree

    'Result of ArcTan is always between 0° and 90°, so check the relative position for correct angle
    If X1 < X2 Then
        If Y1 = Y2 Then Engine_GetAngle = 180
        If Y1 < Y2 Then Engine_GetAngle = 180 - Engine_GetAngle
        If Y1 > Y2 Then Engine_GetAngle = 180 + Engine_GetAngle
    End If

    'Check if directly up
    If X1 > X2 Then
        If Y1 > Y2 Then Engine_GetAngle = 360 - Engine_GetAngle
    End If

    'Check if the X values are the same
    If X1 = X2 Then
        If Y1 < Y2 Then Engine_GetAngle = 90
        If Y1 > Y2 Then Engine_GetAngle = 270
    End If

    'Add 90 to give us the axis we need - the 0/360 degree line points straight up
    Engine_GetAngle = Engine_GetAngle + 90

    'Insure the angle is between 0 and 360
    Engine_GetAngle = Abs(Engine_GetAngle Mod 360)

    'Reverse the direction the degrees add up (go from Counter-Clockwise to Clockwise)
    Engine_GetAngle = 360 - Engine_GetAngle

End Function

Public Sub Engine_Init()

'************************************************************
'Initialize the engine - required to be called before using DirectX
'************************************************************

    'Set the effects array size
    ReDim Effect(1 To NumEffects)

    'Create the DirectX8 devices
    Set DX = New DirectX8
    Set D3D = DX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    'Set the timer to the highest frequency possible
    timeBeginPeriod 1

    'Set up the debugging line list
    LineListA(0).Rhw = 1
    LineListA(0).Specular = 0
    LineListA(0).Z = 1
    LineListA(1) = LineListA(0)

    LineListB(0) = LineListA(0)
    LineListB(1) = LineListA(1)
    LineListC(0) = LineListA(0)
    LineListC(1) = LineListA(1)

    'Create the D3D Device
    If Engine_Init_D3DDevice(D3DCREATE_PUREDEVICE) = 0 Then
        If Engine_Init_D3DDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) = 0 Then
            If Engine_Init_D3DDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) = 0 Then
                MsgBox "There was an error creating the Direct3D device. :(", vbOKOnly
                Unload frmMain
                Set DX = Nothing
                Set D3D = Nothing
                Set D3DX = Nothing
                End
            End If
        End If
    End If

    'Set the render states of the D3DDevice
    Engine_SetRenderStates

    'Initialize the graphics
    Engine_Init_Gfx

    'Initialize DirectSound
    Engine_Init_Sfx
    
    'Clear out the DirectX device (we have finished using it)
    Set DX = Nothing

    'Set FPS values to current time
    FPSLastSecond = timeGetTime
    LastCheck = FPSLastSecond

End Sub

Public Function Engine_Init_D3DDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS)

'************************************************************
'Initialize the Direct3D Device - start off trying with the
'best settings and move to the worst until one works
'************************************************************
Dim D3DWindow As D3DPRESENT_PARAMETERS  'Describes the viewport
Dim DispMode As D3DDISPLAYMODE          'Describes the display mode

    'When there is an error, destroy the D3D device and get ready to make a new one
    On Error GoTo ErrOut

    'Retrieve current display mode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode

    'Set format for windowed mode
    D3DWindow.Windowed = 1  'State that using windowed mode
    D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC 'Refresh with the monitor
    D3DWindow.BackBufferFormat = DispMode.Format    'Use format just retrieved
    D3DWindow.BackBufferWidth = frmMain.ScaleWidth
    D3DWindow.BackBufferHeight = frmMain.ScaleHeight

    'Set the D3DDevice
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATEFLAGS, D3DWindow)
    
    'Display which rendering method made it through
    If DEBUGMODE_ShowUsedRenderMethod Then
        If D3DCREATEFLAGS = D3DCREATE_PUREDEVICE Then
            MsgBox "Using PureDevice"
        ElseIf D3DCREATEFLAGS = D3DCREATE_HARDWARE_VERTEXPROCESSING Then
            MsgBox "Using Hardware"
        Else
            MsgBox "Using Software"
        End If
    End If
    
    'Everything was successful
    Engine_Init_D3DDevice = 1

Exit Function

ErrOut:

    'Destroy the D3DDevice so it can be remade
    Set D3DDevice = Nothing
    
    'Return a failure - 0
    Engine_Init_D3DDevice = 0

End Function

Public Sub Engine_Init_Gfx()

'************************************************************
'Load the graphics/textures into memory
'Remember that textures NEED to be by a power of 2, though Width and Height can be different dimensions
'************************************************************
Dim i As Long

    'Hexagon
    Engine_Init_Texture TexHexagon, HexVertex(), "hex", 0, 0, HexWidth, HexHeight

    'Hexagon rotation bullet
    Engine_Init_Texture TexBullet, BulVertex(), "bullet", 0, 0, BulWidth, BulHeight
    
    'Hexagon star
    Engine_Init_Texture TexStar, StarVertex(), "star", 0, 0, StarWidth, StarHeight
    
    'Numbers
    For i = 0 To 9
        Engine_Init_Texture TexNumber(i), PointVertex(), "a_" & i, 0, 0, 16, 16
    Next i
    
    'Init the particles - dont use Engine_Init_Texture because we only need to init the texture
    For i = 1 To NumParticles
        Set TexParticle(i) = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & GfxPath & "particle" & i & ".png", D3DX_DEFAULT, D3DX_DEFAULT, _
            0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_NONE, ColorKey, ByVal 0, ByVal 0)
    Next i
    
End Sub

Public Sub Engine_Init_Sfx()

'************************************************************
'Initialize the sound effects
'************************************************************

Dim DSBDesc As DSBUFFERDESC

    'Create the DirectSound devices
    Set DSEnum = DX.GetDSEnum
    Set DS = DX.DirectSoundCreate(DSEnum.GetGuid(1))

    'Set the cooperative level
    DS.SetCooperativeLevel frmMain.hWnd, DSSCL_PRIORITY

    'Set the flags to be able to control the frequency and volume
    DSBDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLVOLUME

    'Set the sound buffers
    Set SfxRotate = DS.CreateSoundBufferFromFile(App.Path & SfxPath & "rotate.wav", DSBDesc)
    Set SfxClick = DS.CreateSoundBufferFromFile(App.Path & SfxPath & "click.wav", DSBDesc)
    Set SfxPop = DS.CreateSoundBufferFromFile(App.Path & SfxPath & "pop.wav", DSBDesc)

End Sub

Public Sub Engine_Init_Texture(Texture As Direct3DTexture8, VertexArray() As TLVERTEX, ByVal FileName As String, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single)

'************************************************************
'Set up an individual texture and the VertexArray() assigned to it
'X/Y/Width/Height are the dimensions you want of the picture in the
'texture, not of the whole texture itself
'************************************************************
Dim TexInfo As D3DXIMAGE_INFO_A

    'Set the texture
    Set Texture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & GfxPath & FileName & ".PNG", D3DX_DEFAULT, D3DX_DEFAULT, _
                  0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_NONE, ColorKey, TexInfo, ByVal 0)

    'Set the VertexArray() information - Tu, Tv, Rhw and Specular values
    'Upper-left corner
    With VertexArray(0)
        .Specular = 0
        .Rhw = 1
        .Tu = X / TexInfo.Width
        .Tv = Y / TexInfo.Height
    End With

    'Upper-right corner
    With VertexArray(1)
        .Specular = 0
        .Rhw = 1
        .Tu = (X + Width) / TexInfo.Width
        .Tv = Y / TexInfo.Height
    End With

    'Buttom-left corner
    With VertexArray(2)
        .Specular = 0
        .Rhw = 1
        .Tu = X / TexInfo.Width
        .Tv = (Y + Height) / TexInfo.Height
    End With

    'Bottom-right corner
    With VertexArray(3)
        .Specular = 0
        .Rhw = 1
        .Tu = (X + Width) / TexInfo.Width
        .Tv = (Y + Height) / TexInfo.Height
    End With

End Sub

Private Function Engine_MoveToPoint(X As Single, Y As Single, ByVal TargetX As Single, ByVal TargetY As Single, ByVal Speed As Single, Optional ByVal MoveRadVal As Single = 0) As Byte

'************************************************************
'Move X/Y point towards target point based by a certain speed
'************************************************************

Dim MoveRad As Single
Dim MoveX As Single
Dim MoveY As Single
Dim NeedMove As Byte

'Calculate the angle between the two points
    If MoveRadVal = 0 Then
        MoveRad = Engine_GetAngle(X, Y, TargetX, TargetY) * DegreeToRadian
    Else
        MoveRad = MoveRadVal
    End If
    
    'Check if the X axis needs updating
    If X <> TargetX Then
    
        'Mark pieces as needed to move
        NeedMove = 1

        'Calculate the move distance along the axis
        MoveX = Sin(MoveRad) * Speed

        'Check if the distance to move is greater then the distance needed
        If Abs(MoveX) > Abs(X - TargetX) Then
            'Place at the target position
            X = TargetX
        Else
            'Move the object towards the target position by the sin/cos of the angle multiplied by speed
            X = X + MoveX
        End If

    End If

    'Check if the Y axis needs updating
    If Y <> TargetY Then

        'Mark pieces as needed to move
        NeedMove = 1

        'Calculate the move distance along the axis
        MoveY = -Cos(MoveRad) * Speed

        'Check if the distance to move is greater then the distance needed
        If Abs(MoveY) > Abs(Y - TargetY) Then
            'Place at the target position
            Y = TargetY
        Else
            'Move the object towards the target position by the sin/cos of the angle multiplied by speed
            Y = Y + MoveY
        End If
        
    End If
    
    'If pieces have finished moving AND actually moved, return a 1
    If NeedMove = 1 Then
        If X = TargetX Then
            If Y = TargetY Then
                Engine_MoveToPoint = 1
            End If
        End If
    End If

End Function

Public Function Engine_PointInTriangle(ByVal pX As Single, ByVal pY As Single, ByVal aX As Single, ByVal aY As Single, ByVal bX As Single, ByVal bY As Single, ByVal cX As Single, ByVal cY As Single) As Byte

'************************************************************
'Returns if a point is inside a triangle
'************************************************************

Dim bc As Single
Dim ca As Single
Dim ab As Single
Dim ap As Single
Dim bp As Single
Dim cp As Single
Dim abc As Single

'Get all the calculations

    bc = bX * cY - bY * cX
    ca = cX * aY - cY * aX
    ab = aX * bY - aY * bX
    ap = aX * pY - aY * pX
    bp = bX * pY - bY * pX
    cp = cX * pY - cY * pX
    abc = Sgn(bc + ca + ab)

    'Check if the point is inside the triangle
    If (abc * (bc - bp + cp) > 0) Then
        If (abc * (ca - cp + ap) > 0) Then
            If (abc * (ab - ap + bp) > 0) Then
                Engine_PointInTriangle = 1
            End If
        End If
    End If

End Function

Sub Engine_Render_Rectangle(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal Degrees As Single, VertexArray() As TLVERTEX, Optional ByRef Texture As Direct3DTexture8, Optional ByVal Color0 As Long = -1, Optional ByVal Color1 As Long = -1, Optional ByVal Color2 As Long = -1, Optional ByVal Color3 As Long = -1)

'************************************************************
'Render a square/rectangle based on the specified values then rotate it
'VertexArray() MUST be from 0 to 3.
'************************************************************
Dim RadAngle As Single 'The angle in Radians
Dim CenterX As Single
Dim CenterY As Single
Dim Index As Integer
Dim NewX As Single
Dim NewY As Single
Dim SinRad As Single
Dim CosRad As Single

    'Set the top-left corner
    With VertexArray(0)
        .X = X
        .Y = Y
        .Color = Color0
    End With

    'Set the top-right corner
    With VertexArray(1)
        .X = X + Width
        .Y = Y
        .Color = Color1
    End With

    'Set the bottom-left corner
    With VertexArray(2)
        .X = X
        .Y = Y + Height
        .Color = Color2
    End With

    'Set the bottom-right corner
    With VertexArray(3)
        .X = X + Width
        .Y = Y + Height
        .Color = Color3
    End With

    'Check if a rotation is required
    If Degrees <> 0 Or Degrees <> 360 Then

        'Converts the angle to rotate by into radians
        RadAngle = Degrees * DegreeToRadian

        'Set the CenterX and CenterY values
        CenterX = X + (Width * 0.5)
        CenterY = Y + (Height * 0.5)

        'Pre-calculate the cosine and sine of the radiant
        SinRad = Sin(RadAngle)
        CosRad = Cos(RadAngle)

        'Loops through the passed vertex buffer
        For Index = 0 To 3

            'Calculates the new X and Y co-ordinates of the vertices for the given angle around the center co-ordinates
            NewX = CenterX + (VertexArray(Index).X - CenterX) * CosRad - (VertexArray(Index).Y - CenterY) * SinRad
            NewY = CenterY + (VertexArray(Index).Y - CenterY) * CosRad + (VertexArray(Index).X - CenterX) * SinRad

            'Applies the new co-ordinates to the buffer
            VertexArray(Index).X = NewX
            VertexArray(Index).Y = NewY

        Next Index

    End If

    'Set the texture
    If Not Texture Is Nothing Then D3DDevice.SetTexture 0, Texture

    'Render the texture to the device
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), Len(VertexArray(0))

End Sub

Public Sub Game_Points_Create(ByVal X As Single, ByVal Y As Single, ByVal Points As Long, ByVal ColorID As Byte, ByVal Magnifier As Single)

'************************************************************
'Create a number-display for points
'************************************************************
Dim PointsStrLen As Integer
Dim PointsStr As String
Dim OffsetX As Integer
Dim A As Long
Dim B As Long

    'Convert the points to a string
    PointsStr = Str$(Points)
    PointsStr = Right$(PointsStr, Len(PointsStr) - 1)
    
    'Raise the user's points
    User.Points = User.Points + Points
    
    '//TEMP
    frmMain.Caption = "VBHexic - Points: " & User.Points
    
    'Store the length of the points string
    PointsStrLen = Len(PointsStr)
    
    'Set the offsetx
    OffsetX = -PointsStrLen * 7
    
    'Loop through the points
    For A = 1 To PointsStrLen
        
        'Get an open PointDisplay array place to use
        B = 0
        Do
            B = B + 1
            If B > NumPointDisplays Then Exit Sub
        Loop While PointDisplay(B).Used = 1
        
        'Set the values
        With PointDisplay(B)
            .Color = HexColorARGB(ColorID)
            .Number = Val(Mid$(PointsStr, A, 1))
            .X = X + OffsetX
            .Y = Y
            .Used = 1
            .Magnifier = Magnifier
        End With
        
        'Increaes the offsetx value
        OffsetX = OffsetX + 14
        
    Next A
    

End Sub

Public Sub Engine_SetRenderStates()

'************************************************************
'Set the render states of the Direct3D Device
'This is in a seperate sub since if using Fullscreen and device is lost
'this is eventually called to restore settings.
'************************************************************

'Set the shader to be used

    D3DDevice.SetVertexShader FVF

    'Set the render states
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE

    'Particle engine settings
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0

    'Set the texture stage information (filters)
    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR

End Sub

Public Sub Engine_Sfx_Play(SfxBuffer As DirectSoundSecondaryBuffer8, Optional ByVal Flags As CONST_DSBPLAYFLAGS = DSBPLAY_DEFAULT)

'************************************************************
'Play a sound from the selected buffer
'************************************************************

    'Play the sound buffer
    SfxBuffer.Stop
    SfxBuffer.SetCurrentPosition 0
    If PlaySounds = True Then SfxBuffer.Play Flags

End Sub

Public Function Game_HexAreMoving() As Boolean

'************************************************************
'Checks if hexagons are moving or not
'************************************************************
Dim X As Long
Dim Y As Long

    For X = 1 To HexFieldWidth
        For Y = 1 To HexFieldHeight
            If HexPiece(X, Y).X <> HexPiece(X, Y).TargetX Or HexPiece(X, Y).Y <> HexPiece(X, Y).TargetY Then
                Game_HexAreMoving = True
                Exit Function
            End If
        Next Y
    Next X

End Function

Public Function Game_Bullet_GetPos() As Point

'************************************************************
'Return the central position of the selected hexagons
'Key thing to keep in mind is that the center of the hexagons are found by triangles
'Create a triangle from the 3 central points of the hexagon to find the center
'ASCII Example:
'
' [1]               Central point "x" is found if the mouse collides
'    x [3]    ===>    with a triangle made by connecting the center of
' [2]                 hexagon 1, 2 and 3. Center of the triangel will equal "x".
'************************************************************

Dim GeneralPos As Point
Dim LoopX As Long
Dim LoopY As Long
Dim EmptyPoint As Point 'Empty Point to clear hexagon points
Dim ETLV As TLVERTEX    'Empty TLVertex to clear debug lines

'Calculate the row and column the mouse is in

    GeneralPos.X = ((MousePos.X) - HexFieldOffsetY) / HexWidth
    GeneralPos.Y = ((MousePos.Y) - HexFieldOffsetY) / HexHeight

    'Clear points if in debug mode
    If DEBUGMODE_ShowSelectionTriangles Then
        LineListA(0) = ETLV
        LineListA(1) = ETLV
        LineListB(0) = ETLV
        LineListB(1) = ETLV
        LineListC(0) = ETLV
        LineListC(1) = ETLV
    End If

    'Clear selected hexagons
    SelectedHex1 = EmptyPoint
    SelectedHex2 = EmptyPoint
    SelectedHex3 = EmptyPoint

    'Check for the triangle which collides with the mouse pointer
    'Start from the roughly estimated point calculated above, then brach outwords until it's found
    For LoopX = GeneralPos.X - 1 To GeneralPos.X + 2 'We will check the central point and up to 2 points away

        'Only check if the values are in range of the field
        If LoopX > 0 Then
            If LoopX < HexFieldWidth Then

                For LoopY = GeneralPos.Y - 2 To GeneralPos.Y

                    'Only check if the values are in range of the field
                    If LoopY > 0 Then
                        If LoopY < HexFieldHeight Then

                            D3DDevice.SetTexture 0, Nothing

                            'Check if there is collision using the hexagon center table created earlier
                            If LoopX / 2 = LoopX \ 2 Then

                                'Odd column
                                If Engine_PointInTriangle(MousePos.X, MousePos.Y, HexCenter(LoopX, LoopY).X, HexCenter(LoopX, LoopY).Y, HexCenter(LoopX + 1, LoopY).X, HexCenter(LoopX + 1, LoopY).Y, HexCenter(LoopX, LoopY + 1).X, HexCenter(LoopX, LoopY + 1).Y) = 1 Then

                                    'Display debug information - selected triangles
                                    If DEBUGMODE_ShowSelectionTriangles Then
                                        LineListA(0).Color = D3DColorARGB(255, 255, 0, 0)
                                        LineListA(1).Color = LineListA(0).Color
                                        LineListB(0).Color = LineListA(0).Color
                                        LineListB(1).Color = LineListA(0).Color
                                        LineListC(0).Color = LineListA(0).Color
                                        LineListC(1).Color = LineListA(0).Color
                                        LineListA(0).X = HexCenter(LoopX, LoopY).X
                                        LineListA(0).Y = HexCenter(LoopX, LoopY).Y
                                        LineListA(1).X = HexCenter(LoopX + 1, LoopY).X
                                        LineListA(1).Y = HexCenter(LoopX + 1, LoopY).Y
                                        LineListB(0).X = HexCenter(LoopX + 1, LoopY).X
                                        LineListB(0).Y = HexCenter(LoopX + 1, LoopY).Y
                                        LineListB(1).X = HexCenter(LoopX, LoopY + 1).X
                                        LineListB(1).Y = HexCenter(LoopX, LoopY + 1).Y
                                        LineListC(0).X = HexCenter(LoopX, LoopY + 1).X
                                        LineListC(0).Y = HexCenter(LoopX, LoopY + 1).Y
                                        LineListC(1).X = HexCenter(LoopX, LoopY).X
                                        LineListC(1).Y = HexCenter(LoopX, LoopY).Y
                                    End If

                                    'Store the selected hexagons
                                    SelectedHex1.X = LoopX
                                    SelectedHex1.Y = LoopY
                                    SelectedHex2.X = (LoopX + 1)
                                    SelectedHex2.Y = LoopY
                                    SelectedHex3.X = LoopX
                                    SelectedHex3.Y = (LoopY + 1)

                                    'Return the bullet position
                                    Game_Bullet_GetPos.X = HexCenter(LoopX + 1, LoopY).X - (HexWidth * 0.5)
                                    Game_Bullet_GetPos.Y = HexCenter(LoopX + 1, LoopY).Y

                                ElseIf Engine_PointInTriangle(MousePos.X, MousePos.Y, HexCenter(LoopX, LoopY + 1).X, HexCenter(LoopX, LoopY + 1).Y, HexCenter(LoopX + 1, LoopY).X, HexCenter(LoopX + 1, LoopY).Y, HexCenter(LoopX + 1, LoopY + 1).X, HexCenter(LoopX + 1, LoopY + 1).Y) = 1 Then

                                    'Display debug information - selected triangles
                                    If DEBUGMODE_ShowSelectionTriangles Then
                                        LineListA(0).Color = D3DColorARGB(255, 0, 255, 0)
                                        LineListA(1).Color = LineListA(0).Color
                                        LineListB(0).Color = LineListA(0).Color
                                        LineListB(1).Color = LineListA(0).Color
                                        LineListC(0).Color = LineListA(0).Color
                                        LineListC(1).Color = LineListA(0).Color
                                        LineListA(0).X = HexCenter(LoopX, LoopY + 1).X
                                        LineListA(0).Y = HexCenter(LoopX, LoopY + 1).Y
                                        LineListA(1).X = HexCenter(LoopX + 1, LoopY).X
                                        LineListA(1).Y = HexCenter(LoopX + 1, LoopY).Y
                                        LineListB(0).X = HexCenter(LoopX + 1, LoopY).X
                                        LineListB(0).Y = HexCenter(LoopX + 1, LoopY).Y
                                        LineListB(1).X = HexCenter(LoopX + 1, LoopY + 1).X
                                        LineListB(1).Y = HexCenter(LoopX + 1, LoopY + 1).Y
                                        LineListC(0).X = HexCenter(LoopX + 1, LoopY + 1).X
                                        LineListC(0).Y = HexCenter(LoopX + 1, LoopY + 1).Y
                                        LineListC(1).X = HexCenter(LoopX, LoopY + 1).X
                                        LineListC(1).Y = HexCenter(LoopX, LoopY + 1).Y
                                    End If

                                    'Store the selected hexagons
                                    SelectedHex1.X = LoopX
                                    SelectedHex1.Y = (LoopY + 1)
                                    SelectedHex2.X = (LoopX + 1)
                                    SelectedHex2.Y = LoopY
                                    SelectedHex3.X = (LoopX + 1)
                                    SelectedHex3.Y = (LoopY + 1)

                                    'Return the bullet position
                                    Game_Bullet_GetPos.X = HexCenter(LoopX, LoopY + 1).X + (HexWidth * 0.5)
                                    Game_Bullet_GetPos.Y = HexCenter(LoopX, LoopY + 1).Y

                                End If
                            Else
                                'Even column
                                If Engine_PointInTriangle(MousePos.X, MousePos.Y, HexCenter(LoopX, LoopY).X, HexCenter(LoopX, LoopY).Y, HexCenter(LoopX + 1, LoopY + 1).X, HexCenter(LoopX + 1, LoopY + 1).Y, HexCenter(LoopX, LoopY + 1).X, HexCenter(LoopX, LoopY + 1).Y) = 1 Then

                                    'Display debug information - selected triangles
                                    If DEBUGMODE_ShowSelectionTriangles Then
                                        LineListA(0).Color = D3DColorARGB(255, 255, 255, 255)
                                        LineListA(1).Color = LineListA(0).Color
                                        LineListB(0).Color = LineListA(0).Color
                                        LineListB(1).Color = LineListA(0).Color
                                        LineListC(0).Color = LineListA(0).Color
                                        LineListC(1).Color = LineListA(0).Color
                                        LineListA(0).X = HexCenter(LoopX, LoopY).X
                                        LineListA(0).Y = HexCenter(LoopX, LoopY).Y
                                        LineListA(1).X = HexCenter(LoopX + 1, LoopY + 1).X
                                        LineListA(1).Y = HexCenter(LoopX + 1, LoopY + 1).Y
                                        LineListB(0).X = HexCenter(LoopX + 1, LoopY + 1).X
                                        LineListB(0).Y = HexCenter(LoopX + 1, LoopY + 1).Y
                                        LineListB(1).X = HexCenter(LoopX, LoopY + 1).X
                                        LineListB(1).Y = HexCenter(LoopX, LoopY + 1).Y
                                        LineListC(0).X = HexCenter(LoopX, LoopY + 1).X
                                        LineListC(0).Y = HexCenter(LoopX, LoopY + 1).Y
                                        LineListC(1).X = HexCenter(LoopX, LoopY).X
                                        LineListC(1).Y = HexCenter(LoopX, LoopY).Y
                                    End If

                                    'Store the selected hexagons
                                    SelectedHex1.X = LoopX
                                    SelectedHex1.Y = LoopY
                                    SelectedHex2.X = (LoopX + 1)
                                    SelectedHex2.Y = (LoopY + 1)
                                    SelectedHex3.X = LoopX
                                    SelectedHex3.Y = (LoopY + 1)

                                    'Return the bullet position
                                    Game_Bullet_GetPos.X = HexCenter(LoopX + 1, LoopY + 1).X - (HexWidth * 0.5)
                                    Game_Bullet_GetPos.Y = HexCenter(LoopX + 1, LoopY + 1).Y

                                ElseIf Engine_PointInTriangle(MousePos.X, MousePos.Y, HexCenter(LoopX, LoopY).X, HexCenter(LoopX, LoopY).Y, HexCenter(LoopX + 1, LoopY + 1).X, HexCenter(LoopX + 1, LoopY + 1).Y, HexCenter(LoopX + 1, LoopY).X, HexCenter(LoopX + 1, LoopY).Y) = 1 Then

                                    'Display debug information - selected triangles
                                    If DEBUGMODE_ShowSelectionTriangles Then
                                        LineListA(0).Color = D3DColorARGB(255, 0, 0, 255)
                                        LineListA(1).Color = LineListA(0).Color
                                        LineListB(0).Color = LineListA(0).Color
                                        LineListB(1).Color = LineListA(0).Color
                                        LineListC(0).Color = LineListA(0).Color
                                        LineListC(1).Color = LineListA(0).Color
                                        LineListA(0).X = HexCenter(LoopX, LoopY).X
                                        LineListA(0).Y = HexCenter(LoopX, LoopY).Y
                                        LineListA(1).X = HexCenter(LoopX + 1, LoopY + 1).X
                                        LineListA(1).Y = HexCenter(LoopX + 1, LoopY + 1).Y
                                        LineListB(0).X = HexCenter(LoopX + 1, LoopY + 1).X
                                        LineListB(0).Y = HexCenter(LoopX + 1, LoopY + 1).Y
                                        LineListB(1).X = HexCenter(LoopX + 1, LoopY).X
                                        LineListB(1).Y = HexCenter(LoopX + 1, LoopY).Y
                                        LineListC(0).X = HexCenter(LoopX + 1, LoopY).X
                                        LineListC(0).Y = HexCenter(LoopX + 1, LoopY).Y
                                        LineListC(1).X = HexCenter(LoopX, LoopY).X
                                        LineListC(1).Y = HexCenter(LoopX, LoopY).Y
                                    End If

                                    'Store the selected hexagons
                                    SelectedHex1.X = LoopX
                                    SelectedHex1.Y = LoopY
                                    SelectedHex2.X = (LoopX + 1)
                                    SelectedHex2.Y = LoopY
                                    SelectedHex3.X = (LoopX + 1)
                                    SelectedHex3.Y = (LoopY + 1)

                                    'Return the bullet position
                                    Game_Bullet_GetPos.X = HexCenter(LoopX, LoopY).X + (HexWidth * 0.5)
                                    Game_Bullet_GetPos.Y = HexCenter(LoopX, LoopY).Y

                                End If
                            End If

                        End If
                    End If

                Next LoopY

            End If
        End If

    Next LoopX

End Function

Private Sub Game_HexPiece_Collision()

'************************************************************
'Checks the entire game-field for hexagon patterns of any type
'Most advanced combinations are checked for first
'Optimized to call only when CheckForCollision flag is raised
'Flag is raised only when a piece finishes moving
'************************************************************

Dim X As Long
Dim Y As Long

    'Loop through all hexagons
    For X = 1 To HexFieldWidth
        For Y = 1 To HexFieldHeight

            'Make sure the hex piece is valid
            If HexPiece(X, Y).Color <> 0 Then
                If HexPiece(X, Y).IsShrink = 0 Then
    
                    'Only check pieces which are done moving
                    If Game_HexPiece_Inactive(HexPiece(X, Y)) Then
    
                        'Check for 3-piece collision
                        If Game_HexPiece_Collision_FivePiece(X, Y) Then RotateDir = 0
                        If Game_HexPiece_Collision_FourPiece(X, Y) Then RotateDir = 0
                        If Game_HexPiece_Collision_ThreePiece(X, Y) Then RotateDir = 0
    
                    End If
                
                End If
            End If

        Next Y
    Next X

End Sub

Private Function Game_HexPiece_Collision_FivePiece(ByVal X As Long, ByVal Y As Long) As Byte

'************************************************************
'Check for five-piece collision
'************************************************************
Dim NumStars As Long

    'Set the initial value to "true"
    Game_HexPiece_Collision_FivePiece = 1

    'Odd Row
    If X / 2 <> X \ 2 Then

        'Look for the combination of (X,Y)=(X,Y+1)=(X+1,Y)=(X+1,Y+1)=(X+1,Y+2)
        If X < HexFieldWidth Then
            If Y + 1 < HexFieldHeight Then
                If HexPiece(X, Y).Color = HexPiece(X, Y + 1).Color Then
                    If HexPiece(X, Y).Color = HexPiece(X + 1, Y).Color Then
                        If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 1).Color Then
                            If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 2).Color Then
                                If Game_HexPiece_Inactive(HexPiece(X, Y + 1)) Then
                                    If Game_HexPiece_Inactive(HexPiece(X + 1, Y)) Then
                                        If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 1)) Then
                                            If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 2)) Then
                                                HexPiece(X, Y).IsShrink = 1
                                                HexPiece(X, Y + 1).IsShrink = 1
                                                HexPiece(X + 1, Y).IsShrink = 1
                                                HexPiece(X + 1, Y + 1).IsShrink = 1
                                                HexPiece(X + 1, Y + 2).IsShrink = 1
                                                NumStars = 1 + HexPiece(X, Y).Star + HexPiece(X, Y + 1).Star + HexPiece(X + 1, Y).Star + HexPiece(X + 1, Y + 1).Star + HexPiece(X + 1, Y + 2).Star
                                                Game_Points_Create HexPiece(X, Y).X + HexWidth, HexPiece(X, Y).Y, Points_FivePiece * NumStars, HexPiece(X, Y).Color, Size_FivePiece
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Look for the combination of (X,Y)=(X,Y+1)=(X,Y+2)=(X+1,Y+1)=(X+1,Y+2)
        If X < HexFieldWidth Then
            If Y + 1 < HexFieldHeight Then
                If HexPiece(X, Y).Color = HexPiece(X, Y + 1).Color Then
                    If HexPiece(X, Y).Color = HexPiece(X, Y + 2).Color Then
                        If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 1).Color Then
                            If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 2).Color Then
                                If Game_HexPiece_Inactive(HexPiece(X, Y + 1)) Then
                                    If Game_HexPiece_Inactive(HexPiece(X, Y + 2)) Then
                                        If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 1)) Then
                                            If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 2)) Then
                                                HexPiece(X, Y).IsShrink = 1
                                                HexPiece(X, Y + 1).IsShrink = 1
                                                HexPiece(X, Y + 2).IsShrink = 1
                                                HexPiece(X + 1, Y + 1).IsShrink = 1
                                                HexPiece(X + 1, Y + 2).IsShrink = 1
                                                NumStars = 1 + HexPiece(X, Y).Star + HexPiece(X, Y + 1).Star + HexPiece(X, Y + 2).Star + HexPiece(X + 1, Y + 1).Star + HexPiece(X + 1, Y + 2).Star
                                                Game_Points_Create HexPiece(X, Y).X + HexWidth, HexPiece(X, Y).Y, Points_FivePiece * NumStars, HexPiece(X, Y).Color, Size_FivePiece
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Look for the combination of (X,Y)=(X+1,Y)=(X+2,Y-1)=(X+1,Y+1)=(X+2,Y)
        If Y > 1 Then
            If X + 1 < HexFieldWidth Then
                If Y < HexFieldHeight Then
                    If HexPiece(X, Y).Color = HexPiece(X + 1, Y).Color Then
                        If HexPiece(X, Y).Color = HexPiece(X + 2, Y - 1).Color Then
                            If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 1).Color Then
                                If HexPiece(X, Y).Color = HexPiece(X + 2, Y).Color Then
                                    If Game_HexPiece_Inactive(HexPiece(X + 1, Y)) Then
                                        If Game_HexPiece_Inactive(HexPiece(X + 2, Y - 1)) Then
                                            If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 1)) Then
                                                If Game_HexPiece_Inactive(HexPiece(X + 2, Y)) Then
                                                    HexPiece(X, Y).IsShrink = 1
                                                    HexPiece(X + 1, Y).IsShrink = 1
                                                    HexPiece(X + 2, Y - 1).IsShrink = 1
                                                    HexPiece(X + 1, Y + 1).IsShrink = 1
                                                    HexPiece(X + 2, Y).IsShrink = 1
                                                    NumStars = 1 + HexPiece(X, Y).Star + HexPiece(X + 1, Y).Star + HexPiece(X + 2, Y - 1).Star + HexPiece(X + 1, Y + 1).Star + HexPiece(X + 2, Y).Star
                                                    Game_Points_Create HexPiece(X, Y).X + HexWidth, HexPiece(X, Y).Y, Points_FivePiece * NumStars, HexPiece(X, Y).Color, Size_FivePiece
                                                    Exit Function
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Look for the combination of (X,Y)=(X+1,Y+1)=(X+2,Y+1)=(X,Y+1)=(X+1,Y+2)
        If X + 1 < HexFieldWidth Then
            If Y + 1 < HexFieldHeight Then
                If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 1).Color Then
                    If HexPiece(X, Y).Color = HexPiece(X + 2, Y + 1).Color Then
                        If HexPiece(X, Y).Color = HexPiece(X, Y + 1).Color Then
                            If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 2).Color Then
                                If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 1)) Then
                                    If Game_HexPiece_Inactive(HexPiece(X + 2, Y + 1)) Then
                                        If Game_HexPiece_Inactive(HexPiece(X, Y + 1)) Then
                                            If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 2)) Then
                                                HexPiece(X, Y).IsShrink = 1
                                                HexPiece(X + 1, Y + 1).IsShrink = 1
                                                HexPiece(X + 2, Y + 1).IsShrink = 1
                                                HexPiece(X, Y + 1).IsShrink = 1
                                                HexPiece(X + 1, Y + 2).IsShrink = 1
                                                NumStars = 1 + HexPiece(X, Y).Star + HexPiece(X + 1, Y + 2).Star + HexPiece(X, Y + 1).Star + HexPiece(X + 2, Y + 1).Star + HexPiece(X + 1, Y + 1).Star
                                                Game_Points_Create HexPiece(X, Y).X + HexWidth, HexPiece(X, Y).Y, Points_FivePiece * NumStars, HexPiece(X, Y).Color, Size_FivePiece
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Look for the combination of (X,Y)=(X+1,Y)=(X+1,Y+1)=(X+2,Y)=(X+2,Y+1)
        If X + 1 < HexFieldWidth Then
            If Y < HexFieldHeight Then
                If HexPiece(X, Y).Color = HexPiece(X + 1, Y).Color Then
                    If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 1).Color Then
                        If HexPiece(X, Y).Color = HexPiece(X + 2, Y).Color Then
                            If HexPiece(X, Y).Color = HexPiece(X + 2, Y + 1).Color Then
                                If Game_HexPiece_Inactive(HexPiece(X + 1, Y)) Then
                                    If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 1)) Then
                                        If Game_HexPiece_Inactive(HexPiece(X + 2, Y)) Then
                                            If Game_HexPiece_Inactive(HexPiece(X + 2, Y + 1)) Then
                                                HexPiece(X, Y).IsShrink = 1
                                                HexPiece(X + 1, Y).IsShrink = 1
                                                HexPiece(X + 1, Y + 1).IsShrink = 1
                                                HexPiece(X + 2, Y).IsShrink = 1
                                                HexPiece(X + 2, Y + 1).IsShrink = 1
                                                NumStars = 1 + HexPiece(X, Y).Star + HexPiece(X + 2, Y + 1).Star + HexPiece(X + 2, Y).Star + HexPiece(X + 1, Y).Star + HexPiece(X + 1, Y + 1).Star
                                                Game_Points_Create HexPiece(X, Y).X + HexWidth, HexPiece(X, Y).Y, Points_FivePiece * NumStars, HexPiece(X, Y).Color, Size_FivePiece
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Look for the combination of (X,Y)=(X,Y+1)=(X+1,Y)=(X+1,Y+1)=(X+2,Y)
        If X + 1 < HexFieldWidth Then
            If Y < HexFieldHeight Then
                If HexPiece(X, Y).Color = HexPiece(X, Y + 1).Color Then
                    If HexPiece(X, Y).Color = HexPiece(X + 1, Y).Color Then
                        If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 1).Color Then
                            If HexPiece(X, Y).Color = HexPiece(X + 2, Y).Color Then
                                If Game_HexPiece_Inactive(HexPiece(X, Y + 1)) Then
                                    If Game_HexPiece_Inactive(HexPiece(X + 1, Y)) Then
                                        If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 1)) Then
                                            If Game_HexPiece_Inactive(HexPiece(X + 2, Y)) Then
                                                HexPiece(X, Y).IsShrink = 1
                                                HexPiece(X, Y + 1).IsShrink = 1
                                                HexPiece(X + 1, Y).IsShrink = 1
                                                HexPiece(X + 1, Y + 1).IsShrink = 1
                                                HexPiece(X + 2, Y).IsShrink = 1
                                                NumStars = 1 + HexPiece(X, Y).Star + HexPiece(X + 2, Y).Star + HexPiece(X + 1, Y + 1).Star + HexPiece(X + 1, Y).Star + HexPiece(X, Y + 1).Star
                                                Game_Points_Create HexPiece(X, Y).X + HexWidth, HexPiece(X, Y).Y, Points_FivePiece * NumStars, HexPiece(X, Y).Color, Size_FivePiece
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Even Row
    Else

        'Look for the combination of (X,Y)=(X,Y+1)=(X,Y+2)=(X+1,Y)=(X+1,Y+1)
        If X < HexFieldWidth Then
            If Y + 1 < HexFieldHeight Then
                If HexPiece(X, Y).Color = HexPiece(X, Y + 1).Color Then
                    If HexPiece(X, Y).Color = HexPiece(X, Y + 2).Color Then
                        If HexPiece(X, Y).Color = HexPiece(X + 1, Y).Color Then
                            If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 1).Color Then
                                If Game_HexPiece_Inactive(HexPiece(X, Y + 1)) Then
                                    If Game_HexPiece_Inactive(HexPiece(X, Y + 2)) Then
                                        If Game_HexPiece_Inactive(HexPiece(X + 1, Y)) Then
                                            If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 1)) Then
                                                HexPiece(X, Y).IsShrink = 1
                                                HexPiece(X, Y + 1).IsShrink = 1
                                                HexPiece(X, Y + 2).IsShrink = 1
                                                HexPiece(X + 1, Y).IsShrink = 1
                                                HexPiece(X + 1, Y + 1).IsShrink = 1
                                                NumStars = 1 + HexPiece(X, Y + 1).Star + HexPiece(X + 1, Y + 1).Star + HexPiece(X + 1, Y).Star + HexPiece(X, Y + 2).Star + HexPiece(X, Y).Star
                                                Game_Points_Create HexPiece(X, Y).X + HexWidth - HexOffsetX, HexPiece(X, Y).Y, Points_FivePiece * NumStars, HexPiece(X, Y).Color, Size_FivePiece
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Look for the combination of (X,Y)=(X,Y+1)=(X+1,Y)=(X+1,Y+1)=(X+2,Y+1)
        If X + 1 < HexFieldWidth Then
            If Y < HexFieldHeight Then
                If HexPiece(X, Y).Color = HexPiece(X, Y + 1).Color Then
                    If HexPiece(X, Y).Color = HexPiece(X + 1, Y).Color Then
                        If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 1).Color Then
                            If HexPiece(X, Y).Color = HexPiece(X + 2, Y + 1).Color Then
                                If Game_HexPiece_Inactive(HexPiece(X, Y + 1)) Then
                                    If Game_HexPiece_Inactive(HexPiece(X + 1, Y)) Then
                                        If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 1)) Then
                                            If Game_HexPiece_Inactive(HexPiece(X + 2, Y + 1)) Then
                                                HexPiece(X, Y).IsShrink = 1
                                                HexPiece(X, Y + 1).IsShrink = 1
                                                HexPiece(X + 1, Y).IsShrink = 1
                                                HexPiece(X + 1, Y + 1).IsShrink = 1
                                                HexPiece(X + 2, Y + 1).IsShrink = 1
                                                NumStars = 1 + HexPiece(X + 2, Y + 1).Star + HexPiece(X + 1, Y + 1).Star + HexPiece(X + 1, Y).Star + HexPiece(X, Y).Star + HexPiece(X, Y + 1).Star
                                                Game_Points_Create HexPiece(X, Y).X + HexWidth - HexOffsetX, HexPiece(X, Y).Y, Points_FivePiece * NumStars, HexPiece(X, Y).Color, Size_FivePiece
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Look for the combination of (X,Y)=(X+1,Y-1)=(X+1,Y)=(X+2,Y-1)=(X+2,Y)
        If Y > 1 Then
            If X + 1 < HexFieldWidth Then
                If HexPiece(X, Y).Color = HexPiece(X + 1, Y - 1).Color Then
                    If HexPiece(X, Y).Color = HexPiece(X + 1, Y).Color Then
                        If HexPiece(X, Y).Color = HexPiece(X + 2, Y - 1).Color Then
                            If HexPiece(X, Y).Color = HexPiece(X + 2, Y).Color Then
                                If Game_HexPiece_Inactive(HexPiece(X + 1, Y - 1)) Then
                                    If Game_HexPiece_Inactive(HexPiece(X + 1, Y)) Then
                                        If Game_HexPiece_Inactive(HexPiece(X + 2, Y - 1)) Then
                                            If Game_HexPiece_Inactive(HexPiece(X + 2, Y)) Then
                                                HexPiece(X, Y).IsShrink = 1
                                                HexPiece(X + 1, Y - 1).IsShrink = 1
                                                HexPiece(X + 1, Y).IsShrink = 1
                                                HexPiece(X + 2, Y - 1).IsShrink = 1
                                                HexPiece(X + 2, Y).IsShrink = 1
                                                NumStars = 1 + HexPiece(X, Y).Star + HexPiece(X + 2, Y).Star + HexPiece(X + 2, Y - 1).Star + HexPiece(X + 1, Y).Star + HexPiece(X + 1, Y - 1).Star
                                                Game_Points_Create HexPiece(X, Y).X + HexWidth - HexOffsetX, HexPiece(X, Y).Y, Points_FivePiece * NumStars, HexPiece(X, Y).Color, Size_FivePiece
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Look for the combination of (X,Y)=(X,Y+1)=(X+1,Y-1)=(X+1,Y)=(X+1,Y+1)
        If Y > 1 Then
            If X < HexFieldWidth Then
                If Y < HexFieldHeight Then
                    If HexPiece(X, Y).Color = HexPiece(X, Y + 1).Color Then
                        If HexPiece(X, Y).Color = HexPiece(X + 1, Y - 1).Color Then
                            If HexPiece(X, Y).Color = HexPiece(X + 1, Y).Color Then
                                If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 1).Color Then
                                    If Game_HexPiece_Inactive(HexPiece(X, Y + 1)) Then
                                        If Game_HexPiece_Inactive(HexPiece(X + 1, Y - 1)) Then
                                            If Game_HexPiece_Inactive(HexPiece(X + 1, Y)) Then
                                                If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 1)) Then
                                                    HexPiece(X, Y).IsShrink = 1
                                                    HexPiece(X, Y + 1).IsShrink = 1
                                                    HexPiece(X + 1, Y - 1).IsShrink = 1
                                                    HexPiece(X + 1, Y).IsShrink = 1
                                                    HexPiece(X + 1, Y + 1).IsShrink = 1
                                                    NumStars = 1 + HexPiece(X + 1, Y + 1).Star + HexPiece(X + 1, Y).Star + HexPiece(X + 1, Y - 1).Star + HexPiece(X, Y + 1).Star + HexPiece(X, Y).Star
                                                    Game_Points_Create HexPiece(X, Y).X + HexWidth - HexOffsetX, HexPiece(X, Y).Y, Points_FivePiece * NumStars, HexPiece(X, Y).Color, Size_FivePiece
                                                    Exit Function
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Look for the combination of (X,Y)=(X+1,Y-1)=(X+1,Y)=(X+2,Y)=(X+2,Y+1)
        If Y > 1 Then
            If X + 1 < HexFieldWidth Then
                If Y < HexFieldHeight Then
                    If HexPiece(X, Y).Color = HexPiece(X + 1, Y - 1).Color Then
                        If HexPiece(X, Y).Color = HexPiece(X + 1, Y).Color Then
                            If HexPiece(X, Y).Color = HexPiece(X + 2, Y).Color Then
                                If HexPiece(X, Y).Color = HexPiece(X + 2, Y + 1).Color Then
                                    If Game_HexPiece_Inactive(HexPiece(X + 1, Y - 1)) Then
                                        If Game_HexPiece_Inactive(HexPiece(X + 1, Y)) Then
                                            If Game_HexPiece_Inactive(HexPiece(X + 2, Y)) Then
                                                If Game_HexPiece_Inactive(HexPiece(X + 2, Y + 1)) Then
                                                    HexPiece(X, Y).IsShrink = 1
                                                    HexPiece(X + 1, Y - 1).IsShrink = 1
                                                    HexPiece(X + 1, Y).IsShrink = 1
                                                    HexPiece(X + 2, Y).IsShrink = 1
                                                    HexPiece(X + 2, Y + 1).IsShrink = 1
                                                    NumStars = 1 + HexPiece(X + 2, Y + 1).Star + HexPiece(X + 2, Y).Star + HexPiece(X + 1, Y).Star + HexPiece(X + 1, Y - 1).Star + HexPiece(X, Y).Star
                                                    Game_Points_Create HexPiece(X, Y).X + HexWidth - HexOffsetX, HexPiece(X, Y).Y, Points_FivePiece * NumStars, HexPiece(X, Y).Color, Size_FivePiece
                                                    Exit Function
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Look for the combination of (X,Y)=(X,Y+1)=(X+1,Y-1)=(X+1,Y)=(X+2,Y)
        If Y > 1 Then
            If X + 1 < HexFieldWidth Then
                If Y < HexFieldHeight Then
                    If HexPiece(X, Y).Color = HexPiece(X, Y + 1).Color Then
                        If HexPiece(X, Y).Color = HexPiece(X + 1, Y - 1).Color Then
                            If HexPiece(X, Y).Color = HexPiece(X + 1, Y).Color Then
                                If HexPiece(X, Y).Color = HexPiece(X + 2, Y).Color Then
                                    If Game_HexPiece_Inactive(HexPiece(X, Y + 1)) Then
                                        If Game_HexPiece_Inactive(HexPiece(X + 1, Y - 1)) Then
                                            If Game_HexPiece_Inactive(HexPiece(X + 1, Y)) Then
                                                If Game_HexPiece_Inactive(HexPiece(X + 2, Y)) Then
                                                    HexPiece(X, Y).IsShrink = 1
                                                    HexPiece(X, Y + 1).IsShrink = 1
                                                    HexPiece(X + 1, Y - 1).IsShrink = 1
                                                    HexPiece(X + 1, Y).IsShrink = 1
                                                    HexPiece(X + 2, Y).IsShrink = 1
                                                    NumStars = 1 + HexPiece(X + 2, Y).Star + HexPiece(X + 1, Y).Star + HexPiece(X + 1, Y - 1).Star + HexPiece(X, Y + 1).Star + HexPiece(X, Y).Star
                                                    Game_Points_Create HexPiece(X, Y).X + HexWidth - HexOffsetX, HexPiece(X, Y).Y, Points_FivePiece * NumStars, HexPiece(X, Y).Color, Size_FivePiece
                                                    Exit Function
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

    End If

    'If we have gotten this far, there was no collision
    Game_HexPiece_Collision_FivePiece = 0

End Function

Private Function Game_HexPiece_Collision_FourPiece(ByVal X As Long, ByVal Y As Long) As Byte

'************************************************************
'Check for four-piece collision
'************************************************************
Dim NumStars As Long

    'Set the initial value to "true"
    Game_HexPiece_Collision_FourPiece = 1

    'Odd Row
    If X / 2 <> X \ 2 Then

        'Look for the combination of (X,Y)=(X,Y+1)=(X+1,Y)=(X+1,Y+1)
        If X < HexFieldWidth Then
            If Y < HexFieldHeight Then
                If HexPiece(X, Y).Color = HexPiece(X, Y + 1).Color Then
                    If HexPiece(X, Y).Color = HexPiece(X + 1, Y).Color Then
                        If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 1).Color Then
                            If Game_HexPiece_Inactive(HexPiece(X, Y + 1)) Then
                                If Game_HexPiece_Inactive(HexPiece(X + 1, Y)) Then
                                    If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 1)) Then
                                        HexPiece(X, Y).IsShrink = 1
                                        HexPiece(X, Y + 1).IsShrink = 1
                                        HexPiece(X + 1, Y).IsShrink = 1
                                        HexPiece(X + 1, Y + 1).IsShrink = 1
                                        NumStars = 1 + HexPiece(X + 1, Y + 1).Star + HexPiece(X + 1, Y).Star + HexPiece(X, Y + 1).Star + HexPiece(X, Y).Star
                                        Game_Points_Create HexPiece(X, Y).X + HexWidth, HexPiece(X, Y).Y, Points_FourPiece * NumStars, HexPiece(X, Y).Color, Size_FourPiece
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Look for the combination of (X,Y)=(X,Y+1)=(X+1,Y+1)=(X+1,Y+2)
        If X < HexFieldWidth Then
            If Y + 1 < HexFieldHeight Then
                If HexPiece(X, Y).Color = HexPiece(X, Y + 1).Color Then
                    If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 1).Color Then
                        If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 2).Color Then
                            If Game_HexPiece_Inactive(HexPiece(X, Y + 1)) Then
                                If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 1)) Then
                                    If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 2)) Then
                                        HexPiece(X, Y).IsShrink = 1
                                        HexPiece(X, Y + 1).IsShrink = 1
                                        HexPiece(X + 1, Y + 1).IsShrink = 1
                                        HexPiece(X + 1, Y + 2).IsShrink = 1
                                        NumStars = 1 + HexPiece(X + 1, Y + 2).Star + HexPiece(X + 1, Y + 1).Star + HexPiece(X, Y + 1).Star + HexPiece(X, Y).Star
                                        Game_Points_Create HexPiece(X, Y).X + HexWidth, HexPiece(X, Y).Y, Points_FourPiece * NumStars, HexPiece(X, Y).Color, Size_FourPiece
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Look for the combination (X,Y)=(X+1,Y)=(X+1,Y+1)=(X+2,Y)
        If X + 1 < HexFieldWidth Then
            If Y < HexFieldHeight Then
                If HexPiece(X, Y).Color = HexPiece(X + 1, Y).Color Then
                    If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 1).Color Then
                        If HexPiece(X, Y).Color = HexPiece(X + 2, Y).Color Then
                            If Game_HexPiece_Inactive(HexPiece(X + 1, Y)) Then
                                If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 1)) Then
                                    If Game_HexPiece_Inactive(HexPiece(X + 2, Y)) Then
                                        HexPiece(X, Y).IsShrink = 1
                                        HexPiece(X + 1, Y).IsShrink = 1
                                        HexPiece(X + 1, Y + 1).IsShrink = 1
                                        HexPiece(X + 2, Y).IsShrink = 1
                                        NumStars = 1 + HexPiece(X, Y).Star + HexPiece(X + 1, Y).Star + HexPiece(X + 1, Y + 1).Star + HexPiece(X + 2, Y).Star
                                        Game_Points_Create HexPiece(X, Y).X + HexWidth, HexPiece(X, Y).Y, Points_FourPiece * NumStars, HexPiece(X, Y).Color, Size_FourPiece
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Even Row
    Else

        'Look for the combination of (X,Y)=(X,Y+1)=(X+1,Y-1)=(X+1,Y)
        If Y > 1 Then
            If X < HexFieldWidth Then
                If Y < HexFieldHeight Then
                    If HexPiece(X, Y).Color = HexPiece(X, Y + 1).Color Then
                        If HexPiece(X, Y).Color = HexPiece(X + 1, Y).Color Then
                            If HexPiece(X, Y).Color = HexPiece(X + 1, Y - 1).Color Then
                                If Game_HexPiece_Inactive(HexPiece(X, Y + 1)) Then
                                    If Game_HexPiece_Inactive(HexPiece(X + 1, Y - 1)) Then
                                        If Game_HexPiece_Inactive(HexPiece(X + 1, Y)) Then
                                            HexPiece(X, Y).IsShrink = 1
                                            HexPiece(X, Y + 1).IsShrink = 1
                                            HexPiece(X + 1, Y - 1).IsShrink = 1
                                            HexPiece(X + 1, Y).IsShrink = 1
                                            NumStars = 1 + HexPiece(X, Y).Star + HexPiece(X, Y + 1).Star + HexPiece(X + 1, Y - 1).Star + HexPiece(X + 1, Y).Star
                                            Game_Points_Create HexPiece(X, Y).X + HexWidth - HexOffsetX, HexPiece(X, Y).Y, Points_FourPiece * NumStars, HexPiece(X, Y).Color, Size_FourPiece
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Look for the combination of (X,Y)=(X,Y+1)=(X+1,Y)=(X+1,Y+1)
        If X < HexFieldWidth Then
            If Y < HexFieldHeight Then
                If HexPiece(X, Y).Color = HexPiece(X, Y + 1).Color Then
                    If HexPiece(X, Y).Color = HexPiece(X + 1, Y).Color Then
                        If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 1).Color Then
                            If Game_HexPiece_Inactive(HexPiece(X, Y + 1)) Then
                                If Game_HexPiece_Inactive(HexPiece(X + 1, Y)) Then
                                    If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 1)) Then
                                        HexPiece(X, Y).IsShrink = 1
                                        HexPiece(X, Y + 1).IsShrink = 1
                                        HexPiece(X + 1, Y).IsShrink = 1
                                        HexPiece(X + 1, Y + 1).IsShrink = 1
                                        NumStars = 1 + HexPiece(X, Y).Star + HexPiece(X, Y + 1).Star + HexPiece(X + 1, Y).Star + HexPiece(X + 1, Y + 1).Star
                                        Game_Points_Create HexPiece(X, Y).X + HexWidth - HexOffsetX, HexPiece(X, Y).Y, Points_FourPiece * NumStars, HexPiece(X, Y).Color, Size_FourPiece
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Look for the combination of (X,Y)=(X+1,Y-1)=(X+1,Y)=(X+2,Y)
        If Y > 1 Then
            If X + 1 < HexFieldWidth Then
                If Y < HexFieldHeight Then
                    If HexPiece(X, Y).Color = HexPiece(X + 1, Y - 1).Color Then
                        If HexPiece(X, Y).Color = HexPiece(X + 1, Y).Color Then
                            If HexPiece(X, Y).Color = HexPiece(X + 2, Y).Color Then
                                If Game_HexPiece_Inactive(HexPiece(X + 1, Y - 1)) Then
                                    If Game_HexPiece_Inactive(HexPiece(X + 1, Y)) Then
                                        If Game_HexPiece_Inactive(HexPiece(X + 2, Y)) Then
                                            HexPiece(X, Y).IsShrink = 1
                                            HexPiece(X + 1, Y - 1).IsShrink = 1
                                            HexPiece(X + 1, Y).IsShrink = 1
                                            HexPiece(X + 2, Y).IsShrink = 1
                                            NumStars = 1 + HexPiece(X, Y).Star + HexPiece(X + 1, Y - 1).Star + HexPiece(X + 1, Y).Star + HexPiece(X + 2, Y).Star
                                            Game_Points_Create HexPiece(X, Y).X + HexWidth - HexOffsetX, HexPiece(X, Y).Y, Points_FourPiece * NumStars, HexPiece(X, Y).Color, Size_FourPiece
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

    End If

    'If we have gotten this far, there was no collision
    Game_HexPiece_Collision_FourPiece = 0

End Function

Private Function Game_HexPiece_Collision_ThreePiece(ByVal X As Long, ByVal Y As Long) As Byte

'************************************************************
'Check for three-piece collision
'************************************************************
Dim NumStars As Long

    'Set the initial value to "true"
    Game_HexPiece_Collision_ThreePiece = 1

    'Odd Row
    If X / 2 <> X \ 2 Then

        'Look for the combination of (X,Y)=(X+1,Y)=(X+1,Y+1)
        If X < HexFieldWidth Then
            If Y < HexFieldWidth Then
                If HexPiece(X, Y).Color = HexPiece(X + 1, Y).Color Then
                    If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 1).Color Then
                        If Game_HexPiece_Inactive(HexPiece(X + 1, Y)) Then
                            If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 1)) Then
                                HexPiece(X, Y).IsShrink = 1
                                HexPiece(X + 1, Y).IsShrink = 1
                                HexPiece(X + 1, Y + 1).IsShrink = 1
                                NumStars = 1 + HexPiece(X, Y).Star + HexPiece(X + 1, Y).Star + HexPiece(X + 1, Y + 1).Star
                                Game_Points_Create HexPiece(X, Y).X + HexWidth, HexPiece(X, Y).Y, Points_ThreePiece * NumStars, HexPiece(X, Y).Color, Size_ThreePiece
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Look for the combination of (X,Y)=(X+1,Y+1)=(X,Y+1)
        If X < HexFieldWidth Then
            If Y < HexFieldWidth Then
                If HexPiece(X, Y).Color = HexPiece(X + 1, Y + 1).Color Then
                    If HexPiece(X, Y).Color = HexPiece(X, Y + 1).Color Then
                        If Game_HexPiece_Inactive(HexPiece(X + 1, Y + 1)) Then
                            If Game_HexPiece_Inactive(HexPiece(X, Y + 1)) Then
                                HexPiece(X, Y).IsShrink = 1
                                HexPiece(X + 1, Y + 1).IsShrink = 1
                                HexPiece(X, Y + 1).IsShrink = 1
                                NumStars = 1 + HexPiece(X, Y).Star + HexPiece(X + 1, Y + 1).Star + HexPiece(X, Y + 1).Star
                                Game_Points_Create HexPiece(X, Y).X + HexWidth, HexPiece(X, Y).Y, Points_ThreePiece * NumStars, HexPiece(X, Y).Color, Size_ThreePiece
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Even Row
    Else

        'Look for the combination of (X,Y)=(X+1,Y)=(X,Y+1)
        If X < HexFieldWidth Then
            If Y < HexFieldWidth Then
                If HexPiece(X, Y).Color = HexPiece(X + 1, Y).Color Then
                    If HexPiece(X, Y).Color = HexPiece(X, Y + 1).Color Then
                        If Game_HexPiece_Inactive(HexPiece(X + 1, Y)) Then
                            If Game_HexPiece_Inactive(HexPiece(X, Y + 1)) Then
                                HexPiece(X, Y).IsShrink = 1
                                HexPiece(X + 1, Y).IsShrink = 1
                                HexPiece(X, Y + 1).IsShrink = 1
                                NumStars = 1 + HexPiece(X, Y).Star + HexPiece(X + 1, Y).Star + HexPiece(X, Y + 1).Star
                                Game_Points_Create HexPiece(X, Y).X + HexWidth - HexOffsetX, HexPiece(X, Y).Y, Points_ThreePiece * NumStars, HexPiece(X, Y).Color, Size_ThreePiece
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Look for the combination of (X,Y)=(X+1,Y-1)=(X+1,Y)
        If Y > 1 Then
            If X < HexFieldWidth Then
                If HexPiece(X, Y).Color = HexPiece(X + 1, Y - 1).Color Then
                    If HexPiece(X, Y).Color = HexPiece(X + 1, Y).Color Then
                        If Game_HexPiece_Inactive(HexPiece(X + 1, Y - 1)) Then
                            If Game_HexPiece_Inactive(HexPiece(X + 1, Y)) Then
                                HexPiece(X, Y).IsShrink = 1
                                HexPiece(X + 1, Y - 1).IsShrink = 1
                                HexPiece(X + 1, Y).IsShrink = 1
                                NumStars = 1 + HexPiece(X, Y).Star + HexPiece(X + 1, Y).Star + HexPiece(X + 1, Y - 1).Star
                                Game_Points_Create HexPiece(X, Y).X + HexWidth - HexOffsetX, HexPiece(X, Y).Y, Points_ThreePiece * NumStars, HexPiece(X, Y).Color, Size_ThreePiece
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If

    End If

    'If we have gotten this far, there was no collision
    Game_HexPiece_Collision_ThreePiece = 0

End Function

Private Sub Game_HexPiece_Create(ByRef HexSlot As HexPiece, ByVal Column As Byte)

'************************************************************
'Create a hexagon to fall down a column to a designated slot
'************************************************************
Dim HasStar As Byte

    'Check to randomly create a star hexagon
    If Int(Rnd * StarChance) = 0 Then HasStar = 1

    'Set the values
    HexSlot.Color = Int(Rnd * 6) + 1
    HexSlot.Degree = 0
    HexSlot.Magnification = 0
    HexSlot.Star = HasStar
    HexSlot.Shrink = 0
    HexSlot.IsShrink = 0
    HexSlot.X = HexPoint(Column, 1).X
    HexSlot.Y = 0 - HexHeight
    HexSlot.MoveRad = Engine_GetAngle(HexSlot.X, HexSlot.Y, HexSlot.TargetX, HexSlot.TargetY)

End Sub

Public Sub Game_HexPiece_CreateTables()

'************************************************************
'Create a lookup table with the center position and normal
'position of the hexagons for later reference
'************************************************************

Dim OffsetX As Single
Dim X As Single
Dim Y As Single

'Loop through the field hexagons

    For X = 1 To HexFieldWidth

        'Set the X-value offset
        If X > 1 Then
            If X / 2 <> X \ 2 Then OffsetX = -(HexOffsetX * (X - 1))
        End If

        For Y = 1 To HexFieldHeight

            'Find position of an even-numbered row
            If X / 2 = X \ 2 Then
                HexPoint(X, Y).X = ((X * HexWidth) + HexFieldOffsetX - HexOffsetX + OffsetX)
                HexPoint(X, Y).Y = ((Y * HexHeight) + HexFieldOffsetY - HexOffsetY)
                'Find position of an odd-numbered row
            Else
                HexPoint(X, Y).X = ((X * HexWidth) + HexFieldOffsetX + OffsetX)
                HexPoint(X, Y).Y = ((Y * HexHeight) + HexFieldOffsetY)
            End If

            'Set the Hex Center point
            HexCenter(X, Y).X = HexPoint(X, Y).X + (HexWidth * 0.5)
            HexCenter(X, Y).Y = HexPoint(X, Y).Y + (HexHeight * 0.5) - Y

        Next Y

    Next X

End Sub

Private Function Game_HexPiece_Inactive(ByRef HexInfo As HexPiece) As Byte

'************************************************************
'Check if a hexagon piece is inactive - will return
'a value if has a color, is not shrinking and is not moving
'************************************************************

    'Make sure the piece is in play
    If HexInfo.Color <> 0 Then
        If HexInfo.IsShrink = 0 Then

            'Check if the coordinates are equal
            If HexInfo.X = HexInfo.TargetX Then
                If HexInfo.Y = HexInfo.TargetY Then
                    Game_HexPiece_Inactive = 1
                End If
            End If
        
        End If
    End If
    
End Function

Private Sub Game_HexPiece_Move(ByRef HexInfo As HexPiece, ByVal TargetX As Single, ByVal TargetY As Single)

'************************************************************
'Move a hexagon towards a new position
'Most always use this sub to update TargetX and TargetY values
'************************************************************

    'Update the values
    With HexInfo
        .TargetX = TargetX
        .TargetY = TargetY
        .MoveRad = Engine_GetAngle(.X, .Y, .TargetX, .TargetY) * DegreeToRadian
    End With

End Sub

Private Sub Game_HexPiece_Remove(ByVal ArrayX As Byte, ByVal ArrayY As Byte, Optional ByVal Points As Integer)

'************************************************************
'Remove a hexagon piece and give the appropriate points
'************************************************************
Dim HexColor As ARGBSet
Dim B As Integer
Dim i As Byte


    'Get the hexagon removal effect index
    Do
        i = i + 1
        If i > NumEffects Then
            i = 0
            Exit Do
        End If
    Loop While Effect(i).Used = True

    'Create the hexagon removal particle effect
    HexColor = HexColorARGB(HexPiece(ArrayX, ArrayY).Color)
    If i > 0 Then Effect_Glitter_Begin HexCenter(ArrayX, ArrayY).X, HexCenter(ArrayX, ArrayY).Y, 70, 1, HexColor.R / 255, HexColor.G / 255, HexColor.B / 255

    'Check if hexagons above need to be moved down
    If ArrayY > 1 Then
        For B = ArrayY To 2 Step -1
            HexPiece(ArrayX, B) = HexPiece(ArrayX, B - 1)
            Game_HexPiece_Move HexPiece(ArrayX, B), HexPoint(ArrayX, B).X, HexPoint(ArrayX, B).Y
        Next B

    End If
    
    'Play pop sound
    Engine_Sfx_Play SfxPop

    'Recreate the hexagon piece
    Game_HexPiece_Create HexPiece(ArrayX, 1), ArrayX

End Sub

Public Sub Game_HexPiece_RotateCluster(ByVal Hex1X As Byte, ByVal Hex1Y As Byte, ByVal Hex2X As Byte, ByVal Hex2Y As Byte, ByVal Hex3X As Byte, ByVal Hex3Y As Byte, Clockwise As Boolean)

'************************************************************
'Rotate a hexagon cluster (3 hexagons) one time in the specified direction
'Make sure the hexagons are defined going clockwise
'************************************************************

Dim TempHex As HexPiece 'Used to assign a hexagon to temporarly

'Check for valid hex values

    If Hex1X <= 0 Then Exit Sub
    If Hex1Y <= 0 Then Exit Sub
    If Hex2X <= 0 Then Exit Sub
    If Hex2Y <= 0 Then Exit Sub
    If Hex3X <= 0 Then Exit Sub
    If Hex3Y <= 0 Then Exit Sub

    'Rotate each cluster down in the array (clockwise)
    If Clockwise Then

        'Set the target positions
        Game_HexPiece_Move HexPiece(Hex1X, Hex1Y), HexPoint(Hex2X, Hex2Y).X, HexPoint(Hex2X, Hex2Y).Y 'Piece 1 --> Piece 2
        Game_HexPiece_Move HexPiece(Hex2X, Hex2Y), HexPoint(Hex3X, Hex3Y).X, HexPoint(Hex3X, Hex3Y).Y 'Piece 2 --> Piece 3
        Game_HexPiece_Move HexPiece(Hex3X, Hex3Y), HexPoint(Hex1X, Hex1Y).X, HexPoint(Hex1X, Hex1Y).Y 'Piece 3 --> Piece 1

        'Switch around the hexagons
        TempHex = HexPiece(Hex1X, Hex1Y)
        HexPiece(Hex1X, Hex1Y) = HexPiece(Hex3X, Hex3Y)
        HexPiece(Hex3X, Hex3Y) = HexPiece(Hex2X, Hex2Y)
        HexPiece(Hex2X, Hex2Y) = TempHex

    Else    'Rotate each cluster up in the array (counter-clockwise)

        'Set the positions
        Game_HexPiece_Move HexPiece(Hex1X, Hex1Y), HexPoint(Hex3X, Hex3Y).X, HexPoint(Hex3X, Hex3Y).Y 'Piece 1 --> Piece 3
        Game_HexPiece_Move HexPiece(Hex2X, Hex2Y), HexPoint(Hex1X, Hex1Y).X, HexPoint(Hex1X, Hex1Y).Y 'Piece 2 --> Piece 1
        Game_HexPiece_Move HexPiece(Hex3X, Hex3Y), HexPoint(Hex2X, Hex2Y).X, HexPoint(Hex2X, Hex2Y).Y 'Piece 3 --> Piece 2

        'Switch around the hexagons
        TempHex = HexPiece(Hex3X, Hex3Y)
        HexPiece(Hex3X, Hex3Y) = HexPiece(Hex1X, Hex1Y)
        HexPiece(Hex1X, Hex1Y) = HexPiece(Hex2X, Hex2Y)
        HexPiece(Hex2X, Hex2Y) = TempHex

    End If

    'Reset any already-going rotation
    HexPiece(Hex1X, Hex1Y).Degree = 0
    HexPiece(Hex2X, Hex2Y).Degree = 0
    HexPiece(Hex3X, Hex3Y).Degree = 0

    'Set rotation values
    HexPiece(Hex1X, Hex1Y).Rotate = True
    HexPiece(Hex2X, Hex2Y).Rotate = True
    HexPiece(Hex3X, Hex3Y).Rotate = True

    'Play rotation sound
    Engine_Sfx_Play SfxRotate

End Sub

Public Sub Game_Init()

'************************************************************
'Init the game variables
'************************************************************

Dim X As Byte
Dim Y As Byte

'Create the hexagon center position lookup table

    Game_HexPiece_CreateTables

    'Set the ARGB colors
    HexColorARGB(HexColorRed).A = 255
    HexColorARGB(HexColorRed).R = 255
    HexColorARGB(HexColorRed).G = 0
    HexColorARGB(HexColorRed).B = 0
    
    HexColorARGB(HexColorGreen).A = 255
    HexColorARGB(HexColorGreen).R = 0
    HexColorARGB(HexColorGreen).G = 255
    HexColorARGB(HexColorGreen).B = 0
    
    HexColorARGB(HexColorBlue).A = 255
    HexColorARGB(HexColorBlue).R = 0
    HexColorARGB(HexColorBlue).G = 0
    HexColorARGB(HexColorBlue).B = 255
    
    HexColorARGB(HexColorYellow).A = 255
    HexColorARGB(HexColorYellow).R = 255
    HexColorARGB(HexColorYellow).G = 255
    HexColorARGB(HexColorYellow).B = 0
    
    HexColorARGB(HexColorAqua).A = 255
    HexColorARGB(HexColorAqua).R = 0
    HexColorARGB(HexColorAqua).G = 200
    HexColorARGB(HexColorAqua).B = 200
    
    HexColorARGB(HexColorDarkPurple).A = 255
    HexColorARGB(HexColorDarkPurple).R = 175
    HexColorARGB(HexColorDarkPurple).G = 0
    HexColorARGB(HexColorDarkPurple).B = 175
    
    'Use the ARGB colors to create the Longs
    HexColor(HexColorRed) = D3DColorARGB(HexColorARGB(HexColorRed).A, HexColorARGB(HexColorRed).R, HexColorARGB(HexColorRed).G, HexColorARGB(HexColorRed).B)
    HexColor(HexColorGreen) = D3DColorARGB(HexColorARGB(HexColorGreen).A, HexColorARGB(HexColorGreen).R, HexColorARGB(HexColorGreen).G, HexColorARGB(HexColorGreen).B)
    HexColor(HexColorBlue) = D3DColorARGB(HexColorARGB(HexColorBlue).A, HexColorARGB(HexColorBlue).R, HexColorARGB(HexColorBlue).G, HexColorARGB(HexColorBlue).B)
    HexColor(HexColorYellow) = D3DColorARGB(HexColorARGB(HexColorYellow).A, HexColorARGB(HexColorYellow).R, HexColorARGB(HexColorYellow).G, HexColorARGB(HexColorYellow).B)
    HexColor(HexColorAqua) = D3DColorARGB(HexColorARGB(HexColorAqua).A, HexColorARGB(HexColorAqua).R, HexColorARGB(HexColorAqua).G, HexColorARGB(HexColorAqua).B)
    HexColor(HexColorDarkPurple) = D3DColorARGB(HexColorARGB(HexColorDarkPurple).A, HexColorARGB(HexColorDarkPurple).R, HexColorARGB(HexColorDarkPurple).G, HexColorARGB(HexColorDarkPurple).B)
    
    

    'Loop through the hexagons
    For X = 1 To HexFieldWidth
        For Y = 1 To HexFieldHeight
            
            'Create the initial pieces
            Game_HexPiece_Create HexPiece(X, Y), X
            Game_HexPiece_Move HexPiece(X, Y), HexPoint(X, Y).X, HexPoint(X, Y).Y

        Next Y
    Next X

End Sub

Sub Game_Loop()

'************************************************************
'General part of the whole game loop
'************************************************************
Dim X As Long
'Loop only while EndGameLoop is false

    Do While EndGameLoop = False

        'Check for hex piece collision if needed - only check once per frame
        If CheckForCollision Then
            Game_HexPiece_Collision
            CheckForCollision = 0
        End If
        
        'Check to rotate
        If RotateDir > 0 Then
            If RotateDelay < timeGetTime Then
                RotateCount = RotateCount + 1
                If RotateCount = 4 Then
                    RotateCount = 0
                    RotateDir = 0
                Else
                    RotateDelay = timeGetTime + 425 'Delay (in milliseconds) between rotate attemps
                    Game_HexPiece_RotateCluster RotateHex1.X, RotateHex1.Y, RotateHex2.X, RotateHex2.Y, RotateHex3.X, RotateHex3.Y, IIf(RotateDir = 1, True, False)
                End If
            End If
        End If
        
        'Draw the game stage
        Draw_Stage
        
        'Calculate the FPS values
        FPSCounter = FPSCounter + 1                 'Raise the FPSCounter since a frame has finished
        CurrTime = timeGetTime                      'Store the current time
        Elapsed = CurrTime - LastCheck              'Get the time that has passed (in miliseconds)
        If CurrTime - 1000 >= FPSLastSecond Then    '1000ms (1 second) has passed
            FPS = FPSCounter                        'Set the FPS to the FPS counter
            FPSCounter = 0                          'Clear the FPS counter back to 0
            FPSLastSecond = CurrTime                'Set to wait for another second
            frmMain.Caption = "VBHexic - Points: " & User.Points
        End If
        LastCheck = CurrTime            'Set the last check to that of the current time

        'Let windows do its events
        DoEvents

    Loop

    'If the game loop has stopped running, it is probably time to unload the engine
    Engine_DeInit

End Sub
