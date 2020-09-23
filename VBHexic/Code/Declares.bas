Attribute VB_Name = "Declares"
Option Explicit

'Debug constants
Public Const DEBUGMODE_ShowSelectionTriangles As Boolean = False
Public Const DEBUGMODE_ShowUsedRenderMethod As Boolean = False

'Game Constants
Public Const Points_FivePiece As Long = 35      'Points for a 5-piece combo
Public Const Points_FourPiece As Long = 15      'Points for a 4-piece combo
Public Const Points_ThreePiece As Long = 5      'Points for a 3-piece combo

Public Const Size_FivePiece As Single = 6       'Size of the text for a 5-piece combo
Public Const Size_FourPiece As Single = 3       'Size of the text for a 4-piece combo
Public Const Size_ThreePiece As Single = 0      'Size of the text for a 3-piece combo

Public Const HexPieceRotationSpeed As Single = 1.3  'Speed in which hex pieces rotate
Public Const HexPieceMoveSpeed As Single = 0.1  'Speed in which hex pieces move from (X,Y) to (TargetX,TargetY)
Public Const MagnificationMax As Single = 6     'Maximum magnification size of a hexagon
Public Const MagnificationRate As Single = 0.07 'Rate in which hexagons magnify - higher value, faster they reach MagnificationMax
Public Const NumEffects As Byte = 40            'Maximum total number of particle effects
Public Const NumPointDisplays As Byte = 50      'Maximum number of point displays at once (each one is one number)
Public Const ShrinkSpeed As Single = 0.1        'Speed in which hexagons srhink
Public Const PlaySounds As Boolean = True       'If sounds are enabled or not

Public Const HexWidth As Single = 38            'Width of the actual hexagon, not the texture
Public Const HexHeight As Single = 34           'Height of the hexagon
Public Const HexOffsetX As Single = 9           'X-offset of the hexagon for the hexagon field (for every other column)
Public Const HexOffsetY As Single = 17          'Y-offset of the hexagon (for every other column)

Public Const BulWidth As Single = 24            'Width of bullet
Public Const BulHeight As Single = 23           'Height of bullet

Public Const StarChance As Long = 10            '1 out of every StarChance stars have a star
Public Const StarWidth As Single = HexWidth     'Width of a hexagon star
Public Const StarHeight As Single = HexHeight   'Height of a hexagon star

Public Const HexFieldOffsetX As Single = 10     'X-offset of the hexagon field from point (0,0)
Public Const HexFieldOffsetY As Single = 10     'Y-offset of the hexagon field from point (0,0)
Public Const HexFieldWidth As Single = 10       'Width of the hexagon field in hexagon tiles
Public Const HexFieldHeight As Single = 10      'Height of the hexagon field in hexagon tiles

Public Const ColorKey As Long = &HFF000000      'Magenta Color Key - RGB(255, 0, 255)
Public Const GfxPath As String = "\Gfx\"        'Path to the graphics folder
Public Const SfxPath As String = "\Sfx\"        'Path to the sound folder

'Hexagon color constants
Public Const HexColorRed As Byte = 1
Public Const HexColorGreen As Byte = 2
Public Const HexColorBlue As Byte = 3
Public Const HexColorYellow As Byte = 4
Public Const HexColorAqua As Byte = 5
Public Const HexColorDarkPurple As Byte = 6
Public HexColor(1 To 6) As Long
Public HexColorARGB(1 To 6) As ARGBSet

'Hexagon information type
Public Type HexPiece
    Magnification As Single 'Magnification value - adds up to reach MagnificationMax by MagnificationRate
    Rotate As Boolean   'If to start rotating automatically by HexPieceRotationSpeed rate - each 360 degree rotation is +1, -1 is infinite
    Degree As Single    'Degree in which the hex piece is rotated
    Color As Byte       'Color ID from HexColor consts
    Star As Byte        'If the hexagon contains a star
    X As Single         'Coordinate of the hex piece
    Y As Single
    TargetX As Single   'Target coordinate to reach - equal to X/Y if not moving
    TargetY As Single
    MoveRad As Single   'The angle the hex piece is moving at in radians
    Shrink As Single    'How much the hexagon has srhinked during removal
    IsShrink As Byte    'If the hexagon is shrinking or not - while in this mode, hexagons are disabled
End Type

'Alpha-RGB Color type
Public Type ARGBSet
    A As Integer
    R As Integer
    G As Integer
    B As Integer
End Type

'Player variables
Public Type User
    Name As String
    Points As Long
End Type
Public User As User

'Points display variables
Public Type PointDisplay
    X As Single
    Y As Single
    Used As Byte
    Color As ARGBSet
    Number As Byte
    Magnifier As Single
End Type
Public PointDisplay(1 To NumPointDisplays) As PointDisplay

'Hexagon data array
Public HexPiece(1 To HexFieldWidth, 1 To HexFieldHeight) As HexPiece

'Selected hexagon information (set in GetBulletPos sub, which is updated whenever the mouse moves)
Public SelectedHex1 As Point
Public SelectedHex2 As Point
Public SelectedHex3 As Point

'Mouse position
Public MousePos As Point

'Bullet position - updated only when the mouse moves
Public BulletPos As Point

'If the game engine is running or not
Public EndGameLoop As Boolean

Public Sub Main()

'************************************************************
'Called when the program is loaded - starts up the engine
'************************************************************

'Get new randomization seed

    Randomize

    'Show the main form

    Load frmMain
    frmMain.Show

    'Initialize the engine
    Engine_Init

    'Set the game variables
    Game_Init

    'Start the game loop
    Game_Loop

End Sub
