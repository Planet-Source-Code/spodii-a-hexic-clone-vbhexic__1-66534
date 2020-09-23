VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBHexic Settings"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton RestoreCmd 
      Caption         =   "Restore Settings"
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton SaveCmd 
      Caption         =   "Save Settings"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   4680
      Width           =   1335
   End
   Begin VB.ComboBox cmbBehaviorFlags 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4200
      Width           =   4455
   End
   Begin VB.ComboBox cmbAdapters 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   4455
   End
   Begin VB.ComboBox cmbDevice 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   4455
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label lblmisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3D Device Settings:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   1725
   End
   Begin VB.Label lblmisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hardware Adapters Available:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblmisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rendering Devices Available:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblmisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Basic Hardware Capabilities"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2385
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Dx As DirectX8
Dim D3D As Direct3D8

Dim BehaviorFlagsVal As Long                'Value of the Behavior Flags
Dim nAdapters As Long                       'How many adapters found
Dim nModes As Long                          'How many display modes found
Dim AdapterInfo As D3DADAPTER_IDENTIFIER8   'Information on the adaptor
Dim HardwareRenderer As Long
Dim AdapterVal As Long

Private Sub cmbBehaviorFlags_Click()

    'Set the value of behavior flags
    Select Case cmbBehaviorFlags.ListIndex
        Case 0
            BehaviorFlagsVal = D3DCREATE_SOFTWARE_VERTEXPROCESSING
        Case 1
            BehaviorFlagsVal = D3DCREATE_HARDWARE_VERTEXPROCESSING
        Case 2
            BehaviorFlagsVal = D3DCREATE_MIXED_VERTEXPROCESSING
        Case 3
            BehaviorFlagsVal = D3DCREATE_PUREDEVICE
        Case 4
            BehaviorFlagsVal = D3DCREATE_FPU_PRESERVE
        Case 5
            BehaviorFlagsVal = D3DCREATE_MULTITHREADED
    End Select
    
End Sub

Private Sub Form_Load()
    Me.Show
    
    'Create the DirectX 8 devices
    Set Dx = New DirectX8
    Set D3D = Dx.Direct3DCreate
    
    'Get the adapter information
    EnumerateAdapters
    
    'Get the device information
    EnumerateDevices
    
    'Set the behavior flag values
    cmbBehaviorFlags.AddItem "Software Vertex Processing"
    cmbBehaviorFlags.AddItem "Hardware Vertex Processing"
    cmbBehaviorFlags.AddItem "Mixed Vertex Processing"
    cmbBehaviorFlags.AddItem "Pure Device"
    cmbBehaviorFlags.AddItem "FPU Preserve"
    cmbBehaviorFlags.AddItem "Multithreaded"
    
    'Load current settings
    RestoreCmd_Click

End Sub

Private Sub EnumerateAdapters()
    Dim sTemp As String
    Dim I As Integer
    Dim J As Integer
    
    'Will be either 1 or 2 adapters found
    nAdapters = D3D.GetAdapterCount
    
    For I = 0 To nAdapters - 1
    
        'Get the relevent details
        D3D.GetAdapterIdentifier I, 0, AdapterInfo
        
        'Decode adapter name
        sTemp = ""
        For J = 0 To 511
            sTemp = sTemp & Chr$(AdapterInfo.Description(J))
        Next J
        sTemp = Replace(sTemp, Chr$(0), " ")
        cmbAdapters.AddItem sTemp
        
    Next I

End Sub

Private Sub EnumerateDevices()
On Local Error Resume Next
Dim Caps As D3DCAPS8

    'Get the device caps
    D3D.GetDeviceCaps cmbAdapters.ListIndex, D3DDEVTYPE_HAL, Caps
    
    'If there is an error, there is no hardware acceleration
    If Err.Number = D3DERR_NOTAVAILABLE Then
        cmbDevice.AddItem "Reference Rasterizer (REF)"  'Reference device will always be available
    Else    'Add hardware acceleration to the list
        cmbDevice.AddItem "Hardware Acceleration (HAL)"
        cmbDevice.AddItem "Reference Rasterizer (REF)"
    End If
    
End Sub

Private Sub EnumerateHardware(Renderer As Long)
Dim Caps As D3DCAPS8

    'Clear the list
    List1.Clear
    
    'Get the capibilities
    HardwareRenderer = Renderer
    D3D.GetDeviceCaps cmbAdapters.ListIndex, Renderer, Caps

    'Display all the information
    List1.AddItem "Maximum Point Vertex size: " & Caps.MaxPointSize
    List1.AddItem "Maximum Texture Size: " & Caps.MaxTextureWidth & "x" & Caps.MaxTextureHeight
    List1.AddItem "Maximum Primatives in one call: " & Caps.MaxPrimitiveCount
    
    If Caps.TextureCaps And D3DPTEXTURECAPS_SQUAREONLY Then
        List1.AddItem "Textures must always be square"
    End If

    If Caps.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT Then
        List1.AddItem "Device supports hardware transform and lighting"
    End If
    
    If Caps.DevCaps And D3DDEVCAPS_HWRASTERIZATION Then
        List1.AddItem "Device can use Hardware Rasterization"
    End If

    If Caps.Caps2 And D3DCAPS2_CANRENDERWINDOWED Then
        List1.AddItem "Device can Render in Windowed Mode"
    End If
    
    If Caps.RasterCaps And D3DPRASTERCAPS_ANISOTROPY Then
        List1.AddItem "Device supports Anisotropic Filtering"
    End If
    
    If Caps.RasterCaps And D3DPRASTERCAPS_ZBUFFERLESSHSR Then
        List1.AddItem "Device does not require a Z-Buffer/Depth Buffer"
    End If

End Sub

Private Sub cmbDevice_Click()
    If UCase(Left(cmbDevice.Text, 3)) = "REF" Then
        EnumerateHardware 2 'Reference device
    Else
        EnumerateHardware 1 'Hardware device
    End If
End Sub

Private Sub RestoreCmd_Click()
Dim FileNum As Byte
Dim TempInt As Integer

    'Restore settings
    FileNum = FreeFile
    Open App.Path & "/config.ini" For Binary As #FileNum
        Get #FileNum, , BehaviorFlagsVal
        Get #FileNum, , HardwareRenderer
        Get #FileNum, , AdapterVal
        
        Get #FileNum, , TempInt
        cmbAdapters.ListIndex = TempInt
        Get #FileNum, , TempInt
        cmbDevice.ListIndex = TempInt
        Get #FileNum, , TempInt
        cmbBehaviorFlags.ListIndex = TempInt
        
    Close #FileNum
    
End Sub

Private Sub SaveCmd_Click()

    'Save settings
    Open App.Path & "/config.ini" For Binary As #1
        Put #1, , BehaviorFlagsVal
        Put #1, , HardwareRenderer
        Put #1, , AdapterVal
        Put #1, , cmbAdapters.ListIndex
        Put #1, , cmbDevice.ListIndex
        Put #1, , cmbBehaviorFlags.ListIndex
    Close #1

End Sub
