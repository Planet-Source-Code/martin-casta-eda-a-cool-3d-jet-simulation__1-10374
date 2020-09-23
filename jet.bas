Attribute VB_Name = "Module1"

Type PIXELFORMATDESCRIPTOR
    nSize As Integer
    nVersion As Integer
    dwFlags As Long
    iPixelType As Byte
    cColorBits As Byte
    cRedBits As Byte
    cRedShift As Byte
    cGreenBits As Byte
    cGreenShift As Byte
    cBlueBits As Byte
    cBlueShift As Byte
    cAlphaBits As Byte
    cAlphaShift As Byte
    cAccumBits As Byte
    cAccumRedBits As Byte
    cAccumGreenBits As Byte
    cAccumBlueBits As Byte
    cAccumAlpgaBits As Byte
    cDepthBits As Byte
    cStencilBits As Byte
    cAuxBuffers As Byte
    iLayerType As Byte
    bReserved As Byte
    dwLayerMask As Long
    dwVisibleMask As Long
    dwDamageMask As Long
End Type

Global Const PFD_TYPE_RGBA = 0
Global Const PFD_TYPE_COLORINDEX = 1
Global Const PFD_MAIN_PLANE = 0
Global Const PFD_DOUBLEBUFFER = 1
Global Const PFD_DRAW_TO_WINDOW = &H4
Global Const PFD_SUPPORT_OPENGL = &H20
Global Const PFD_NEED_PALETTE = &H80

Declare Function ChoosePixelFormat Lib "gdi32" (ByVal hDC As Long, pfd As PIXELFORMATDESCRIPTOR) As Long
Declare Function SetPixelFormat Lib "gdi32" (ByVal hDC As Long, ByVal i As Long, pfd As PIXELFORMATDESCRIPTOR) As Boolean
Declare Sub SwapBuffers Lib "gdi32" (ByVal hDC As Long)
Declare Function wglCreateContext Lib "OpenGL32" (ByVal hDC As Long) As Long
Declare Sub wglDeleteContext Lib "OpenGL32" (ByVal hContext As Long)
Declare Sub wglMakeCurrent Lib "OpenGL32" (ByVal l1 As Long, ByVal l2 As Long)


Global xSw As Integer
Global ySw As Integer
Global zSw As Integer
