VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Martin"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7080
   Icon            =   "jet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   472
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Roll"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pitch"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yaw"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   6120
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim hGLRC As Long

Dim xAngle As GLfloat
Dim yAngle As GLfloat
Dim zAngle As GLfloat

Sub SetupPixelFormat(ByVal hDC As Long)

    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim PixelFormat As Integer
    pfd.nSize = Len(pfd)
    pfd.nVersion = 1
    pfd.dwFlags = PFD_SUPPORT_OPENGL Or PFD_DRAW_TO_WINDOW Or PFD_DOUBLEBUFFER Or PFD_TYPE_RGBA
    pfd.iPixelType = PFD_TYPE_RGBA
    pfd.cColorBits = 16
    pfd.cDepthBits = 16
    pfd.iLayerType = PFD_MAIN_PLANE
    PixelFormat = ChoosePixelFormat(hDC, pfd)
    If PixelFormat = 0 Then MsgBox ("Could not retrieve pixel format!")
    SetPixelFormat hDC, PixelFormat, pfd
End Sub


Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     ySw = 1
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     ySw = 0
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     xSw = 1
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     xSw = 0
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     zSw = 1
End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     zSw = 0
End Sub


Private Sub Form_Load()
   
    Dim hGLRC As Long
    
    xAngle = 0
    yAngle = 0
    zAngle = 0
    SetupPixelFormat hDC
    hGLRC = wglCreateContext(hDC)
    wglMakeCurrent hDC, hGLRC
    glEnable GL_DEPTH_TEST
    glFrontFace GL_CCW
             
    Dim LightPos(3) As GLfloat
    Dim SpecRef(3) As GLfloat
        
    LightPos(0) = 0
    LightPos(1) = 0
    LightPos(2) = 0
    LightPos(3) = 1
    SpecRef(0) = 0.5
    SpecRef(1) = 0.5
    SpecRef(2) = 0.5
    SpecRef(3) = 0.5
        
    glEnable GL_LIGHTING
    glLightfv GL_LIGHT0, GL_POSITION, LightPos(0)
          
    glEnable GL_LIGHT0
    
    glEnable GL_COLOR_MATERIAL
    glColorMaterial GL_FRONT, GL_AMBIENT_AND_DIFFUSE
    glMaterialfv GL_FRONT, GL_SPECULAR, SpecRef(0)
    glMateriali GL_FRONT, GL_SHININESS, 128
    
    glClearColor 0, 0, 0, 0
    
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
End Sub
Private Sub Form_Paint()
         
        glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
        
        glLoadIdentity
  
        glScalef 0.03, 0.03, 0.03
        glTranslatef 0, 0, -3
        glRotatef xAngle, 1, 0, 0
        glRotatef yAngle, 0, 1, 0
        glRotatef zAngle, 0, 0, 1
        
        glBegin GL_TRIANGLES
        'Nose
        
        glColor3f 0, 0, 1
        glVertex3f 0, 0, -30
        glVertex3f 0, 10, -10
        glVertex3f -10, 0, -10
        
        glVertex3f 0, 0, -30
        glVertex3f 10, 0, -10
        glVertex3f 0, 10, -10
        
        glVertex3f 0, 0, -30
        glVertex3f -10, 0, -10
        glVertex3f 10, 0, -10
        
        'Cockpit
        glColor3f 0, 0.03, 0.03
        glVertex3f 0, 5, -20
        glVertex3f 5, 5, -10
        glVertex3f 0, 11, -10
        
        glVertex3f 0, 5, -20
        glVertex3f 0, 11, -10
        glVertex3f -5, 5, -10
                
        glVertex3f 0, 11, -10
        glVertex3f 5, 5, -10
        glVertex3f 0, 10, 0
        
        glVertex3f 0, 11, -10
        glVertex3f 0, 10, 0
        glVertex3f -5, 5, -10
         
        'Jet
        glColor3f 225, 0, 0
        glVertex3f 0, 10, 17
        glVertex3f 10, 0, 17
        glVertex3f -10, 0, 17
                
        glEnd
        
        glBegin GL_QUADS
        
        'Body
        glColor3f 0, 0, 3
        glVertex3f -10, 0, -10
        glVertex3f 10, 0, -10
        glVertex3f 10, 0, 20
        glVertex3f -10, 0, 20
        
        glVertex3f -10, 0, -10
        glVertex3f 0, 10, -10
        glVertex3f 0, 10, 20
        glVertex3f -10, 0, 20
        
        glVertex3f 0, 10, -10
        glVertex3f 10, 0, -10
        glVertex3f 10, 0, 20
        glVertex3f 0, 10, 20
       
        'Wings
        glColor3f 0, 0, 3
        glVertex3f 10, 0, -10
        glVertex3f 25, 0, 0
        glVertex3f 25, 0, 10
        glVertex3f 10, 0, 10
        
        glVertex3f -10, 0, -10
        glVertex3f -10, 0, 10
        glVertex3f -25, 0, 10
        glVertex3f -25, 0, 0
       
        'Tail
        glVertex3f 10, 0, 10
        glVertex3f 15, 0, 15
        glVertex3f 15, 0, 20
        glVertex3f 10, 0, 20
        
        glVertex3f -10, 0, 10
        glVertex3f -10, 0, 20
        glVertex3f -15, 0, 20
        glVertex3f -15, 0, 15
        
        glVertex3f 0, 10, 10
        glVertex3f 0, 20, 15
        glVertex3f 0, 20, 20
        glVertex3f 0, 10, 20
          
        glEnd
        
        SwapBuffers hDC
        
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If hGLRC <> 0 Then
        wglMakeCurrent 0, 0
        wglDeleteContext hGLRC
    End If

End Sub

Private Sub Timer1_Timer()
     If xSw = 1 Then
           xAngle = xAngle + 5
           If xAngle > 360 Then xAngle = 0
           Form_Paint
     End If
     If ySw = 1 Then
           yAngle = yAngle + 5
           If yAngle > 360 Then yAngle = 0
           Form_Paint
     End If
     If zSw = 1 Then
           zAngle = zAngle + 5
           If zAngle > 360 Then zAngle = 0
           Form_Paint
     End If
End Sub
