VERSION 5.00
Begin VB.Form frmCanvas 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EGL_GlassCube"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   ClipControls    =   0   'False
   Icon            =   "frmCanvas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCanvas.frx":000C
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   50
      Left            =   1080
      Max             =   500
      Min             =   50
      SmallChange     =   25
      TabIndex        =   2
      Top             =   240
      Value           =   250
      Width           =   1815
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   1080
      Max             =   10
      TabIndex        =   1
      Top             =   0
      Value           =   5
      Width           =   1815
   End
   Begin VB.PictureBox picTex 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   0
      Picture         =   "frmCanvas.frx":6327
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Scale"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transparency"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "frmCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rX, rY As Single

Private Sub Form_Load()
    
    Dim idx As Integer
    Dim bmp As BITMAPINFO
    
    OriginX = Me.ScaleWidth / 2
    OriginY = Me.ScaleHeight / 2
    Camera = VertexSet(0, 0, 1)
    TraRate1 = 0.5
    TraRate2 = 1 - TraRate1
    
    Call CreateCube

'Texture array
    ReDim TexArray(picTex.ScaleWidth - 1, picTex.ScaleHeight - 1)
    With bmp.bmiHeader
        .biSize = Len(bmp.bmiHeader)
        .biWidth = picTex.ScaleWidth
        .biHeight = picTex.ScaleHeight
        .biPlanes = 1
        .biBitCount = 32
    End With
    Call GetDIBits(picTex.hdc, picTex.Picture.handle, 0, picTex.ScaleHeight, TexArray(0, 0), bmp, 0)

'Canvas array
    ReDim ClrArray(frmCanvas.ScaleWidth - 1, frmCanvas.ScaleHeight - 1)
    With bmp.bmiHeader
        .biSize = Len(bmp.bmiHeader)
        .biWidth = frmCanvas.ScaleWidth
        .biHeight = frmCanvas.ScaleHeight
        .biPlanes = 1
        .biBitCount = 32
    End With
    Call GetDIBits(frmCanvas.hdc, frmCanvas.Picture.handle, 0, frmCanvas.ScaleHeight, ClrArray(0, 0), bmp, 0)
    
    
    Me.Show
    
    Do
        DoEvents
        Call UpdateMeshParameters
        With Meshs
'Vertices and screen coords
            For idx = 0 To NumVertex
                .VerticesT(idx) = MatrixMultVertex(.World, .Vertices(idx))
                .Screen(idx).X = .VerticesT(idx).X + OriginX
                .Screen(idx).Y = .VerticesT(idx).Y + OriginY
            Next idx
'Faces
            For idx = 0 To NumFaces
                .Faces(idx).NormalT = MatrixMultVertex(.World, .Faces(idx).Normal)
            Next
'Render
            Call SortVisibleFaces
            PicArray = ClrArray
            For idx = 0 To UBound(.FaceV)
                Call DrawTexTriangle(.FaceV(idx).index)
            Next
            SetDIBits Me.hdc, Me.Picture.handle, 0, Me.ScaleHeight, PicArray(0, 0), bmp, 0
            Me.Refresh
        End With
    Loop
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button Then
        rX = X: rY = Y
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        With Meshs
            .Rotation.X = Y - rY: .Rotation.X = .Rotation.X Mod 360
            .Rotation.Y = X - rX: .Rotation.Y = .Rotation.Y Mod 360
        End With
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    End

End Sub

Private Sub HScroll1_Change()
    'transparency ratio
    '0.5-0.5 : 50% - 50%
    '0.3-0.7 : 30% - 70%
    TraRate1 = HScroll1.Value / 10
    TraRate2 = 1 - TraRate1

End Sub

Private Sub HScroll1_Scroll()
HScroll1_Change
End Sub

Private Sub HScroll2_Change()
   
    With Meshs
        .Scale = VertexSet(HScroll2.Value, HScroll2.Value, HScroll2.Value)
    End With

End Sub

Private Sub HScroll2_Scroll()
HScroll2_Change
End Sub
