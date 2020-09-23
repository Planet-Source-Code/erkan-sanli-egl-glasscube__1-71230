Attribute VB_Name = "modData"
Option Explicit
Public Const NumVertex  As Integer = 7
Public Const NumFaces   As Integer = 5

Public Type VERTEX
    X                   As Single
    Y                   As Single
    Z                   As Single
    W                   As Single
End Type

Public Type FACE
    A                   As Integer
    B                   As Integer
    C                   As Integer
    D                   As Integer
    Normal              As VERTEX
    NormalT             As VERTEX
End Type

Public Type ORDER
    ZValue              As Single
    index            As Integer
End Type

Public Type POINTAPI
    X                   As Long
    Y                   As Long
End Type

Public Type MATRIX
    rc11 As Single: rc12 As Single: rc13 As Single: rc14 As Single
    rc21 As Single: rc22 As Single: rc23 As Single: rc24 As Single
    rc31 As Single: rc32 As Single: rc33 As Single: rc34 As Single
    rc41 As Single: rc42 As Single: rc43 As Single: rc44 As Single
End Type

Public Type MESH
    Vertices(NumVertex) As VERTEX
    VerticesT(NumVertex) As VERTEX
    Screen(NumVertex)   As POINTAPI
    Faces(NumFaces)     As FACE
    FaceV()             As ORDER
    Scale               As VERTEX
    Rotation            As VERTEX
    Translation         As VERTEX
    World               As MATRIX
End Type

Public Meshs            As MESH
Public Camera           As VERTEX

Public TexArray()       As RGBQUAD
Public PicArray()       As RGBQUAD
Public ClrArray()       As RGBQUAD

Public OriginX As Long
Public OriginY As Long
Public TraRate1 As Single 'transparency ratio
Public TraRate2 As Single

Public Sub CreateCube()
    
    Dim idx As Integer
    
    'Cube lenght 1 unit
    With Meshs
        .Vertices(0).X = -0.5:   .Vertices(0).Y = -0.5:   .Vertices(0).Z = 0.5:     .Vertices(0).W = 1
        .Vertices(1).X = 0.5:    .Vertices(1).Y = -0.5:   .Vertices(1).Z = 0.5:     .Vertices(1).W = 1
        .Vertices(2).X = 0.5:    .Vertices(2).Y = 0.5:    .Vertices(2).Z = 0.5:     .Vertices(2).W = 1
        .Vertices(3).X = -0.5:   .Vertices(3).Y = 0.5:    .Vertices(3).Z = 0.5:     .Vertices(3).W = 1
        .Vertices(4).X = -0.5:   .Vertices(4).Y = -0.5:   .Vertices(4).Z = -0.5:    .Vertices(4).W = 1
        .Vertices(5).X = 0.5:    .Vertices(5).Y = -0.5:   .Vertices(5).Z = -0.5:    .Vertices(5).W = 1
        .Vertices(6).X = 0.5:    .Vertices(6).Y = 0.5:    .Vertices(6).Z = -0.5:    .Vertices(6).W = 1
        .Vertices(7).X = -0.5:   .Vertices(7).Y = 0.5:    .Vertices(7).Z = -0.5:    .Vertices(7).W = 1
        
        .Faces(0).A = 0:        .Faces(0).B = 1:        .Faces(0).C = 2:          .Faces(0).D = 3
        .Faces(1).A = 1:        .Faces(1).B = 5:        .Faces(1).C = 6:          .Faces(1).D = 2
        .Faces(2).A = 5:        .Faces(2).B = 4:        .Faces(2).C = 7:          .Faces(2).D = 6
        .Faces(3).A = 4:        .Faces(3).B = 0:        .Faces(3).C = 3:          .Faces(3).D = 7
        .Faces(4).A = 3:        .Faces(4).B = 2:        .Faces(4).C = 6:          .Faces(4).D = 7
        .Faces(5).A = 1:        .Faces(5).B = 0:        .Faces(5).C = 4:          .Faces(5).D = 5
        
        For idx = 0 To NumFaces
            .Faces(idx).Normal = VertexNormalize( _
                                 CrossProduct( _
                                 VertexSub(.Vertices(.Faces(idx).C), .Vertices(.Faces(idx).B)), _
                                 VertexSub(.Vertices(.Faces(idx).A), .Vertices(.Faces(idx).B))))
        Next
        .Scale = VertexSet(250, 250, 250)
    End With
    
End Sub

