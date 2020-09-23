Attribute VB_Name = "modMatrix"
Option Explicit

Public Const sPIDiv180 As Single = 0.0174533 'PI / 180

Public Function MatrixIdentity() As MATRIX

    With MatrixIdentity
        .rc11 = 1: .rc12 = 0: .rc13 = 0: .rc14 = 0
        .rc21 = 0: .rc22 = 1: .rc23 = 0: .rc24 = 0
        .rc31 = 0: .rc32 = 0: .rc33 = 1: .rc34 = 0
        .rc41 = 0: .rc42 = 0: .rc43 = 0: .rc44 = 1
    End With

End Function

Public Function MatrixMultiply(M1 As MATRIX, M2 As MATRIX) As MATRIX

    MatrixMultiply = MatrixIdentity
    With MatrixMultiply
        .rc11 = M1.rc11 * M2.rc11 + M1.rc21 * M2.rc12 + M1.rc31 * M2.rc13 + M1.rc41 * M2.rc14
        .rc12 = M1.rc12 * M2.rc11 + M1.rc22 * M2.rc12 + M1.rc32 * M2.rc13 + M1.rc42 * M2.rc14
        .rc13 = M1.rc13 * M2.rc11 + M1.rc23 * M2.rc12 + M1.rc33 * M2.rc13 + M1.rc43 * M2.rc14
        .rc14 = M1.rc14 * M2.rc11 + M1.rc24 * M2.rc12 + M1.rc34 * M2.rc13 + M1.rc44 * M2.rc14
        .rc21 = M1.rc11 * M2.rc21 + M1.rc21 * M2.rc22 + M1.rc31 * M2.rc23 + M1.rc41 * M2.rc24
        .rc22 = M1.rc12 * M2.rc21 + M1.rc22 * M2.rc22 + M1.rc32 * M2.rc23 + M1.rc42 * M2.rc24
        .rc23 = M1.rc13 * M2.rc21 + M1.rc23 * M2.rc22 + M1.rc33 * M2.rc23 + M1.rc43 * M2.rc24
        .rc24 = M1.rc14 * M2.rc21 + M1.rc24 * M2.rc22 + M1.rc34 * M2.rc23 + M1.rc44 * M2.rc24
        .rc31 = M1.rc11 * M2.rc31 + M1.rc21 * M2.rc32 + M1.rc31 * M2.rc33 + M1.rc41 * M2.rc34
        .rc32 = M1.rc12 * M2.rc31 + M1.rc22 * M2.rc32 + M1.rc32 * M2.rc33 + M1.rc42 * M2.rc34
        .rc33 = M1.rc13 * M2.rc31 + M1.rc23 * M2.rc32 + M1.rc33 * M2.rc33 + M1.rc43 * M2.rc34
        .rc34 = M1.rc14 * M2.rc31 + M1.rc24 * M2.rc32 + M1.rc34 * M2.rc33 + M1.rc44 * M2.rc34
        .rc41 = M1.rc11 * M2.rc41 + M1.rc21 * M2.rc42 + M1.rc31 * M2.rc43 + M1.rc41 * M2.rc44
        .rc42 = M1.rc12 * M2.rc41 + M1.rc22 * M2.rc42 + M1.rc32 * M2.rc43 + M1.rc42 * M2.rc44
        .rc43 = M1.rc13 * M2.rc41 + M1.rc23 * M2.rc42 + M1.rc33 * M2.rc43 + M1.rc43 * M2.rc44
        .rc44 = M1.rc14 * M2.rc41 + M1.rc24 * M2.rc42 + M1.rc34 * M2.rc43 + M1.rc44 * M2.rc44
    End With

End Function

Public Function MatrixMultVertex(M As MATRIX, V As VERTEX) As VERTEX

    MatrixMultVertex.X = M.rc11 * V.X + M.rc12 * V.Y + M.rc13 * V.Z + M.rc14
    MatrixMultVertex.Y = M.rc21 * V.X + M.rc22 * V.Y + M.rc23 * V.Z + M.rc24
    MatrixMultVertex.Z = M.rc31 * V.X + M.rc32 * V.Y + M.rc33 * V.Z + M.rc34
    MatrixMultVertex.W = 1

End Function

Public Function MatrixWorld() As MATRIX
    
    Dim CosX As Single
    Dim SinX As Single
    Dim CosY As Single
    Dim SinY As Single
    Dim CosZ As Single
    Dim SinZ As Single
    
    With Meshs
        With .Rotation
            CosX = Cos(.X * sPIDiv180)
            SinX = Sin(.X * sPIDiv180)
            CosY = Cos(.Y * sPIDiv180)
            SinY = Sin(.Y * sPIDiv180)
            CosZ = Cos(.Z * sPIDiv180)
            SinZ = Sin(.Z * sPIDiv180)
        End With
        MatrixWorld.rc11 = .Scale.X * CosY * CosZ
        MatrixWorld.rc12 = .Scale.Y * (SinX * SinY * CosZ + CosX * -SinZ)
        MatrixWorld.rc13 = .Scale.Z * (CosX * SinY * CosZ + SinX * SinZ)
        MatrixWorld.rc14 = .Translation.X
        MatrixWorld.rc21 = .Scale.X * CosY * SinZ
        MatrixWorld.rc22 = .Scale.Y * (SinX * SinY * SinZ + CosX * CosZ)
        MatrixWorld.rc23 = .Scale.Z * (CosX * SinY * SinZ + -SinX * CosZ)
        MatrixWorld.rc24 = .Translation.Y
        MatrixWorld.rc31 = .Scale.X * -SinY
        MatrixWorld.rc32 = .Scale.Y * SinX * CosY
        MatrixWorld.rc33 = .Scale.Z * CosX * CosY
        MatrixWorld.rc34 = .Translation.Z
        MatrixWorld.rc41 = 0
        MatrixWorld.rc42 = 0
        MatrixWorld.rc43 = 0
        MatrixWorld.rc44 = 1
    End With
    
End Function

