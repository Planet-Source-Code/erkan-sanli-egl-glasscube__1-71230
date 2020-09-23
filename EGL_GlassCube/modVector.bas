Attribute VB_Name = "modVector"
Option Explicit

Public Function VertexSet(X As Single, Y As Single, Z As Single) As VERTEX

    VertexSet.X = X
    VertexSet.Y = Y
    VertexSet.Z = Z

End Function

Public Function VertexSub(V1 As VERTEX, V2 As VERTEX) As VERTEX

    VertexSub.X = V1.X - V2.X
    VertexSub.Y = V1.Y - V2.Y
    VertexSub.Z = V1.Z - V2.Z
    VertexSub.W = 1

End Function

'Public Function VertexAdd(V1 As VERTEX, V2 As VERTEX) As VERTEX
'
'    VertexAdd.X = V1.X + V2.X
'    VertexAdd.Y = V1.Y + V2.Y
'    VertexAdd.Z = V1.Z + V2.Z
'    VertexAdd.W = 1
'
'End Function

'Public Function VertexScale(V As VERTEX, S As Single) As VERTEX
'
'    VertexScale.X = V.X * S
'    VertexScale.Y = V.Y * S
'    VertexScale.Z = V.Z * S
'    VertexScale.W = 1
'
'End Function

Public Function CrossProduct(V1 As VERTEX, V2 As VERTEX) As VERTEX
     
    CrossProduct.X = (V1.Y * V2.Z) - (V1.Z * V2.Y)
    CrossProduct.Y = (V1.Z * V2.X) - (V1.X * V2.Z)
    CrossProduct.Z = (V1.X * V2.Y) - (V1.Y * V2.X)
    CrossProduct.W = 1

End Function

Public Function VertexNormalize(V As VERTEX) As VERTEX

    Dim VLength As Single
    
    VLength = Sqr((V.X * V.X) + (V.Y * V.Y) + (V.Z * V.Z))
    If VLength = 0 Then VLength = 1
    VertexNormalize.X = V.X / VLength
    VertexNormalize.Y = V.Y / VLength
    VertexNormalize.Z = V.Z / VLength
    VertexNormalize.W = 1

End Function

'Public Function VertexLength(V As VERTEX) As Single
'
'    VertexLength = Sqr((V.X * V.X) + (V.Y * V.Y) + (V.Z * V.Z))
'
'End Function

Public Function DotProduct(V1 As VERTEX, V2 As VERTEX) As Single

    DotProduct = (V1.X * V2.X) + _
                 (V1.Y * V2.Y) + _
                 (V1.Z * V2.Z)

End Function


