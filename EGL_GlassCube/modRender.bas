Attribute VB_Name = "modRender"
Option Explicit

Private Type TEXEL
    Y1      As Single
    U1      As Single
    V1      As Single
    Y2      As Single
    U2      As Single
    V2      As Single
    Used    As Boolean
End Type

Dim Texels() As TEXEL

Public Sub DrawTexTriangle(idx As Integer)
    
    Dim StartX As Single, EndX As Single
    Dim X1 As Long, Y1 As Long
    Dim X2 As Long, Y2 As Long
    Dim X3 As Long, Y3 As Long
    Dim X4 As Long, Y4 As Long
        
    With Meshs
        X1 = .Screen(.Faces(idx).A).X
        Y1 = .Screen(.Faces(idx).A).Y
        X2 = .Screen(.Faces(idx).B).X
        Y2 = .Screen(.Faces(idx).B).Y
        X3 = .Screen(.Faces(idx).C).X
        Y3 = .Screen(.Faces(idx).C).Y
        X4 = .Screen(.Faces(idx).D).X
        Y4 = .Screen(.Faces(idx).D).Y
        
        If X1 < X2 Then
            StartX = X1:    EndX = X2
        Else
            StartX = X2:    EndX = X1
        End If
        If X3 < StartX Then StartX = X3
        If X4 < StartX Then StartX = X4
        If X3 > EndX Then EndX = X3
        If X4 > EndX Then EndX = X4
        ReDim Texels(StartX To EndX)
        AffineTexLine X1, Y1, X2, Y2, 0, 0, 199, 0
        AffineTexLine X2, Y2, X3, Y3, 199, 0, 199, 199
        AffineTexLine X3, Y3, X4, Y4, 199, 199, 0, 199
        AffineTexLine X4, Y4, X1, Y1, 0, 199, 0, 0
    End With
    
    If StartX < 0 Then StartX = 0
    If EndX > frmCanvas.ScaleWidth - 1 Then EndX = frmCanvas.ScaleWidth - 1
    
    For StartX = StartX To EndX
        With Texels(StartX)
            FillTexLine .Y1, StartX, .Y2, .U1, .V1, .U2, .V2
        End With
    Next
    
End Sub

Private Function Ratio(ByVal R1 As Single, ByVal R2 As Single) As Single

    If R2 = 0 Then
        Ratio = 0
    Else
        Ratio = R1 / R2
    End If

End Function

Private Sub FillTexLine(ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, _
                        ByVal U1 As Single, ByVal V1 As Single, ByVal U2 As Single, ByVal V2 As Single)
    
    Dim DeltaY  As Single
    Dim StepU   As Single, StepV  As Single
    Dim StartY As Single, EndY As Single
    Dim TexPix As Long
    
    On Error Resume Next
    
    DeltaY = Y1 - Y2
    StepU = Ratio(U1 - U2, DeltaY)
    StepV = Ratio(V1 - V2, DeltaY)
    
    If Y2 < 0 Then
        StartY = 0
        U2 = U2 + (StepU * Abs(Y2))
        V2 = V2 + (StepV * Abs(Y2))
    Else
        StartY = Y2
    End If

    EndY = IIf(Y1 > frmCanvas.ScaleHeight - 1, frmCanvas.ScaleHeight - 1, Y1)
    
    For StartY = StartY To EndY
        PicArray(X2, StartY).rgbRed = (PicArray(X2, StartY).rgbRed * TraRate1) + (TexArray(U2, V2).rgbRed * TraRate2)
        PicArray(X2, StartY).rgbGreen = (PicArray(X2, StartY).rgbGreen * TraRate1) + (TexArray(U2, V2).rgbGreen * TraRate2)
        PicArray(X2, StartY).rgbBlue = (PicArray(X2, StartY).rgbBlue * TraRate1) + (TexArray(U2, V2).rgbBlue * TraRate2)
        U2 = U2 + StepU
        V2 = V2 + StepV
    Next
    
End Sub

Private Sub AffineTexLine(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, _
                          ByVal U1 As Single, ByVal V1 As Single, ByVal U2 As Single, ByVal V2 As Single)
    
    Dim DeltaX As Single
    Dim StepY, StepU, StepV As Single
        
    If X1 < X2 Then
        DeltaX = X2 - X1
        StepY = Ratio(Y2 - Y1, DeltaX)
        StepU = Ratio(U2 - U1, DeltaX)
        StepV = Ratio(V2 - V1, DeltaX)
        For X1 = X1 To X2
            With Texels(X1)
                If .Used Then
                    If .Y1 < Fix(Y1) Then .Y1 = Fix(Y1): .U1 = U1:  .V1 = V1
                    If .Y2 > Fix(Y1) Then .Y2 = Fix(Y1): .U2 = U1:  .V2 = V1
                Else
                    .Y1 = Fix(Y1): .U1 = U1:  .V1 = V1
                    .Y2 = Fix(Y1): .U2 = U1:  .V2 = V1
                    .Used = True
                End If
            End With
            Y1 = Y1 + StepY
            U1 = U1 + StepU
            V1 = V1 + StepV
        Next
    Else
        DeltaX = X1 - X2
        StepY = Ratio(Y1 - Y2, DeltaX)
        StepU = Ratio(U1 - U2, DeltaX)
        StepV = Ratio(V1 - V2, DeltaX)
        For X2 = X2 To X1
            With Texels(X2)
                If .Used Then
                    If .Y1 < Fix(Y2) Then .Y1 = Fix(Y2): .U1 = U2:  .V1 = V2
                    If .Y2 > Fix(Y2) Then .Y2 = Fix(Y2): .U2 = U2:  .V2 = V2
                Else
                    .Y1 = Fix(Y2): .U1 = U2:  .V1 = V2
                    .Y2 = Fix(Y2): .U2 = U2:  .V2 = V2
                    .Used = True
                End If
            End With
            Y2 = Y2 + StepY
            U2 = U2 + StepU
            V2 = V2 + StepV
        Next
   End If
End Sub
