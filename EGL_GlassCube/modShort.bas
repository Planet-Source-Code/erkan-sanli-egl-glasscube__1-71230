Attribute VB_Name = "modShort"
Option Explicit

Public Function SortVisibleFaces() As Integer
    
    Dim i As Long
    Dim iV As Long
    With Meshs
        iV = -1
        Erase .FaceV
        For i = 0 To NumFaces
            'If IIf(DotProduct(.NormalsT(i), Camera) > 0, True, False) Then
                iV = iV + 1
                ReDim Preserve .FaceV(iV)
                .FaceV(iV).ZValue = ( _
                    .VerticesT(.Faces(i).A).Z + _
                    .VerticesT(.Faces(i).B).Z + _
                    .VerticesT(.Faces(i).C).Z)
                .FaceV(iV).index = i
            'End If
        Next
        If iV > -1 Then SortFaces 0, iV
        SortVisibleFaces = iV
    End With

End Function

Private Sub SortFaces(ByVal First As Long, ByVal Last As Long)

    Dim FirstIdx  As Long
    Dim MidIdx As Long
    Dim LastIdx  As Long
    Dim MidVal As Single
    Dim TempOrder  As ORDER
    
    If (First < Last) Then
        With Meshs
            MidIdx = (First + Last) \ 2
            MidVal = .FaceV(MidIdx).ZValue
            FirstIdx = First
            LastIdx = Last
            Do
                Do While .FaceV(FirstIdx).ZValue < MidVal
                    FirstIdx = FirstIdx + 1
                Loop
                Do While .FaceV(LastIdx).ZValue > MidVal
                    LastIdx = LastIdx - 1
                Loop
                If (FirstIdx <= LastIdx) Then
                    TempOrder = .FaceV(LastIdx)
                    .FaceV(LastIdx) = .FaceV(FirstIdx)
                    .FaceV(FirstIdx) = TempOrder
                    FirstIdx = FirstIdx + 1
                    LastIdx = LastIdx - 1
                End If
            Loop Until FirstIdx > LastIdx

            If (LastIdx <= MidIdx) Then
                SortFaces First, LastIdx
                SortFaces FirstIdx, Last
            Else
                SortFaces FirstIdx, Last
                SortFaces First, LastIdx
            End If
        End With
    End If

End Sub


