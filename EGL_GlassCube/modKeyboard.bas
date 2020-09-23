Attribute VB_Name = "modKeyboard"

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public rX, rY, rZ As Single

Private Function State(Key As Long) As Boolean
    
    Dim KeyState As Integer
    
    KeyState = GetKeyState(Key)
    State = IIf(KeyState And &H8000, True, False)

End Function

Public Sub UpdateMeshParameters()

    Const Value = 0.3
    
    With Meshs
        If State(vbKeyEscape) Then Unload frmCanvas
        If State(vbKeyX) Then Call ResetMeshParameters
        If State(vbKeyS) Then rX = rX + Value
        If State(vbKeyW) Then rX = rX - Value
        If State(vbKeyD) Then rY = rY + Value
        If State(vbKeyA) Then rY = rY - Value
        If State(vbKeyE) Then rZ = rZ + Value
        If State(vbKeyQ) Then rZ = rZ - Value
        .Rotation.X = .Rotation.X + rX: .Rotation.X = .Rotation.X Mod 360
        .Rotation.Y = .Rotation.Y + rY: .Rotation.Y = .Rotation.Y Mod 360
        .Rotation.Z = .Rotation.Z + rZ: .Rotation.Z = .Rotation.Z Mod 360
        .World = MatrixWorld()
    End With
    
End Sub

Public Sub ResetMeshParameters()

    With Meshs
        .Rotation = VertexSet(0, 0, 0)
        .Translation = VertexSet(0, 0, 0)
        .Scale = VertexSet(250, 250, 250)
        rX = 0
        rY = 0
        rZ = 0
    End With
    
End Sub

