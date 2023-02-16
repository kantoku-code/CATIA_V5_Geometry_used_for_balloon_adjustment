VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BBox2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'vba BBox2D.cls

Option Explicit

Private points_ As Variant
Private center_ As Pnt2D
Private origin_ As Pnt2D

Private Sub Class_Initialize()
    points_ = Empty
End Sub


Private Sub Class_Terminate()
    
End Sub


'**�v���p�e�B**
Public Property Get center_point() As Pnt2D
    Set center_point = center_
End Property


Public Property Get origin_point() As Pnt2D
    Set origin_point = origin_
End Property


'**���\�b�h-��j��**

'�z��
Public Function as_array() As Variant
    Dim ary As Variant
    ary = points_(0).as_array()
    
    extend_array ary, points_(1).as_array()
    extend_array ary, points_(2).as_array()
    extend_array ary, points_(3).as_array()

    as_array = ary
End Function


'����
Public Function str() As String
    str = _
        points_(0).str() & "," & _
        points_(1).str() & "," & _
        points_(2).str() & "," & _
        points_(3).str()

End Function


'���S�_�Ɗ�_
Private Sub set_center()
    Dim center As Pnt2D
    Set center = New Pnt2D
    
    center.with_array Array( _
        (points_(0).x + points_(3).x) * 0.5, _
        (points_(0).y + points_(3).y) * 0.5 _
    )
    
    Set center_ = center

End Sub


'**���\�b�h-�j��**

'�ړ�
Public Sub translate_by(ByVal vector As Vec2D)
    Dim i As Long
    For i = 0 To UBound(points_)
        points_(i).translate_by vector
    Next

    set_center
End Sub


'�z��Őݒ�
Public Sub with_array(ByVal ary As Variant)
    If Not is_boundary_box2d(ary) Then
        Err.Raise Number:=600, Description:="�z�񂪕s���ł�"
    End If

    Dim pnts(3) As Variant

    Dim pnt As Pnt2D
    Set pnt = New Pnt2D

    Dim i As Long
    For i = 0 To UBound(ary) Step 2
        pnt.with_array Array(ary(i), ary(i + 1))
        Set pnts(i / 2) = pnt.clone()
    Next
    
    points_ = pnts

    set_center
End Sub


'�r���[�݂̂Ŏg�p����r���[�̌��_
Public Sub set_origin_with_array(ByVal ary As Variant)
    Dim pnt As Pnt2D
    Set pnt = New Pnt2D
    pnt.with_array ary
    
    Set origin_ = pnt
End Sub


'BBox2D�Ƃ��Ĉ�����z��?
Private Function is_boundary_box2d( _
    ary As Variant) _
    As Boolean

    is_boundary_box2d = True

    Select Case True
        Case Not IsArray(ary)
            is_boundary_box2d = False
        Case UBound(ary) <> 7
            is_boundary_box2d = False
    End Select
    
    If Not is_boundary_box2d Then
        Exit Function
    End If
    
    Dim i As Long
    For i = 0 To UBound(ary)
        If Not IsNumeric(ary(i)) Then
            is_boundary_box2d = False
            Exit Function
        End If
    Next

End Function


'�z��̌���
Private Sub extend_array(ByRef ary1 As Variant, ByVal ary2 As Variant)
    Dim ary1Count As Long
    ary1Count = UBound(ary1)
    
    ReDim Preserve ary1(ary1Count + UBound(ary2) + 1)
    
    Dim i As Long
    For i = 0 To UBound(ary2)
        ary1(ary1Count + 1 + i) = ary2(i)
    Next

End Sub
