VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vec2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Vec2D

Option Explicit

Private x_ As Double
Private y_ As Double
Private length_ As Double
Private zero_ As Boolean

Private Sub Class_Initialize()
    x_ = 0
    y_ = 0
    length_ = 0
    zero_ = True
End Sub


Private Sub Class_Terminate()
    
End Sub


'**�v���p�e�B**
Public Property Get x() As Double
    x = x_
End Property


Public Property Let x(ByVal value As Double)
    x_ = value
End Property


Public Property Get y() As Double
    y = y_
End Property


Public Property Let y(ByVal value As Double)
    y_ = value
End Property


Public Property Get length() As Double
    length = length_
End Property


Public Property Get is_zero() As Double
    is_zero = zero_
End Property


'**���\�b�h-��j��**

'�N���[��
Public Function clone() As Pnt2D
    Dim vec As Vec2D
    Set vec = New Vec2D
    
    vec.with_array Array(x_, y_)
    
    Set clone = vec
End Function


'�|�C���g��
Public Function as_point() As Pnt2D
    Dim pnt As Pnt2D
    Set pnt = New Pnt2D
    
    pnt.with_array Array(x_, y_)
    
    Set as_point = pnt
End Function


'�z��
Public Function as_array() As Variant
    as_array = Array(x_, y_)
End Function


'�O��
Public Function cross(ByVal vector As Vec2D) As Double
    cross = x_ * vector.x - y_ * vector.y
End Function


'����
Public Function dot(ByVal vector As Vec2D) As Double
    dot = x_ * vector.x + y_ * vector.y
End Function


'**���\�b�h-�j��**

'�����ݒ�-����
Private Sub set_length()
    Dim length As Double
    length = Sqr(x_ * x_ + y_ * y_)
    
    If length <= 0 Then
        zero_ = True
    Else
        zero_ = False
    End If

    length_ = length
End Sub


'��
Public Sub subtract(ByVal vector As Vec2D)
    x_ = x_ - vector.x
    y_ = y_ - vector.y
    set_length
End Sub


'�a
Public Sub add(ByVal vector As Vec2D)
    x_ = x_ + vector.x
    y_ = y_ + vector.y
    set_length
End Sub


'�X�J���[�{
Public Sub scale_by(ByVal ratio As Double)
    x_ = x_ * ratio
    y_ = y_ * ratio
    set_length
End Sub


'�P�ʉ�
Public Sub normalize()
    x_ = x_ / length_
    y_ = y_ / length_
    set_length
End Sub


'�z��Őݒ�
Public Sub with_array(ByVal ary As Variant)
    x_ = ary(0)
    y_ = ary(1)
    set_length
End Sub


'����
Public Function str() As String
    str = x_ & "," & y_
End Function
