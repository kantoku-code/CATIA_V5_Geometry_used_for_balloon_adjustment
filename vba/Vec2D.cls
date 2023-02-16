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


'**プロパティ**
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


'**メソッド-非破壊**

'クローン
Public Function clone() As Pnt2D
    Dim vec As Vec2D
    Set vec = New Vec2D
    
    vec.with_array Array(x_, y_)
    
    Set clone = vec
End Function


'ポイント化
Public Function as_point() As Pnt2D
    Dim pnt As Pnt2D
    Set pnt = New Pnt2D
    
    pnt.with_array Array(x_, y_)
    
    Set as_point = pnt
End Function


'配列化
Public Function as_array() As Variant
    as_array = Array(x_, y_)
End Function


'外積
Public Function cross(ByVal vector As Vec2D) As Double
    cross = x_ * vector.x - y_ * vector.y
End Function


'内積
Public Function dot(ByVal vector As Vec2D) As Double
    dot = x_ * vector.x + y_ * vector.y
End Function


'**メソッド-破壊**

'長さ設定-内部
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


'差
Public Sub subtract(ByVal vector As Vec2D)
    x_ = x_ - vector.x
    y_ = y_ - vector.y
    set_length
End Sub


'和
Public Sub add(ByVal vector As Vec2D)
    x_ = x_ + vector.x
    y_ = y_ + vector.y
    set_length
End Sub


'スカラー倍
Public Sub scale_by(ByVal ratio As Double)
    x_ = x_ * ratio
    y_ = y_ * ratio
    set_length
End Sub


'単位化
Public Sub normalize()
    x_ = x_ / length_
    y_ = y_ / length_
    set_length
End Sub


'配列で設定
Public Sub with_array(ByVal ary As Variant)
    x_ = ary(0)
    y_ = ary(1)
    set_length
End Sub


'文字
Public Function str() As String
    str = x_ & "," & y_
End Function

