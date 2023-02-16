VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pnt2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pnt2D

Option Explicit

Private x_ As Double
Private y_ As Double


Private Sub Class_Initialize()
    x_ = 0
    y_ = 0
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


'**メソッド**
'2点間ベクトル
Public Function vector_to(ByVal point As Pnt2D) As Vec2D
    Dim vec As Vec2D
    Set vec = New Vec2D
    
    vec.with_array Array(point.x - x_, point.y - y_)
    
    Set vector_to = vec
End Function


'2点間距離
Public Function distance_to(ByVal point As Pnt2D) As Double
    distance_to = Sqr(point.x - x_ + point.y - y_)
End Function


'クローン
Public Function clone() As Pnt2D
    Dim pnt As Pnt2D
    Set pnt = New Pnt2D
    pnt.with_array (Array(x_, y_))

    Set clone = pnt
End Function


'ベクトル化
Public Function as_vector() As Vec2D
    Dim vec As Vec2D
    Set vec = New Vec2D
    
    vec.with_array Array(x_, y_)
    
    Set as_vector = vec
End Function


'配列化
Public Function as_array() As Variant
    as_array = Array(x_, y_)
End Function


'移動
Public Sub translate_by(ByVal vector As Vec2D)
    x_ = x_ + vector.x
    y_ = y_ + vector.y
End Sub



'配列で設定
Public Sub with_array(ByVal ary As Variant)
    If Not is_point2d(ary) Then
        Err.Raise Number:=600, Description:="配列が不正です"
    End If

    x_ = ary(0)
    y_ = ary(1)
End Sub


'文字
Public Function str() As String
    str = x_ & "," & y_
End Function


'ベクトル2Dとして扱える配列か?
Private Function is_point2d( _
    ary As Variant) _
    As Boolean

    Select Case True
        Case Not IsArray(ary)
            is_point2d = False
        Case UBound(ary) <> 1
            is_point2d = False
        Case Not IsNumeric(ary(0))
            is_point2d = False
        Case Not IsNumeric(ary(1))
            is_point2d = False
        Case Else
            is_point2d = True
    End Select

End Function

