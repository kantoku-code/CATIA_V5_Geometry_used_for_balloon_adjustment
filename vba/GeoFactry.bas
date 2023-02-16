Attribute VB_Name = "GeoFactry"
'vba

Option Explicit

Sub CATMain()

End Sub


'寸法からバウンダリボックス
Function create_boundary_box_by_dimension( _
    ByVal drawDim As DrawingDimension) _
    As BBox2D

    Dim view As DrawingView
    Set view = KCL.GetParent_Of_T(drawDim, "DrawingView")

    Dim viewScale As Double
    viewScale = view.Scale

    Dim valDim As Variant
    Set valDim = drawDim
    
    Dim ary(7) As Variant
    valDim.GetBoundaryBox ary

    Dim aryCount As Long
    aryCount = UBound(ary) \ 2

    Dim i As Long
    For i = 0 To aryCount
        ary(i * 2) = ary(i * 2) * viewScale
        ary(i * 2 + 1) = ary(i * 2 + 1) * viewScale
    Next
    
    Dim bBox As BBox2D
    Set bBox = New BBox2D
    bBox.with_array ary
    
    Set create_boundary_box_by_dimension = bBox

End Function


'ビューからバウンダリボックス
Function create_boundary_box_by_view( _
    ByVal view As DrawingView) _
    As BBox2D

    set_view_axis_visible view, False

    Dim variView As Variant
    Set variView = view
    
    Dim size(3) As Variant
    variView.size size

    Dim bBox As BBox2D
    Set bBox = New BBox2D

    bBox.with_array Array( _
        size(0), _
        size(2), _
        size(1), _
        size(2), _
        size(0), _
        size(3), _
        size(1), _
        size(3) _
    )
    
    bBox.set_origin_with_array Array( _
        view.xAxisData, _
        view.yAxisData _
    )

    Set create_boundary_box_by_view = bBox

    set_view_axis_visible view, True
    
End Function


'ビューの原点矢印の表示/非表示
Private Sub set_view_axis_visible( _
    ByVal view As DrawingView, _
    ByVal isShow As Boolean)

    Dim showAttr As Long
    If isShow Then
        showAttr = catVisPropertyShowAttr
    Else
        showAttr = catVisPropertyNoShowAttr
    End If

    Dim dDoc As DrawingDocument
    Set dDoc = KCL.GetParent_Of_T(view, "DrawingDocument")

    Dim sel As Selection
    Set sel = dDoc.Selection

    Dim searchWord As String
    searchWord = "(" & _
        "CATSketchSearch.2DAxis_HDirection" & _
        " + " & _
        "CATSketchSearch.2DAxis_VDirection" & _
        ")"

    view.Activate

    sel.Clear
    sel.add view
    sel.Search searchWord & ",sel"

    Dim vis As VisPropertySet
    Set vis = sel.VisProperties
    vis.SetShow showAttr

    sel.Clear
    
End Sub


'**********************
'debug
Sub dump_bbox( _
    ByVal bBox As BBox2D)

    Dim dDoc As DrawingDocument
    Set dDoc = CATIA.ActiveDocument

    Dim sheet As DrawingSheet
    Set sheet = dDoc.sheets.ActiveSheet
    
    Dim view As DrawingView
    Set view = get_view_by_name("dump")
    
    Dim fact As Factory2D
    Set fact = view.Factory2D

    view.Activate

    Dim bBoxAry As Variant
    bBoxAry = bBox.as_array()

    Dim ary As Variant
    ary = Array( _
        bBoxAry(0), _
        bBoxAry(1), _
        bBoxAry(2), _
        bBoxAry(3), _
        bBoxAry(6), _
        bBoxAry(7), _
        bBoxAry(4), _
        bBoxAry(5), _
        bBoxAry(0), _
        bBoxAry(1) _
    )

    Dim i As Long
    Dim line As Line2D
    For i = 0 To UBound(ary) - 2 Step 2
        Set line = fact.CreateLine( _
            ary(i), _
            ary(i + 1), _
            ary(i + 2), _
            ary(i + 3) _
        )
    Next

End Sub


'ビューを名前で取得
'Optional isCreate - true:なきゃ作る false:なきゃnothing
Private Function get_view_by_name( _
    ByVal name As String, _
    Optional isCreate = True) _
    As DrawingView
    
    Dim dDoc As DrawingDocument
    Set dDoc = CATIA.ActiveDocument

    Dim views As DrawingViews
    Set views = dDoc.sheets.ActiveSheet.views

    Dim view As DrawingView
    For Each view In views
        If Not view.name = name Then
            GoTo continue
        End If
        
        Set get_view_by_name = view
        Exit Function
        
continue:
    Next

    If isCreate Then
        Set get_view_by_name = views.add(name)
    Else
        Set get_view_by_name = Nothing
    End If
    
End Function

