Attribute VB_Name = "create_balloons_test"
'vba  by Kantoku
'using-'KCL0.1.0','GeoFactry','BBox2D','Pnt2D','Vec2D'
'アクティブなシートの全寸法に番号バルーンを付ける

'バルーンレポートビュー名
Private Const BALLOON_VIEW_NAME = "BALLOON"

'テンプレートバルーン名
Private Const TEMPLATE_SHEET_NAME = "template"
Private Const TEMPLATE_BALLOON_NAME = "balloon"

'true/false　vbaとCATIA API食い違い代案
Private Enum BOOL
    boolFalse = 0
    boolTrue = 1
End Enum

'関連テキスト用
Private Enum Dimension_Value
    Main_Value = 1
    Dual_Value = 2
End Enum

'寸法値情報用キー
Private Enum Dimension_Info_Keys
    value = 0 '寸法値
    Tolerance_upper '公差上
    Tolerance_lower '公差下
End Enum

Option Explicit


Sub CATMain()

    'エクスポートパス
    Dim path As String
    path = get_export_path()

    If path = vbNullString Then
        Exit Sub
    End If
    

    '除外する関連テキストフィルター
    'Before, After, Upper, Lower
    Dim blackList As Variant
    blackList = Array( _
        Array("(", ")", "", "") _
    )
    
    'バルーン作成・寸法値情報書き出し
    create_balloons blackList, 5, path
    
End Sub


'バルーンの作成
Private Sub create_balloons( _
    ByVal blackList As Variant, _
    ByVal startNumber As Integer, _
    ByVal exportPath As String)

    Dim dDoc As DrawingDocument
    Set dDoc = CATIA.ActiveDocument

    Dim sheet As DrawingSheet
    Set sheet = dDoc.sheets.ActiveSheet

    KCL.SW_Start

    'バルーンテンプレート コピー
    preparation_balloon_paste

    'バルーン用のビュー
    Dim balloonView As DrawingView
    Set balloonView = get_view_by_name( _
        BALLOON_VIEW_NAME)

    Dim drawDims As Variant
    Dim view As DrawingView
    Dim i As Long
    Dim balloon As DrawingText

    Dim balloonIdx As Long 'ナンバリング用インデックス
    balloonIdx = startNumber

    Dim dimBBox As BBox2D
    Dim viewBBox As BBox2D
    Dim viewVec As Vec2D
    Dim textVec As Vec2D
    Dim textPnt As Pnt2D
    Dim leaderPnt As Pnt2D

    Dim dimValuesDict As Object
    Set dimValuesDict = KCL.InitDic()

    CATIA.HSOSynchronized = True

    For Each view In sheet.views
        '寸法検索
        drawDims = get_dimensions_by_view(view, blackList)
        If UBound(drawDims) < 1 Then
            GoTo continue
        End If

        'ビューバウンダリボックス
        Set viewBBox = GeoFactry.create_boundary_box_by_view( _
            view _
        )
        Set viewVec = viewBBox.origin_point().as_vector()
        
        For i = 0 To UBound(drawDims)
            '寸法バウンダリボックス
            Set dimBBox = GeoFactry.create_boundary_box_by_dimension( _
                drawDims(i) _
            )
            dimBBox.translate_by viewVec

            'バルーンのペースト
            Set balloon = pre_copide_and_paste(balloonView)

            'テキスト位置算出
            Set textPnt = dimBBox.center_point.clone()
            Set textVec = viewBBox.center_point.vector_to(textPnt)
            textVec.normalize
            textVec.scale_by 20
            textPnt.translate_by textVec

            'リーダー位置算出
            Set leaderPnt = dimBBox.center_point.clone()
            textVec.normalize
            textVec.scale_by 5
            leaderPnt.translate_by textVec

            'バルーン位置調整
            move_balloon _
                balloon, _
                textPnt, _
                leaderPnt

            'バルーンテキスト修正
            balloon.text = CStr(balloonIdx)
            
            '寸法値情報
            dimValuesDict.add balloonIdx, init_dimension_info(drawDims(i))

            'カウンタ更新
            balloonIdx = balloonIdx + 1
        Next

        CATIA.RefreshDisplay = True

continue:
    Next

    CATIA.HSOSynchronized = False

    '寸法情報書き出し
    export_csv exportPath, dimValuesDict

    MsgBox "Done : " & KCL.SW_GetTime

End Sub


'検索し配列で取得
Private Function get_search_items( _
    ByVal searchWord As String, _
    Optional selectEntity = Nothing) _
    As Variant

    Dim dDoc As DrawingDocument
    Set dDoc = CATIA.ActiveDocument
    
    Dim sel As Selection
    Set sel = dDoc.Selection

    sel.Clear
    If Not selectEntity Is Nothing Then
        sel.add selectEntity
    End If
    
    sel.Search searchWord
    
    If sel.Count2 < 1 Then
        get_search_items = Array()
        Exit Function
    End If
    
    Dim drawDims() As Variant
    ReDim drawDims(sel.Count2 - 1)
    
    Dim i As Long
    For i = 1 To sel.Count2
        Set drawDims(i - 1) = sel.Item(i).value
    Next
    
    sel.Clear
    
    get_search_items = drawDims

End Function


'ビュー内の寸法取得-フィルター付き
Private Function get_dimensions_by_view( _
    ByVal view As DrawingView, _
    ByVal filterBlack As Variant) _
    As Variant

    Dim dims As Variant
    dims = get_search_items( _
        "CATDrwSearch.DrwDimension,sel", _
        view _
    )
    
    Dim lst As Collection
    Set lst = New Collection
    
    Dim i As Long
    For i = 0 To UBound(dims)
        If Not is_match_bault_text(dims(i), filterBlack) Then
            lst.add dims(i)
        End If
    Next

    get_dimensions_by_view = collection_to_array_by_obj(lst)

End Function


'関連テキストのメイン値でフィルターと一致するか？
Private Function is_match_bault_text( _
    ByVal drawDim As DrawingDimension, _
    ByVal filterList As Variant) _
    As Boolean
    
    Dim dimValue As Variant 'DrawingDimValue
    Set dimValue = drawDim.GetValue()

    Dim before, after, upper, lower
    dimValue.GetBaultText _
        Dimension_Value.Main_Value, _
        before, _
        after, _
        upper, _
        lower

    Dim stateBaultText As String
    stateBaultText = Join( _
        Array(before, after, upper, lower), _
        "@" _
    )

    Dim i As Long
    Dim filter As String
    For i = 0 To UBound(filterList)
        filter = Join( _
            filterList(i), _
            "@" _
        )
        If stateBaultText = filter Then
            is_match_bault_text = True
            Exit Function
        End If
    Next

    is_match_bault_text = False

End Function


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


'シートを名前で取得
Private Function get_sheet_by_name( _
    ByVal name As String, _
    Optional isDetail As Integer) _
    As DrawingSheet
    
    Dim dDoc As DrawingDocument
    Set dDoc = CATIA.ActiveDocument

    Dim sheets As DrawingSheets
    Set sheets = dDoc.sheets

    Dim sheet As DrawingSheet
    For Each sheet In sheets
        If Not sheet.name = name Then
            GoTo continue
        End If

        If sheet.isDetail <> isDetail Then
            GoTo continue
        End If

        Set get_sheet_by_name = sheet
        Exit Function
        
continue:
    Next

    Set get_sheet_by_name = Nothing
    
End Function


'テンプレートバルーンの取得
'return array(DrawingSheet, DrawingView, drawingtext)
Private Function get_template_balloon() _
    As Variant

    get_template_balloon = Array()
    
    Dim dDoc As DrawingDocument
    Set dDoc = CATIA.ActiveDocument

    Dim backupSheet As DrawingSheet
    Set backupSheet = dDoc.sheets.ActiveSheet.views

    Dim sheet As DrawingSheet
    Set sheet = get_sheet_by_name(TEMPLATE_SHEET_NAME, BOOL.boolTrue)
    If sheet Is Nothing Then
        MsgBox "テンプレートディテールシートがありません"
        Exit Function
    End If
    sheet.Activate
    
    Dim view As DrawingView
    Set view = get_view_by_name(TEMPLATE_BALLOON_NAME)
    If view Is Nothing Then
        MsgBox "テンプレートビューがありません"
        backupSheet.Activate
        Exit Function
    End If

    view.Activate
    Dim items As Variant
    items = get_search_items( _
        "CATDrwSearch.DrwBalloon,sel", _
        view _
    )
    
    If UBound(items) < 0 Then
        MsgBox "テンプレートバルーンがありません"
        backupSheet.Activate
        Exit Function
    End If

    '最初にHitしたバルーン
    get_template_balloon = Array(sheet, view, items(0))

    backupSheet.Activate

End Function


'コピー済みの状態からペースト
Private Function pre_copide_and_paste( _
    ByVal targetView As DrawingView) _
    As DrawingText
    
    Dim targetSheet As DrawingSheet
    Set targetSheet = KCL.GetParent_Of_T(targetView, "DrawingSheet")

    Dim dDoc As DrawingDocument
    Set dDoc = CATIA.ActiveDocument

    Dim sel As Selection
    Set sel = dDoc.Selection

    targetSheet.Activate
    
    sel.add targetView
    sel.Paste
    
    Set pre_copide_and_paste = sel.Item2(1).value
    sel.Clear

End Function


'バルーン情報から要素をコピーのみ
Private Sub copy_entity( _
    ByVal balloonInfo As Variant)
    
    Dim sheet As DrawingSheet
    Set sheet = balloonInfo(0)

    Dim view As DrawingView
    Set view = balloonInfo(1)
    
    Dim balloon As DrawingText
    Set balloon = balloonInfo(2)
    
    Dim dDoc As DrawingDocument
    Set dDoc = CATIA.ActiveDocument
    
    Dim sel As Selection
    Set sel = dDoc.Selection

    sel.Clear
    sheet.Activate
    view.Activate
    sel.add balloon
    sel.Copy
    sel.Clear

End Sub


'バルーンの移動
Private Sub move_balloon( _
    ByVal balloon As DrawingText, _
    textPnt As Pnt2D, _
    leaderPnt As Pnt2D)

    Dim drawLeader As DrawingLeader
    Set drawLeader = balloon.leaders.Item(1)

    balloon.x = textPnt.x
    balloon.y = textPnt.y
    
    drawLeader.ModifyPoint 1, leaderPnt.x, leaderPnt.y

End Sub


'オブジェクトコレクション->配列
Private Function collection_to_array_by_obj( _
    lst As Collection) _
    As Variant

    If lst.Count < 1 Then
        collection_to_array_by_obj = Array()
        Exit Function
    End If

    Dim ary() As Variant
    ReDim ary(lst.Count - 1)

    Dim i As Long
    For i = 1 To lst.Count
        Set ary(i - 1) = lst(i)
    Next

    collection_to_array_by_obj = ary

End Function


'辞書の値をカンマ区切り文字列に変換
Private Function dict2str( _
    ByVal dict As Object) _
    As String

    Dim info() As Variant
    ReDim info(dict.Count - 1)

    Dim i As Long
    i = 0

    Dim key As Variant
    For Each key In dict.keys
        info(i) = dict(key)
        i = i + 1
    Next

    dict2str = Join(info, ",")

End Function


'寸法値情報をファイルに書き出し
Private Sub export_csv( _
    ByVal path As String, _
    ByVal dict As Object)

    Dim infos() As Variant
    ReDim infos(dict.Count)

    Dim key As Variant
    Dim i As Long
    i = 0
    For Each key In dict.keys()
        infos(i) = key & "," & dict2str(dict(key))
        i = i + 1
    Next
    
    KCL.WriteFile path, Join(infos, vbCrLf)
    
End Sub


'ファイル書き出し先のパス取得
Private Function get_export_path() _
    As String

    Dim path As String
    Dim msg As String
    msg = _
        "寸法値をCSVファイルに書き出します。" & vbCrLf & _
        "ファイル名を指定してください｡"

    path = CATIA.FileSelectionBox(msg, "*.csv", CatFileSelectionModeSave)
    If path = vbNullString Then
        MsgBox "中止します"
        Exit Function
    End If

    Dim fso As Object
    Set fso = KCL.GetFSO()
    
    If fso.FileExists(path) Then
        '上書きを確認
        msg = "ファイルを上書きしますか?"
        If MsgBox(msg, vbOKCancel + vbQuestion) = vbCancel Then
            MsgBox "中止します"
            path = vbNullString
            Exit Function
        End If
    End If

    get_export_path = path

End Function


'寸法値情報取得
Private Function init_dimension_info( _
    ByVal drawDim As DrawingDimension) _
    As Object

    Dim view As DrawingView
    Set view = KCL.GetParent_Of_T(drawDim, "DrawingView")

    Dim dict As Object
    Set dict = KCL.InitDic()

    Dim valueInfo As Variant
    valueInfo = get_dimension_values(drawDim)

    Dim roundCount As Long
    roundCount = get_decimal_places(valueInfo(1))

    dict.add _
        Dimension_Info_Keys.value, _
        valueInfo(2) & Round(valueInfo(0), roundCount) & valueInfo(3)

    Dim tolInfo As Variant
    tolInfo = get_dimension_tolerances(drawDim)
    
    
    If IsEmpty(tolInfo(0)) Then
        dict.add _
            Dimension_Info_Keys.Tolerance_upper, _
            Empty

        dict.add _
            Dimension_Info_Keys.Tolerance_lower, _
            Empty
    Else
        dict.add _
            Dimension_Info_Keys.Tolerance_upper, _
            Round(tolInfo(0), roundCount)

        dict.add _
            Dimension_Info_Keys.Tolerance_lower, _
            Round(tolInfo(1), roundCount)
    End If

    Set init_dimension_info = dict

End Function


'小数点以下の桁数取得
Private Function get_decimal_places( _
    ByVal number As Double) _
    As Long
    
    Dim numStr As String
    numStr = str(number)
    
    Dim decimalPoint As Long
    decimalPoint = InStr(numStr, ".")

    If decimalPoint < 1 Then
        get_decimal_places = 0
    Else
        get_decimal_places = Len(numStr) - decimalPoint
    End If

End Function


'寸法値情報取得
Private Function get_dimension_values( _
    ByVal drawDim As DrawingDimension) _
    As Variant
    
    Dim dimValue As DrawingDimValue
    Set dimValue = drawDim.GetValue()

    Dim prefix As String, suffix As String
    dimValue.GetPSText 1, prefix, suffix

    get_dimension_values = Array( _
        dimValue.value, _
        dimValue.GetFormatPrecision(1), _
        convert_prefix_suffix_str(prefix), _
        convert_prefix_suffix_str(suffix) _
    )
    
End Function


Private Function convert_prefix_suffix_str( _
    text As String) _
    As String
    
    Select Case text
        Case "<THREADPREFIX>"
            'M
            convert_prefix_suffix_str = "M"
        Case "<SQUARE>"
            '□
            convert_prefix_suffix_str = "□"
        Case "<DIAMETER>"
            'φ
            convert_prefix_suffix_str = "φ"
        Case Else
            'その他
            convert_prefix_suffix_str = text
    End Select
    
End Function


'トレランス情報取得
Private Function get_dimension_tolerances( _
    ByVal drawDim As DrawingDimension) _
    As Variant
    
    Dim variDim As Variant
    Set variDim = drawDim
    
    Dim tolType, tolName, upTol, lowTol, dUpTol, dLowTol, displayMode
    variDim.GetTolerances _
        tolType, _
        tolName, _
        upTol, _
        lowTol, _
        dUpTol, _
        dLowTol, _
        displayMode

    If tolType = 0 Then
        dUpTol = Empty
        dLowTol = Empty
    End If

    get_dimension_tolerances = Array( _
        dUpTol, dLowTol _
    )

End Function


'テンプレートバルーンのコピーのみ
Private Function preparation_balloon_paste() _
    As Boolean

    preparation_balloon_paste = False

    Dim sheet As DrawingSheet
    Set sheet = CATIA.ActiveDocument.sheets.ActiveSheet

    'バルーンテンプレート
    Dim balloonInfo As Variant
    balloonInfo = get_template_balloon()
    If UBound(balloonInfo) < 2 Then
        Exit Function
    End If

    copy_entity balloonInfo
    
    preparation_balloon_paste = True

    sheet.Activate
    
End Function

