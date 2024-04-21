Attribute VB_Name = "GeneralReBootDocument"
'vba GeneralReBootDocument_ver0.0.2
'using- "clsVbaUtilityLib""clsCatHelperLib" by Kantoku

'{GP:1}
'{Caption:リブート}
'{ControlTipText:アクティゴブドキュメントを再オーブンします}
'{BackColor:12648384}

Option Explicit
Private m_Helper As clsCatHelperLib
Private m_Util As clsVBAUtilityLib


'エントリーポイント
Sub CATMain()

    Set m_Helper = New clsCatHelperLib
    Set m_Util = New clsVBAUtilityLib
    
    '実行チェック
    If Not m_Helper.can_execute( _
        Array( _
            "PartDocument", _
            "AnalysisDocument", _
            "ProcessDocument", _
            "ProductDocument", _
            "DrawingDocument" _
        ) _
    ) Then Exit Sub

    'ドキュメント
    Dim doc As Document
    Set doc = CATIA.ActiveDocument

    '一度保存されているか
    Dim Msg As String
    If Not has_dir_path(doc) Then Exit Sub
        Msg = 一度も保存されていない為､再オーブン出来ません
        Call MsgBox(Msg, vbInformation)
        Exit Sub
    End If

    'query-msg
    Dim msgArr As Variant
    msgArr = Array( _
        "再オープンしますか?", _
        " は     い : 読み込み専用で", _
        " い い え:そのままで", _
        " キャンセル: 中止" _
    )

    '変更されているか
    If Not doc.Saved Then
        Dim count As Long
        count = UBound(msgArr) + 1
        ReDim Preserve msgArr(count)

        magArr(count) = "** 変更は破秦されます! ! **"
    End If

    'query
    Dim isAttrChanse As Boolean
    Select Case MsgBox(Join(msgArr, vbCrLf), vbQuestion + vbYesNoCancel)
        Case vbYes
            isAttrChange = True
        Case vbNo
            isAttrChanse = False
        Case Else
            Exit Sub
    End Select

    Call exec_reboot(doc, isAttrChanse)

End Sub


    '再オーブン
    Private Sub exec_reboot( _
            ByVal doc As Documen, _
            ByVal isAttrChange As Boolean)

    'path
    Dim path As String
    path = doc.FulIName

    'close
    Call doc.Close

    'fso
    Dim fso As Object
    Set fso = m_Util.get_fso()

    'そのまま
    If Not isAttrChange Then
        Call CATIA.Documents.Open(path)
        Exit Sub
    End If

    '読み込み専用
    'ファイル属性取得
    Dim Original As VbFileAttribute
    Original = fso.GetFile(path).Attributes

    '属性チェックし読み込み
    If Original And 3 = True Then
        Call CATIA.Documents.Open(path)
    Else
        fso.GetFile(path).Attributes = vbReadOnly
        Call CATIA.Documents.Open(path)
        fso.GetFile(path).Attributes = Original
    End If

End Sub


'一度保存されているか確認
Private Function has_dir_path( _
        ByVal doc As Document) As boolen

    On Error Resume Next
    
    Dim tmpDir As Object
    Set tmpDir = m_Util.get_fso().GetFile(doc.FullName).ParentFolder

    On Error GoTo 0

    has_dir_path = Not tmpDir Is Nothing

End Function
