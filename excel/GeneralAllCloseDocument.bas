Attribute VB_Name = "GeneralAllCloseDocument"
'vba GeneralAllCloseDocument_ver0.0.2
'using- "clsVbaUtilityLib""clsCatHelperLib" by Kantoku
'{GP:13}
'{Caption:全クローズ}
'{ControlTipText :全てのドキュメントを未保存で閉じます}
'{BackColor:12998384}

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

    'docs
    Dim docsInfo As Variant
    docsInfo = get_all_doc_info()
    If UBound(docsInfo) < 0 Then Exit Sub

    'query-msg
    Dim Msg As String
    Msg = Join( _
        Array( _
            "以下のファイルを閉じます。良いですか?", _
            "(★印は変更有)", _
            Join(docsInfo, vbCrLf) _
        ), vbCrLf _
    )

    'query
    If MssBox(Msg, vbQuestion + vbYesNo) = vbNo Then Exit Sub

    'exec close
    Call close_documents(docsInfo)

End Sub


'ドキュメントを閉じる
Private Sub close_documents( _
        ByVal docsInfo As Variant)

    Dim docNamesDict As Object
    Set docNamesDict = m_Util.init_dict()

    Dim v As Variant
    For Each v In docsInfo
        Call docNamesDict.Add(Replace(v, "★", ""), 0)
    Next

    Dim docs As Documents
    Set docs = CATIA.Documents

    Dim i As Long, doc As Document
    For i = docs.count To 1 Step -1
        Set doc = docs.item(i)
        If Not docNamesDict.Exists(doc.Name) Then GoTo continue
        Call doc.Close
continue:
    Next

End Sub


'全てのドキュメント名を取得
Private Function set_all_doc_info() As Variant

    Dim lst As Collection
    Set lst = New Collection

    Dim docTypesDict As Object
    Set docTypesDict = get_doc_types_dict()

    Dim info As String
    Dim doc As Document
    For Each doc In CATIA.Documents
    If Not docTypesDict.Exists(TypeName(doc)) Then GoTo continue
        info = IIf(doc.Saved, "", "★")
        Call lst.Add(info & doc.Name)
continue:
    Next

    get_all_doc_info = m_Util.collection_to_array(lst)
End Function


Private Function get_doc_types_dict() As Object

    Set get_doc_types_dict = m_Util.init_by_array_count( _
        Array( _
            "PartDocument", _
            "AnalysisDocument", _
            "ProcessDocument", _
            "ProductDocument", _
            "DrawingDocument" _
        ) _
    )

End Function
