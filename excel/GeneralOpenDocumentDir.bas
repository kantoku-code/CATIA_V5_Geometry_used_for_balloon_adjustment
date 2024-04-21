Attribute VB_Name = "GeneralOpenDocumentDir"
'vba GeneralOpenDocumentDir_ver0.0.2
'using -"clsVbaUtilityLib" "clsCatHelperLib" by Kantoku

'{GP:1}
'{Caption:Docフォルダ}
'{ControlTipText :アクティブドキュメントのフォルダを開きます}
'{BackColor:12990094}

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
    
    'open folder
    Call open_dir(doc)

End Sub


'フォルダを開く
Private Sub open_dir( _
        ByVal doc As Document)

    'path
    Dim path As String
    path = m_Util.get_fso().GetFile(doc.FullName).ParentFolder.path

    'shell
    Shell "C:\Windows\Explorer.exe" & path, vbNormalFocus

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

