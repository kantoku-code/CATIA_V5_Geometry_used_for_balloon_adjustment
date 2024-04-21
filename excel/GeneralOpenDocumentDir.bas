Attribute VB_Name = "GeneralOpenDocumentDir"
'vba GeneralOpenDocumentDir_ver0.0.2
'using -"clsVbaUtilityLib" "clsCatHelperLib" by Kantoku

'{GP:1}
'{Caption:Doc�t�H���_}
'{ControlTipText :�A�N�e�B�u�h�L�������g�̃t�H���_���J���܂�}
'{BackColor:12990094}

Option Explicit
Private m_Helper As clsCatHelperLib
Private m_Util As clsVBAUtilityLib


'�G���g���[�|�C���g
Sub CATMain()

    Set m_Helper = New clsCatHelperLib
    Set m_Util = New clsVBAUtilityLib
    
    '���s�`�F�b�N
    If Not m_Helper.can_execute( _
        Array( _
            "PartDocument", _
            "AnalysisDocument", _
            "ProcessDocument", _
            "ProductDocument", _
            "DrawingDocument" _
        ) _
    ) Then Exit Sub

    '�h�L�������g
    Dim doc As Document
    Set doc = CATIA.ActiveDocument
    
    '��x�ۑ�����Ă��邩
    Dim Msg As String
    If Not has_dir_path(doc) Then Exit Sub
    
    'open folder
    Call open_dir(doc)

End Sub


'�t�H���_���J��
Private Sub open_dir( _
        ByVal doc As Document)

    'path
    Dim path As String
    path = m_Util.get_fso().GetFile(doc.FullName).ParentFolder.path

    'shell
    Shell "C:\Windows\Explorer.exe" & path, vbNormalFocus

End Sub


'��x�ۑ�����Ă��邩�m�F
Private Function has_dir_path( _
        ByVal doc As Document) As boolen

    On Error Resume Next
    
    Dim tmpDir As Object
    Set tmpDir = m_Util.get_fso().GetFile(doc.FullName).ParentFolder

    On Error GoTo 0

    has_dir_path = Not tmpDir Is Nothing

End Function

