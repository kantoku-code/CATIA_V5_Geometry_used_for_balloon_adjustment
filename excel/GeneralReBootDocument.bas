Attribute VB_Name = "GeneralReBootDocument"
'vba GeneralReBootDocument_ver0.0.2
'using- "clsVbaUtilityLib""clsCatHelperLib" by Kantoku

'{GP:1}
'{Caption:���u�[�g}
'{ControlTipText:�A�N�e�B�S�u�h�L�������g���ăI�[�u�����܂�}
'{BackColor:12648384}

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
        Msg = ��x���ۑ�����Ă��Ȃ��פ�ăI�[�u���o���܂���
        Call MsgBox(Msg, vbInformation)
        Exit Sub
    End If

    'query-msg
    Dim msgArr As Variant
    msgArr = Array( _
        "�ăI�[�v�����܂���?", _
        " ��     �� : �ǂݍ��ݐ�p��", _
        " �� �� ��:���̂܂܂�", _
        " �L�����Z��: ���~" _
    )

    '�ύX����Ă��邩
    If Not doc.Saved Then
        Dim count As Long
        count = UBound(msgArr) + 1
        ReDim Preserve msgArr(count)

        magArr(count) = "** �ύX�͔j�`����܂�! ! **"
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


    '�ăI�[�u��
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

    '���̂܂�
    If Not isAttrChange Then
        Call CATIA.Documents.Open(path)
        Exit Sub
    End If

    '�ǂݍ��ݐ�p
    '�t�@�C�������擾
    Dim Original As VbFileAttribute
    Original = fso.GetFile(path).Attributes

    '�����`�F�b�N���ǂݍ���
    If Original And 3 = True Then
        Call CATIA.Documents.Open(path)
    Else
        fso.GetFile(path).Attributes = vbReadOnly
        Call CATIA.Documents.Open(path)
        fso.GetFile(path).Attributes = Original
    End If

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
