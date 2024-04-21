VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVBAUtilityLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'/_/
'/_/ clsVbaUtilityLib
'/_/ CATIA V5�p���[�e�B���e�B�N���X
'/_/ ��ɔz��A����
'/_/
'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Option Explicit

Private Const VER = "0.0.3"

'�f�o�b�O���[�h
Private m_debug_mode As Boolean


Private Sub Class_Initialize()
    m_debug_mode = True
End Sub


Private Sub Class_Terminate()

End Sub


Property Get debug_mode() As Boolean
    debug_mode = m_debug_mode
End Property


Property Let debug_mode(ByVal mode As Boolean)
    m_debug_mode = mode
End Property


Function get_ver() As String
    get_ver = VER
End Function


'�_���v
'param:string
'return:
Sub dump( _
        ByVal Txt As String)
    
    If Not m_debug_mode Then Exit Sub
    Debug.Print Txt
End Sub

'*****���тȊ֐�*****
'APC�擾
'param:
'return:Object(Of IApc)
Function get_apc() As Object
    Dim COMObjectName As String
    #If VBA7 Then
        COMObjectName = "MSAPC.Apc.7.1"
    #ElseIf VBA6 Then
        COMObjectName = "MSAPC.Apc.6.2"
    #End If
    
    Dim oApc As Object
    On Error Resume Next
    Set oApc = CreateObject(COMObjectName)
    On Error GoTo 0
    
    If oApc Is Nothing Then
        Set oApc = CreateObject("MSAPC.Apc")
    End If
    
    Set get_apc = oApc
End Function


'VBE�擾
'param:
'return:Object(Of IApc)
Function get_vbe() As Object
    Set get_vbe = get_apc().vbe
End Function


'���s����vbproject�̃t�@���_
'param:
'return:string
Function get_this_dir() As String
    Dim oApc As Object
    Set oApc = get_apc()

    Dim projFilePath As String
    projFilePath = oApc.ExecutingProject.VBProject.fileName
    
    get_this_dir = get_fso.GetParentFolderName(projFilePath)
    
End Function


'Scripting.Dictionary
'param:compareMode-Long
'return:Object(Of Dictionary)
Function init_dict( _
        Optional compareMode As Long = vbBinaryCompare) As Object

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    dict.compareMode = compareMode
    
    Set init_dict = dict
End Function


'�����̔z����L�[�Ƃ��A�����o�����[�Ƃ��������쐬
'param:txtArray-array(Of string)
'param:compareMode-Long
'return:Object(Of Dictionary)
Function init_dict_by_array_count( _
        ByVal txtArray As Variant, _
        Optional compareMode As Long = vbBinaryCompare) As Object

    Dim dict As Object
    Set dict = init_dict(compareMode)

    Dim key As Variant
    For Each key In txtArray
        If Not dict.Exists(key) Then
            dict(key) = dict(key) + 1
        Else
            Call dict.Add(key, 0)
        End If
    Next

    Set init_dict_by_array_count = dict
End Function


'����̎����� - �L�[�ƃo�����[�Ŕ�r�@�K�w�ł�OK
'param:dict1-Object(Of Dictionary)
'param:dict2-Object(Of Dictionary
'return:boolean
Function is_equal_dict( _
        ByVal dict1 As Object, _
        ByVal dict2 As Object) As Boolean

    If dict1.count <> dict2.count Then
        is_equal_dict = False Exit Function
    End If
    
    Dim key As Variant
    For Each key In dict1.keys
        If Not dict2.Exists(key) Then
            is_equal_dict = False Exit Function
        End If
        If TypeName(dict1(key)) = "Dictionary" And TypeName(dict2(key)) = "Dictionary" Then
            ' If both values are dictionaries, recursively compare them
            If Not is_equal_dict(dict1(key), dict2(key)) Then
                is_equal_dict = False
                Exit Function
            End If
        ElseIf dict1(key) <> dict2(key) Then
            is_equal_dict = False
            Exit Function
        End If
    Next key

    is_equal_dict = True
    
End Function


'�����̃f�B�[�v�R�s�[
'param:source-Object(Of Dictionary)
'return:Object(Of Dictionary)
Function deep_copy_dict( _
        ByVal source As Variant) As Object

    Dim key As Variant
    Dim newDict As Object
    Dim value As Variant
    Set newDict = CreateObject("Scripting.Dictionary")
    For Each key In source.keys
        value = source.item(key)
        If TypeName(value) = "Dictionary" Then
            Set value = deep_copy_dict(value)
        End If

        newDict.Add key, value
    Next key

    Set deep_copy_dict = newDict

End Function


'*****�z��Ȋ֐�*****
'�T�C�Y�w��̔z��쐬 - �����l��let�̂�
'param:count-long
'param:value
'return:Variant(Of Array)
Function init_array( _
        ByVal count As Long, _
        Optional ByVal value As Variant = Empty) As Variant

    Dim ary() As Variant
    If count < 0 Then
        ary = Array()
    Else
        ReDim ary(count)
    End If
    
    If IsEmpty(value) Then
        init_array = ary
    End If
    
    Dim i As Long
    For i = 0 To UBound(ary)
        ary(i) = value
    Next
    init_array = ary

End Function


'�A�������z��̍쐬 - �A������Ȃ��V���Ȕz��쐬
'param:ary1-Variant(Of Array)
'param:ary2-Variant(Of Array)
'return:Variant(Of Array)
Function joint_array( _
        ByVal ary1 As Variant, _
        ByVal ary2 As Variant) As Variant

    Select Case True
        Case Not IsArray(ary1) And Not IsArray(ary2)
            joint_array = Empty
            Exit Function
        Case Not IsArray(ary1)
            joint_array = ary2
            Exit Function
        Case Not IsArray(ary2)
            joint_array = ary1
            Exit Function
        Case UBound(ary1) < 0 And UBound(ary2) > -1
            joint_array = ary2
            Exit Function
        Case UBound(ary1) > -1 And UBound(ary2) < 0
            joint_array = ary1
            Exit Function
        Case UBound(ary1) < 0 And UBound(ary2) < 0
            joint_array = Empty
            Exit Function
        End Select
        
        Dim StCount&: StCount = UBound(ary1)
        ReDim Preserve ary1(UBound(ary1) + UBound(ary2) + 1)
        
        Dim i&
        If IsObject(ary2(0)) Then
            For i = StCount + 1 To UBound(ary1)
                Set ary1(i) = ary2(i - StCount - 1)
            Next
        Else
            For i = StCount + 1 To UBound(ary1)
                ary1(i) = ary2(i - StCount - 1)
            Next
        End If
        
        joint_array = ary1

End Function


'�X���C�X
'param:ary-Variant(Of Array)
'param:startIdx-Long
'param:endIdx-Long
'return:Variant(Of Array)
Function slice( _
        ByVal ary As Variant, _
        ByVal startIdx As Long, _
        ByVal endIdx As Long) As Variant

    If Not IsArray(ary) Then Exit Function
    If endIdx - startIdx < 0 Then Exit Function
    If startIdx < 0 Then Exit Function
    If endIdx > UBound(ary) Then Exit Function
    
    Dim rngAry() As Variant
    ReDim rngAry(endIdx - startIdx)

    Dim i As Long
    For i = startIdx To endIdx
        rngAry(i - startIdx) = ary(i)
    Next
    
    slice = rngAry

End Function


'�z��̃f�B�[�v�R�s�[
'param:ary-Variant(Of Array)
'return:Variant(Of Array)
Function deep_copy_array( _
        ByVal ary As Variant) As Variant

    If Not IsArray(ary) Then Exit Function

    deep_copy_array = slice(ary, 0, UBound(ary))

End Function


'�z�񂪈�v���Ă��邩�H
'param:ary1-Variant(Of Array)
'param:ary2-Variant(Of Array)
'return:Boolean
Function is_equal_array( _
        ByVal ary1 As Variant, _
        ByVal ary2 As Variant) As Boolean

    is_equal_array = False
    
    If Not IsArray(ary1) Or Not IsArray(ary2) Then Exit Function
    If Not UBound(ary1) = UBound(ary2) Then Exit Function
    
    Dim i As Long
    For i = 0 To UBound(ary1)
        If Not ary1(i) = ary2(i) Then Exit Function
    Next
    
    is_equal_array = True

End Function


'�R���N�V�����̔z�� - let/set�͍ŏ��̗v�f�Ŕ��f
'�R���N�V��������Let��Set�̗v�f�������Ȃ���!!
'param:lst-collection
'return:Variant(Of Array)
Function collection_to_array( _
        ByVal lst As Variant) As Variant

    If lst.count < 1 Then
        collection_to_array = Array()
        Exit Function
    End If
    
    If IsObject(lst.item(1)) Then
        collection_to_array = collection_to_array_set(lst)
    Else
        collection_to_array = collection_to_array_let(lst)
    End If

End Function


'-�R���N�V�����̔z��-set
'param:lst-collection
'return:Variant(Of Array)
Function collection_to_array_set( _
        ByVal lst As Variant) As Variant

    Dim lstCount As Long
    lstCount = lst.count

    If lstCount < 1 Then
        collection_to_array_set = Array()
        Exit Function
    End If

    Dim ary() As Variant
    ReDim ary(lstCount - 1)

    Dim i As Long
    For i = 1 To lst.count
        Set ary(i - 1) = lst(i)
    Next

    collection_to_array_set = ary

End Function


'-�R���N�V�����̔z��-let
'param:lst-collection
'return:Variant(Of Array)
Function collection_to_array_let( _
        ByVal lst As Variant) As Variant

    Dim lstCount As Long
    lstCount = lst.count

    If lstCount < 1 Then
        collection_to_array_let = Array()
        Exit Function
    End If

    Dim ary() As Variant
    ReDim ary(lstCount - 1)

    Dim i As Long
    For i = 1 To lst.count
        Let ary(i - 1) = lst(i)
    Next

    collection_to_array_let = ary

End Function
'-コレクションの配�?�化-let 'param:lst-collection 'return:Variant(Of Array) Private Function collection_to_array_let( _ ByVal lst As Variant) As Variant Dim lstCount As Long lstCount = lst.count If lstCount < 1 Then collection_to_array_let = Array() Exit Function End If Dim ary() As Variant ReDim ary(lstCount - 1) Dim i As Long For i = 1 To lst.count ary(i - 1) = lst(i) Next collection_to_array_let = ary End Function

'�z��̃R���N�V������
'param:Variant(Of Array)
'return:collection
Function array_to_collection( _
        ByVal ary As Variant) As Collection

    Dim lst As Collection
    Set lst = New Collection

    If UBound(ary) < 0 Then Exit Function
    
    Dim v As Variant
    For Each v In ary
        lst.Add v
    Next

    Set array_to_collection = lst


End Function


'*****IO�Ȋ֐�*****
'FileSystemObject
'return:Object(Of FileSystemObject)
Function get_fso() As Object
    Set get_fso = CreateObject("Scripting.FileSystemObject")
End Function


'Dir�߽/̧�ٖ�/�g���q ����
'param:FullPath-�t�@�C���p�X
'return:Variant(Of Array(Of String)) (0-Path 1-BaseName 2-Extension)
Function split_path_name( _
        ByVal fullpath As String) As Variant

    Dim path(2) As String
    With get_fso
        path(0) = .GetParentFolderName(fullpath)
        path(1) = .GetBaseName(fullpath)
        path(2) = .GetExtensionName(fullpath)
    End With

    split_path_name = path

End Function


'Dir�߽/̧�ٖ�/�g���q �A��
'param:Path-Variant(Of Array(Of String)) (0-Path 1-BaseName 2-Extension)
'return:�t�@�C���p�X
Function joint_path_name( _
        ByVal path As Variant) As String

    joint_path_name = ""

    If Not IsArray(path) Then Exit Function '���Ή�
    If Not UBound(path) = 2 Then Exit Function '���Ή�
    
    joint_path_name = get_fso().BuildPath path(0), path(1) & "." & path(2)

End Function


'�d�����Ȃ��p�X�̎擾
'�d������ꍇ (�t�@�C���x�[�X��)_(�i���o�����O) �Ƃ���
'param:Path-�t�@�C���p�X
'return:�d�������p�X
Function get_unique_path( _
        ByVal oldPath As String) As String
    Dim path() As Varian
    path = split_path_name(oldPath)
    path(2) = "." & path(2)

    Dim newPath As String
    newPath = path(0) + "\" + path(1)
    If Not fso.FileExists(newPath + path(2)) Then
        get_unique_path = newPath + path(2)
        Exit Function
    End If

    Dim fso As Object
    Set fso = get_fso()

    Dim tempName As String
    Dim i As Long
    i = 0
    Do
        i = i + 1
        tempName = newPath + "_" + CStr(i) + path(2)
        If Not fso.FileExists(tempName) Then
            get_unique_path = tempName
            Exit Function
        End If
    Loop

End Function


'�t�@�C���̓ǂݍ���
'1�s�œǂݍ���ł���̂ŉ��s��split����
'param:path-string
'return:string
Function read_file( _
        ByVal path As String) As String

    read_file = ""

    On Error Resume Next

    With get_fso
        read_file = .OpenTextFile(path).ReadAll
        .Close
    End With

    On Error GoTo 0

End Function


'�t�@�C���̏����o��
'param:path-string
'param:txt-string
Sub write_file( _
        ByVal path As String, _
        ByVal Txt As String)

    With get_fso().OpenTextFile(path, 2, True)
        .Write Txt
        .Close
    End With

End Sub
