Attribute VB_Name = "PartXYZPlaneShowHide"
'vba PartXYZPlaneShowHide_ver0.0.1
'using- "clsCatHelperLib�h by Kantoku

'{GP:5}
'{Caption:XYZ����}
'{ControlTipText :XY,YZ,ZX���ʂ�\��/��\�����܂��B}
'{BackColor:02608384}

Option Explicit
Private m_Helper As clsCatHelperLib


'�G���g���[�|�C���g
Sub CATMain()
    Set m_Helper = New clsCatHelperLib

    '���s�`�F�b�N
    If Not m_Helper.can_execute( _
        Array( _
            "PartDocument" _
        ) _
    ) Then Exit Sub

    'xyz plane
    Dim doc As PartDocument
    Set doc = CATIA.ActiveDocument

    Dim pt As Part
    Set pt = doc.Part

    Dim arrPlanes() As Variant
    arrPlanes = Array( _
        pt.OriginElements.PlaneXY, _
        pt.OriginElements.PlaneYZ, _
        pt.OriginElements.PlaneZX _
    )

    'show/hide xy���ʂŔ��f���؂�ւ�
    With CATIA.ActiveDocument.Selection
        .Clear
        .Add arrPlanes(0)

        Dim showState As CatVisPropertyShow
        .VisProperties.GetShow showState

        Dim ent As Variant
        For Each ent In arrPlanes
            Add ent
        Next

        If showState = catVisPropertyShowAttr Then
            .VisProperties.SetShow catVisPropertyNoShowAttr
        Else
            .VisProperties.SetShow catVisPropertyShowAttr
        End If

        .Clear
    End With

End Sub
