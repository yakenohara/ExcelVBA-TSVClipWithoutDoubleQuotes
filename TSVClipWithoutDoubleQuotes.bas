Attribute VB_Name = "�P���R�s�["
'�I���Z���̓��e���N���b�v�{�[�h�ɃR�s�[����
'Ctrl+C�ŃR�s�[�������e���_�u���N�H�[�e�[�V�����ň͂܂��̂��������ꍇ�Ɏg�p���Ă�������
'
Sub �P���R�s�[()
    
    '�ϐ��錾
    Dim startOfRow As Long
    Dim lastOfRow As Long
    Dim startOfCol As Long
    Dim lastOfCol As Long
    Dim buf As String
    Dim CB As New DataObject
    
    '������
    startOfRow = Selection.Row
    lastOfRow = startOfRow + Selection.Rows.Count - 1
    startOfCol = Selection.Column
    lastOfCol = startOfCol + Selection.Columns.Count - 1
    buf = ""
    
    '�������荞�݃��[�v
    rowFocus = startOfRow
    Do '�s���[�v
        
        colFocus = startOfCol
        Do '�񃋁[�v
                
            If colFocus > startOfCol Then '2��ڈȍ~�Ȃ�
                buf = buf & vbTab '�^�u�}��
                
            End If
            
            '�������荞��
            buf = buf & Application.ActiveSheet.Cells(rowFocus, colFocus).Text
            
            colFocus = colFocus + 1
        
        Loop While colFocus <= lastOfCol
        
        buf = buf & vbCrLf '���s�}��
        
        rowFocus = rowFocus + 1
    
    Loop While rowFocus <= lastOfRow
    
    '�N���b�v�{�[�h����
    With CB
        .SetText buf        '�ϐ��̃f�[�^��DataObject�Ɋi�[����
        .PutInClipboard     'DataObject�̃f�[�^���N���b�v�{�[�h�Ɋi�[����
    End With
    
End Sub


