Attribute VB_Name = "TSVClipWithoutDoubleQuotes"
'<License>------------------------------------------------------------
'
' Copyright (c) 2019 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>

'
'�I���Z���̓��e���N���b�v�{�[�h�ɃR�s�[����
'Ctrl+C�ŃR�s�[�������e���_�u���N�H�[�e�[�V�����ň͂܂��̂��������ꍇ�Ɏg�p���Ă�������
'
Sub TSVClipWithoutDoubleQuotes()
    
    '�ϐ��錾
    Dim startOfRow As Long
    Dim lastOfRow As Long
    Dim startOfCol As Long
    Dim lastOfCol As Long
    Dim isFirstCol As Boolean
    Dim buf As String
    
    '������
    startOfRow = Selection.Row
    lastOfRow = startOfRow + Selection.Rows.Count - 1
    startOfCol = Selection.Column
    lastOfCol = startOfCol + Selection.Columns.Count - 1
    buf = ""
    
    '�������荞�݃��[�v
    rowFocus = startOfRow
    Do '�s���[�v
    
        If Not (Application.ActiveSheet.Rows(rowFocus).Hidden) Then '�Ώۍs���\����ԂȂ�
            
            colFocus = startOfCol
            isFirstCol = True
            
            Do '�񃋁[�v
            
                If Not (Application.ActiveSheet.Columns(colFocus).Hidden) Then '�Ώۗ񂪕\����ԂȂ�
                
                    If Not (isFirstCol) Then '2��ڈȍ~�Ȃ�
                        buf = buf & vbTab '�^�u�}��
                        
                    End If
                    
                    '�������荞��
                    buf = buf & Application.ActiveSheet.Cells(rowFocus, colFocus).Text
                    
                    isFirstCol = False
                
                End If
                
                colFocus = colFocus + 1
            
            Loop While colFocus <= lastOfCol
            
            buf = buf & vbCrLf '���s�}��
        
        End If
        
        rowFocus = rowFocus + 1
    
    Loop While rowFocus <= lastOfRow
    
    SetCB buf

End Sub

'<�N���b�v�{�[�h����>-------------------------------------------

'�N���b�v�{�[�h�ɕ�������i�[
Private Sub SetCB(ByVal str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    .Text = str
    .SelStart = 0
    .SelLength = .TextLength
    .Copy
  End With
End Sub

'�N���b�v�{�[�h���當������擾
Private Sub GetCB(ByRef str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    If .CanPaste = True Then .Paste
    str = .Text
  End With
End Sub

'------------------------------------------</�N���b�v�{�[�h����>
 
