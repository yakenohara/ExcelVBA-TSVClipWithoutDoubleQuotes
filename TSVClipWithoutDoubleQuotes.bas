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
'選択セルの内容をクリップボードにコピーする
'Ctrl+Cでコピーした内容がダブルクォーテーションで囲まれるのを回避する場合に使用してください
'
Sub TSVClipWithoutDoubleQuotes()
    
    '変数宣言
    Dim startOfRow As Long
    Dim lastOfRow As Long
    Dim startOfCol As Long
    Dim lastOfCol As Long
    Dim isFirstCol As Boolean
    Dim buf As String
    
    '初期化
    startOfRow = Selection.Row
    lastOfRow = startOfRow + Selection.Rows.Count - 1
    startOfCol = Selection.Column
    lastOfCol = startOfCol + Selection.Columns.Count - 1
    buf = ""
    
    '文字列取り込みループ
    rowFocus = startOfRow
    Do '行ループ
    
        If Not (Application.ActiveSheet.Rows(rowFocus).Hidden) Then '対象行が表示状態なら
            
            colFocus = startOfCol
            isFirstCol = True
            
            Do '列ループ
            
                If Not (Application.ActiveSheet.Columns(colFocus).Hidden) Then '対象列が表示状態なら
                
                    If Not (isFirstCol) Then '2列目以降なら
                        buf = buf & vbTab 'タブ挿入
                        
                    End If
                    
                    '文字列取り込み
                    buf = buf & Application.ActiveSheet.Cells(rowFocus, colFocus).Text
                    
                    isFirstCol = False
                
                End If
                
                colFocus = colFocus + 1
            
            Loop While colFocus <= lastOfCol
            
            buf = buf & vbCrLf '改行挿入
        
        End If
        
        rowFocus = rowFocus + 1
    
    Loop While rowFocus <= lastOfRow
    
    SetCB buf

End Sub

'<クリップボード操作>-------------------------------------------

'クリップボードに文字列を格納
Private Sub SetCB(ByVal str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    .Text = str
    .SelStart = 0
    .SelLength = .TextLength
    .Copy
  End With
End Sub

'クリップボードから文字列を取得
Private Sub GetCB(ByRef str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    If .CanPaste = True Then .Paste
    str = .Text
  End With
End Sub

'------------------------------------------</クリップボード操作>
 
