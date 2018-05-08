Attribute VB_Name = "単純コピー"
'選択セルの内容をクリップボードにコピーする
'Ctrl+Cでコピーした内容がダブルクォーテーションで囲まれるのを回避する場合に使用してください
'
Sub 単純コピー()
    
    '変数宣言
    Dim startOfRow As Long
    Dim lastOfRow As Long
    Dim startOfCol As Long
    Dim lastOfCol As Long
    Dim buf As String
    Dim CB As New DataObject
    
    '初期化
    startOfRow = Selection.Row
    lastOfRow = startOfRow + Selection.Rows.Count - 1
    startOfCol = Selection.Column
    lastOfCol = startOfCol + Selection.Columns.Count - 1
    buf = ""
    
    '文字列取り込みループ
    rowFocus = startOfRow
    Do '行ループ
        
        colFocus = startOfCol
        Do '列ループ
                
            If colFocus > startOfCol Then '2列目以降なら
                buf = buf & vbTab 'タブ挿入
                
            End If
            
            '文字列取り込み
            buf = buf & Application.ActiveSheet.Cells(rowFocus, colFocus).Text
            
            colFocus = colFocus + 1
        
        Loop While colFocus <= lastOfCol
        
        buf = buf & vbCrLf '改行挿入
        
        rowFocus = rowFocus + 1
    
    Loop While rowFocus <= lastOfRow
    
    'クリップボード操作
    With CB
        .SetText buf        '変数のデータをDataObjectに格納する
        .PutInClipboard     'DataObjectのデータをクリップボードに格納する
    End With
    
End Sub


