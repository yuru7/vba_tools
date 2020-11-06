Sub Main()
    Dim Path As String: Path = ThisWorkbook.Worksheets("Sheet1").Range("targetDir").Value   '// 対象フォルダパス
    Const cnsTitle = "フォルダ内のエクセルファイル提出状態化" 'ダイアログのタイトル
    
    Dim processYesNo As Integer
    
    ' フォルダの存在確認
    If Dir(Path, vbDirectory) = "" Then
        MsgBox "指定のフォルダは存在しません。", vbExclamation, cnsTitle
        Exit Sub
    End If
    
    processYesNo = MsgBox(Path & vbCrLf & vbCrLf & "上記フォルダ内の Excel ファイルを処理します。" _
            & vbCrLf & "サブフォルダも対象になります。" _
            & vbCrLf & vbCrLf & "実行しますか？", vbYesNo + vbExclamation, cnsTitle)
    If processYesNo = vbNo Then
        Exit Sub
    End If
    
    '// A1＆拡大率100%を設定する
    Call setA1And100Per(Path)

    MsgBox "処理が終了しました。", vbOKOnly + vbInformation, cnsTitle
End Sub

'// A1＆拡大率100%を設定する
Private Sub setA1And100Per(Path)
    '// A1＆拡大率100％を設定する
    Call executeA1And100Per(Path)
    
    '// サブフォルダを再帰する
    Dim f As Object
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(Path).SubFolders
            Call setA1And100Per(f.Path)
        Next f
    End With
End Sub


'// Excel系ファイルのみ、A1かつ拡大率100％にして、最初のシートを指定して保存する
Private Sub executeA1And100Per(Path)
    Dim book As Workbook
    Dim sheet As Object
    
    Const cnsDIR = "\*.*"
    Dim strFileName As String '処理中のファイル名を格納する変数
    Dim fileAndPath As String '処理中のファイル名（パス含む）を格納する変数
    Dim pos As Long
    
    ' 先頭のファイル名の取得
    strFileName = Dir(Path & cnsDIR, vbNormal)
    ' ファイルが見つからなくなるまで繰り返す
    Do While strFileName <> ""
    
        ' エクセルファイルのみを処理対象とする
        pos = InStrRev(strFileName, ".")
        If Not LCase(Mid(strFileName, pos + 1)) Like "xls*" Then
            ' 次のファイル名を取得
            GoTo Continue
        End If
        
        ' 自ファイル（A1&拡大率100％.xlsm）は除く
        If strFileName = ThisWorkbook.Name Then
            GoTo Continue
        End If
    
        ' エクセルファイルを開く
        fileAndPath = Path + "\" + strFileName
        Set book = Workbooks.Open(fileAndPath)
        
        '一番先頭のシートから順にループ処理を行う
        For Each sheet In book.Sheets
            sheet.Activate                 '対象のシートをアクティブにする
            ActiveSheet.Range("A1").Select 'シートのA1を選択する
            ActiveWindow.Zoom = 100        '拡大倍率を100に設定する
        Next sheet
        book.Sheets(1).Select
    
        ' エクセルファイルを保存して閉じる
        book.Save
        book.Close
    
Continue:
    
        ' 次のファイル名を取得
        strFileName = Dir()
    Loop
End Sub
