Attribute VB_Name = "LogVisualization"
' This VBA program is licensed as MIT License
' About MIT License, please see LICENSE file in same GitHub repository (Umemaru/UiPath_LogVisualization)
' Copyright (c) 2018 Umemaru (@UmegayaRollcake)
' https://github.com/Umemaru/UiPath_LogVisualization/blob/master/LICENSE

' 2018/08/30 Initial release - Umemaru

' UiPathの実行ログを収集しExcelシートで可視化するVBAマクロです。
' GitHubの同リポジトリにあるExcelマクロファイル[UiPath実行ログ集計.xlsm]のうち、
' 標準モジュールをエクスポートしたものです。（GitHubで公開・バージョン管理するため）

Option Explicit

Const SETUP_SHEET_NAME = "設定"
Const RAWLOG_SHEET_NAME = "実行ログ"
Const PIVOT_SHEET_NAME = "Pivot"
Const TIMESPAN_SHEET_NAME = "時間帯別"

Const PIVOT_TABLE_NAME = "実行ログ集計用"

' "ログ集計"ボタン押下時の処理
Sub ProcessStart()
    Call CollectLog '①指定期間の実行ログを集めて
    Call ArrangeLog_TimeSpan '②時間帯別に稼働率を集計する
End Sub


' ①指定期間の実行ログを収集＞
' パラメータ「ログフォルダパス」の先に格納されているUiPathの実行ログファイルを
' 今日から過去「集計期間（日）」日分だけ収集し、
' 中身を新規"実行ログ"シートに整形した上で貼り付ける。
' 最後にそれをデータソースとしている"Pivot"シート中のピボットテーブルを更新する。
Sub CollectLog()
    Dim i As Integer 'ループカウンター
    Dim logPath As String '実行ログファイルの格納先
    Dim collectPeriod As Integer '集計期間(日)
    
    Dim fso As FileSystemObject 'ユーザーフォルダ取得用
    Dim rootFolder As Folder
    Dim userFolder As Folder
    
    ' パラメータ取得
    logPath = Worksheets(SETUP_SHEET_NAME).Range("C4").value
    collectPeriod = Worksheets(SETUP_SHEET_NAME).Range("C5").value
    
    ' 最後尾に"実行ログ"シート追加（既に同名のシートがあれば削除してから追加）
    DeleteSheet (RAWLOG_SHEET_NAME)
    With Worksheets.Add(after:=Worksheets(Worksheets.Count))  '
        .Name = RAWLOG_SHEET_NAME
        .Cells(1, 1).value = "message" 'ヘッダ行を設定
        .Cells(1, 2).value = "level"
        .Cells(1, 3).value = "logType"
        .Cells(1, 4).value = "timeStamp"
        .Cells(1, 5).value = "fingerprint"
        .Cells(1, 6).value = "windowsIdentity"
        .Cells(1, 7).value = "machineName"
        .Cells(1, 8).value = "processName"
        .Cells(1, 9).value = "processVersion"
        .Cells(1, 10).value = "fileName"
        .Cells(1, 11).value = "jobId"
        .Cells(1, 12).value = "robotName"
        ' .Cells(1, 13).value = "totalExecutionTimeInSeconds"
        ' .Cells(1, 14).value = "totalExecutionTime"
    End With
    
    ' C:\Users\以下の各ユーザーフォルダにある実行ログを取得する
    Set fso = New FileSystemObject
    Set rootFolder = fso.GetFolder("C:\Users\")
    For Each userFolder In rootFolder.SubFolders
        ' 今日から過去「集計期間（日）」日分のログファイルを収集→"実行ログ"シートに貼り付け
        For i = 0 To collectPeriod - 1
            Dim logDate As Date
            Dim logFileName As String
            
            ' 取りたいログファイル名を設定
            logDate = DateAdd("d", -1 * i, Date)
            logFileName = Format(logDate, "YYYY-MM-DD") & "_Execution.log" ' 例:2018-08-07_Execution.log
            
            ' ログファイルが存在したら中身を"実行ログ"シートに貼り付け
            If Dir(userFolder.Path & logPath & "\" & logFileName) <> "" Then
                ListupLogText (userFolder.Path & logPath & "\" & logFileName)
            End If
        Next
    Next
        
    ' 最後に"Pivot"シート中のピボットテーブルを更新
    Worksheets(PIVOT_SHEET_NAME).Activate
    ActiveSheet.PivotTables(PIVOT_TABLE_NAME).PivotCache.Refresh

End Sub


' ②時間帯別に稼働率を集計＞
' パラメータ「見える化時間帯From」～「見える化時間帯To」までの時間を「見える化間隔（分）」で区切って
' それぞれの時間帯にロボットが実行されていたかを判定し、それを「集計期間（日）」日分繰り返す事で
' "時間帯別"シートに時間帯ごとの稼働率を表示する。
Sub ArrangeLog_TimeSpan()

    Dim collectPeriod As Integer '集計期間(日)
    Dim from_TimeSpan As Date '見える化時間帯From
    Dim to_TimeSpan As Date '見える化時間帯To
    Dim interval_TimeSpan As Date '見える化間隔(分)
    
    Dim arrayTimeSpan() As Date '見える化時間帯の時刻(From)の配列 (6:00, 6:30, 7:00, ..., 21:30)
    Dim arrayExecuting() As Boolean 'それぞれの時間帯でロボットが稼働しているかどうかのフラグ
    Dim arrayExecuteCount() As Integer 'それぞれの時間帯でロボットが稼働していた回数
    Dim size_arrayTimeSpan As Integer '時間帯配列の要素数
    
    Dim rowCntPivot As Integer 'ピボットテーブル"実行ログ集計用"の行数
    Dim startDate As Date '実行開始日
    Dim startDate_prev As Date '１つ前の実行開始日(日付ブレーク判定用)
    Dim startTime As Date '実行開始時刻
    Dim endTime As Date '実行終了時刻
    Dim i As Integer 'ループカウンター
    Dim j As Integer 'ループカウンター
    
    ' パラメータ取得
    collectPeriod = Worksheets(SETUP_SHEET_NAME).Range("C5").value
    from_TimeSpan = Worksheets(SETUP_SHEET_NAME).Range("C6").value
    to_TimeSpan = Worksheets(SETUP_SHEET_NAME).Range("C7").value
    interval_TimeSpan = Worksheets(SETUP_SHEET_NAME).Range("C8").value
        
    ' 時間帯配列の要素数の決定
    size_arrayTimeSpan = DateDiff("n", from_TimeSpan, to_TimeSpan) \ interval_TimeSpan '割り算で"\"を使えば整数部だけ取れる
    ReDim arrayTimeSpan(size_arrayTimeSpan)
    ReDim arrayExecuting(size_arrayTimeSpan)
    ReDim arrayExecuteCount(size_arrayTimeSpan)
        
    ' 時間帯配列の用意
    For i = 0 To size_arrayTimeSpan
        arrayTimeSpan(i) = DateAdd("n", interval_TimeSpan * i, from_TimeSpan)
        arrayExecuting(i) = False
    Next
    
    rowCntPivot = Worksheets(PIVOT_SHEET_NAME).PivotTables(PIVOT_TABLE_NAME).DataBodyRange.Rows.Count 'ピボットテーブル"実行ログ集計用"の件数
    ' ピボットテーブルの件数分、時間帯別の実行状況を集計する
    For i = 0 To rowCntPivot - 1
        ' ロボット実行の開始日/開始時刻/終了時刻
        startDate = DateValue(Worksheets(PIVOT_SHEET_NAME).Cells(5 + i, 5).value)
        startTime = TimeValue(Worksheets(PIVOT_SHEET_NAME).Cells(5 + i, 5).value)
        endTime = TimeValue(Worksheets(PIVOT_SHEET_NAME).Cells(5 + i, 6).value)
        
        ' 時間帯配列の先頭から順に、その時間帯にロボット実行時刻が含まれるか判定する
        For j = 0 To UBound(arrayTimeSpan) - 1
            arrayExecuting(j) = IsWithinTimeSpan(startTime, endTime, arrayTimeSpan(j), arrayTimeSpan(j + 1))
        Next
        
        If startDate <> startDate_prev Then
            ' 各時間帯毎に実行フラグを見て、Trueなら実行回数を+1する
            For j = 0 To UBound(arrayTimeSpan) - 1
                If arrayExecuting(j) Then
                    arrayExecuteCount(j) = arrayExecuteCount(j) + 1
                End If
            Next
        End If
        
        startDate_prev = startDate
    Next
    
    ' "時間帯別"シートの値を全クリア
    Worksheets(TIMESPAN_SHEET_NAME).Activate
    Worksheets(TIMESPAN_SHEET_NAME).Cells.ClearContents
    
    ' "時間帯別"シートのヘッダ行を設定
    For i = 0 To UBound(arrayTimeSpan)
        Worksheets(TIMESPAN_SHEET_NAME).Cells(1, i + 1).value = arrayTimeSpan(i)
    Next
    
    ' 最後に、各時間帯でのロボット稼働率(=実行回数/集計期間)を"時間帯別"シートに記入
    For i = 0 To UBound(arrayTimeSpan)
        Worksheets(TIMESPAN_SHEET_NAME).Cells(2, i + 1).value = arrayExecuteCount(i) / collectPeriod
    Next

End Sub


' 指定したパスのログファイルの中身を"実行ログ"シートに追記貼り付け
Sub ListupLogText(LogFilePath As String)
    Dim i As Integer 'ループカウンター
    Dim j As Integer 'ループカウンター
    Dim lastRowNum As Integer
    Dim buf As String
    Dim value As String
    Dim arrayLogItem As Variant 'ログテキストを[","]毎に分割した配列
    
    ' "実行ログ"シートの最終行番号(この下にログを追記する)
    lastRowNum = Worksheets(RAWLOG_SHEET_NAME).UsedRange.Rows.Count
    
    i = 1
    Open LogFilePath For Input As #1 'ログファイルの中身を１行ずつ処理
        Do Until EOF(1)
            Line Input #1, buf ' 行の中身をbufに代入
            '→ 07:11:00.7212 Info {"message":"XXXXX execution started","level":"Information","logType":"Default",....,"robotName":"ZZZZ"}
            
            buf = Mid(buf, InStr(buf, " {")) ' ログテキスト中の" {"から右(=json部分)を取り出す
            '→ "message":"XXXXX execution started","level":"Information","logType":"Default",....,"robotName":"ZZZZ"}
            
            buf = Left(buf, Len(buf) - 1) ' あと最後の"}"は要らないので除く
            '→ "message":"XXXXX execution started","level":"Information","logType":"Default",....,"robotName":"ZZZZ"
            
            ' 行の中身を[","]毎に分割してセルに貼り付け（本当はスマートにjson→excelに落とし込みたいけど）
            arrayLogItem = Split(buf, """,""", -1, vbTextCompare)
            ' 分割した要素数だけ繰り返し
            For j = 0 To UBound(arrayLogItem) ' UBoundは配列の要素数ではなく最大インデックス値を返す(自分メモ)
                value = Mid(arrayLogItem(j), InStr(arrayLogItem(j), ":") + 1) ' 要素key:valueのうち、":"より右のvalue部分だけ取り出す
                value = Replace(value, """", "") ' 値からダブルクォーテーションを排除
                
                If j = 3 Then ' timestamp(4要素目)はそのままだと使えないので、日付形式に変換
                    Worksheets(RAWLOG_SHEET_NAME).Cells(lastRowNum + i, j + 1) = TimeStamp2Date(value)
                Else
                    Worksheets(RAWLOG_SHEET_NAME).Cells(lastRowNum + i, j + 1) = value
                End If
            Next
                        
            i = i + 1
        Loop
    Close #1
End Sub


' タイムスタンプをDate形式に変換(2018-08-26T08:27:03.7277346+09:00→2018/08/26 08:27:03)
Function TimeStamp2Date(TimeStamp As String) As Variant
    TimeStamp2Date = DateValue(Mid(TimeStamp, 1, 10)) + TimeValue(Mid(TimeStamp, 12, 8))
End Function

' 指定した名前のシートの削除（無ければ何もしない）
Sub DeleteSheet(SheetName As String)
   Dim sheet As Excel.Worksheet
   Application.DisplayAlerts = False ' シート削除時の警告メッセージを出さない
   
   ' 指定したシートを削除（無ければエラーが発生するが、それを握りつぶしてスルー）
   On Error Resume Next
   Set sheet = Worksheets(SheetName)
   sheet.Delete
   On Error GoTo 0 ' 何もしない
   
   Application.DisplayAlerts = True
   
End Sub

' 処理の実行時間(Start~End)が、指定したタイムスパン(From~To)に含まれれるかどうか判定
Function IsWithinTimeSpan(startTime As Date, endTime As Date, fromTime As Date, toTime As Date) As Boolean
    If (fromTime <= startTime And startTime < toTime) Or _
       (fromTime <= endTime And endTime < toTime) Or _
       (startTime < fromTime And toTime < endTime) Then
        IsWithinTimeSpan = True
    Else
        IsWithinTimeSpan = False
    End If
End Function

