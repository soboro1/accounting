Attribute VB_Name = "M03_実務"
Option Explicit

'==============================================================
' Lv3 実務でよく使う操作
'==============================================================

'--- 共通: シートが存在するか調べる -------------------------
Public Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function

'--- Q10: オートフィルタで抽出して別シートへ ----------------
Sub Q10_東京本社を抽出()
    Dim wsS As Worksheet, wsOut As Worksheet
    Dim lastRow As Long
    Const OUT_NAME As String = "東京本社_抽出"

    Set wsS = ThisWorkbook.Worksheets("売上明細")
    lastRow = wsS.Cells(wsS.Rows.Count, "A").End(xlUp).Row

    ' 出力先シートを作り直す
    Application.DisplayAlerts = False
    If SheetExists(OUT_NAME) Then ThisWorkbook.Worksheets(OUT_NAME).Delete
    Application.DisplayAlerts = True
    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOut.Name = OUT_NAME

    ' 既存フィルタを解除してからかける
    If wsS.AutoFilterMode Then wsS.AutoFilterMode = False
    wsS.Range("A1:K" & lastRow).AutoFilter Field:=3, Criteria1:="東京本社"

    ' 可視セルだけコピー（該当0件だと落ちるので件数を確認）
    If wsS.Range("A1:A" & lastRow).SpecialCells(xlCellTypeVisible).Count > 1 Then
        wsS.Range("A1:K" & lastRow).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=wsOut.Range("A1")
    Else
        MsgBox "該当データがありません。", vbExclamation
    End If

    wsS.AutoFilterMode = False
    wsOut.Columns("A:K").AutoFit
    wsOut.Activate
End Sub

'--- Q11: シートを安全に追加する ----------------------------
Sub Q11_月次シートを作る()
    Dim ws As Worksheet
    Dim sheetName As String

    sheetName = Format(Date, "yyyymm") & "_月次"

    If SheetExists(sheetName) Then
        If MsgBox(sheetName & " は既にあります。作り直しますか？", _
                  vbYesNo + vbQuestion) = vbNo Then Exit Sub
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(sheetName).Delete
        Application.DisplayAlerts = True
    End If

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
    ws.Range("A1").Value = sheetName & " 集計表"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Tab.Color = RGB(31, 78, 121)
End Sub

'--- Q12: 並べ替え ------------------------------------------
Sub Q12_金額降順で並べ替え()
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets("売上明細")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("K2:K" & lastRow), Order:=xlDescending
        .SortFields.Add Key:=ws.Range("A2:A" & lastRow), Order:=xlAscending
        .SetRange ws.Range("A1:K" & lastRow)
        .Header = xlYes
        .Apply
    End With

    MsgBox "金額の降順（同額は日付の昇順）で並べ替えました。", vbInformation
End Sub

'--- Q13: マスタから単価を転記（VLOOKUP相当） ---------------
Sub Q13_単価を転記()
    Dim wsM As Worksheet, wsI As Worksheet
    Dim lastM As Long, lastI As Long, i As Long
    Dim v As Variant

    Set wsM = ThisWorkbook.Worksheets("商品マスタ")
    Set wsI = ThisWorkbook.Worksheets("在庫")
    lastM = wsM.Cells(wsM.Rows.Count, "A").End(xlUp).Row
    lastI = wsI.Cells(wsI.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastI
        ' 見つからないとエラーになるので、いったん Variant で受けて判定する
        v = Application.VLookup(wsI.Cells(i, "A").Value, _
                                wsM.Range("A2:D" & lastM), 4, False)
        If IsError(v) Then
            wsI.Cells(i, "F").Value = ""
            wsI.Cells(i, "F").Interior.Color = RGB(255, 235, 235)
        Else
            wsI.Cells(i, "F").Value = v
        End If
    Next i

    MsgBox "単価を転記しました。", vbInformation
End Sub

'--- Q14: 書式を整える定型処理 ------------------------------
Sub Q14_表を整形する()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long, lastCol As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Then Exit Sub

    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' 罫線
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(191, 191, 191)
    End With

    ' 見出し
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' 表示・印刷設定
    ws.Cells.EntireColumn.AutoFit
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True

    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$1:$1"
    End With
End Sub
