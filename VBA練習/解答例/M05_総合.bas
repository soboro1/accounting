Attribute VB_Name = "M05_総合"
Option Explicit

'==============================================================
' Lv5 総合課題
'   データ フォルダの CSV を全部取り込む
'     → 支店 × カテゴリ のクロス集計
'     → 請求書テンプレートへ書き出し → PDF 出力
'
' ※ Mac の場合、Dir でフォルダを走査する前に一度
'    GrantAccessToMultipleFiles でアクセス許可が必要になることがあります。
'==============================================================

Private Const CSV_FOLDER As String = "データ"

'--- Q20-1: フォルダ内の CSV をすべて1シートに取り込む ------
Sub Q20_1_CSVを全部取り込む()
    Dim folderPath As String, fileName As String
    Dim wbCsv As Workbook
    Dim wsOut As Worksheet
    Dim lastRowSrc As Long, nextRow As Long
    Dim fileCount As Long
    Const OUT_NAME As String = "取込データ"

    folderPath = ThisWorkbook.Path & Application.PathSeparator & CSV_FOLDER & Application.PathSeparator

    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    ' 出力シートを作り直す
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(OUT_NAME).Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True
    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOut.Name = OUT_NAME

    nextRow = 1
    fileName = Dir(folderPath & "*.csv")

    If fileName = "" Then
        MsgBox "CSV が見つかりません。" & vbLf & folderPath, vbExclamation
        GoTo CleanUp
    End If

    Do While fileName <> ""
        Set wbCsv = Workbooks.Open(fileName:=folderPath & fileName, ReadOnly:=True)
        With wbCsv.Worksheets(1)
            lastRowSrc = .Cells(.Rows.Count, "A").End(xlUp).Row
            If nextRow = 1 Then
                ' 1ファイル目は見出しごとコピー
                .Range("A1:K" & lastRowSrc).Copy wsOut.Cells(1, 1)
                nextRow = lastRowSrc + 1
            Else
                ' 2ファイル目以降は見出しを飛ばす
                .Range("A2:K" & lastRowSrc).Copy wsOut.Cells(nextRow, 1)
                nextRow = nextRow + lastRowSrc - 1
            End If
        End With
        wbCsv.Close SaveChanges:=False
        fileCount = fileCount + 1
        fileName = Dir                      ' ← 次のファイルへ
    Loop

    wsOut.Columns("A:K").AutoFit
    MsgBox fileCount & " 個の CSV を取り込みました（" & (nextRow - 2) & " 行）。", vbInformation

CleanUp:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "取込中にエラー: " & Err.Description, vbCritical
    If Not wbCsv Is Nothing Then wbCsv.Close SaveChanges:=False
    Resume CleanUp
End Sub

'--- Q20-2: 支店 × カテゴリ のクロス集計 --------------------
Sub Q20_2_クロス集計()
    Dim wsIn As Worksheet, wsOut As Worksheet
    Dim data As Variant
    Dim branches As Collection, cats As Collection
    Dim grid() As Currency
    Dim i As Long, bi As Long, ci As Long
    Const IN_NAME As String = "取込データ"
    Const OUT_NAME As String = "クロス集計"

    On Error Resume Next
    Set wsIn = ThisWorkbook.Worksheets(IN_NAME)
    On Error GoTo 0
    If wsIn Is Nothing Then
        MsgBox "先に Q20_1_CSVを全部取り込む を実行してください。", vbExclamation
        Exit Sub
    End If

    data = wsIn.Range("A1").CurrentRegion.Value

    Set branches = New Collection
    Set cats = New Collection
    For i = 2 To UBound(data, 1)
        AddUnique branches, CStr(data(i, 3))    ' C列: 支店
        AddUnique cats, CStr(data(i, 8))        ' H列: カテゴリ
    Next i

    ReDim grid(1 To branches.Count, 1 To cats.Count)
    For i = 2 To UBound(data, 1)
        bi = FindIndex(branches, CStr(data(i, 3)))
        ci = FindIndex(cats, CStr(data(i, 8)))
        grid(bi, ci) = grid(bi, ci) + data(i, 11)
    Next i

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(OUT_NAME).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOut.Name = OUT_NAME

    wsOut.Range("A1").Value = "支店＼カテゴリ"
    For ci = 1 To cats.Count
        wsOut.Cells(1, ci + 1).Value = cats(ci)
    Next ci
    wsOut.Cells(1, cats.Count + 2).Value = "合計"

    For bi = 1 To branches.Count
        wsOut.Cells(bi + 1, 1).Value = branches(bi)
        For ci = 1 To cats.Count
            wsOut.Cells(bi + 1, ci + 1).Value = grid(bi, ci)
        Next ci
        wsOut.Cells(bi + 1, cats.Count + 2).Formula = _
            "=SUM(B" & (bi + 1) & ":" & Chr(65 + cats.Count) & (bi + 1) & ")"
    Next bi

    With wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(branches.Count + 1, cats.Count + 2))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 191, 191)
    End With
    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, cats.Count + 2)).Interior.Color = RGB(31, 78, 121)
    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, cats.Count + 2)).Font.Color = RGB(255, 255, 255)
    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, cats.Count + 2)).Font.Bold = True
    wsOut.Range(wsOut.Cells(2, 2), wsOut.Cells(branches.Count + 1, cats.Count + 2)).NumberFormat = "#,##0"
    wsOut.Columns.AutoFit
    wsOut.Activate
End Sub

Private Sub AddUnique(ByVal col As Collection, ByVal key As String)
    On Error Resume Next
    col.Add key, key      ' キー重複ならエラー → 追加されない
    Err.Clear
End Sub

Private Function FindIndex(ByVal col As Collection, ByVal key As String) As Long
    Dim i As Long
    For i = 1 To col.Count
        If col(i) = key Then
            FindIndex = i
            Exit Function
        End If
    Next i
End Function

'--- Q20-3: 請求書テンプレートへ書き出して PDF 出力 ---------
Sub Q20_3_請求書を発行()
    Dim wsT As Worksheet, wsS As Worksheet
    Dim data As Variant
    Dim targetSlip As String
    Dim i As Long, r As Long
    Dim total As Currency
    Dim pdfPath As String

    Set wsT = ThisWorkbook.Worksheets("請求書テンプレート")
    Set wsS = ThisWorkbook.Worksheets("売上明細")

    targetSlip = InputBox("請求書にする伝票番号を入力してください（例: S10001）", "伝票番号", "S10001")
    If targetSlip = "" Then Exit Sub

    ' 明細欄をクリア
    wsT.Range("B10:F24").ClearContents

    data = wsS.Range("A1").CurrentRegion.Value
    r = 10
    total = 0

    For i = 2 To UBound(data, 1)
        If CStr(data(i, 2)) = targetSlip Then
            If r > 24 Then
                MsgBox "明細が15行を超えました。", vbExclamation
                Exit For
            End If
            wsT.Cells(r, "B").Value = r - 9            ' No
            wsT.Cells(r, "C").Value = data(i, 7)       ' 商品名
            wsT.Cells(r, "D").Value = data(i, 10)      ' 数量
            wsT.Cells(r, "E").Value = data(i, 9)       ' 単価
            wsT.Cells(r, "F").Value = data(i, 11)      ' 金額
            wsT.Cells(r, "E").NumberFormat = "#,##0"
            wsT.Cells(r, "F").NumberFormat = "#,##0"
            total = total + data(i, 11)

            wsT.Range("C4").Value = data(i, 5) & " 様"  ' 担当者名を宛先に流用
            wsT.Range("C5").Value = data(i, 3) & " 分 売上"
            wsT.Range("F4").Value = data(i, 1)
            wsT.Range("F4").NumberFormat = "yyyy/mm/dd"
            wsT.Range("F5").Value = targetSlip
            r = r + 1
        End If
    Next i

    If total = 0 Then
        MsgBox "伝票番号 " & targetSlip & " が見つかりませんでした。", vbExclamation
        Exit Sub
    End If

    wsT.Range("C7").Value = Int(total * 1.1)

    ' PDF 出力
    pdfPath = ThisWorkbook.Path & Application.PathSeparator & "請求書_" & targetSlip & ".pdf"
    wsT.PageSetup.Orientation = xlPortrait
    wsT.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, Quality:=xlQualityStandard

    MsgBox "請求書を出力しました。" & vbLf & pdfPath, vbInformation
End Sub

'--- 一括実行 -----------------------------------------------
Sub Q20_すべて実行()
    Q20_1_CSVを全部取り込む
    Q20_2_クロス集計
End Sub
