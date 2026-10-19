Attribute VB_Name = "M04_応用"
Option Explicit

'==============================================================
' Lv4 応用
' ※ Scripting.Dictionary は Windows 版 Excel 専用です。
'    Mac の場合は Q15_担当者別集計_Mac版（Collection 方式）を使ってください。
'==============================================================

'--- Q15: Dictionary で担当者別集計（Windows） --------------
Sub Q15_担当者別集計()
    Dim ws As Worksheet, wsOut As Worksheet
    Dim dic As Object
    Dim data As Variant
    Dim i As Long, r As Long
    Dim key As String

    Set ws = ThisWorkbook.Worksheets("売上明細")
    Set wsOut = ThisWorkbook.Worksheets("集計結果")
    Set dic = CreateObject("Scripting.Dictionary")

    ' セルを1回だけ読んで配列に入れる（これが速さの肝）
    data = ws.Range("A1").CurrentRegion.Value

    For i = 2 To UBound(data, 1)
        key = data(i, 5)                       ' E列: 担当者名
        dic(key) = dic(key) + data(i, 11)      ' K列: 金額
    Next i

    wsOut.Range("G:H").ClearContents
    wsOut.Range("G1").Value = "担当者"
    wsOut.Range("H1").Value = "売上合計"
    wsOut.Range("G1:H1").Font.Bold = True

    r = 2
    Dim k As Variant
    For Each k In dic.Keys
        wsOut.Cells(r, "G").Value = k
        wsOut.Cells(r, "H").Value = dic(k)
        wsOut.Cells(r, "H").NumberFormat = "#,##0"
        r = r + 1
    Next k

    wsOut.Columns("G:H").AutoFit
    MsgBox dic.Count & " 人分を集計しました。", vbInformation
End Sub

'--- Q15': Mac でも動く Collection 版 -----------------------
Sub Q15_担当者別集計_Mac版()
    Dim ws As Worksheet, wsOut As Worksheet
    Dim names As Collection
    Dim data As Variant
    Dim sums() As Currency
    Dim i As Long, idx As Long, r As Long
    Dim key As String

    Set ws = ThisWorkbook.Worksheets("売上明細")
    Set wsOut = ThisWorkbook.Worksheets("集計結果")
    Set names = New Collection
    ReDim sums(1 To 1000)

    data = ws.Range("A1").CurrentRegion.Value

    For i = 2 To UBound(data, 1)
        key = CStr(data(i, 5))
        idx = IndexOfKey(names, key)
        If idx = 0 Then
            names.Add key
            idx = names.Count
        End If
        sums(idx) = sums(idx) + data(i, 11)
    Next i

    wsOut.Range("G:H").ClearContents
    wsOut.Range("G1").Value = "担当者"
    wsOut.Range("H1").Value = "売上合計"
    wsOut.Range("G1:H1").Font.Bold = True

    For r = 1 To names.Count
        wsOut.Cells(r + 1, "G").Value = names(r)
        wsOut.Cells(r + 1, "H").Value = sums(r)
        wsOut.Cells(r + 1, "H").NumberFormat = "#,##0"
    Next r

    wsOut.Columns("G:H").AutoFit
    MsgBox names.Count & " 人分を集計しました。", vbInformation
End Sub

Private Function IndexOfKey(ByVal col As Collection, ByVal key As String) As Long
    Dim i As Long
    For i = 1 To col.Count
        If col(i) = key Then
            IndexOfKey = i
            Exit Function
        End If
    Next i
    IndexOfKey = 0
End Function

'--- Q16: 配列で高速化＋速度比較 ----------------------------
Sub Q16_速度比較()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim t1 As Double, t2 As Double
    Dim data As Variant
    Dim total As Currency

    Set ws = ThisWorkbook.Worksheets("売上明細")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' 方法1: 1セルずつ読む（遅い）
    t1 = Timer
    total = 0
    For i = 2 To lastRow
        total = total + ws.Cells(i, "K").Value
    Next i
    t1 = Timer - t1

    ' 方法2: いったん配列に入れる（速い）
    t2 = Timer
    data = ws.Range("K2:K" & lastRow).Value
    total = 0
    For i = 1 To UBound(data, 1)
        total = total + data(i, 1)
    Next i
    t2 = Timer - t2

    MsgBox "セル直接 : " & Format(t1, "0.000") & " 秒" & vbLf & _
           "配列経由 : " & Format(t2, "0.000") & " 秒" & vbLf & vbLf & _
           "合計金額 : " & Format(total, "#,##0") & " 円"
End Sub

'--- Q16': 一括書き戻しの型 ---------------------------------
Sub Q16b_税込列を一括追加()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim src As Variant, dst() As Variant

    Set ws = ThisWorkbook.Worksheets("売上明細")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    src = ws.Range("K2:K" & lastRow).Value
    ReDim dst(1 To UBound(src, 1), 1 To 1)

    For i = 1 To UBound(src, 1)
        dst(i, 1) = Int(src(i, 1) * 1.1)
    Next i

    ws.Range("L1").Value = "税込金額"
    ws.Range("L1").Font.Bold = True
    ws.Range("L2").Resize(UBound(dst, 1), 1).Value = dst   ' ← 1回で書き戻す
    ws.Range("L2:L" & lastRow).NumberFormat = "#,##0"
End Sub

'--- Q17: エラー処理と高速化のおまじない --------------------
Sub Q17_安全に処理する()
    Dim ws As Worksheet
    Dim savedCalc As XlCalculation

    ' 開始処理
    savedCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets("存在しないシート")   ' ← わざとエラー
    ws.Range("A1").Value = "ここには来ない"

CleanUp:
    ' 何があっても必ず元に戻す
    Application.Calculation = savedCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "エラーが発生しました。" & vbLf & vbLf & _
           "番号: " & Err.Number & vbLf & _
           "内容: " & Err.Description, vbCritical, "処理中断"
    Resume CleanUp
End Sub

'--- Q18: ユーザー定義関数 ----------------------------------
' セルに =税込(A1) のように書いて使えます
Public Function 税込(ByVal amount As Double, Optional ByVal rate As Double = 0.1) As Double
    税込 = Int(amount * (1 + rate))
End Function

' 全角英数字・カナを半角に、前後の空白を除去
Public Function 正規化(ByVal s As Variant) As String
    If IsError(s) Or IsEmpty(s) Then
        正規化 = ""
        Exit Function
    End If
    Dim t As String
    t = CStr(s)
    t = StrConv(t, vbNarrow)          ' 全角 → 半角
    t = Replace(t, vbTab, " ")
    t = Replace(t, "　", " ")          ' 全角スペース → 半角
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")      ' 連続スペースを1つに
    Loop
    正規化 = Trim(t)
End Function

' 年度（4月始まり）を返す
Public Function 年度(ByVal d As Date) As Long
    If Month(d) >= 4 Then
        年度 = Year(d)
    Else
        年度 = Year(d) - 1
    End If
End Function

'--- Q19: 汚いデータのクレンジング --------------------------
Sub Q19_データを整える()
    Dim wsSrc As Worksheet, wsOut As Worksheet
    Dim lastRow As Long, i As Long, r As Long
    Dim seen As Collection
    Dim code As String
    Const OUT_NAME As String = "整形済みデータ"

    Set wsSrc = ThisWorkbook.Worksheets("要整形データ")
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(OUT_NAME).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOut.Name = OUT_NAME
    wsOut.Range("A1:G1").Value = Array("顧客番号", "会社名", "電話番号", "郵便番号", "住所", "契約金額", "契約日")
    wsOut.Range("A1:G1").Font.Bold = True

    Set seen = New Collection
    r = 2

    For i = 2 To lastRow
        code = UCase(正規化(wsSrc.Cells(i, "A").Value))

        If code <> "" Then
            ' 重複チェック（Collection のキー重複エラーを利用する定番手法）
            On Error Resume Next
            seen.Add code, code
            If Err.Number = 0 Then
                On Error GoTo 0
                wsOut.Cells(r, "A").Value = code
                wsOut.Cells(r, "B").Value = 正規化(wsSrc.Cells(i, "B").Value)
                wsOut.Cells(r, "C").Value = Replace(Replace(正規化(wsSrc.Cells(i, "C").Value), "‐", "-"), "−", "-")
                wsOut.Cells(r, "D").Value = 郵便番号整形(wsSrc.Cells(i, "D").Value)
                wsOut.Cells(r, "E").Value = 正規化(wsSrc.Cells(i, "E").Value)
                wsOut.Cells(r, "F").Value = 金額に変換(wsSrc.Cells(i, "F").Value)
                wsOut.Cells(r, "G").Value = 日付に変換(wsSrc.Cells(i, "G").Value)
                wsOut.Cells(r, "F").NumberFormat = "#,##0"
                wsOut.Cells(r, "G").NumberFormat = "yyyy/mm/dd"
                r = r + 1
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next i

    wsOut.Columns("A:G").AutoFit
    MsgBox (r - 2) & " 件に整形しました（空行・重複を除去）。", vbInformation
End Sub

Private Function 郵便番号整形(ByVal v As Variant) As String
    Dim s As String
    s = Replace(正規化(v), "-", "")
    If Len(s) = 7 Then
        郵便番号整形 = Left(s, 3) & "-" & Right(s, 4)
    Else
        郵便番号整形 = s
    End If
End Function

Private Function 金額に変換(ByVal v As Variant) As Variant
    Dim s As String
    s = 正規化(v)
    s = Replace(s, ",", "")
    s = Replace(s, "，", "")
    s = Replace(s, "円", "")
    s = Replace(s, " ", "")
    If IsNumeric(s) Then
        金額に変換 = CDbl(s)
    Else
        金額に変換 = s
    End If
End Function

Private Function 日付に変換(ByVal v As Variant) As Variant
    Dim s As String
    s = 正規化(v)
    s = Replace(s, "年", "/")
    s = Replace(s, "月", "/")
    s = Replace(s, "日", "")
    s = Replace(s, ".", "/")

    ' 和暦（R7 → 2025）
    If UCase(Left(s, 1)) = "R" Then
        Dim parts As Variant
        parts = Split(Mid(s, 2), "/")
        If UBound(parts) >= 2 Then
            s = (2018 + CLng(parts(0))) & "/" & parts(1) & "/" & parts(2)
        End If
    End If

    If IsDate(s) Then
        日付に変換 = CDate(s)
    Else
        日付に変換 = s
    End If
End Function
