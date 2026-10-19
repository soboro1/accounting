Attribute VB_Name = "M01_基礎"
Option Explicit

'==============================================================
' Lv1 基礎
'==============================================================

'--- Q2: 売上明細の行数を表示する ---------------------------
Sub Q2_行数を表示()
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets("売上明細")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    MsgBox "売上明細は見出しを除いて " & (lastRow - 1) & " 行あります。", _
           vbInformation, "行数チェック"
End Sub

'--- Q3: セルの読み書き -------------------------------------
Sub Q3_セルに書き込む()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("集計結果")

    ws.Range("A1").Value = "作成日"
    ws.Range("B1").Value = Date
    ws.Range("B1").NumberFormat = "yyyy/mm/dd"

    ws.Range("A2").Value = "作成者"
    ws.Range("B2").Value = Application.UserName

    ws.Range("A1:A2").Font.Bold = True
    ws.Columns("A:B").AutoFit
End Sub

'--- Q4: 変数・型・計算 -------------------------------------
Sub Q4_金額を計算()
    Dim ws As Worksheet
    Dim price As Long          ' 単価（整数）
    Dim qty As Long            ' 数量（整数）
    Dim amount As Long         ' 金額
    Dim taxIncluded As Double  ' 税込（小数が出るので Double）
    Const TAX_RATE As Double = 0.1

    Set ws = ThisWorkbook.Worksheets("売上明細")
    price = ws.Range("I2").Value
    qty = ws.Range("J2").Value

    amount = price * qty
    taxIncluded = amount * (1 + TAX_RATE)

    MsgBox "単価: " & Format(price, "#,##0") & " 円" & vbLf & _
           "数量: " & qty & vbLf & _
           "金額: " & Format(amount, "#,##0") & " 円" & vbLf & _
           "税込: " & Format(taxIncluded, "#,##0") & " 円"
End Sub
