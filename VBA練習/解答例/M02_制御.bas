Attribute VB_Name = "M02_制御"
Option Explicit

'==============================================================
' Lv2 制御構文
'==============================================================

'--- Q5: If で在庫の発注判定 --------------------------------
Sub Q5_発注判定()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long

    Set ws = ThisWorkbook.Worksheets("在庫")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, "D").Value = 0 Then
            ws.Cells(i, "G").Value = "欠品"
            ws.Cells(i, "G").Font.Color = RGB(192, 0, 0)
        ElseIf ws.Cells(i, "D").Value <= ws.Cells(i, "E").Value Then
            ws.Cells(i, "G").Value = "要発注"
            ws.Cells(i, "G").Font.Color = RGB(255, 140, 0)
        Else
            ws.Cells(i, "G").Value = "OK"
            ws.Cells(i, "G").Font.Color = RGB(0, 128, 0)
        End If
    Next i
End Sub

'--- Q6: For で合計を求める ---------------------------------
Sub Q6_売上合計()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim total As Currency   ' 金額の合計は Currency か Double が安全

    Set ws = ThisWorkbook.Worksheets("売上明細")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    total = 0
    For i = 2 To lastRow
        total = total + ws.Cells(i, "K").Value
    Next i

    ' 検算: WorksheetFunction を使えば1行で書ける
    MsgBox "ループ集計 : " & Format(total, "#,##0") & " 円" & vbLf & _
           "SUM関数   : " & Format(WorksheetFunction.Sum(ws.Range("K2:K" & lastRow)), "#,##0") & " 円"
End Sub

'--- Q7: Do While で最終行まで走る --------------------------
Sub Q7_DoWhileで走査()
    Dim ws As Worksheet
    Dim i As Long
    Dim cnt As Long

    Set ws = ThisWorkbook.Worksheets("売上明細")

    i = 2
    Do While ws.Cells(i, "A").Value <> ""
        If ws.Cells(i, "J").Value >= 10 Then cnt = cnt + 1
        i = i + 1
    Loop

    MsgBox "数量10個以上の明細は " & cnt & " 件です。" & vbLf & _
           "（走査した行数: " & (i - 2) & " 行）"
End Sub

'--- Q8: Select Case でランク付け ---------------------------
Sub Q8_給与ランク()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim salary As Long

    Set ws = ThisWorkbook.Worksheets("社員名簿")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Range("J1").Value = "給与ランク"
    ws.Range("J1").Font.Bold = True

    For i = 2 To lastRow
        salary = ws.Cells(i, "I").Value
        Select Case salary
            Case Is >= 500000
                ws.Cells(i, "J").Value = "A"
            Case 400000 To 499999
                ws.Cells(i, "J").Value = "B"
            Case 300000 To 399999
                ws.Cells(i, "J").Value = "C"
            Case Else
                ws.Cells(i, "J").Value = "D"
        End Select
    Next i

    ws.Columns("J").AutoFit
End Sub

'--- Q9: 二重ループで支店別集計 -----------------------------
Sub Q9_支店別集計()
    Dim wsS As Worksheet, wsOut As Worksheet
    Dim branches As Variant
    Dim lastRow As Long, i As Long, b As Long
    Dim total As Currency

    Set wsS = ThisWorkbook.Worksheets("売上明細")
    Set wsOut = ThisWorkbook.Worksheets("集計結果")
    lastRow = wsS.Cells(wsS.Rows.Count, "A").End(xlUp).Row

    branches = Array("東京本社", "大阪支店", "名古屋支店", "福岡支店", "札幌支店")

    wsOut.Range("D1").Value = "支店"
    wsOut.Range("E1").Value = "売上合計"
    wsOut.Range("D1:E1").Font.Bold = True

    For b = LBound(branches) To UBound(branches)
        total = 0
        For i = 2 To lastRow
            If wsS.Cells(i, "C").Value = branches(b) Then
                total = total + wsS.Cells(i, "K").Value
            End If
        Next i
        wsOut.Cells(b + 2, "D").Value = branches(b)
        wsOut.Cells(b + 2, "E").Value = total
        wsOut.Cells(b + 2, "E").NumberFormat = "#,##0"
    Next b

    wsOut.Columns("D:E").AutoFit
    MsgBox "支店別集計が終わりました。", vbInformation
End Sub
