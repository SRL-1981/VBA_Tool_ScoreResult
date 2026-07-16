Attribute VB_Name = "Module1"
Sub ScoreResult()

  Dim wbThis As Workbook
  Set wbThis = ThisWorkbook
  
'既存の「試験結果」シートを定義
  Dim wsResult As Worksheet
  Set wsResult = wbThis.Worksheets("試験結果")

  Dim wsScore As Worksheet
  Dim arr As Variant
  Dim i As Long
  Dim j As Long
  Dim last As Long
  Dim TotalScore As Long
  Dim Avg As Double
  Dim JudgeResult As String

'新規作成する「成績一覧」シートが重複して存在していた場合の処理
On Error Resume Next
  wbThis.Worksheets("成績一覧").Delete
On Error GoTo 0

  Set wsScore = wbThis.Worksheets.Add
  wsScore.Name = "成績一覧"
  Debug.Print "「成績一覧」シート作成"

  last = wsResult.Cells(wsResult.Rows.Count, 1).End(xlUp).Row
  
  '「試験結果」シートのデータを配列へ格納
  arr = wsResult.Range("a2:e" & last).Value
  
'「成績一覧」シートのヘッダー入力
  wsScore.Range("a1:d1").Value = Array("受験番号", "氏名", "平均点", "判定")
    
  j = 2
  
'格納した配列から「成績一覧」シートへ各列の項目を行単位で出力
For i = 1 To UBound(arr, 1)

  Avg = CalculateAvg(arr(i, 3), arr(i, 4), arr(i, 5))
  JudgeResult = Judge(Avg)

  wsScore.Cells(j, 1).Value = arr(i, 1)   '受験番号
  wsScore.Cells(j, 2).Value = arr(i, 2)   '氏名
  wsScore.Cells(j, 3).Value = Avg         '3教科の平均点
  wsScore.Cells(j, 4).Value = JudgeResult '平均点から合格・不合格の判定

  j = j + 1

Next i

End Sub

'3教科の平均点を算出する関数
Function CalculateAvg(Japanese As Variant, Math As Variant, Eng As Variant) As Double

  CalculateAvg = Round(Application.WorksheetFunction.Average(Japanese, Math, Eng), 1)
  
End Function

'3教科の平均点から合格・不合格の判定を行う関数
Function Judge(Avg As Double) As String

  If Avg >= 70 Then
    Judge = "合格"
  
  Else
    Judge = "不合格"
    
  End If
  
End Function
