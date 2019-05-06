Option Explicit

Private tryCnt As Long

Sub reset()

Dim clear As VbMsgBoxResult
  clear = MsgBox("本当に数独をリセットしても大丈夫ですか？数独開始時の初期状態に戻ります。", vbYesNo, Title:="リセット")

If clear = vbYes Then
'セルをリセット1回目
Columns("A:AZ").ColumnWidth = 8.38
Rows("1:300").RowHeight = 25
Range("A1:AZ300").clear
'セルをリセット2回目
Columns("A:I").ColumnWidth = 8.38
Rows("1:9").RowHeight = 54

'罫線を引く
   Dim bs As Borders
   Set bs = Range("A1:I9").Borders
   bs.LineStyle = xlContinuous
   bs.Weight = xlThin
   Dim cs As Border
   Set cs = Range("A3:I3").Borders(xlEdgeBottom)
   cs.LineStyle = xlDouble
   Dim ds As Border
   Set ds = Range("A6:I6").Borders(xlEdgeBottom)
   ds.LineStyle = xlDouble
   Dim es As Border
   Set es = Range("C1:C9").Borders(xlEdgeRight)
   es.LineStyle = xlDouble
   Dim fs As Border
   Set fs = Range("F1:F9").Borders(xlEdgeRight)
   fs.LineStyle = xlDouble
   
   '中央揃え&フォントサイズ
   With Range("A1: I9")
        .HorizontalAlignment = xlCenter
    End With
    
   With Range("A1:I9").Font
        .Size = 22
    End With

   MsgBox ("適当に罫線が引かれているセルに'1から9までの数字を'入れてください。入れ終わったら解くボタンを押してください。自動で解きはじめます。"), Buttons:=vbInformation, Title:="数字をどうぞ"
End If

End Sub

   Sub sudoku()
   
    Debug.Print Timer
    Dim Ar(1 To 9, 1 To 9) As Integer
    Dim i1 As Integer
    Dim i2 As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim d As Integer
    Dim count As Integer
    Dim time As Double
    count = 0
    tryCnt = 0
    time = Timer
    Erase Ar
    Range("A1:I9").Font.ColorIndex = xlAutomatic
    
    For i = 1 To 9
     For j = 1 To 9
      For k = 1 To 9
       For l = 1 To 9
       
           '初期化
           a = i
           b = j
           c = k
           d = l
       
         '同行・同列に同じ数字があった場合エラー吐かせるようにする
          If (Cells(i, j) = Cells(k, l) And Cells(i, j) <> "" And i <> k And j = l) Then
            count = count + 1
          End If
      
          If (Cells(i, j) = Cells(k, l) And Cells(i, j) <> "" And i = k And j <> l) Then
            count = count + 1
          End If
      
        '3*3枠内チェック
          If (Cells(i, j) = Cells(k, l) And Cells(i, j) <> "" And i <> k And i <= 3 And k <= 3 And j <> l And j <= 3 And l <= 3) Then '33
            count = count + 1
          End If
     
          If (Cells(i, j) = Cells(k, l) And Cells(i, j) <> "" And i <> k And i > 3 And k > 3 And i <= 6 And k <= 6 And j <> l And j > 3 And l > 3 And j <= 6 And l <= 6) Then '66
            count = count + 1
          End If
      
          If (Cells(i, j) = Cells(k, l) And Cells(i, j) <> "" And i <> k And i > 6 And k > 6 And j <> l And j > 6 And l > 6) Then '99
            count = count + 1
          End If
      
          If (Cells(i, j) = Cells(k, l) And Cells(i, j) <> "" And i <> k And i > 3 And k > 3 And i <= 6 And k <= 6 And j <> l And j <= 3 And l <= 3) Then  '63
            count = count + 1
          End If
     
          If (Cells(i, j) = Cells(k, l) And Cells(i, j) <> "" And i <> k And i <= 3 And k <= 3 And j <> l And j > 3 And l > 3 And j <= 6 And l <= 6) Then  '36
            count = count + 1
          End If
      
          If (Cells(i, j) = Cells(k, l) And Cells(i, j) <> "" And i <> k And i > 6 And k > 6 And j <> l And j <= 3 And l <= 3) Then '93
            count = count + 1
          End If
      
          If (Cells(i, j) = Cells(k, l) And Cells(i, j) <> "" And i <> k And i <= 3 And k <= 3 And j <> l And j > 6 And l > 6) Then '39
            count = count + 1
          End If
     
          If (Cells(i, j) = Cells(k, l) And Cells(i, j) <> "" And i <> k And i > 6 And k > 6 And j <> l And j > 3 And l > 3 And j <= 6 And l <= 6) Then '96
            count = count + 1
          End If
      
          If (Cells(i, j) = Cells(k, l) And Cells(i, j) <> "" And i <> k And i > 3 And k > 3 And i <= 6 And k <= 6 And j <> l And j > 6 And l > 6) Then  '69
            count = count + 1
          End If
      
        Next
       Next
      Next
     Next
     
  If (count = 0) Then
    For i1 = 1 To 9
        For i2 = 1 To 9
            If Cells(i1, i2) = "" Then
               Cells(i1, i2).Font.Color = vbBlue
            Else
               Ar(i1, i2) = Cells(i1, i2)
            End If
        Next
    Next
  
    Call trySu(Ar)
  
    Range("A1:I9").Value = Ar
    Debug.Print Timer
    Else
        MsgBox "解読できません。解けない数独(同じ列・行・3*3のマス目内に同じ数字が入っている等)を解かせようとしていませんか？" & Chr(13) & "同じものを連続して解かせようとしてもこのメッセージが出ます。リセットしてください。", Title:="かなしいおしらせ"
  End If
  
  If getBlank(Ar(), i1, i2) = False Then
        MsgBox "解読成功しました。" & Chr(13) & tryCnt & "手試行しました。" & Chr(13) & "かかった時間は" & Round(Timer - time, 1) & "秒です。", Title:="うれしいおしらせ"
    End If
  
End Sub

Sub readme()

MsgBox "これは9*9のマス目の数独を数秒くらいで解いてくれる(はず)のプログラムです。" & Chr(13) & "リセットボタンを押すとセルがあるべき状態に戻り、罫線が引かれます。そうしたら、出てくるメッセージボックスの指示に従ってください。" & Chr(13) & "正答が複数ある場合、その1つを返します。" & Chr(13) & "このプログラムは速く解くということは念頭に置いていません。難しいと(全探索するので)、最長で(テンプレート)2分越えします(簡単なのは一瞬)。", Title:="説明"

End Sub

Sub quote_note()

MsgBox "リセットの部分のMarshmello: https://goo.gl/ETA6vk" & Chr(13) & "りいどみいの本:https://goo.gl/Y2bn69" & Chr(13) & "テンプレートの数独:フィンランド人数学者のArto Inkala氏作", Title:="引用注"

End Sub

Sub error()

MsgBox "(数字はエラー番号でこの中に無いのも出るかもしれません)" & Chr(13) & "9:数字を1から9の間で入力していない。" & Chr(13) & "6:オーバーフロー。上に同じ。" & Chr(13) & "(無いようにしたつもりだが)リセットしてもまだ数字を埋め込んでいる:数字の入っているセル1つを空にしてからリセット。"

End Sub

Sub temp()

Range("A1").Value = 8
Range("C2").Value = 3
Range("D2").Value = 6
Range("B3").Value = 7
Range("E3").Value = 9
Range("G3").Value = 2
Range("B4").Value = 5
Range("F4").Value = 7
Range("E5").Value = 4
Range("F5").Value = 5
Range("G5").Value = 7
Range("D6").Value = 1
Range("H6").Value = 3
Range("C7").Value = 1
Range("H7").Value = 6
Range("I7").Value = 8
Range("C8").Value = 8
Range("D8").Value = 5
Range("H8").Value = 1
Range("B9").Value = 9
Range("G9").Value = 4

End Sub

Function trySu(ByRef Ar() As Integer) As Boolean

    Dim i1 As Integer
    Dim i2 As Integer
    Dim su As Integer
    Dim tryAry() As Integer
    
    If getBlank(Ar(), i1, i2) = False Then
        trySu = True
        Exit Function
    End If
    
    If chkSu(Ar(), i1, i2, tryAry()) <> 0 Then
        For su = 1 To 9
            If tryAry(su) <> 0 Then
                Ar(i1, i2) = su
                tryCnt = tryCnt + 1
                Cells(i1, i2) = su
                If trySu(Ar) = True Then
                    trySu = True
                    Exit Function
                End If
            End If
        Next
    End If
    
    Ar(i1, i2) = 0
    Cells(i1, i2) = ""
    DoEvents
    trySu = False
    
End Function

Function getBlank(ByRef Ar() As Integer, ByRef i1 As Integer, ByRef i2 As Integer) As Boolean

    Dim cnt As Integer
    Dim tryMin As Integer
    Dim i1Min As Integer
    Dim i2Min As Integer
    Dim tryAry() As Integer
    Dim chkAry1(1 To 9, 1 To 9) As Integer
    Dim chkAry2(1 To 9, 1 To 9) As Integer
    tryMin = 10
    
    For i1 = 1 To 9
        For i2 = 1 To 9
            If Ar(i1, i2) = 0 Then
                chkAry1(i1, i2) = chkSu(Ar, i1, i2, tryAry)
            End If
        Next
    Next
  
    Dim ix1 As Integer
    Dim ix2 As Integer
    Dim i1S As Integer
    Dim i2S As Integer
    
    For i1 = 1 To 9
        For i2 = 1 To 9
            If Ar(i1, i2) = 0 Then
                cnt = 0
                '横を合計
                For ix2 = 1 To 9
                    If ix2 <> i2 Then
                        If chkAry1(i1, ix2) <> 0 Then cnt = cnt + 1
                    End If
                Next
                '縦を合計
                For ix1 = 1 To 9
                    If ix1 <> i1 Then
                        If chkAry1(ix1, i2) <> 0 Then cnt = cnt + 1
                    End If
                Next
                '枠内を合計
                i1S = (Int((i1 + 2) / 3) - 1) * 3 + 1
                i2S = (Int((i2 + 2) / 3) - 1) * 3 + 1
                For ix1 = i1S To i1S + 2
                    For ix2 = i2S To i2S + 2
                        If ix1 <> i1 And ix2 <> i2 Then
                            If chkAry1(ix1, ix2) <> 0 Then cnt = cnt + 1
                        End If
                    Next
                Next
                chkAry2(i1, i2) = chkAry1(i1, i2) * 1000 + cnt
            End If
        Next
    Next
  
    tryMin = 9999
    
    For i1 = 1 To 9
        For i2 = 1 To 9
            If Ar(i1, i2) = 0 Then
                If tryMin > chkAry2(i1, i2) Then
                    i1Min = i1
                    i2Min = i2
                    tryMin = chkAry2(i1, i2)
                End If
            End If
        Next
    Next
  
    If tryMin = 9999 Then
        getBlank = False
    Else
        i1 = i1Min
        i2 = i2Min
        getBlank = True
    End If
    
End Function

Function chkSu(ByRef Ar() As Integer, ByVal i1 As Integer, ByVal i2 As Integer, ByRef tryAry() As Integer) As Integer

    Dim ix1 As Integer
    Dim ix2 As Integer
    Dim i1S As Integer
    Dim i2S As Integer
    chkSu = False
    ReDim tryAry(1 To 9)
    
    For ix1 = 1 To 9
        tryAry(ix1) = ix1
    Next
  
    '横チェック
    For ix2 = 1 To 9
        If ix2 <> i2 Then
            If Ar(i1, ix2) <> 0 Then
                tryAry(Ar(i1, ix2)) = 0
            End If
        End If
    Next
    
    '縦チェック
    For ix1 = 1 To 9
        If ix1 <> i1 Then
            If Ar(ix1, i2) <> 0 Then
                tryAry(Ar(ix1, i2)) = 0
            End If
        End If
    Next
    
    '枠内チェック
    i1S = (Int((i1 + 2) / 3) - 1) * 3 + 1
    i2S = (Int((i2 + 2) / 3) - 1) * 3 + 1
    
    For ix1 = i1S To i1S + 2
        For ix2 = i2S To i2S + 2
            If ix1 <> i1 Or ix2 <> i2 Then
                If Ar(ix1, ix2) <> 0 Then
                    tryAry(Ar(ix1, ix2)) = 0
                End If
            End If
        Next
    Next
    
    chkSu = 0
    
    For ix1 = 1 To 9
        If tryAry(ix1) <> 0 Then
            chkSu = chkSu + 1
        End If
    Next
    
End Function
