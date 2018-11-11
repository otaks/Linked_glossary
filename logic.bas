Attribute VB_Name = "logic"
'Option Explicit
Option Base 0

'HTML出力
Sub outputHtml()

    Call outputList 'リストHTML出力
    Call EmptyFolder(ThisWorkbook.Path + "\words\")
    Call outputWords    '単語群HTML出力
    
    MsgBox ("完了しました。用語集.htmlを開いてください。")

End Sub


Public Sub EmptyFolder(strFolder As String)
'フォルダ内を空にする

  Dim strFile As String

  '最初のファイルを検索
  strFile = Dir(strFolder & "*.*")
  Do While strFile <> ""
    'ファイルがあれば削除
    Kill strFolder & strFile
    '次のファイルを検索
    strFile = Dir
  Loop

End Sub

'リストHTML出力
Sub outputList()
    strFileName = ThisWorkbook.Path + "\list.html"  'リストHTML
    
    Dim output As ADODB.Stream
    Set output = New ADODB.Stream
    
    With output
        .Type = adTypeText
        .Charset = "UTF-8"
        .LineSeparator = adLF
        .Open
    End With

    GYO = 5
    
    HTML_code = ""
    HTML_code = HTML_code + "<html>"
    HTML_code = HTML_code + "<head><meta charset=""UTF-8""></head>"
    
    Do Until Worksheets("用語集").Cells(GYO, 4).Value = ""
        If IsOutputed(Worksheets("用語集").Cells(GYO, 4).Value, GYO) Then
            GoTo CONTINUE1
        End If
        aa = Worksheets("用語集").Cells(GYO, 4).Value
        HTML_code = HTML_code + "<a href=""./words/" + aa + ".html"" target=""migi"">" + aa + "</a><br>"
CONTINUE1:
        GYO = GYO + 1
    Loop
    
    HTML_code = HTML_code + "</html>"

    output.WriteText HTML_code, adWriteLine
    output.SaveToFile strFileName, adSaveCreateOverWrite
    output.Close
End Sub

'既にリスト項目出力済みか
Function IsOutputed(name As String, curRow) As Boolean
    GYO = 5
    Do While GYO < curRow
        If Worksheets("用語集").Cells(GYO, 4).Value = name Then
            IsOutputed = True
            Exit Function
        End If
        GYO = GYO + 1
    Loop
    IsOutputed = False
End Function

'単語群HTML出力
Sub outputWords()
    
    Dim words() As String   '単語配列
    
    Call setWords(words)    '単語配列設定
    

    GYO = 5
    Do Until Worksheets("用語集").Cells(GYO, 4).Value = ""
    
        aa = Worksheets("用語集").Cells(GYO, 4).Value
        bb = Worksheets("用語集").Cells(GYO, 5).Value
    
        strFileName = ThisWorkbook.Path + "\words\" + aa + ".html"
        
        If Dir(strFileName) <> "" Then      '既存ファイル有り → 追記
            Dim output As ADODB.Stream
            Set output = New ADODB.Stream
            
            With output
                .Type = adTypeText
                .Charset = "UTF-8"
                .LineSeparator = adLF
                .Open
                .LoadFromFile (strFileName)
                .Position = .Size
                'HTML_code = .ReadText
            End With
            
            'HTML_code = HTML_code +
    
            GYO = GYO + 1
    
            'HTML_code = HTML_code + "</body></html>"
            'output.WriteText HTML_code, adWriteLine
            output.WriteText "<br>---------------------------------<br>" + setLink(bb, words), adWriteLine
            output.SaveToFile strFileName, adSaveCreateOverWrite
            output.Close
            Set output = Nothing
        
        
        Else    '既存ファイル無し
        
            'Dim output As ADODB.Stream
            Set output = New ADODB.Stream
            
            With output
                .Type = adTypeText
                .Charset = "UTF-8"
                .LineSeparator = adLF
                .Open
            End With
            HTML_code = ""
            HTML_code = HTML_code + "<html><body>"
            HTML_code = HTML_code + "<h1>" + aa + "</h1>---------------------------------<br>" + setLink(bb, words)
    
            GYO = GYO + 1
    
            'HTML_code = HTML_code + "</body></html>"
            output.WriteText HTML_code, adWriteLine
            output.SaveToFile strFileName, adSaveCreateOverWrite
            output.Close
            Set output = Nothing
        End If
    Loop
    
    
    
    GYO = 5
    Do Until Worksheets("用語集").Cells(GYO, 4).Value = ""
    
        aa = Worksheets("用語集").Cells(GYO, 4).Value
        bb = Worksheets("用語集").Cells(GYO, 5).Value
    
        strFileName = ThisWorkbook.Path + "\words\" + aa + ".html"
        

        'Dim output As ADODB.Stream
        Set output = New ADODB.Stream
        HTML_code = ""
        With output
            .Type = adTypeText
            .Charset = "UTF-8"
            .LineSeparator = adLF
            .Open
            .LoadFromFile (strFileName)
            .Position = .Size
            'HTML_code = .ReadText
            
        End With

        GYO = GYO + 1

        'HTML_code = HTML_code + "</body></html>"
        'output.WriteText "</body></html>", adWriteLine
        output.WriteText HTML_code, adWriteLine
        output.SaveToFile strFileName, adSaveCreateOverWrite
        output.Close
        Set output = Nothing
    Loop
End Sub

'キーワードリンク設定
Function setLink(bb, words)
    Dim replaceTarget()
    ReDim replaceTarget(UBound(words, 2))
    
    Call getTarget(bb, words, replaceTarget)      '置換対象取得

    GYO = 5
    
    For i = 0 To UBound(words, 2)
        aa = words(0, i)
        
        If replaceTarget(i) = 1 Then
            bb = Replace(bb, aa, "<a href=""" + aa + ".html"">" + aa + "</a>")
        End If
        GYO = GYO + 1
    'Loop
    Next i
    setLink = bb
End Function

Sub getTarget(bb, words, replaceTarget)    '置換対象取得

    Dim tmp()
    ReDim tmp(Len(bb))
    
    For i = 0 To UBound(words, 2)
        If InStr(bb, words(0, i)) Then  '当該単語を含むか
            If haveReplaced(tmp, bb, words(0, i)) = False Then  '塗られていないか？
                Call doReplace(tmp, bb, words(0, i))    '塗る
                replaceTarget(i) = 1
            End If
        End If
    
    Next i

End Sub

'塗られていないか？
Function haveReplaced(tmp, bb, word)

    pos = InStr(bb, word)
    flg = 0
    For i = 0 To Len(word) - 1
        If tmp(pos - 1 + i) = 1 Then
            flg = 1
            Exit For
        End If
    Next i
    
    If flg = 1 Then
        haveReplaced = True
    Else
        haveReplaced = False
    End If
End Function

'塗る
Sub doReplace(tmp, bb, word)

    pos = InStr(bb, word)
    For i = 0 To Len(word) - 1
        tmp(pos - 1 + i) = 1
    Next i
    
End Sub

'単語配列設定
Sub setWords(words)
    GYO = 5
    cnt = 0
    
    '配列に取り込む
    Do Until Worksheets("用語集").Cells(GYO, 4).Value = ""
        aa = Worksheets("用語集").Cells(GYO, 4).Value
        bb = Worksheets("用語集").Cells(GYO, 5).Value

        ReDim Preserve words(1, cnt)
        words(0, cnt) = aa
        words(1, cnt) = bb
        cnt = cnt + 1
        GYO = GYO + 1
    Loop
    
    '単語の長さでソート
    For i = 0 To UBound(words, 2)
        For j = i + 1 To UBound(words, 2)
            If Len(words(0, i)) < Len(words(0, j)) Then
                swap1 = words(0, i)
                swap2 = words(1, i)
                words(0, i) = words(0, j)
                words(1, i) = words(1, j)
                words(0, j) = swap1
                words(1, j) = swap2
            End If
        Next j
    Next i
End Sub
