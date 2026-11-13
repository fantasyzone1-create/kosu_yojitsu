Attribute VB_Name = "MonthSheetTools"
Option Explicit

' 「シート追加」シートの年月を読み取り、月シート(YY年MM月)を追加するマクロ。
' 既存の月シートがあれば、最新シートのメンバー名を引き継ぐ。

Private Const CTRL_SHEET As String = "シート追加"
Private Const IMPORT_SHEET As String = "実績取込"
Private Const MEMBER_BLOCKS As Long = 12
Private Const HEADER_ROW As Long = 6
Private Const FIRST_ROW As Long = 7
Private Const FONT_NAME As String = "游ゴシック"
Private Const HOLIDAY_SHEET As String = "祝日"
Private Const HOLIDAY_NAME As String = "祝日リスト"   ' 稼働日数の計算で除外する日付の範囲名
Private Const DEFAULT_DAILY As Double = 7.5            ' 1日の定時工数(h)の初期値
Private Const DAILY_LABEL As String = "1日の定時工数(h)"

' 実績取込で使う正規表現(期間 / 工数計)。モジュール先頭で宣言する必要がある
Private reP As Object
Private reH As Object
Private diagSrc As String   ' 取込結果に付ける診断情報

' ボタンに割り当てるマクロ
Public Sub AddMonthSheet()
    Dim wsC As Worksheet, msg As String
    Set wsC = ThisWorkbook.Worksheets(CTRL_SHEET)
    msg = CreateMonthSheet(wsC.Range("C4").Value, wsC.Range("C5").Value)
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation
    Else
        MsgBox "シートを追加しました。", vbInformation
    End If
End Sub

' 月シートを追加する。成功なら空文字、失敗ならエラーメッセージを返す(テスト用に公開)
Public Function CreateMonthSheet(ByVal y As Variant, ByVal m As Variant) As String
    Dim nm As String, members As Variant

    If IsEmpty(y) Or IsEmpty(m) Then
        CreateMonthSheet = "年と月を入力してください。": Exit Function
    End If
    If Not IsNumeric(y) Or Not IsNumeric(m) Then
        CreateMonthSheet = "年と月は数値で入力してください。": Exit Function
    End If
    y = CLng(y): m = CLng(m)
    If y < 2000 Or y > 2099 Or m < 1 Or m > 12 Then
        CreateMonthSheet = "年は2000～2099、月は1～12で入力してください。": Exit Function
    End If

    nm = Format(y Mod 100, "00") & "年" & Format(m, "00") & "月"
    If SheetExists(nm) Then
        CreateMonthSheet = "シート「" & nm & "」は既に存在します。": Exit Function
    End If

    members = ReadMembers(LatestMonthSheet())

    Application.ScreenUpdating = False
    BuildSheet CLng(y), CLng(m), nm, members
    Application.ScreenUpdating = True
End Function

' 「シート追加」シートにボタンを設置する(最初に1回だけ実行)
Public Sub SetupButton()
    Dim ws As Worksheet, btn As Object, r As Range
    Set ws = ThisWorkbook.Worksheets(CTRL_SHEET)
    ws.Buttons.Delete
    Set r = ws.Range("B8:C9")
    Set btn = ws.Buttons.Add(r.Left, r.Top, r.Width, r.Height)
    btn.Caption = "シートを追加"
    btn.OnAction = "AddMonthSheet"
    btn.Font.Size = 12
    SetupImportSheet
    SetupHolidaySheet
    AddCapacityToAllMonthSheets   ' 旧バージョンの月シートに定時工数欄を追加(追加済みなら何もしない)
End Sub

' ===== 実績取込(元ファイルの工数計をテンプレートの「実績」行へ反映) =====

' 実績取込シートを作る(無ければ)。SetupButtonから呼ばれる
Private Sub SetupImportSheet()
    Dim ws As Worksheet, btn As Object, r As Range
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(IMPORT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(CTRL_SHEET))
        ws.Name = IMPORT_SHEET
        ws.Cells.Font.Name = FONT_NAME
        ws.Range("A1").Value = "実績の取込"
        ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
        ws.Range("A2").Value = "元ファイル(SharePoint)の「工数計」を、各月シートの「実績」行へ反映します。"
        ws.Range("A2").Font.Size = 9: ws.Range("A2").Font.Color = RGB(89, 89, 89)
        ws.Range("B4").Value = "元ファイルのURL/パス"
        StyleCell ws.Range("B4"), True, 0, RGB(242, 242, 242)
        StyleCell ws.Range("C4"), False, RGB(0, 0, 255), RGB(255, 242, 204)
        ws.Range("C4").HorizontalAlignment = xlLeft
        ws.Range("B5").Value = "※ SharePointのファイルURLを貼り付けます。空欄の場合はファイル選択ダイアログが開きます。"
        ws.Range("B5").Font.Size = 9: ws.Range("B5").Font.Color = RGB(89, 89, 89)
        ws.Columns(1).ColumnWidth = 3
        ws.Columns(2).ColumnWidth = 22
        ws.Columns(3).ColumnWidth = 80
        ws.Activate
        ActiveWindow.DisplayGridlines = False
    End If
    ws.Buttons.Delete
    Set r = ws.Range("B7:B8")
    Set btn = ws.Buttons.Add(r.Left, r.Top, r.Width, r.Height)
    btn.Caption = "実績を取り込む"
    btn.OnAction = "ImportActuals"
    btn.Font.Size = 12
End Sub

' ボタンに割り当てるマクロ
Public Sub ImportActuals()
    Dim src As String
    src = Trim(CStr(ThisWorkbook.Worksheets(IMPORT_SHEET).Range("C4").Value))
    If Len(src) = 0 Then
        src = CStr(Application.GetOpenFilename("Excelファイル (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , "元ファイルを選択"))
        If src = "False" Then Exit Sub
    End If
    ' 末尾の「?web=1」などのクエリは除去(ファイル直接URLの場合のみ)
    If InStr(src, "?") > 0 Then
        Dim head As String
        head = Left(src, InStr(src, "?") - 1)
        If LCase(Right(head, 5)) = ".xlsx" Or LCase(Right(head, 5)) = ".xlsm" Then src = head
    End If
    ' 閲覧用ページのURL(Doc.aspx / _layouts)はExcelファイルとして開けない
    If InStr(src, "_layouts/") > 0 Or InStr(LCase(src), "sourcedoc=") > 0 Then
        MsgBox "貼り付けられたURLは閲覧用ページのリンクのため、Excelファイルとして開けません。" & vbLf & vbLf & _
               "【直接URLの取得方法】" & vbLf & _
               "  Excelでそのファイルを開き、「ファイル → 情報 → パスをコピー」で取得したURLを貼り付けてください。" & vbLf & _
               "  (https://～/～.xlsx の形になります)" & vbLf & vbLf & _
               "または、ファイルをダウンロードしてURL欄を空にし、ファイル選択で指定してください。", vbExclamation
        Exit Sub
    End If
    ShowResult ImportActualsCore(src, 0)
End Sub

' 結果を「実績取込」シートに書き出し、ダイアログには冒頭だけ表示する
' (MsgBoxは約1000文字で切れるため、全文はシートで確認する)
Private Sub ShowResult(ByVal res As String)
    MsgBox WriteResultSheet(res), vbInformation
End Sub

' 結果をシートに書き出し、ダイアログ用の短いメッセージを返す(テスト用に公開)
Public Function WriteResultSheet(ByVal res As String) As String
    Dim ws As Worksheet, lines As Variant, i As Long, shortMsg As String
    Set ws = ThisWorkbook.Worksheets(IMPORT_SHEET)
    ws.Range("B10:B600").ClearContents
    ws.Range("B10:B600").NumberFormat = "@"
    ws.Range("B10:B600").Font.Name = FONT_NAME
    ws.Range("B10").Value = "実行結果 (" & Format(Now, "yyyy/m/d hh:mm") & ")"
    ws.Range("B10").Font.Bold = True
    lines = Split(res, vbLf)
    For i = 0 To UBound(lines)
        If 11 + i > 600 Then Exit For
        ws.Cells(11 + i, 2).Value = lines(i)
    Next i
    If Len(res) > 500 Then
        shortMsg = Left(res, 500) & vbLf & "…" & vbLf & vbLf & "(全文は「" & IMPORT_SHEET & "」シートの10行目以降に出力しました)"
    Else
        shortMsg = res & vbLf & vbLf & "(詳細は「" & IMPORT_SHEET & "」シートの10行目以降)"
    End If
    WriteResultSheet = shortMsg
End Function

' 診断用: セルの文章に「期間」「工数計」が何回出るか、「期間」の直後の文字コードを返す
Public Function AnalyzeText(ByVal s As String) As String
    Dim nP As Long, nH As Long, pos As Long, i As Long, codes As String, seg As String
    nP = (Len(s) - Len(Replace(s, "期間", ""))) / 2
    nH = (Len(s) - Len(Replace(s, "工数計", ""))) / 3
    AnalyzeText = "「期間」" & nP & "回 / 「工数計」" & nH & "回 / 「期」" & _
                  ((Len(s) - Len(Replace(s, "期", ""))) / 1) & "回 / 「~」系記号あり=" & _
                  (InStr(StrConv(s, vbNarrow), "~") > 0 Or InStr(s, ChrW(65374)) > 0 Or InStr(s, ChrW(12316)) > 0)
    pos = InStr(s, "期間")
    If pos = 0 Then pos = InStr(s, "期")
    If pos > 0 Then
        seg = Mid(s, pos, 28)
        For i = 1 To Len(seg)
            codes = codes & Hex(AscW(Mid(seg, i, 1))) & " "
        Next i
        AnalyzeText = AnalyzeText & vbLf & "      「期間」付近(" & Flat(seg, 28) & ")の文字コード: " & codes
    End If
End Function

' セルの文章を1行に整形する(改行・タブを記号に置換)
Private Function Flat(ByVal s As String, ByVal n As Long) As String
    s = Replace(s, vbCrLf, "<改>")
    s = Replace(s, vbLf, "<改>")
    s = Replace(s, vbCr, "<改>")
    s = Replace(s, Chr(11), "<改>")
    s = Replace(s, vbTab, "<tab>")
    Flat = Left(s, n)
End Function

' dlgMode: 0=上書き時に確認ダイアログを出す / 1=確認せず実行(テスト用) / 2=確認で「いいえ」扱い(テスト用)
Public Function ImportActualsCore(ByVal src As String, ByVal dlgMode As Long) As String
    Dim wbS As Workbook
    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    On Error Resume Next
    Application.DisplayAlerts = False
    Set wbS = Workbooks.Open(Filename:=src, ReadOnly:=True, UpdateLinks:=0)
    Application.DisplayAlerts = prevAlerts
    On Error GoTo 0
    If wbS Is Nothing Then
        ImportActualsCore = "元ファイルを開けませんでした。" & vbLf & _
            "ファイルをダウンロードし、URL欄を空にして実行(ファイル選択)してください。"
        Exit Function
    End If

    On Error GoTo EH
    Application.ScreenUpdating = False
    ImportActualsCore = DoImport(wbS, dlgMode)
    wbS.Close SaveChanges:=False
    Application.ScreenUpdating = True
    ThisWorkbook.Activate
    Exit Function
EH:
    ImportActualsCore = "エラーが発生しました: " & Err.Description & " (Err " & Err.Number & ", 行" & Erl & ")"
    On Error Resume Next
    wbS.Close SaveChanges:=False
    Application.ScreenUpdating = True
End Function

' 元ファイルから「YYYY年度」シートを探す。全角数字・空白の違いは無視する
Private Function FindYearSheet(ByVal wbS As Workbook, ByVal fy As Long) As Object
    Dim sh As Object, n As String
    For Each sh In wbS.Sheets
        n = StrConv(NormName(sh.Name), vbNarrow)
        If n = fy & "年度" Then Set FindYearSheet = sh: Exit Function
    Next sh
    For Each sh In wbS.Sheets
        n = StrConv(NormName(sh.Name), vbNarrow)
        If InStr(n, CStr(fy)) > 0 And InStr(n, "年度") > 0 Then Set FindYearSheet = sh: Exit Function
    Next sh
End Function

Private Function SheetNameList(ByVal wb As Workbook) As String
    Dim sh As Object
    For Each sh In wb.Sheets
        SheetNameList = SheetNameList & "「" & sh.Name & "」"
    Next sh
End Function

Private Function NormName(ByVal s As String) As String
    s = Replace(s, " ", "")
    s = Replace(s, ChrW(12288), "")
    s = Replace(s, ChrW(160), "")
    s = Replace(s, vbLf, "")
    NormName = Trim(s)
End Function

Private Function WeekStarts(ByVal y As Long, ByVal m As Long) As Variant
    Dim s As Date, lastD As Date, n As Long, arr() As Long
    s = DateSerial(y, m, 1): lastD = DateSerial(y, m + 1, 0)
    Do While s <= lastD
        ReDim Preserve arr(0 To n)
        arr(n) = CLng(s)
        n = n + 1
        s = s + (8 - Weekday(s, vbMonday))
    Loop
    WeekStarts = arr
End Function

Private Function DoImport(ByVal wbS As Workbook, ByVal dlgMode As Long) As String
    Dim wsT As Worksheet, wsS As Object
    Dim wanted As Object, srcCache As Object, matched As Object, notFound As Object
    Dim pending As Collection, item As Variant
    Dim y As Long, m As Long, fy As Long, sn As String
    Dim b As Long, j As Long, r0 As Long, nm As String, key As String
    Dim starts As Variant, md As Object, data As Object
    Dim skipped As String, k As Variant, k2 As Variant
    Dim cur As Variant, overwrites As Long, written As Long, unmatched As Long
    Dim d As Long, msg As String

    Set wanted = CreateObject("Scripting.Dictionary")
    Set srcCache = CreateObject("Scripting.Dictionary")
    Set matched = CreateObject("Scripting.Dictionary")
    Set notFound = CreateObject("Scripting.Dictionary")
    Set pending = New Collection
    diagSrc = ""

    ' 1) 対象メンバー(テンプレート記載済み)を集める
    For Each wsT In ThisWorkbook.Worksheets
        If IsMonthSheet(wsT.Name) Then
            For b = 0 To MEMBER_BLOCKS - 1
                nm = CStr(wsT.Cells(FIRST_ROW + b * 3, 1).Value)
                If Len(NormName(nm)) > 0 Then wanted(NormName(nm)) = nm
            Next b
        End If
    Next wsT
    If wanted.Count = 0 Then DoImport = "テンプレートにメンバーが記載されていません。": Exit Function

    ' 2) 月シートごとに、対応する年度シートから値を引く
    For Each wsT In ThisWorkbook.Worksheets
        If IsMonthSheet(wsT.Name) Then
            y = 2000 + CLng(Left(wsT.Name, 2)): m = CLng(Mid(wsT.Name, 4, 2))
            fy = IIf(m >= 4, y, y - 1)
            sn = fy & "年度"
            If Not srcCache.Exists(sn) Then
                Set wsS = FindYearSheet(wbS, fy)
                If wsS Is Nothing Then
                    srcCache.Add sn, ""
                Else
                    srcCache.Add sn, LoadSource(wsS, wanted)
                End If
            End If
            If Not IsObject(srcCache(sn)) Then
                skipped = skipped & vbLf & "  ・" & wsT.Name & "(元ファイルに「" & sn & "」シートなし)"
                If InStr(skipped, "[元ファイルのシート]") = 0 Then
                    skipped = skipped & vbLf & "  [元ファイルのシート] " & wbS.Name & " : " & SheetNameList(wbS)
                End If
            Else
                Set data = srcCache(sn)
                starts = WeekStarts(y, m)
                For b = 0 To MEMBER_BLOCKS - 1
                    r0 = FIRST_ROW + b * 3
                    nm = CStr(wsT.Cells(r0, 1).Value)
                    key = NormName(nm)
                    If Len(key) > 0 Then
                        If Not data.Exists(key) Then
                            notFound(nm) = 1
                        Else
                            Set md = data(key)
                            For j = 0 To UBound(starts)
                                d = starts(j)
                                If md.Exists(d) Then
                                    pending.Add Array(wsT, r0 + 1, 3 + j, md(d))
                                    matched(sn & "|" & key & "|" & d) = 1
                                End If
                            Next j
                        End If
                    End If
                Next b
            End If
        End If
    Next wsT

    ' 3) 週に対応しなかった期間(対象月のシートがあるもの)を数える
    For Each k In srcCache.Keys
        If IsObject(srcCache(k)) Then
            For Each k2 In srcCache(k).Keys
                For Each item In srcCache(k)(k2).Keys
                    If SheetExists(Format(CDate(item), "yy") & "年" & Format(CDate(item), "mm") & "月") Then
                        If Not matched.Exists(k & "|" & k2 & "|" & item) Then unmatched = unmatched + 1
                    End If
                Next item
            Next k2
        End If
    Next k

    ' 4) 上書き確認
    For Each item In pending
        cur = item(0).Cells(item(1), item(2)).Value
        If Not IsEmpty(cur) Then
            If CStr(cur) <> "" Then
                If Abs(Val(cur) - item(3)) > 0.0001 Then overwrites = overwrites + 1
            End If
        End If
    Next item
    If overwrites > 0 Then
        Dim ans As VbMsgBoxResult
        Select Case dlgMode
            Case 0
                Application.ScreenUpdating = True
                ans = MsgBox("入力済みの実績 " & overwrites & " セルが元ファイルの値で上書きされます。" & vbLf & _
                             "実行してよろしいですか?", vbYesNo + vbExclamation, "上書き確認")
                Application.ScreenUpdating = False
            Case 1: ans = vbYes
            Case Else: ans = vbNo
        End Select
        If ans <> vbYes Then DoImport = "キャンセルしました(変更なし)。": Exit Function
    End If

    ' 5) 書き込み
    For Each item In pending
        cur = item(0).Cells(item(1), item(2)).Value
        If IsEmpty(cur) Or CStr(cur) = "" Or Abs(Val(cur) - item(3)) > 0.0001 Then
            item(0).Cells(item(1), item(2)).Value = item(3)
            written = written + 1
        End If
    Next item

    msg = "取込が完了しました。" & vbLf & "元ファイルの値と対応した週: " & pending.Count & " セル / 書き込み: " & _
          written & " セル(うち上書き " & overwrites & " セル、同じ値は書き込み不要)"
    If notFound.Count > 0 Then
        msg = msg & vbLf & "元ファイルに見つからないメンバー:"
        For Each k In notFound.Keys: msg = msg & vbLf & "  ・" & k: Next k
    End If
    If Len(skipped) > 0 Then msg = msg & vbLf & "対象外のシート:" & skipped
    If unmatched > 0 Then msg = msg & vbLf & "テンプレートの週に対応しなかった期間: " & unmatched & " 件"
    If Len(diagSrc) > 0 Then msg = msg & vbLf & vbLf & "[診断]" & diagSrc
    DoImport = msg
End Function

' 年度シートから、対象メンバーごとの「期間開始日 -> 工数計」を読む
Private Function LoadSource(ByVal wsS As Object, ByVal wanted As Object) As Object
    Dim res As Object, lastRow As Long, lastCol As Long, r As Long
    Dim key As String, v As Variant, vals As Variant, c As Long, md As Object
    Set res = CreateObject("Scripting.Dictionary")
    lastRow = wsS.Cells(wsS.Rows.Count, 2).End(xlUp).Row
    lastCol = wsS.Cells(3, wsS.Columns.Count).End(xlToLeft).Column
    diagSrc = diagSrc & vbLf & "「" & wsS.Name & "」 最終行=" & lastRow & " 最終列=" & lastCol & _
              "(3行目の日付ヘッダ基準)"
    If lastCol < 4 Then Set LoadSource = res: Exit Function
    Dim nStr As Long, rng As String, smp As String, maxLen As Long, longTxt As String
    For r = 5 To lastRow
        v = wsS.Cells(r, 2).Value
        If Not IsError(v) Then
            key = NormName(CStr(v))
            If wanted.Exists(key) And Not res.Exists(key) Then
                ' 1セルだけの範囲は配列にならないため、右隣を1列含めて必ず2次元配列で受ける
                vals = wsS.Range(wsS.Cells(r, 4), wsS.Cells(r, lastCol + 1)).Value
                Set md = CreateObject("Scripting.Dictionary")
                nStr = 0: maxLen = 0: smp = ""
                For c = 1 To UBound(vals, 2)      ' 左から右(古い提出→新しい提出)の順
                    If Not IsError(vals(1, c)) Then
                        If VarType(vals(1, c)) = vbString Then
                            nStr = nStr + 1
                            If Len(vals(1, c)) > maxLen Then
                                maxLen = Len(vals(1, c))
                                smp = "(列" & (c + 3) & ", " & Len(vals(1, c)) & "文字) " & Flat(CStr(vals(1, c)), 150)
                                longTxt = CStr(vals(1, c))
                            End If
                            ParseCell CStr(vals(1, c)), md
                        End If
                    End If
                Next c
                res.Add key, md
                rng = ""
                If md.Count > 0 Then
                    Dim kk As Variant, dMin As Long, dMax As Long
                    dMin = 2958465: dMax = 0
                    For Each kk In md.Keys
                        If kk < dMin Then dMin = kk
                        If kk > dMax Then dMax = kk
                    Next kk
                    rng = " 期間 " & Format(dMin, "yyyy/m/d") & "～" & Format(dMax, "yyyy/m/d")
                End If
                diagSrc = diagSrc & vbLf & "  ・" & CStr(v) & " (行" & r & "): 文字列セル " & nStr & " 個 / 読取った期間 " & _
                          md.Count & " 件" & rng
                If md.Count = 0 And Len(smp) > 0 Then
                    diagSrc = diagSrc & vbLf & "      最長セルの冒頭: " & smp
                    diagSrc = diagSrc & vbLf & "      最長セルの末尾: " & Flat(Right(longTxt, 120), 120)
                    diagSrc = diagSrc & vbLf & "      " & AnalyzeText(longTxt)
                End If
            End If
        End If
    Next r
    Set LoadSource = res
End Function

' セル内の「期間 ～ ... 工数計 xx.xxH」ブロックを読み取り md に反映する。
' 同じ期間は、0以外の値のうち一番右(最新)を採用。0しかなければ0。
Private Sub ParseCell(ByVal t As String, ByVal md As Object)
    Dim ms As Object, hm As Object, i As Long, st As Long, ln As Long
    Dim blk As String, d As Long, h As Double
    ' 全角数字・全角記号などを半角にそろえてから解析する
    t = StrConv(t, vbNarrow)
    If reP Is Nothing Then
        Set reP = CreateObject("VBScript.RegExp")
        reP.Global = True
        ' 区切り文字は「. / - 年月日」などの1文字の非数字を許容
        reP.Pattern = "期\s*間[^0-9]{0,20}(\d{4})[^0-9](\d{1,2})[^0-9](\d{1,2})[^0-9]{0,2}[^0-9]{1,10}(\d{4})[^0-9](\d{1,2})[^0-9](\d{1,2})"
        Set reH = CreateObject("VBScript.RegExp")
        reH.Global = False
        reH.IgnoreCase = True
        reH.Pattern = "工\s*数\s*計[^0-9]*([0-9]+\.?[0-9]*)"
    End If
    Set ms = reP.Execute(t)
    For i = 0 To ms.Count - 1
        st = ms(i).FirstIndex + 1
        If i < ms.Count - 1 Then ln = ms(i + 1).FirstIndex + 1 - st Else ln = Len(t) - st + 1
        blk = Mid(t, st, ln)
        d = CLng(DateSerial(CLng(ms(i).SubMatches(0)), CLng(ms(i).SubMatches(1)), CLng(ms(i).SubMatches(2))))
        Set hm = reH.Execute(blk)
        If hm.Count = 0 Then h = 0 Else h = Val(hm(0).SubMatches(0))
        If h <> 0 Then
            md(d) = h
        ElseIf Not md.Exists(d) Then
            md(d) = 0
        End If
    Next i
End Sub

Private Function IsMonthSheet(ByVal nm As String) As Boolean
    IsMonthSheet = (nm Like "##年##月")
End Function

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nm Then SheetExists = True: Exit Function
    Next ws
End Function

' 年月が最新の月シートを返す(無ければNothing)
Private Function LatestMonthSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsMonthSheet(ws.Name) Then
            If LatestMonthSheet Is Nothing Then
                Set LatestMonthSheet = ws
            ElseIf StrComp(ws.Name, LatestMonthSheet.Name, vbBinaryCompare) > 0 Then
                Set LatestMonthSheet = ws
            End If
        End If
    Next ws
End Function

Private Function ReadMembers(ByVal ws As Worksheet) As Variant
    Dim arr(1 To MEMBER_BLOCKS) As String, i As Long
    If ws Is Nothing Then ReadMembers = Empty: Exit Function
    For i = 1 To MEMBER_BLOCKS
        arr(i) = CStr(ws.Cells(FIRST_ROW + (i - 1) * 3, 1).Value)
    Next i
    ReadMembers = arr
End Function

Private Function ColL(ByVal c As Long) As String
    ColL = Split(Cells(1, c).Address(True, False), "$")(0)
End Function

Private Sub StyleCell(ByVal rng As Range, Optional ByVal bold As Boolean = False, _
                      Optional ByVal fontColor As Long = 0, Optional ByVal fillColor As Long = -1, _
                      Optional ByVal wrap As Boolean = False)
    With rng
        .Font.Name = FONT_NAME
        .Font.Size = 10
        .Font.Bold = bold
        .Font.Color = fontColor
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = wrap
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(153, 153, 153)
        If fillColor <> -1 Then .Interior.Color = fillColor
    End With
End Sub

' 実績が予定を下回るセルを太赤字にする条件付き書式を設定する(何度実行しても同じ結果)
Private Sub ApplyShortfallFormat(ByVal ws As Worksheet)
    Dim colFirst As Long, colTotal As Long, lastRow As Long, b As Long
    Dim area As Range
    colFirst = 3
    colTotal = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    lastRow = FIRST_ROW + MEMBER_BLOCKS * 3 - 1
    Set area = ws.Range(ws.Cells(FIRST_ROW, colFirst), ws.Cells(lastRow + 3, colTotal))
    area.FormatConditions.Delete

    ws.Activate
    For b = 0 To MEMBER_BLOCKS
        Dim rPlan As Long
        rPlan = FIRST_ROW + b * 3
        If b = MEMBER_BLOCKS Then rPlan = lastRow + 1   ' 全体合計(予定・実績・差異の3行)
        AddShortfallRule ws, rPlan, colFirst, colTotal
    Next b
End Sub

' rPlan=予定行、rPlan+1=実績行、rPlan+2=差異行。
' 条件式の相対参照は選択セルが基準になるため、範囲の先頭セルを選択してから追加する
Private Sub AddShortfallRule(ByVal ws As Worksheet, ByVal rPlan As Long, ByVal c1 As Long, ByVal c2 As Long)
    Dim cl As String, fc As Object
    cl = ColL(c1)
    ' 実績行: 実績・予定がともに数値で、実績 < 予定
    ws.Cells(rPlan + 1, c1).Select
    Set fc = ws.Range(ws.Cells(rPlan + 1, c1), ws.Cells(rPlan + 1, c2)).FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=AND(ISNUMBER(" & cl & (rPlan + 1) & "),ISNUMBER(" & cl & rPlan & ")," & _
        cl & (rPlan + 1) & "<" & cl & rPlan & ")")
    fc.Font.Bold = True
    fc.Font.Color = RGB(255, 0, 0)
    ' 差異行: 差異がマイナス(=実績が予定より小さい)
    ws.Cells(rPlan + 2, c1).Select
    Set fc = ws.Range(ws.Cells(rPlan + 2, c1), ws.Cells(rPlan + 2, c2)).FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=AND(ISNUMBER(" & cl & (rPlan + 2) & ")," & cl & (rPlan + 2) & "<0)")
    fc.Font.Bold = True
    fc.Font.Color = RGB(255, 0, 0)
End Sub

' 既存の全ての月シートに上記の書式を設定する(旧バージョンで作ったシート用)
Public Sub FormatAllMonthSheets()
    Dim ws As Worksheet, cur As Object
    Set cur = ActiveSheet
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        If IsMonthSheet(ws.Name) Then ApplyShortfallFormat ws
    Next ws
    If Not cur Is Nothing Then cur.Activate
    Application.ScreenUpdating = True
End Sub

' ===== 1日の定時工数 / 稼働日数(日本の平日-祝日) / 定時予定工数 =====

' 祝日データ(2025~2035年。「日付 名称;」の並び)。これ以降の年は「祝日」シートに追記する
Private Function HolidayData() As String
    Dim s As String
    s = s & "2025/1/1 元日;2025/1/13 成人の日;2025/2/11 建国記念の日;2025/2/23 天皇誕生日;2025/2/24 振替休日;2025/3/20 春分の日;"
    s = s & "2025/4/29 昭和の日;2025/5/3 憲法記念日;2025/5/4 みどりの日;2025/5/5 こどもの日;2025/5/6 振替休日;2025/7/21 海の日;"
    s = s & "2025/8/11 山の日;2025/9/15 敬老の日;2025/9/23 秋分の日;2025/10/13 スポーツの日;2025/11/3 文化の日;2025/11/23 勤労感謝の日;"
    s = s & "2025/11/24 振替休日;2026/1/1 元日;2026/1/12 成人の日;2026/2/11 建国記念の日;2026/2/23 天皇誕生日;2026/3/20 春分の日;"
    s = s & "2026/4/29 昭和の日;2026/5/3 憲法記念日;2026/5/4 みどりの日;2026/5/5 こどもの日;2026/5/6 振替休日;2026/7/20 海の日;"
    s = s & "2026/8/11 山の日;2026/9/21 敬老の日;2026/9/22 国民の休日;2026/9/23 秋分の日;2026/10/12 スポーツの日;2026/11/3 文化の日;"
    s = s & "2026/11/23 勤労感謝の日;2027/1/1 元日;2027/1/11 成人の日;2027/2/11 建国記念の日;2027/2/23 天皇誕生日;2027/3/21 春分の日;"
    s = s & "2027/3/22 振替休日;2027/4/29 昭和の日;2027/5/3 憲法記念日;2027/5/4 みどりの日;2027/5/5 こどもの日;2027/7/19 海の日;"
    s = s & "2027/8/11 山の日;2027/9/20 敬老の日;2027/9/23 秋分の日;2027/10/11 スポーツの日;2027/11/3 文化の日;2027/11/23 勤労感謝の日;"
    s = s & "2028/1/1 元日;2028/1/10 成人の日;2028/2/11 建国記念の日;2028/2/23 天皇誕生日;2028/3/20 春分の日;2028/4/29 昭和の日;"
    s = s & "2028/5/3 憲法記念日;2028/5/4 みどりの日;2028/5/5 こどもの日;2028/7/17 海の日;2028/8/11 山の日;2028/9/18 敬老の日;"
    s = s & "2028/9/22 秋分の日;2028/10/9 スポーツの日;2028/11/3 文化の日;2028/11/23 勤労感謝の日;2029/1/1 元日;2029/1/8 成人の日;"
    s = s & "2029/2/11 建国記念の日;2029/2/12 振替休日;2029/2/23 天皇誕生日;2029/3/20 春分の日;2029/4/29 昭和の日;2029/4/30 振替休日;"
    s = s & "2029/5/3 憲法記念日;2029/5/4 みどりの日;2029/5/5 こどもの日;2029/7/16 海の日;2029/8/11 山の日;2029/9/17 敬老の日;"
    s = s & "2029/9/23 秋分の日;2029/9/24 振替休日;2029/10/8 スポーツの日;2029/11/3 文化の日;2029/11/23 勤労感謝の日;2030/1/1 元日;"
    s = s & "2030/1/14 成人の日;2030/2/11 建国記念の日;2030/2/23 天皇誕生日;2030/3/20 春分の日;2030/4/29 昭和の日;2030/5/3 憲法記念日;"
    s = s & "2030/5/4 みどりの日;2030/5/5 こどもの日;2030/5/6 振替休日;2030/7/15 海の日;2030/8/11 山の日;2030/8/12 振替休日;"
    s = s & "2030/9/16 敬老の日;2030/9/23 秋分の日;2030/10/14 スポーツの日;2030/11/3 文化の日;2030/11/4 振替休日;2030/11/23 勤労感謝の日;"
    s = s & "2031/1/1 元日;2031/1/13 成人の日;2031/2/11 建国記念の日;2031/2/23 天皇誕生日;2031/2/24 振替休日;2031/3/21 春分の日;"
    s = s & "2031/4/29 昭和の日;2031/5/3 憲法記念日;2031/5/4 みどりの日;2031/5/5 こどもの日;2031/5/6 振替休日;2031/7/21 海の日;"
    s = s & "2031/8/11 山の日;2031/9/15 敬老の日;2031/9/23 秋分の日;2031/10/13 スポーツの日;2031/11/3 文化の日;2031/11/23 勤労感謝の日;"
    s = s & "2031/11/24 振替休日;2032/1/1 元日;2032/1/12 成人の日;2032/2/11 建国記念の日;2032/2/23 天皇誕生日;2032/3/20 春分の日;"
    s = s & "2032/4/29 昭和の日;2032/5/3 憲法記念日;2032/5/4 みどりの日;2032/5/5 こどもの日;2032/7/19 海の日;2032/8/11 山の日;"
    s = s & "2032/9/20 敬老の日;2032/9/21 国民の休日;2032/9/22 秋分の日;2032/10/11 スポーツの日;2032/11/3 文化の日;2032/11/23 勤労感謝の日;"
    s = s & "2033/1/1 元日;2033/1/10 成人の日;2033/2/11 建国記念の日;2033/2/23 天皇誕生日;2033/3/20 春分の日;2033/3/21 振替休日;"
    s = s & "2033/4/29 昭和の日;2033/5/3 憲法記念日;2033/5/4 みどりの日;2033/5/5 こどもの日;2033/7/18 海の日;2033/8/11 山の日;"
    s = s & "2033/9/19 敬老の日;2033/9/23 秋分の日;2033/10/10 スポーツの日;2033/11/3 文化の日;2033/11/23 勤労感謝の日;2034/1/1 元日;"
    s = s & "2034/1/2 振替休日;2034/1/9 成人の日;2034/2/11 建国記念の日;2034/2/23 天皇誕生日;2034/3/20 春分の日;2034/4/29 昭和の日;"
    s = s & "2034/5/3 憲法記念日;2034/5/4 みどりの日;2034/5/5 こどもの日;2034/7/17 海の日;2034/8/11 山の日;2034/9/18 敬老の日;"
    s = s & "2034/9/23 秋分の日;2034/10/9 スポーツの日;2034/11/3 文化の日;2034/11/23 勤労感謝の日;2035/1/1 元日;2035/1/8 成人の日;"
    s = s & "2035/2/11 建国記念の日;2035/2/12 振替休日;2035/2/23 天皇誕生日;2035/3/21 春分の日;2035/4/29 昭和の日;2035/4/30 振替休日;"
    s = s & "2035/5/3 憲法記念日;2035/5/4 みどりの日;2035/5/5 こどもの日;2035/7/16 海の日;2035/8/11 山の日;2035/9/17 敬老の日;"
    s = s & "2035/9/23 秋分の日;2035/9/24 振替休日;2035/10/8 スポーツの日;2035/11/3 文化の日;2035/11/23 勤労感謝の日;"
    HolidayData = s
End Function

' 祝日シートと範囲名「祝日リスト」を用意する(シートが既にあれば内容は変更しない)
Private Sub SetupHolidaySheet()
    Dim ws As Worksheet, cur As Object, items As Variant, p As Variant, d As Variant
    Dim i As Long, anchor As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HOLIDAY_SHEET)
    Set anchor = ThisWorkbook.Worksheets(IMPORT_SHEET)
    If anchor Is Nothing Then Set anchor = ThisWorkbook.Worksheets(CTRL_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set cur = ActiveSheet
        If anchor Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        Else
            Set ws = ThisWorkbook.Worksheets.Add(After:=anchor)
        End If
        ws.Name = HOLIDAY_SHEET
        ws.Cells.Font.Name = FONT_NAME
        ws.Range("A1").Value = "日付": ws.Range("B1").Value = "名称"
        StyleCell ws.Range("A1:B1"), True, RGB(255, 255, 255), RGB(31, 78, 120)
        items = Split(HolidayData(), ";")
        For i = 0 To UBound(items)
            If Len(items(i)) > 0 Then
                p = Split(items(i), " ")
                d = Split(p(0), "/")
                ws.Cells(i + 2, 1).Value = DateSerial(CLng(d(0)), CLng(d(1)), CLng(d(2)))
                ws.Cells(i + 2, 2).Value = p(1)
            End If
        Next i
        ws.Range("A2:A400").NumberFormat = "yyyy/m/d"
        ws.Range("D1").Value = "稼働日数の計算で除外する日付です。会社休業日(年末年始など)や次の年の祝日は末尾に追記してください(A2:A400が対象)。"
        ws.Range("D1").Font.Size = 9: ws.Range("D1").Font.Color = RGB(89, 89, 89)
        ws.Columns(1).ColumnWidth = 12
        ws.Columns(2).ColumnWidth = 18
        If Not cur Is Nothing Then cur.Activate
    End If
    ThisWorkbook.Names.Add Name:=HOLIDAY_NAME, RefersTo:="='" & HOLIDAY_SHEET & "'!$A$2:$A$400"
End Sub

' 月シートの3~5行目に、1日の定時工数の入力欄・週ごとの稼働日数・定時予定工数を作る
Private Sub BuildCapacity(ByVal ws As Worksheet, ByVal y As Long, ByVal m As Long, _
                          ByVal colFirst As Long, ByVal n As Long)
    Dim colTotal As Long, j As Long, r As Long, mon As String, cl As String
    Dim cCalc As Long
    cCalc = RGB(242, 242, 242)
    colTotal = colFirst + n

    ws.Range("A3").Value = DAILY_LABEL
    StyleCell ws.Range("A3"), True, 0, cCalc
    ws.Range("B3").Value = DEFAULT_DAILY
    StyleCell ws.Range("B3"), False, RGB(0, 0, 255), RGB(255, 242, 204)
    ws.Range("B3").NumberFormat = "0.0"
    With ws.Range("B3").Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="0", Formula2:="24"
        .ErrorTitle = "入力エラー"
        .ErrorMessage = "0~24の数値(h)を入力してください"
    End With
    ' 稼働日数の計算に使う月初日
    ws.Range("C3").Value = DateSerial(y, m, 1)
    ws.Range("C3").NumberFormat = """月初:""yyyy/m/d"
    ws.Range("C3").Font.Name = FONT_NAME
    ws.Range("C3").Font.Size = 9
    ws.Range("C3").Font.Color = RGB(153, 153, 153)
    ws.Range("C3").HorizontalAlignment = xlCenter

    For r = 4 To 5
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Merge
        ws.Cells(r, 1).Value = IIf(r = 4, "稼働日数(平日)", "定時予定工数(h)")
        StyleCell ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)), True, 0, cCalc
        StyleCell ws.Range(ws.Cells(r, colFirst), ws.Cells(r, colTotal)), False, 0, cCalc
        ws.Cells(r, colTotal).Font.Bold = True
    Next r
    For j = 0 To n - 1
        cl = ColL(colFirst + j)
        mon = "$C$3-WEEKDAY($C$3,3)+7*" & j
        ws.Cells(4, colFirst + j).Formula = "=NETWORKDAYS(MAX($C$3," & mon & "),MIN(EOMONTH($C$3,0)," & _
            mon & "+6)," & HOLIDAY_NAME & ")"
        ws.Cells(5, colFirst + j).Formula = "=" & cl & "4*$B$3"
    Next j
    ws.Cells(4, colTotal).Formula = "=SUM(" & ColL(colFirst) & "4:" & ColL(colTotal - 1) & "4)"
    ws.Cells(5, colTotal).Formula = "=SUM(" & ColL(colFirst) & "5:" & ColL(colTotal - 1) & "5)"
    ws.Range(ws.Cells(4, colFirst), ws.Cells(4, colTotal)).NumberFormat = "0"
    ws.Range(ws.Cells(5, colFirst), ws.Cells(5, colTotal)).NumberFormat = "0.0"
End Sub

' 定時工数欄のない既存の月シートに欄を追加する(追加済みのシートは何もしない)
Public Sub AddCapacityToAllMonthSheets()
    Dim ws As Worksheet, note As String, colTotal As Long
    SetupHolidaySheet
    For Each ws In ThisWorkbook.Worksheets
        If IsMonthSheet(ws.Name) Then
            If CStr(ws.Range("A3").Value) <> DAILY_LABEL Then
                note = CStr(ws.Range("A3").Value)
                ws.Range("A3").ClearContents
                ws.Range("E3").Value = note
                ws.Range("E3").Font.Name = FONT_NAME
                ws.Range("E3").Font.Size = 9
                ws.Range("E3").Font.Color = RGB(89, 89, 89)
                colTotal = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
                BuildCapacity ws, 2000 + CLng(Left(ws.Name, 2)), CLng(Mid(ws.Name, 4, 2)), 3, colTotal - 3
            End If
        End If
    Next ws
End Sub

Private Sub BuildSheet(ByVal y As Long, ByVal m As Long, ByVal nm As String, ByVal members As Variant)
    Dim ws As Worksheet, before As Worksheet, sh As Worksheet
    Dim firstD As Date, lastD As Date, s As Date, e As Date
    Dim n As Long, i As Long, j As Long, b As Long, k As Long
    Dim r0 As Long, r As Long, colFirst As Long, colTotal As Long, lastRow As Long
    Dim cl As String, rng As String, labels As Variant
    Dim hasMembers As Boolean
    Dim cBlue As Long, cInput As Long, cCalc As Long, cTotal As Long, cHead As Long

    cBlue = RGB(0, 0, 255): cInput = RGB(255, 242, 204): cCalc = RGB(242, 242, 242)
    cTotal = RGB(221, 235, 247): cHead = RGB(31, 78, 120)
    labels = Array("予定", "実績", "差異")
    hasMembers = Not IsEmpty(members)

    SetupHolidaySheet   ' 祝日シート(稼働日数の計算用)が無ければ作る
    ' 年月順になる位置にシートを挿入
    For Each sh In ThisWorkbook.Worksheets
        If IsMonthSheet(sh.Name) Then
            If before Is Nothing And StrComp(sh.Name, nm, vbBinaryCompare) > 0 Then Set before = sh
        End If
    Next sh
    If before Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    Else
        Set ws = ThisWorkbook.Worksheets.Add(Before:=before)
    End If
    ws.Name = nm

    ' 週(月曜始まり、月内に切り詰め)
    firstD = DateSerial(y, m, 1)
    lastD = DateSerial(y, m + 1, 0)
    n = 0
    s = firstD
    Do While s <= lastD
        n = n + 1
        s = s + (8 - Weekday(s, vbMonday))
    Loop
    colFirst = 3
    colTotal = colFirst + n
    lastRow = FIRST_ROW + MEMBER_BLOCKS * 3 - 1

    ' タイトル・凡例
    ws.Cells.Font.Name = FONT_NAME
    ws.Range("A1").Value = y & "年" & m & "月 工数予実管理表"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "単位: 人時(h)。黄色のセルに入力してください(青文字=入力値、黒文字=自動計算)。"
    If hasMembers Then
        ws.Range("E3").Value = "差異 = 実績 - 予定(プラスは超過)。メンバー名は直近のシートから引き継ぎました。"
    Else
        ws.Range("E3").Value = "差異 = 実績 - 予定(プラスは超過)。"
    End If
    ws.Range("A2").Font.Size = 9
    ws.Range("E3").Font.Size = 9
    ws.Range("A2").Font.Color = RGB(89, 89, 89)
    ws.Range("E3").Font.Color = RGB(89, 89, 89)

    ' ヘッダー
    ws.Cells(HEADER_ROW, 1).Value = "メンバー"
    ws.Cells(HEADER_ROW, 2).Value = "区分"
    s = firstD
    For i = 1 To n
        e = s + (7 - Weekday(s, vbMonday))
        If e > lastD Then e = lastD
        ws.Cells(HEADER_ROW, colFirst + i - 1).Value = "第" & i & "週" & vbLf & _
            Month(s) & "/" & Day(s) & "~" & Month(e) & "/" & Day(e)
        s = e + 1
    Next i
    ws.Cells(HEADER_ROW, colTotal).Value = "月合計"
    StyleCell ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(HEADER_ROW, colTotal)), True, RGB(255, 255, 255), cHead, True
    ws.Rows(HEADER_ROW).RowHeight = 32

    ' 定時工数欄(3~5行目)
    BuildCapacity ws, y, m, colFirst, n

    ' メンバーブロック
    For b = 0 To MEMBER_BLOCKS - 1
        r0 = FIRST_ROW + b * 3
        ws.Range(ws.Cells(r0, 1), ws.Cells(r0 + 2, 1)).Merge
        If hasMembers Then ws.Cells(r0, 1).Value = members(b + 1)
        StyleCell ws.Range(ws.Cells(r0, 1), ws.Cells(r0 + 2, 1)), False, cBlue, cInput, True
        For k = 0 To 2
            r = r0 + k
            ws.Cells(r, 2).Value = labels(k)
            StyleCell ws.Cells(r, 2), True, 0, IIf(k = 2, cCalc, -1)
            For j = 0 To n - 1
                cl = ColL(colFirst + j)
                If k = 2 Then
                    ws.Cells(r, colFirst + j).Formula = "=IF(AND(" & cl & r0 & "=""""," & cl & (r0 + 1) & _
                        "=""""),""""," & cl & (r0 + 1) & "-" & cl & r0 & ")"
                    StyleCell ws.Cells(r, colFirst + j), False, 0, cCalc
                Else
                    StyleCell ws.Cells(r, colFirst + j), False, cBlue, cInput
                End If
            Next j
            rng = ColL(colFirst) & r & ":" & ColL(colTotal - 1) & r
            ws.Cells(r, colTotal).Formula = "=IF(COUNT(" & rng & ")=0,"""",SUM(" & rng & "))"
            StyleCell ws.Cells(r, colTotal), True, 0, cCalc
        Next k
        ' 入力規則(0以上の数値)
        With ws.Range(ws.Cells(r0, colFirst), ws.Cells(r0 + 1, colTotal - 1)).Validation
            .Delete
            .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlGreaterEqual, Formula1:="0"
            .ErrorTitle = "入力エラー"
            .ErrorMessage = "0以上の数値(h)を入力してください"
        End With
    Next b

    ' 全体合計
    ws.Range(ws.Cells(lastRow + 1, 1), ws.Cells(lastRow + 3, 1)).Merge
    ws.Cells(lastRow + 1, 1).Value = "全体合計"
    StyleCell ws.Range(ws.Cells(lastRow + 1, 1), ws.Cells(lastRow + 3, 1)), True, 0, cTotal
    For k = 0 To 2
        r = lastRow + 1 + k
        ws.Cells(r, 2).Value = labels(k)
        StyleCell ws.Cells(r, 2), True, 0, cTotal
        For j = 0 To n
            cl = ColL(colFirst + j)
            If k < 2 Then
                ws.Cells(r, colFirst + j).Formula = "=SUMIF($B$" & FIRST_ROW & ":$B$" & lastRow & _
                    ",$B" & r & "," & cl & "$" & FIRST_ROW & ":" & cl & "$" & lastRow & ")"
            Else
                ws.Cells(r, colFirst + j).Formula = "=" & cl & (r - 1) & "-" & cl & (r - 2)
            End If
            StyleCell ws.Cells(r, colFirst + j), True, 0, cTotal
        Next j
    Next k

    ' 体裁
    ws.Range(ws.Cells(FIRST_ROW, colFirst), ws.Cells(lastRow + 3, colTotal)).NumberFormat = "0.0;-0.0;0.0"
    ApplyShortfallFormat ws
    ws.Columns(1).ColumnWidth = 14
    ws.Columns(2).ColumnWidth = 8
    ws.Range(ws.Columns(colFirst), ws.Columns(colTotal)).ColumnWidth = 12
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(FIRST_ROW, colFirst).Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
End Sub
