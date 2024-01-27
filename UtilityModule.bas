Attribute VB_Name = "UtilityModule"
Option Explicit

Dim ReportString As String

Sub KillHiddenText()
'
'   Removes hidden text from active document
'

    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Font.Hidden = True
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .Text = "^?"
        .replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub PasteAsIs()
'
'   Pastes retaining the original formatting
'
'
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
End Sub

Public Function convert_date(ByVal dt As String) As String
'
'   Converts a single date from EU to US and vice versa
'   Now the separator in the converted date is "/"
'
    
    Dim split_date() As String
    Dim sep, newSep As String
    
    newSep = "/" ' As per new requirements
    
    If InStr(dt, ".") Then sep = "."
    If InStr(dt, "/") Then sep = "/"
    If InStr(dt, "-") Then sep = "-"
    
    
    split_date = Split(dt, sep)
    convert_date = split_date(1) + newSep + split_date(0) + newSep + split_date(2)

End Function

Sub ConvertDates()
'
'   Converts all dates in active document from EU to US and vice versa
'   Now the separator in the converted date is "/"
'

    Dim dates() As String
    Dim replacement As String
    Dim RegExp As Object
    Dim regExp_Matches, Match As Object
    Dim S As String
    
    
    ReportString = ""
    Set RegExp = CreateObject("vbscript.regexp")
    
    With RegExp
        .Pattern = "\b(\d|\d\d)([\./-])(\d|\d\d)[\./-](\d\d\d\d|\d\d)\b"
        .Global = True
        If .Test(ActiveDocument.Content) Then
            Set regExp_Matches = .Execute(ActiveDocument.Content)
            Log ("Changes made:")
        End If
    End With
    
    For Each Match In regExp_Matches
        Selection.Find.ClearFormatting
        Selection.Find.replacement.ClearFormatting
        S = Match & " -> " & convert_date(Match) & Chr(13) & Chr(10)
        Log (S)
        With Selection.Find
            .Text = Match
            .replacement.Text = convert_date(Match)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next
    SaveReport ("ChangeReport.docx")

End Sub

Public Function URL_Encode(ByRef txt As String) As String
'
'   Utility function taken from https://excelvba.ru/code/URLEncode
'   No need to bind a key combo to
'
    Dim buffer As String, i As Long, c As Long, n As Long
    buffer = String$(Len(txt) * 12, "%")
    For i = 1 To Len(txt)
        c = AscW(Mid$(txt, i, 1)) And 65535
 
        Select Case c
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95  ' Unescaped 0-9A-Za-z-._ '
                n = n + 1
                Mid$(buffer, n) = ChrW(c)
            Case Is <= 127            ' Escaped UTF-8 1 bytes U+0000 to U+007F '
                n = n + 3
                Mid$(buffer, n - 1) = Right$(Hex$(256 + c), 2)
            Case Is <= 2047           ' Escaped UTF-8 2 bytes U+0080 to U+07FF '
                n = n + 6
                Mid$(buffer, n - 4) = Hex$(192 + (c \ 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
            Case 55296 To 57343       ' Escaped UTF-8 4 bytes U+010000 to U+10FFFF '
                i = i + 1
                c = 65536 + (c Mod 1024) * 1024 + (AscW(Mid$(txt, i, 1)) And 1023)
                n = n + 12
                Mid$(buffer, n - 10) = Hex$(240 + (c \ 262144))
                Mid$(buffer, n - 7) = Hex$(128 + ((c \ 4096) Mod 64))
                Mid$(buffer, n - 4) = Hex$(128 + ((c \ 64) Mod 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
            Case Else                 ' Escaped UTF-8 3 bytes U+0800 to U+FFFF '
                n = n + 9
                Mid$(buffer, n - 7) = Hex$(224 + (c \ 4096))
                Mid$(buffer, n - 4) = Hex$(128 + ((c \ 64) Mod 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
        End Select
    Next
    URL_Encode = Left$(buffer, n)
End Function

Sub CopyStartupPath()
'
'   Saves the path where word startup files are stored to the clipboard
'
    Dim DataObj As New MSForms.DataObject
    Dim S As String
    S = Application.StartupPath
    DataObj.SetText S
    DataObj.PutInClipboard
End Sub

Sub Log(RepText)
'
'   Logs a single entry
'
    ReportString = ReportString & vbNewLine & RepText
End Sub

Sub SaveReport(ReportTitle)
'
'   Saves the report with the ReportTitle name in the active document folder
'
    Dim Report As Document
    Dim ReportName As String
    
    ReportName = ActiveDocument.Path & "\" & ReportTitle
    Set Report = Application.Documents.Add
    Report.Activate
    Selection.TypeText Text:=ReportString
    Report.SaveAs FileName:=ReportName
    
End Sub
