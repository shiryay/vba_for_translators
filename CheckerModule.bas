Attribute VB_Name = "CheckerModule"
Option Explicit

Public StopWordRules As Object
Dim ReportString As String

Private Function DocInUSEnglish() As Boolean
'
' Returns True if the ENTIRE document is in US English
'
    With ActiveDocument
        .LanguageDetected = False
        .DetectLanguage
        
        If .Range.LanguageID = wdEnglishUS Then
            DocInUSEnglish = True
        Else
            DocInUSEnglish = False
        End If
    End With

End Function

Private Function DocHasLetterSize() As Boolean
'
' Returns True if the paper size is Letter
'
    Dim PageParams
    Set PageParams = ActiveDocument.PageSetup
    DocHasLetterSize = (PageParams.PaperSize = 2)
End Function

Private Function EndTagPresent() As Boolean
'
' Returns true if there is -end of document- tag in the document
'
    Dim RegExp As Object
    
    Set RegExp = CreateObject("VBScript.RegExp")
    
    With RegExp
        .Pattern = "-end of document-\s*$"
        .Global = True
    End With
    
    EndTagPresent = RegExp.Test(ActiveDocument.Content)
End Function

Private Function Detected(Target) As Boolean
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = Target
        .replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Detected = Selection.Find.Found
End Function

Private Function PageBreakPresent() As Boolean
    PageBreakPresent = Detected("^m")
End Function

Private Function SectionBreakPresent() As Boolean
    SectionBreakPresent = Detected("^b")
End Function

Private Function WfTagsPresent() As Boolean
'
' Returns true if there is a wordfast tag in the document
'
    Dim RegExp As Object
    Dim OpenPattern, MidPattern, ClosePattern As String
    Dim OpenTagPresent, MidTagPresent, CloseTagPresent As Boolean
    
    Set RegExp = CreateObject("VBScript.RegExp")
    
    OpenPattern = "{0>"
    MidPattern = "<}\d+{>"
    ClosePattern = "<0}"
    
    With RegExp
        .Pattern = OpenPattern
        .Global = True
    End With
    
    OpenTagPresent = RegExp.Test(ActiveDocument.Content)
    
    With RegExp
        .Pattern = MidPattern
        .Global = True
    End With
    
    MidTagPresent = RegExp.Test(ActiveDocument.Content)
    
    With RegExp
        .Pattern = ClosePattern
        .Global = True
    End With
    
    CloseTagPresent = RegExp.Test(ActiveDocument.Content)
    
    If OpenTagPresent Or MidTagPresent Or CloseTagPresent Then
        WfTagsPresent = True
    Else
        WfTagsPresent = False
    End If
    
End Function

Private Function HighlightText(Target) As Boolean
    Dim retval As Boolean
    retval = False
    
    Application.ScreenUpdating = False
    Options.DefaultHighlightColorIndex = wdRed
    With ActiveDocument.Content.Find
      .ClearFormatting
      .Text = Target
      With .replacement
        .Text = "^&"
        .ClearFormatting
        .Highlight = True
      End With
      .Forward = True
      .Wrap = wdFindContinue
      .Format = True
      .MatchWildcards = True
      .Execute Replace:=wdReplaceAll
      If .Found Then
        retval = True
      End If
    End With
    Application.ScreenUpdating = True
    Log ("Highlighted " & Target)
    HighlightText = retval
End Function

Private Function BadDatesPresent() As Boolean
'
' Check if there are dates separated by anything else than slashes
'
    Dim RegExp As Object
    Dim regExp_Matches, Match As Object
    Dim BadDatesFound As Boolean
    
    Set RegExp = CreateObject("vbscript.regexp")
    BadDatesFound = False
    
    With RegExp
        .Pattern = "\b(\d|\d\d)([\./-])(\d|\d\d)([\./-])(\d\d\d\d|\d\d)\b"
        .Global = True
        If .Test(ActiveDocument.Content) Then
            Set regExp_Matches = .Execute(ActiveDocument.Content)
            For Each Match In regExp_Matches
                If (Match.SubMatches.Item(1) <> "/" Or Match.SubMatches.Item(3) <> "/") Or CInt(Match.SubMatches.Item(0)) > 12 Then
                    BadDatesFound = True
                    HighlightText (Match)
'                    Log ("Highlighted " & Match)
                End If
            Next
        End If
    End With

    BadDatesPresent = BadDatesFound
End Function

Private Function BritishSpellingPresent() As Boolean
'
' Highlight words with British spelling
'
    Dim RegExp As Object
    Dim regExp_Matches, Match As Object
    Dim BritishFound As Boolean
    Dim regexPatterns
    Dim ptn As Variant
    
    
    Set RegExp = CreateObject("vbscript.regexp")
    BritishFound = False
    
    regexPatterns = Array("\w{2,}(our|ours)", "\w*[^w](ise|ised|ising|yse|ysed|ysing)", _
                        "\w{2,}(ogue|ogues)", "fulfil\s", "fulfil(ed|ing|ment)", "\w+(geing)", "\w+ae\w+", _
                        "\w{2,}(tre|tres)\b", "storey(s*)", "offence\w*", "of the (one|other) part")
                        
    For Each ptn In regexPatterns
        With RegExp
            .Pattern = ptn
            .Global = True
            .IgnoreCase = True
            If .Test(ActiveDocument.Content) Then
                Set regExp_Matches = .Execute(ActiveDocument.Content)
                BritishFound = True
                For Each Match In regExp_Matches
                    HighlightText (Match)
'                    Log ("Highlighted " & Match)
                Next
            End If
        End With
    Next
    
    BritishSpellingPresent = BritishFound
End Function


Private Function StopWordsPresentWithRegex() As Boolean
'
' Finds words that are not to be used and highlights them in red using regular expressions
'
    Dim RegExp As Object
    Dim regExp_Matches, Match As Object
    Dim StopWord As Variant
    Dim StopWordsFound As Boolean
    Dim key As Variant
    Set RegExp = CreateObject("vbscript.regexp")
    
    ' Setting the found flag
    StopWordsFound = False
    
    ' Find and highlight stop words
    For Each key In StopWordRules.Keys
        With RegExp
            .Pattern = StopWordRules(key)(0)
            .Global = True
            .IgnoreCase = True
            .MultiLine = False
            If .Test(ActiveDocument.Content) Then
                Set regExp_Matches = .Execute(ActiveDocument.Content)
                StopWordsFound = True
                For Each Match In regExp_Matches
                    HighlightText (Match)
                Next
                Log ("> " & StopWordRules(key)(1))
            End If
        End With
    Next

    StopWordsPresentWithRegex = StopWordsFound
End Function


Public Sub RunAllTests()
    ' Check if the document is in US English
    If Not DocInUSEnglish() Then
        Log ("=====Set doc language to US English!=====")
    End If
    
    ' Check if document has letter size
    If Not DocHasLetterSize() Then
        Log ("=====Set paper size to US Letter!=====")
    End If
    
    ' Check for end of document tag
    If Not EndTagPresent() Then
        Log ("=====No end of document tag!=====")
    End If
    
    ' Check for end of document tag
    If PageBreakPresent() Or SectionBreakPresent() Then
        Log ("=====Remove page/section breaks!=====")
    End If
    
    ' Check for presence of wordfast tags
    If WfTagsPresent() Then
        Log ("=====Remove wordfast tags!=====")
    End If
    
    ' Check for poorly formated dates
    If BadDatesPresent() Then
        Log ("=====Please check the dates highlighted in red!=====")
    End If
    
    ' Check for stop words
    If StopWordsPresentWithRegex() Then
        Log ("=====Please check stop words highlighted in red!=====")
    End If
    
    ' Check for British spelling
    If BritishSpellingPresent() Then
        Log ("=====Please check British spelling highlighted in red!=====")
    End If
    
    ' Alert user of the report location
    MsgBox ("Please see CheckReport.docx in the active document folder")
    
End Sub

Public Sub DeleteReport()
'
' Clears the log file
'
    Dim FileNum As Integer
    Dim LogFileName As String
    Dim fso As FileSystemObject
    Dim answer As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")
    LogFileName = ActiveDocument.Path & "\" & "CheckReport.docx"
    
    If fso.FileExists(LogFileName) Then
        answer = MsgBox("Delete report? Are you sure?", vbOKCancel)
        If answer = vbOK Then
            fso.DeleteFile (LogFileName)
            MsgBox ("Report deleted!")
        End If
    End If
End Sub

Sub LoadRegexes()
'
' Populating the stop words regular expression dictionary
'
    Dim RegexFileName As String
    RegexFileName = Application.StartupPath & "\" & "regexes.csv"
    Dim DataLine As String
    Dim DataChunks As Variant
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim txtStream
    Set txtStream = fso.OpenTextFile(RegexFileName, ForReading, False)
    
    Set StopWordRules = CreateObject("Scripting.Dictionary")
    
    Do While Not txtStream.AtEndOfStream
        DataLine = txtStream.ReadLine
        DataChunks = Split(DataLine, ",")
        StopWordRules.Add DataChunks(0), Array(DataChunks(1), DataChunks(2))
    Loop
    txtStream.Close

End Sub

Sub Log(RepText)
    ReportString = ReportString & vbNewLine & RepText
End Sub

Sub SaveReport()
    Dim Report As Document
    Dim ReportName As String
    
    ReportName = ActiveDocument.Path & "\" & "CheckReport.docx"
    Set Report = Application.Documents.Add
    Report.Activate
    Selection.TypeText Text:=ReportString
    Report.SaveAs FileName:=ReportName
End Sub

Sub CheckMain()
'
' Main procedure
'
    ReportString = ""
    Log ("Testing " & ActiveDocument.FullName)
    Call LoadRegexes
    Call RunAllTests
    Call SaveReport

End Sub
