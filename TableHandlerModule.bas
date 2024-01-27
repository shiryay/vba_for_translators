Attribute VB_Name = "TableHandlerModule"
Option Explicit

Sub SmartcatTableToWfNotation()
    Dim rg As Range
    Dim tbl As Table
    
    Set rg = ActiveDocument.Range
    Set tbl = rg.Tables(1)
    tbl.Rows(1).Delete
    tbl.Columns(1).Delete
    tbl.Columns(3).Delete
    Selection.Rows.ConvertToText Separator:=wdSeparateByTabs, NestedTables:= _
        True
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "<0}^p{0>"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = "<}0{>"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="{0>"
    Selection.HomeKey Unit:=wdStory
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = "<}0{>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = "<}0{>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.EndKey Unit:=wdStory
    Selection.TypeText Text:="<0}"
End Sub

