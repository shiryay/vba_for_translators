Attribute VB_Name = "SearchModule"
'
' Bind key combos to these functions to trigger the relevant search
'
Sub Google()
    Search_flag = "Google"
    Search (Search_flag)
End Sub

Sub GoogleBooks()
    Search_flag = "GoogleBooks"
    Search (Search_flag)
End Sub

Sub LingueeDe()
    Search_flag = "LingueeDeEn"
    Search (Search_flag)
End Sub

Sub LingueeRu()
    Search_flag = "LingueeRuEn"
    Search (Search_flag)
End Sub

Sub LingueeEs()
    Search_flag = "LingueeEsEn"
    Search (Search_flag)
End Sub

Sub LingueeFr()
    Search_flag = "LingueeFrEn"
    Search (Search_flag)
End Sub

Sub GoogleTranslate()
    Search_flag = "GoogleTr"
    Search (Search_flag)
End Sub

Sub SearchProz()
    Search_flag = "Proz"
    Search (Search_flag)
End Sub

Sub SearchInsurinfo()
    Search_flag = "Insur"
    Search (Search_flag)
End Sub

Sub SearchColloc()
    Search_flag = "Colloc"
    Search (Search_flag)
End Sub

Sub SearchMultitran()
    Search_flag = "Multitran"
    Search (Search_flag)
End Sub

Sub Abkuerzungen()
    Search_flag = "Abkuerzungen"
    Search (Search_flag)
End Sub

Sub Acronymfinder()
    Search_flag = "Acronymfinder"
    Search (Search_flag)
End Sub

Sub Wox()
    Search_flag = "Wox"
    Search (Search_flag)
End Sub

Sub Sokr()
    Search_flag = "Sokr"
    Search (Search_flag)
End Sub

Private Function selected()
'
' Borrowed from the Internet, selects word under cursor if none is selected
'
   If WordBasic.GetSelStartPos() = WordBasic.GetSelEndPos() Then
        selected = 0
     Else
        selected = 1
     End If
End Function

Sub mul()
'
' Calls multitran
'
 Dim retval
 If selected = 0 Then
     WordBasic.SelectCurWord
 End If

 If selected = 1 Then
     WordBasic.EditCopy
 End If

retval = Shell("D:\mt\network\multitran.exe", 1) ' This path might require changing

End Sub

Public Function Search(ByVal flag As String)
'
' Search selection or open a dialog for entering the search query
'
    Dim urls
    Dim arg As String, url As String
    
    Set urls = CreateObject("Scripting.Dictionary")
    
    urls.Add "Google", "https://www.google.ru/search?q=%22{query}%22"
    urls.Add "GoogleBooks", "https://www.google.com/search?tbm=bks&q=%22{query}%22"
    urls.Add "GoogleTr", "https://translate.google.ru/?sl=auto&tl=en&text={query}&op=translate&hl=en"
    urls.Add "LingueeDeEn", "https://www.linguee.de/deutsch-englisch/search?source=auto&query=%22{query}%22"
    urls.Add "LingueeRuEn", "https://www.linguee.ru/russian-english/search?source=auto&query={query}"
    urls.Add "LingueeEsEn", "https://www.linguee.com/english-spanish/search?source=spanish&query={query}"
    urls.Add "LingueeFrEn", "https://www.linguee.fr/francais-anglais/search?source=auto&query={query}"
    ' urls.Add "Proz", "https://www.proz.com/search/?term={query}&from=rus&to=eng&es=1"
    urls.Add "Proz", "https://www.google.ru/search?q=%22{query}%22+english+proz"
    urls.Add "Insur", "https://www.insur-info.ru/dictionary/search/?q={query}&btnFind=%C8%F1%EA%E0%F2%FC%21&q_far"
    urls.Add "Colloc", "http://www.ozdic.com/collocation-dictionary/{query}"
    urls.Add "Multitran", "https://www.multitran.com/c/m.exe?CL=1&s={query}&l1=1&l2=2"
    urls.Add "Abkuerzungen", "http://abkuerzungen.de/result.php?searchterm={query}&language=de"
    urls.Add "Acronymfinder", "https://www.acronymfinder.com/{query}.html"
    urls.Add "Wox", "https://abkuerzungen.woxikon.de/abkuerzung/{query}.php"
    urls.Add "Sokr", "http://sokr.ru/{query}/"

    If selected = 1 Then
        arg = Replace(Selection.Text, vbNewLine, "", , , vbTextCompare) 'new line stripping
        arg = Replace(Selection.Text, "/", "%2F", , , vbTextCompare) 'replacing forward slash to make query address-bar-friendly
        arg = RTrim(arg)
    End If
    If selected = 0 Then
        arg = InputBox("Enter query")
    End If
    If arg = "" Then
        Exit Function
    End If
    
    url = Replace(urls(flag), "{query}", arg, , , vbTextCompare)
    url = UtilityModule.URL_Encode(url)
    
    On Error Resume Next
    ActiveDocument.FollowHyperlink Address:=url

End Function
