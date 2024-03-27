#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; A script by Zack Webster
; Modified and appended by shiryay

Search(engine) {
	urls := {"Google" : "https://www.google.ru/search?q=`%22{query}`%22", "GoogleBooks" : "https://www.google.com/search?tbm=bks&q=`%22{query}`%22", "GoogleTr" : "https://translate.google.ru/?sl=auto&tl=en&text={query}&op=translate&hl=en", "LingueeDeEn" : "https://www.linguee.de/deutsch-englisch/search?source=auto&query=`%22{query}`%22", "LingueeRuEn" : "https://www.linguee.ru/russian-english/search?source=auto&query={query}", "LingueeEsEn" : "https://www.linguee.com/english-spanish/search?source=spanish&query={query}", "LingueeFrEn" : "https://www.linguee.fr/francais-anglais/search?source=auto&query={query}", "Proz" : "https://www.google.ru/search?q=`%22{query}`%22+english+proz", "Insur" : "https://www.insur-info.ru/dictionary/search/?q={query}&btnFind=`%C8`%F1`%EA`%E0`%F2`%FC`%21&q_far", "MultitranWeb" : "https://www.multitran.com/c/m.exe?CL=1&s={query}&l1=1&l2=2", "MultitranLocal" : "d:\mt\network\multitran.exe", "Abkuerzungen" : "http://abkuerzungen.de/result.php?searchterm={query}&language=de", "Acronymfinder" : "https://www.acronymfinder.com/{query}.html", "Webster" : "https://www.merriam-webster.com/dictionary/{query}", "Wox" : "https://abkuerzungen.woxikon.de/abkuerzung/{query}.php", "Sokr" : "http://sokr.ru/{query}/", "Yandex" : "https://yandex.ru/search/?text=`%22{query}`%22"}
	url := urls[engine]
	send,^c
	sleep 150
	url := StrReplace(url, "{query}", clipboard)
	Run, %url%
}

!g::
	provider := "Google"
	Search(provider)
	return

!b::
	provider := "GoogleBooks"
	Search(provider)
	return

!t::
	provider := "GoogleTr"
	Search(provider)
	return

!d::
	provider := "LingueeDeEn"
	Search(provider)
	return

!r::
	provider := "LingueeRuEn"
	Search(provider)
	return

!s::
	provider := "LingueeEsEn"
	Search(provider)
	return

!f::
	provider := "LingueeFrEn"
	Search(provider)
	return

!p::
	provider := "Proz"
	Search(provider)
	return

!i::
	provider := "Insur"
	Search(provider)
	return

!m::
	provider := "MultitranWeb"
	Search(provider)
	return

!a::
	provider := "Acronymfinder"
	Search(provider)
	return

!k::
	provider := "Sokr"
	Search(provider)
	return

!u::
	provider := "Abkuerzungen"
	Search(provider)
	return

!w::
	provider := "Webster"
	Search(provider)
	return

!x::
	provider := "Wox"
	Search(provider)
	return

!y::
	provider := "Yandex"
	Search(provider)
	return

!z::
	provider := "MultitranLocal"
	Search(provider)
	return