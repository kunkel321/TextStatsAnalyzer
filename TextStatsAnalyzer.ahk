#Requires AutoHotkey v2.0
#SingleInstance Force
SetWorkingDir(A_ScriptDir)
TraySetIcon("imageres.dll", 305)

; ============================================================
;  TextStatsAnalyzer.ahk
;  Human: kunkel321
;  AI: Claude
;  Version Date: 4-19-2026 
;  GitHub: https://github.com/kunkel321/TextStatsAnalyzer
;  AutoHotkey: https://www.autohotkey.com/boards/viewtopic.php?f=83&t=140562
;  Analyzes .txt files for letter/word/ngram/sentence stats.
;  Drop a file onto the GUI or use Browse button.
;  Color theme loaded from colorThemeSettings.ini if present.
;  Tips: 
;  - Right-click a word in the Words tab. 
;  - Double-click some items in Summary for info.
;  - Not recommended for files larger than a few MB. 
; ============================================================

; ============================================================
;  USER CONFIGURATION  (edit these to customize behavior)
; ============================================================
global autoAnalyzeOnDrop   := true   ; true = analyze immediately when file is dropped
global autoAnalyzeOnBrowse := true   ; true = analyze immediately after Browse dialog
global defaultFilterStop   := true   ; true = stopword filter ON by default (Words tab)
global defaultMinFreq      := 1      ; default minimum word frequency (Words tab)
global defaultNgramSize    := 4      ; default n-gram size: 2=2-words, 3=3-words, etc.
global maxTableRows        := 5000   ; max rows exported to Tables tab edit box
                                     ; set to 0 for unlimited (warning: very large outputs
                                     ; may cause the edit box to become slow or unresponsive)

; ============================================================
;  STOPWORDS
; ============================================================
global stopwords := Map()
For w in StrSplit("a,an,the,and,or,but,in,on,at,to,for,of,with,by,from,is,it,its,"
    . "was,are,were,be,been,being,have,has,had,do,does,did,will,would,shall,should,"
    . "may,might,must,can,could,not,no,nor,so,yet,both,either,neither,each,few,more,"
    . "most,other,some,such,than,too,very,just,that,this,these,those,i,you,he,she,"
    . "we,they,me,him,her,us,them,my,your,his,our,their,what,which,who,whom,how,"
    . "when,where,why,if,as,up,out,about,into,through,after,before,between,then,"
    . "there,here,all,any,because,also,only,over,same,s,t,don,doesn,didn,won,isn", ",")
    stopwords[w] := 1

; ---------- Theme ----------
global useTheme := false
global fontColor := "", listColor := "", formColor := ""
themeFile := A_ScriptDir "\colorThemeSettings.ini"
If FileExist(themeFile) {
    fontColor := IniRead(themeFile, "ColorSettings", "fontColor", "31FFE7")
    listColor := IniRead(themeFile, "ColorSettings", "listColor", "003E67")
    formColor := IniRead(themeFile, "ColorSettings", "formColor", "00233A")
    useTheme  := true
    fontColor := "c" SubStr(fontColor, -5)
}

; ---------- Global state ----------
global currentFile := ""
global analyzedText := ""
global wordList    := []   ; all words lowercased, in order
global wordListRaw := []   ; all words original case, in order (for case-sensitive mode)
global sentenceList := []  ; array of sentence strings
global lvSortState := Map()
global summaryExplanations := Map()
global wordsMenu := Menu()

; ============================================================
;  BUILD GUI
; ============================================================
global g := Gui("+Resize +MinSize500x400", "Text Stats Analyzer")
g.SetFont("s10", "Segoe UI")
If useTheme
    g.BackColor := formColor

; --- Top bar ---
g.Add("Text", "x10 y12 w30", "File:")
global ctrlPath := g.Add("Edit", "x40 y9 w340 h24 ReadOnly vFilePath")
global ctrlBrowse  := g.Add("Button", "x388 y8 w70 h26", "Browse")
global ctrlAnalyze := g.Add("Button", "x466 y8 w80 h26 Default", "Analyze")
global ctrlClip    := g.Add("Button", "x554 y8 w90 h26", "Clipboard")
global ctrlClear   := g.Add("Button", "x652 y8 w60 h26", "Clear")

; --- Tab control ---
global tabs := g.Add("Tab3", "x5 y40 w700 h520 vMainTabs",
    ["Summary", "Letters", "Words", "N-Grams", "Sentences", "Concordance", "Charts", "Tables"])

ApplyTabTheme()

; ===== TAB 1 — Summary =====
tabs.UseTab(1)
global lvSummary := g.Add("ListView", 
    "x10 y70 w680 h480 -Multi NoSort vLvSummary", ["Statistic", "Value"])
lvSummary.ModifyCol(1, 320)
lvSummary.ModifyCol(2, 340)

; ===== TAB 2 — Letters =====
tabs.UseTab(2)
g.Add("Text", "x10 y70", "Case-insensitive letter frequencies:")
global lvLetters := g.Add("ListView",
    "x10 y90 w680 h460 -Multi vLvLetters", ["Letter", "Count", "Percent"])
lvLetters.ModifyCol(1, 80)
lvLetters.ModifyCol(2, 100)
lvLetters.ModifyCol(3, 100)

; ===== TAB 3 — Words =====
tabs.UseTab(3)
global ctrlStopwords := g.Add("CheckBox", "x10 y70 w160 h22 vFilterStop",
    "Filter stopwords")
ctrlStopwords.Value := defaultFilterStop
global ctrlCaseSensitive := g.Add("CheckBox", "x178 y70 w130 h22 vCaseSensitive",
    "Case-sensitive")
g.Add("Text", "x318 y72", "Min frequency:")
global ctrlMinFreq := g.Add("Edit", "x420 y69 w40 h22 vMinFreq", defaultMinFreq)
global ctrlMinFreqUp := g.Add("UpDown", "Range1-999", defaultMinFreq)
global ctrlWordGo := g.Add("Button", "x470 y68 w60 h24", "Refresh")
global lvWords := g.Add("ListView",
    "x10 y98 w680 h452 -Multi vLvWords", ["Word", "Count", "Percent"])
lvWords.ModifyCol(1, 220)
lvWords.ModifyCol(2, 100)
lvWords.ModifyCol(3, 100)

; ===== TAB 4 — N-Grams =====
tabs.UseTab(4)
g.Add("Text", "x10 y70", "Group size:")
global ctrlNgramSize := g.Add("DropDownList", "x88 y67 w80 vNgramSize Choose" (defaultNgramSize - 1),
    ["2 words", "3 words", "4 words", "5 words", "6 words"])
global ctrlNgramGo := g.Add("Button", "x178 y67 w70 h24", "Refresh")
global lvNgrams := g.Add("ListView",
    "x10 y98 w680 h452 -Multi vLvNgrams", ["Phrase", "Count", "Percent"])
lvNgrams.ModifyCol(1, 400)
lvNgrams.ModifyCol(2, 100)
lvNgrams.ModifyCol(3, 100)

; ===== TAB 5 — Sentences =====
tabs.UseTab(5)
g.Add("Text", "x10 y70", "All sentences detected in the text:")
global lvSentences := g.Add("ListView",
    "x10 y90 w680 h460 -Multi vLvSentences", ["#", "Sentence", "Words"])
lvSentences.ModifyCol(1, 40)
lvSentences.ModifyCol(2, 570)
lvSentences.ModifyCol(3, 60)

tabs.UseTab(0)

; ===== TAB 6 — Concordance =====
tabs.UseTab(6)
g.Add("Text", "x10 y72", "Search word:")
global ctrlConcordWord  := g.Add("Edit",    "x98 y69 w160 h24 vConcordWord")
global ctrlConcordCase  := g.Add("CheckBox","x262 y71 w130 h22 vConcordCase", "Case-sensitive")
global ctrlConcordGo    := g.Add("Button",  "x400 y69 w80 h24", "Search")
global ctrlConcordClear := g.Add("Button",  "x490 y69 w60 h24", "Clear")
global lvConcord := g.Add("ListView",
    "x10 y98 w680 h452 -Multi vLvConcord", ["#", "Sentence", "Sent #"])
lvConcord.ModifyCol(1, 40)
lvConcord.ModifyCol(2, 570)
lvConcord.ModifyCol(3, 55)

tabs.UseTab(0)

; ===== TAB 7 — Charts =====
tabs.UseTab(7)
global ctrlChartType := g.Add("DropDownList", "x10 y67 w220 vChartType Choose1",
    ["Word Length Distribution", "Sentence Length Distribution"])
global ctrlChartGo := g.Add("Button", "x240 y67 w100 h24", "Open Chart")
g.Add("Text", "x10 y110 w680",
    "Charts open in your default browser using Chart.js."
    . "`n`nWord Length Distribution — bar chart showing how many words of each character length appear in the text."
    . "`n`nSentence Length Distribution — bar chart showing how many sentences of each word-count length appear in the text."
    . "`n`nBoth charts give a visual sense of the text's complexity and rhythm."
    . "`n`nTip: Double-click or F1 on any Summary row with an explanation marker to learn more about that statistic.")

tabs.UseTab(0)

; ===== TAB 8 — Tables =====
tabs.UseTab(8)
global ctrlTableSource := g.Add("DropDownList", "x10 y67 w160 vTableSource Choose1",
    ["Summary", "Letters", "Words", "N-Grams", "Sentences"])
global ctrlTableFmt := g.Add("DropDownList", "x180 y67 w100 vTableFmt Choose1",
    ["Tab-delimited", "CSV", "Plain text"])
global ctrlTableGo   := g.Add("Button", "x290 y67 w80 h24", "Generate")
global ctrlTableCopy := g.Add("Button", "x380 y67 w80 h24", "Copy All")
global ctrlTableEdit := g.Add("Edit",
    "x10 y98 w680 h452 ReadOnly -Wrap HScroll VScroll vTableEdit")
ctrlTableEdit.SetFont("s9", "Courier New")

tabs.UseTab(0)

; --- Status bar ---
global sb := g.Add("StatusBar")
sb.SetText("  Drop a .txt file onto this window, or use Browse / Clipboard.")

; --- Apply theme to ListViews and controls ---
ApplyControlTheme()

; ============================================================
;  EVENTS
; ============================================================
g.OnEvent("DropFiles", OnDrop)
g.OnEvent("Size",      OnResize)
g.OnEvent("Close",     OnClose)
OnClose(*) => ExitApp()

ctrlBrowse.OnEvent("Click",  OnBrowse)
ctrlAnalyze.OnEvent("Click", OnAnalyze)
ctrlClip.OnEvent("Click",    OnClipboard)
ctrlClear.OnEvent("Click",   OnClear)
ctrlWordGo.OnEvent("Click",  OnWordRefresh)
ctrlNgramGo.OnEvent("Click", OnNgramRefresh)
ctrlChartGo.OnEvent("Click", OnDrawChart)
ctrlTableGo.OnEvent("Click",   OnGenerateTable)
ctrlTableCopy.OnEvent("Click", OnCopyTable)
ctrlConcordGo.OnEvent("Click",    OnConcordSearch)
ctrlConcordClear.OnEvent("Click", OnConcordClear)
ctrlConcordWord.OnEvent("Change", (*) => (ctrlConcordWord.Value = "" ? OnConcordClear() : ""))
lvWords.OnEvent("ContextMenu", OnWordsContextMenu)
OnWordRefresh(*)  => PopulateWords()
OnNgramRefresh(*) => PopulateNgrams()

; Column-sort events
lvLetters.OnEvent("ColClick",   SortLV)
lvWords.OnEvent("ColClick",     SortLV)
lvNgrams.OnEvent("ColClick",    SortLV)
lvSentences.OnEvent("ColClick", SortLV)

; Summary double-click and F1 explanation
lvSummary.OnEvent("DoubleClick", OnSummaryDblClick)

; ============================================================
;  SHOW
; ============================================================
InitExplanations()
wordsMenu.Add("Search in Concordance", OnWordsMenuConcord)
g.Show("w710 h580")
return

; Words LV right-click context menu handlers

OnWordsContextMenu(lvCtrl, rowNum, isRightClick, x, y) {
    If rowNum = 0
        return
    wordsMenu.Show(x, y)
}

OnWordsMenuConcord(*) {
    rowNum := lvWords.GetNext(0, "F")
    If rowNum = 0
        return
    word := lvWords.GetText(rowNum, 1)
    ; Switch to Concordance tab (tab 6) and run search
    tabs.Value := 6
    ctrlConcordWord.Value := word
    OnConcordSearch()
}

; F1 on Summary ListView shows explanation for focused row
#HotIf WinActive(g)
F1:: {
    If g.FocusedCtrl != lvSummary
        return
    rowNum := lvSummary.GetNext(0, "F")
    If rowNum
        ShowSummaryExplanation(lvSummary.GetText(rowNum, 1))
}
Enter:: {
    If g.FocusedCtrl = ctrlConcordWord
        OnConcordSearch()
}
#HotIf

; ============================================================
;  EVENT HANDLERS
; ============================================================

OnDrop(guiObj, guiCtrlObj, fileArr, x, y, *) {
    global currentFile
    f := fileArr[1]
    If !RegExMatch(f, "i)\.txt$") {
        MsgBox("Please drop a .txt file.", "Wrong file type", 48)
        return
    }
    currentFile := f
    ctrlPath.Value := f
    sb.SetText("  File loaded: " f)
    If autoAnalyzeOnDrop
        OnAnalyze()
}

OnBrowse(*) {
    global currentFile
    f := FileSelect(1,, "Select a text file", "Text Files (*.txt)")
    If f = ""
        return
    currentFile := f
    ctrlPath.Value := f
    sb.SetText("  File loaded: " f)
    If autoAnalyzeOnBrowse
        OnAnalyze()
}

OnClipboard(*) {
    global currentFile, analyzedText
    txt := A_Clipboard
    If Trim(txt) = "" {
        MsgBox("Clipboard is empty or contains no text.", "Clipboard", 48)
        return
    }
    currentFile := ""
    ctrlPath.Value := "(clipboard text)"
    analyzedText := txt
    RunAnalysis()
    sb.SetText("  Analyzed clipboard text  |  " wordList.Length " words")
}

OnClear(*) {
    global currentFile, analyzedText, wordList, wordListRaw, sentenceList, lvSortState
    currentFile  := ""
    analyzedText := ""
    wordList     := []
    wordListRaw  := []
    sentenceList := []
    lvSortState  := Map()   ; reset so next sort starts fresh
    ctrlPath.Value := ""
    ctrlConcordWord.Value := ""
    For lv in [lvSummary, lvLetters, lvWords, lvNgrams, lvSentences, lvConcord]
        lv.Delete()
    ctrlTableEdit.Value := ""
    sb.SetText("  Cleared.  Drop a .txt file onto this window, or use Browse / Clipboard.")
}

OnAnalyze(*) {
    global analyzedText
    If currentFile = "" && ctrlPath.Value != "(clipboard text)" {
        MsgBox("No file loaded.`nDrop a .txt file onto the window or use Browse.", 
            "Nothing to analyze", 48)
        return
    }
    If currentFile != "" {
        Try {
            ; Omit encoding arg so AHK auto-detects BOM (handles UTF-8, UTF-16 LE/BE, ANSI)
            analyzedText := FileRead(currentFile)
        } Catch {
            MsgBox("Could not read file:`n" currentFile, "Error", 16)
            return
        }
    }
    RunAnalysis()
    sb.SetText("  " (currentFile != "" ? currentFile : "(clipboard)")
        "  |  " wordList.Length " words  |  " sentenceList.Length " sentences")
}

; ============================================================
;  SUMMARY EXPLANATIONS
; ============================================================

InitExplanations() {
    global summaryExplanations
    summaryExplanations := Map(
    "Type-Token Ratio (vocabulary %)", 
        "Type-Token Ratio (TTR)`n`n"
        . "Unique words divided by total words, expressed as a percentage.`n`n"
        . "A higher TTR means more varied vocabulary — the author uses a wider "
        . "range of different words. A lower TTR means more repetition.`n`n"
        . "Example: A TTR of 72% means 72 out of every 100 words are unique. "
        . "Song lyrics and legal boilerplate tend to have low TTRs; literary "
        . "prose tends to have higher ones.",

    "Hapax legomena (appear once)",
        "Hapax Legomena`n`n"
        . "From Greek, meaning 'said only once.' These are words that appear "
        . "exactly one time in the entire text.`n`n"
        . "A high hapax count relative to total unique words suggests rich, "
        . "varied vocabulary. In linguistics, hapax legomena are used to study "
        . "authorship and estimate the size of an author's vocabulary.`n`n"
        . "They can also flag rare technical terms, proper nouns, or typos.",

    "Flesch-Kincaid Reading Ease",
        "Flesch-Kincaid Reading Ease`n`n"
        . "A readability formula developed by Rudolf Flesch in the 1940s and "
        . "revised with J. Peter Kincaid for the US Navy in the 1970s.`n`n"
        . "Formula: 206.835 - 1.015*(words/sentences) - 84.6*(syllables/words)`n`n"
        . "Score guide:`n"
        . "  90-100  Very Easy    (children's books)`n"
        . "  70-90   Easy         (popular fiction, news)`n"
        . "  60-70   Standard     (most adult reading)`n"
        . "  50-60   Fairly Difficult`n"
        . "  30-50   Difficult    (academic writing)`n"
        . "   0-30   Very Difficult (legal/scientific)`n`n"
        . "Long sentences and polysyllabic words both lower the score. "
        . "Note: syllable counting here is heuristic, so the score is an estimate.",

    "Average sentence length (words)",
        "Average Sentence Length`n`n"
        . "Total words divided by total sentences.`n`n"
        . "This is a key readability indicator. Short sentences (under 15 words) "
        . "are generally easier to follow. Academic and legal writing often "
        . "averages 25-40+ words per sentence.`n`n"
        . "Most style guides recommend aiming for 15-20 words on average for "
        . "general audiences.",

    "Average word length (chars)",
        "Average Word Length`n`n"
        . "Total characters across all words divided by total word count "
        . "(punctuation excluded).`n`n"
        . "English averages around 4-5 characters per word in everyday text. "
        . "Technical, legal, or scientific writing tends to push this higher "
        . "due to longer specialized vocabulary.",

    "Total syllables",
        "Syllable Count`n`n"
        . "An estimated count of all syllables across every word in the text, "
        . "calculated by counting vowel groups (a heuristic method).`n`n"
        . "Used internally for the Flesch-Kincaid calculation. The estimate "
        . "is generally accurate for common English words but may be off for "
        . "unusual words, proper nouns, or abbreviations."
    )
}

OnSummaryDblClick(lvCtrl, rowNum, *) {
    If rowNum = 0
        return
    ShowSummaryExplanation(lvCtrl.GetText(rowNum, 1))
}

ShowSummaryExplanation(statName) {
    If summaryExplanations.Has(statName)
        MsgBox(summaryExplanations[statName], statName, 64)
}

; ============================================================
;  CORE ANALYSIS
; ============================================================

RunAnalysis() {
    global wordList, wordListRaw, sentenceList, lvSortState
    lvSortState := Map()   ; reset so sort direction starts fresh for each new analysis
    txt := analyzedText

    ; --- Tokenize words (letters/apostrophes only) ---
    wordList    := []
    wordListRaw := []
    pos := 1
    While pos := RegExMatch(txt, "[a-zA-Z']+", &m, pos) {
        w := RegExReplace(m[], "^'+|'+$", "")  ; strip leading/trailing apostrophes
        If w != "" {
            wordListRaw.Push(w)              ; original case
            wordList.Push(StrLower(w))       ; lowercased
        }
        pos += m.Len
    }

    ; --- Tokenize sentences ---
    ; Strategy 1: punctuation-based (works for prose)
    sentenceList := []
    raw := RegExReplace(txt, "\s+", " ")
    pos := 1
    While pos := RegExMatch(raw, "[A-Za-z][^.!?]*[.!?]", &m, pos) {
        s := Trim(m[])
        If StrLen(s) > 3
            sentenceList.Push(s)
        pos += m.Len
    }
    ; Strategy 2: if punctuation-based found nothing (e.g. lyrics, lists),
    ; fall back to non-empty lines as "sentences"
    If sentenceList.Length = 0 {
        Loop Parse, txt, "`n", "`r" {
            s := Trim(A_LoopField)
            If StrLen(s) > 1
                sentenceList.Push(s)
        }
    }

    PopulateSummary()
    PopulateLetters()
    PopulateWords()
    PopulateNgrams()
    PopulateSentences()
}

; ============================================================
;  TAB POPULATION
; ============================================================

PopulateSummary() {
    lvSummary.Delete()
    txt := analyzedText

    ; Counts
    totalChars     := StrLen(txt)
    noSpaceChars   := StrLen(RegExReplace(txt, "\s", ""))
    totalWords     := wordList.Length
    totalSentences := sentenceList.Length
    totalLines     := 0
    Loop Parse, txt, "`n"
        totalLines++
    totalParagraphs := 0
    If Trim(txt) != "" {
        Loop Parse, txt, "`n"
            If Trim(A_LoopField) = "" && A_Index > 1
                totalParagraphs++
        totalParagraphs++   ; at least one
    }

    ; Unique words
    uniqMap := Map()
    For w in wordList
        uniqMap[w] := (uniqMap.Has(w) ? uniqMap[w] + 1 : 1)
    uniqueWords := uniqMap.Count

    ; TTR
    ttr := totalWords > 0 ? Round(uniqueWords / totalWords * 100, 1) : 0

    ; Avg word length
    totalWLen := 0
    longestWord := ""
    For w in wordList {
        totalWLen += StrLen(w)
        If StrLen(w) > StrLen(longestWord)
            longestWord := w
    }
    avgWordLen := totalWords > 0 ? Round(totalWLen / totalWords, 2) : 0

    ; Avg sentence length
    totalSLen := 0
    longestSentWords := 0
    For s in sentenceList {
        wc := CountWords(s)
        totalSLen += wc
        If wc > longestSentWords
            longestSentWords := wc
    }
    avgSentLen := totalSentences > 0 ? Round(totalSLen / totalSentences, 1) : 0

    ; Hapax legomena
    hapax := 0
    For w, c in uniqMap
        If c = 1
            hapax++

    ; Flesch-Kincaid Reading Ease
    ; FK = 206.835 - 1.015*(words/sentences) - 84.6*(syllables/words)
    totalSyllables := 0
    For w in wordList
        totalSyllables += CountSyllables(w)
    fk := 0
    If totalWords > 0 && totalSentences > 0 {
        fk := Round(206.835 
            - 1.015 * (totalWords / totalSentences)
            - 84.6  * (totalSyllables / totalWords), 1)
        fk := Round(Max(0, Min(100, fk)), 1)
    }
    fkLabel := fk >= 90 ? "Very Easy" 
             : fk >= 80 ? "Easy"
             : fk >= 70 ? "Fairly Easy"
             : fk >= 60 ? "Standard"
             : fk >= 50 ? "Fairly Difficult"
             : fk >= 30 ? "Difficult"
             : "Very Difficult"

    rows := [
        ["Total characters (with spaces)",    totalChars],
        ["Total characters (no spaces)",       noSpaceChars],
        ["Total words",                        totalWords],
        ["Total sentences",                    totalSentences],
        ["Total lines",                        totalLines],
        ["Total paragraphs (est.)",            totalParagraphs],
        ["Unique words",                       uniqueWords],
        ["Type-Token Ratio (vocabulary %)",    ttr "%"],
        ["Average word length (chars)",        avgWordLen],
        ["Average sentence length (words)",    avgSentLen],
        ["Longest word",                       longestWord],
        ["Longest sentence (words)",           longestSentWords],
        ["Total syllables",                    totalSyllables],
        ["Hapax legomena (appear once)",       hapax],
        ["Flesch-Kincaid Reading Ease",        fk " / 100  (" fkLabel ")"],
    ]
    For row in rows
        lvSummary.Add("", row[1], row[2])
}

PopulateLetters() {
    lvLetters.Delete()
    txt := StrLower(analyzedText)
    letterMap := Map()
    total := 0
    Loop Parse, txt {
        c := A_LoopField
        If Ord(c) >= 97 && Ord(c) <= 122 {  ; 97='a', 122='z'
            letterMap[c] := (letterMap.Has(c) ? letterMap[c] + 1 : 1)
            total++
        }
    }
    ; Sort a-z
    letters := []
    Loop 26
        letters.Push(Chr(96 + A_Index))
    For l in letters {
        If letterMap.Has(l) {
            cnt := letterMap[l]
            pct := Round(cnt / total * 100, 2)
            lvLetters.Add("", StrUpper(l), cnt, pct "%")
        }
    }
}

PopulateWords() {
    lvWords.Delete()
    filterStop  := ctrlStopwords.Value
    caseSens    := ctrlCaseSensitive.Value
    ; Clamp min-frequency to a valid integer (protects against typed-in garbage)
    val := Trim(ctrlMinFreq.Value)
    minFreq := IsInteger(val) ? Integer(val) : defaultMinFreq
    If minFreq < 1
        minFreq := 1
    If ctrlMinFreq.Value != minFreq
        ctrlMinFreq.Value := minFreq
    sourceList  := caseSens ? wordListRaw : wordList

    freqMap := Map()
    For w in sourceList {
        wLower := StrLower(w)
        If filterStop && stopwords.Has(wLower)
            continue
        key := caseSens ? w : wLower
        freqMap[key] := (freqMap.Has(key) ? freqMap[key] + 1 : 1)
    }
    total := 0
    For w, c in freqMap
        total += c

    ; Collect and sort by count desc
    rows := []
    For w, c in freqMap
        If c >= minFreq
            rows.Push([w, c])
    rows := SortArrayByCol(rows, 2, "desc")

    For row in rows {
        pct := total > 0 ? Round(row[2] / total * 100, 2) : 0
        lvWords.Add("", row[1], row[2], pct "%")
    }
    sb.SetText("  Words: " rows.Length " unique word"
        . (rows.Length = 1 ? "" : "s")
        . (filterStop ? " (stopwords filtered)" : "")
        . (minFreq > 1 ? ", min freq " minFreq : ""))
}

PopulateNgrams() {
    ; Disable button up-front so "Working..." actually covers the expensive work
    ctrlNgramGo.Enabled := false
    ctrlNgramGo.Text := "Working..."
    Try {
        lvNgrams.Delete()
        n := Integer(SubStr(ctrlNgramSize.Text, 1, 1))   ; "2 words" -> 2

        freqMap := Map()
        maxI := wordList.Length - n + 1
        Loop maxI {
            i := A_Index
            phrase := ""
            Loop n
                phrase .= (A_Index = 1 ? "" : " ") . wordList[i + A_Index - 1]
            freqMap[phrase] := (freqMap.Has(phrase) ? freqMap[phrase] + 1 : 1)
        }
        total := 0
        For p, c in freqMap
            total += c

        rows := []
        For p, c in freqMap
            rows.Push([p, c])
        rows := SortArrayByCol(rows, 2, "desc")

        For row in rows {
            pct := total > 0 ? Round(row[2] / total * 100, 3) : 0
            lvNgrams.Add("", row[1], row[2], pct "%")
        }
        sb.SetText("  N-Grams: " rows.Length " unique " n "-word phrase"
            . (rows.Length = 1 ? "" : "s"))
    }
    Finally {
        ; Always restore the button, even if something above threw
        ctrlNgramGo.Text := "Refresh"
        ctrlNgramGo.Enabled := true
    }
}

PopulateSentences() {
    lvSentences.Delete()
    i := 0
    For s in sentenceList {
        i++
        lvSentences.Add("", i, s, CountWords(s))
    }
}

; ============================================================
;  CHARTS
; ============================================================

OnDrawChart(*) {
    If wordList.Length = 0 {
        MsgBox("No text has been analyzed yet.", "Charts", 48)
        return
    }
    choice := ctrlChartType.Text
    If choice = "Word Length Distribution"
        DrawWordLengthChart()
    Else
        DrawSentenceLengthChart()
}

DrawWordLengthChart() {
    ; Build frequency map of word lengths
    lenMap := Map()
    For w in wordList {
        l := StrLen(w)
        lenMap[l] := (lenMap.Has(l) ? lenMap[l] + 1 : 1)
    }
    ; Find range
    minLen := 999, maxLen := 0
    For l, c in lenMap {
        If (l < minLen)
            minLen := l
        If (l > maxLen)
            maxLen := l
    }
    ; Build labels and data arrays for Chart.js
    labels := "", data := ""
    Loop (maxLen - minLen + 1) {
        l := minLen + A_Index - 1
        cnt := lenMap.Has(l) ? lenMap[l] : 0
        labels .= (A_Index = 1 ? "" : ",") . l
        data   .= (A_Index = 1 ? "" : ",") . cnt
    }
    title := "Word Length Distribution"
    xLabel := "Word length (characters)"
    yLabel := "Number of words"
    RenderBarChart(title, labels, data, xLabel, yLabel, "#4A90D9")
}

DrawSentenceLengthChart() {
    If sentenceList.Length = 0 {
        MsgBox("No sentences were detected in this text.", "Charts", 48)
        return
    }
    ; Build frequency map of sentence lengths (in words)
    lenMap := Map()
    For s in sentenceList {
        l := CountWords(s)
        lenMap[l] := (lenMap.Has(l) ? lenMap[l] + 1 : 1)
    }
    minLen := 999, maxLen := 0
    For l, c in lenMap {
        If (l < minLen)
            minLen := l
        If (l > maxLen)
            maxLen := l
    }
    labels := "", data := ""
    Loop (maxLen - minLen + 1) {
        l := minLen + A_Index - 1
        cnt := lenMap.Has(l) ? lenMap[l] : 0
        labels .= (A_Index = 1 ? "" : ",") . l
        data   .= (A_Index = 1 ? "" : ",") . cnt
    }
    title := "Sentence Length Distribution"
    xLabel := "Sentence length (words)"
    yLabel := "Number of sentences"
    RenderBarChart(title, labels, data, xLabel, yLabel, "#E8873A")
}

RenderBarChart(title, labels, data, xLabel, yLabel, color) {
    n := "`n"
    html := "<!DOCTYPE html>" . n
          . "<html><head><meta charset='utf-8'>" . n
          . "<script src='https://cdn.jsdelivr.net/npm/chart.js@4/dist/chart.umd.min.js'></script>" . n
          . "<style>" . n
          . "body{margin:0;padding:16px;background:#f4f4f4;font-family:Segoe UI,sans-serif;}" . n
          . "h2{text-align:center;font-size:16px;color:#333;margin-bottom:12px;}" . n
          . ".wrap{position:relative;width:100%;height:calc(100vh - 80px);}" . n
          . "</style></head><body>" . n
          . "<h2>" . title . "</h2>" . n
          . "<div class='wrap'><canvas id='c'></canvas></div>" . n
          . "<script>" . n
          . "new Chart(document.getElementById('c'),{" . n
          . "  type:'bar'," . n
          . "  data:{" . n
          . "    labels:[" . labels . "]," . n
          . "    datasets:[{" . n
          . "      label:'" . title . "'," . n
          . "      data:[" . data . "]," . n
          . "      backgroundColor:'" . color . "'," . n
          . "      borderColor:'" . color . "'," . n
          . "      borderWidth:1,borderRadius:3" . n
          . "    }]" . n
          . "  }," . n
          . "  options:{" . n
          . "    responsive:true,maintainAspectRatio:false," . n
          . "    plugins:{legend:{display:false}}," . n
          . "    scales:{" . n
          . "      x:{title:{display:true,text:'" . xLabel . "',font:{size:12}}}," . n
          . "      y:{title:{display:true,text:'" . yLabel . "',font:{size:12}}," . n
          . "         beginAtZero:true,ticks:{precision:0}}" . n
          . "    }" . n
          . "  }" . n
          . "});" . n
          . "</script></body></html>"

    ; Use a unique filename per chart so an already-open browser tab
    ; doesn't get silently overwritten when a second chart is generated.
    tmpFile := A_Temp "\TSA_chart_" A_TickCount ".html"
    Try {
        FileAppend(html, tmpFile, "UTF-8")
        Run(tmpFile)
    } Catch Error as e {
        MsgBox("Could not open chart: " e.Message, "Charts", 16)
    }
}

; ============================================================
;  TABLES
; ============================================================

OnGenerateTable(*) {
    If wordList.Length = 0 {
        MsgBox("No text has been analyzed yet.", "Tables", 48)
        return
    }
    source := ctrlTableSource.Text
    fmt    := ctrlTableFmt.Text
    result := ""
    If source = "Summary"
        result := BuildTableFromLV(lvSummary, ["Statistic", "Value"], fmt)
    Else If source = "Letters"
        result := BuildTableFromLV(lvLetters, ["Letter", "Count", "Percent"], fmt)
    Else If source = "Words"
        result := BuildTableFromLV(lvWords, ["Word", "Count", "Percent"], fmt)
    Else If source = "N-Grams"
        result := BuildTableFromLV(lvNgrams, ["Phrase", "Count", "Percent"], fmt)
    Else If source = "Sentences"
        result := BuildTableFromLV(lvSentences, ["#", "Sentence", "Words"], fmt)
    ctrlTableEdit.Value := result
    sb.SetText("  Table generated: " source " — " fmt)
}

OnCopyTable(*) {
    If ctrlTableEdit.Value = "" {
        MsgBox("Nothing to copy. Generate a table first.", "Tables", 48)
        return
    }
    A_Clipboard := ctrlTableEdit.Value
    sb.SetText("  Copied to clipboard.")
}

BuildTableFromLV(lv, headers, fmt) {
    maxRows  := (maxTableRows = 0) ? lv.GetCount() : maxTableRows
    rowCount := lv.GetCount()
    colCount := lv.GetCount("Col")
    sep  := (fmt = "CSV") ? "," : "`t"
    out  := ""

    ; Header row
    If fmt = "Plain text" {
        ; Calculate column widths for plain text padding
        colWidths := []
        Loop colCount
            colWidths.Push(StrLen(headers[A_Index]))
        Loop Min(rowCount, maxRows) {
            r := A_Index
            Loop colCount {
                cellLen := StrLen(lv.GetText(r, A_Index))
                If cellLen > colWidths[A_Index]
                    colWidths[A_Index] := cellLen
            }
        }
        ; Header
        hdr := ""
        Loop colCount {
            hdr .= PadRight(headers[A_Index], colWidths[A_Index] + 2)
        }
        out .= RTrim(hdr) . "`n"
        ; Separator line
        sep_line := ""
        Loop colCount
            sep_line .= RepeatChar("-", colWidths[A_Index] + 2)
        out .= RTrim(sep_line) . "`n"
        ; Rows
        Loop Min(rowCount, maxRows) {
            r := A_Index
            line := ""
            Loop colCount
                line .= PadRight(lv.GetText(r, A_Index), colWidths[A_Index] + 2)
            out .= RTrim(line) . "`n"
        }
    } Else {
        ; TSV or CSV header
        hdr := ""
        Loop colCount
            hdr .= (A_Index = 1 ? "" : sep) . FormatCell(headers[A_Index], fmt)
        out .= hdr . "`n"
        ; Rows
        Loop Min(rowCount, maxRows) {
            r := A_Index
            line := ""
            Loop colCount
                line .= (A_Index = 1 ? "" : sep) . FormatCell(lv.GetText(r, A_Index), fmt)
            out .= line . "`n"
        }
    }
    If rowCount > maxRows
        out .= "`n[Note: output capped at " maxRows " rows. " rowCount - maxRows " rows omitted.]"
    return out
}

FormatCell(val, fmt) {
    If fmt != "CSV"
        return val
    ; CSV: wrap in quotes if value contains comma, quote, or newline
    If InStr(val, ",") || InStr(val, '"') || InStr(val, "`n") {
        val := StrReplace(val, '"', '""')
        return '"' val '"'
    }
    return val
}

PadRight(str, width) {
    Loop (width - StrLen(str))
        str .= " "
    return str
}

RepeatChar(char, count) {
    result := ""
    Loop count
        result .= char
    return result
}

; ============================================================
;  CONCORDANCE
; ============================================================

OnConcordSearch(*) {
    word := Trim(ctrlConcordWord.Value)
    If word = "" {
        OnConcordClear()
        return
    }
    If sentenceList.Length = 0 {
        MsgBox("No text has been analyzed yet.", "Concordance", 48)
        return
    }
    caseSens := ctrlConcordCase.Value
    lvConcord.Delete()
    matchCount := 0
    Loop sentenceList.Length {
        sentNum := A_Index
        s := sentenceList[sentNum]
        found := caseSens
            ? InStr(s, word,, 1)
            : InStr(s, word, false, 1)
        If found {
            matchCount++
            lvConcord.Add("", matchCount, s, sentNum)
        }
    }
    ; Resize Sentence column to fill width
    lvConcord.GetPos(,, &lvW)
    lvConcord.ModifyCol(2, lvW - 105)
    sb.SetText("  Concordance: " matchCount " sentence"
        . (matchCount = 1 ? "" : "s") . " containing '" word "'")
}

OnConcordClear(*) {
    lvConcord.Delete()
    ctrlConcordWord.Value := ""
    sb.SetText("  Concordance cleared.")
}


CountWords(s) {
    cnt := 0
    pos := 1
    While pos := RegExMatch(s, "[a-zA-Z']+", &m, pos) {
        cnt++
        pos += m.Len
    }
    return cnt
}

CountSyllables(word) {
    ; Simple heuristic: count vowel groups
    word := StrLower(RegExReplace(word, "e$", ""))  ; strip trailing e
    cnt := 0
    pos := 1
    While pos := RegExMatch(word, "[aeiouy]+", &m, pos) {
        cnt++
        pos += m.Len
    }
    return Max(1, cnt)
}

SortArrayByCol(arr, col, dir := "asc") {
    ; Serialize rows to tab-delimited lines, sort via callback, deserialize.
    lines := ""
    isNumeric := true
    For row in arr {
        val := RegExReplace(row[col], "%", "")
        If !IsNumber(val)
            isNumeric := false
        payload := ""
        For v in row
            payload .= (payload = "" ? "" : "`t") . v
        lines .= payload . "`n"
    }
    lines := RTrim(lines, "`n")

    colIdx  := col
    isNum   := isNumeric
    reverse := (dir = "desc") ? -1 : 1   ; multiply result to flip direction

    SortCallback(a, b, *) {
        valA := StrSplit(a, "`t")[colIdx]
        valB := StrSplit(b, "`t")[colIdx]
        valA := RegExReplace(valA, "%", "")
        valB := RegExReplace(valB, "%", "")
        If isNum {
            nA := Number(valA), nB := Number(valB)
            return reverse * (nA > nB ? 1 : nA < nB ? -1 : 0)
        }
        return reverse * StrCompare(valA, valB)
    }

    lines := Sort(lines, "", SortCallback)   ; options string empty; R handled in callback

    result := []
    Loop Parse, lines, "`n" {
        If A_LoopField = ""
            continue
        result.Push(StrSplit(A_LoopField, "`t"))
    }
    return result
}

SortLV(lvCtrl, colNum) {
    ; Determine sort direction — toggle if same col clicked again
    key := lvCtrl.Hwnd . "_" . colNum
    dir := (lvSortState.Has(key) && lvSortState[key] = "desc") ? "asc" : "desc"
    lvSortState[key] := dir

    ; Collect all rows from the ListView
    rowCount := lvCtrl.GetCount()
    colCount := lvCtrl.GetCount("Col")
    rows := []
    Loop rowCount {
        r := A_Index
        row := []
        Loop colCount
            row.Push(lvCtrl.GetText(r, A_Index))
        rows.Push(row)
    }
    If rows.Length = 0
        return

    ; Sort and repopulate
    rows := SortArrayByCol(rows, colNum, dir)
    lvCtrl.Delete()
    For row in rows
        lvCtrl.Add("", row*)
}

; ============================================================
;  THEME APPLICATION
; ============================================================

ApplyTabTheme() {
    If !useTheme
        return
    ; Tab control background follows form color automatically in most themes
}

ApplyControlTheme() {
    If !useTheme
        return

    For lv in [lvSummary, lvLetters, lvWords, lvNgrams, lvSentences, lvConcord] {
        lv.Opt("Background" listColor)
        lv.SetFont(fontColor)
    }
    For ctrl in [ctrlPath, ctrlBrowse, ctrlAnalyze, ctrlClip, ctrlClear,
                 ctrlWordGo, ctrlNgramGo, ctrlStopwords, ctrlMinFreq,
                 ctrlCaseSensitive, ctrlTableGo, ctrlTableCopy,
                 ctrlTableSource, ctrlTableFmt, ctrlConcordGo,
                 ctrlConcordClear, ctrlConcordCase, ctrlConcordWord] {
        Try ctrl.SetFont(fontColor)
    }
    Try {
        ctrlTableEdit.Opt("Background" listColor)
        ctrlTableEdit.SetFont(fontColor)
    }
    ; Status bar color not easily themed in AHK — leave as default
}

; ============================================================
;  RESIZE
; ============================================================

OnResize(guiObj, minMax, width, height, *) {
    If minMax = -1   ; minimized
        return
    tabW := width  - 10
    tabH := height - 65
    tabs.Move(,, tabW, tabH)
    lvW := tabW - 25
    lvH := tabH - 55

    ; Suppress redraws during resize to prevent button distortion
    DllCall("SendMessage", "Ptr", g.Hwnd, "UInt", 0x000B, "Ptr", 0, "Ptr", 0)  ; WM_SETREDRAW false

    For lv in [lvSummary, lvLetters, lvWords, lvNgrams, lvSentences, lvConcord]
        lv.Move(,, lvW, lvH)

    ctrlTableEdit.Move(,, lvW, lvH)

    ctrlPath.Move(,, width - 392)
    ctrlBrowse.Move(width  - 334)
    ctrlAnalyze.Move(width - 256)
    ctrlClip.Move(width    - 168)
    ctrlClear.Move(width   -  70)

    ; Re-enable redraws and force repaint
    DllCall("SendMessage", "Ptr", g.Hwnd, "UInt", 0x000B, "Ptr", 1, "Ptr", 0)  ; WM_SETREDRAW true
    WinRedraw(g.Hwnd)

    ; Debounce column resizes — only run 150ms after dragging stops
    SetTimer(ResizeFlexCols, -150)
}

ResizeFlexCols() {
    lvSentences.GetPos(,, &lvW)
    lvSentences.ModifyCol(2, lvW - 125)
    lvConcord.GetPos(,, &lvW)
    lvConcord.ModifyCol(2, lvW - 105)
}
