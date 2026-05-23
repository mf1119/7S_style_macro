REM Macro converted from VBA to LibreOffice Basic

Sub A_SevenSeasElementNumberer()
    Dim oDoc As Object
    Dim oSelection As Object
    Dim oEnum As Object
    Dim oPara As Object
    Dim oCursor As Object
    Dim elementIndex As Integer
    Dim isSpeech As Boolean
    Dim pageNum As Integer
    Dim asterixTable As Object
    Dim shortTable As Object

    oDoc = ThisComponent
    oSelection = oDoc.getCurrentSelection()

    ' A selection in Writer is a collection of text ranges.
    ' We will operate on the first range in the selection.
    ' If there is no text selected (only a cursor), this will be a collapsed range
    ' and the enumeration will yield the single paragraph the cursor is in.
    Dim oRange As Object
    oRange = oSelection.getByIndex(0)
    oEnum = oRange.createEnumeration()

    elementIndex = 1
    isSpeech = False
    pageNum = 0

    Set asterixTable = CreateUnoService("com.sun.star.container.EnumerableMap")
    asterixTable.put("Tho", "Thought by")
    asterixTable.put("Nar", "Narration by")
    asterixTable.put("SNar", "Spoken Narration by")
    
    Set shortTable = CreateUnoService("com.sun.star.container.EnumerableMap")

    Do While oEnum.hasMoreElements()
        oPara = oEnum.nextElement()
        
        ' Skip non-paragraph elements
        If oPara.supportsService("com.sun.star.text.Paragraph") Then
            Dim sParaText As String
            sParaText = Trim(oPara.String)
            
            ' If line is blank, skip
            If sParaText = "" Then
                isSpeech = False
            ' If this line is speech, next line is not
            ElseIf isSpeech = True Then
                isSpeech = False
            ' If this line is escaped
            ElseIf Left(sParaText, 1) = "\" Or Left(sParaText, 1) = "[" Then
                isSpeech = False
            Else
                Dim aWords() As String
                aWords = Split(sParaText, " ")
                
                ' If first word is page, is a page marker
                If LCase(aWords(0)) = "page" Then
                    isSpeech = False
                    elementIndex = 1
                    If UBound(aWords) >= 1 Then
                        If IsNumeric(aWords(1)) Then
                            pageNum = CInt(aWords(1))
                        End If
                    End If
                ' If has number and period as second word, is panel marker
                ElseIf IsNumeric(aWords(0)) And UBound(aWords) >= 1 And aWords(1) = "." Then
                    ' x.y = Panel Numberer
                    If UBound(aWords) >= 3 And IsNumeric(aWords(2)) And aWords(3) = "." Then
                         isSpeech = True 'x.y.z is element
                    Else
                        If IsNumeric(aWords(0)) Then
                            pageNum = CInt(aWords(0))
                        End If
                        isSpeech = False
                    End If
                ' If starts with period and second word is number, then is an unnumbered panel marker
                ElseIf Left(sParaText, 1) = "." And UBound(aWords) >=1 And IsNumeric(aWords(1)) Then
                    oPara.String = pageNum & sParaText
                    isSpeech = False
                ' If starts with number and no period, is a numbered element marker
                ElseIf IsNumeric(aWords(0)) And (UBound(aWords) < 1 Or aWords(1) <> ".") Then
                    If IsNumeric(aWords(0)) Then
                        elementIndex = CInt(aWords(0)) + 1
                    End If
                    isSpeech = True
                ' If starts with word, then is an unnumbered element marker
                ElseIf Not IsNumeric(aWords(0)) And Left(sParaText, 1) <> "\" Then
                    oPara.String = elementIndex & " " & oPara.String
                    elementIndex = elementIndex + 1
                    isSpeech = True
                End If
            End If
            
            ' Handle '*' replacement
            If InStr(oPara.String, "*") > 0 Then
                Dim oSearch As Object
                oSearch = oPara.createSearchDescriptor()
                oSearch.SearchString = "*"
                oSearch.ReplaceString = ""
                oPara.replaceAll(oSearch)
                
                oSearch.SearchString = "  "
                oSearch.ReplaceString = " "
                oPara.replaceAll(oSearch)

                ReplaceShorthand(oPara, asterixTable, False)
            End If

            If isSpeech = True Then
                ReplaceShorthand(oPara, shortTable, True)
            End If
        End If
    Loop
End Sub

Sub ReplaceShorthand(oPara As Object, shortTable As Object, Optional writeMode)
    Dim bWriteMode As Boolean
    If IsMissing(writeMode) Then
        bWriteMode = True
    Else
        bWriteMode = writeMode
    End If

    Dim sParaText As String
    Dim aWords() As String
    Dim i As Integer
    Dim shortKey As String
    Dim val As String
    Dim firstCharIndex As Integer
    Dim lastCharIndex As Integer
    
    sParaText = oPara.String
    aWords = Split(sParaText, " ")

    For i = 0 To UBound(aWords)
        Dim sWord As String
        sWord = Trim(aWords(i))
        
        If shortTable.containsKey(sWord) Then
            Dim replacement As String
            replacement = shortTable.get(sWord)
            sParaText = Replace(sParaText, sWord, replacement)
            oPara.String = sParaText
        ElseIf bWriteMode And sWord = "{" Then
            firstCharIndex = InStr(sParaText, "{")
            lastCharIndex = InStr(firstCharIndex + 1, sParaText, "}")
            
            If firstCharIndex > 0 And lastCharIndex > 0 Then
                val = Mid(sParaText, firstCharIndex + 1, lastCharIndex - firstCharIndex - 1)
                
                ' Find the key before the "{"
                Dim tempSub As String
                tempSub = Left(sParaText, firstCharIndex -1)
                Dim tempWords() As String
                tempWords = Split(Trim(tempSub), " ")
                shortKey = tempWords(UBound(tempWords))

                If shortKey <> "" And val <> "" Then
                    shortTable.put(shortKey, val)
                End If
                ' The original VBA exits loop here, so we do too.
                Exit For 
            End If
        End If
    Next i
End Sub
