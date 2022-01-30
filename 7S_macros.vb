'SevenSeasFormatter v0.0.1 - First version with a number
'Author: M Fulcrum, Masaaki Fukushima
'Please send all inquiries and complaints to laffypost@gmail.com or on Discord: mostly#9069
'I should get a git, but I can't be biffed, aye.
'I claim no responsibility for any losses incurred as a result of using this macro.
'You on your own, mate.


'Does only selection
Sub B_SevenSeasFormatter()

'The guts of the code
'Iterates through each "paragraph" and formats accordingly

'thoughts
'default is !thought
'If detects "Thought by" then next line must be thought - record isThought = true
'Check if this line isThought == true, then italics line, set isThought = false

isThought = False

For Each para In Selection.Paragraphs
    
    If isThought = True Then
        para.Range.Font.Italic = True
        isThought = False
    
    ElseIf para.Range.Words.Count > 1 Then
        If para.Range.Words(2) = "Narration" Or _
            para.Range.Words(2) = "Narration " Then
            isThought = True
        End If
    End If
    
    If para.Range.Words.Count > 3 Then
        If para.Range.Words(2) = "Thought " And _
            para.Range.Words(3) = "by " Then
            isThought = True
        End If
    End If
    
    'Checks if first character is numeric and therefore a panel or element marker
    'Bolds line
    If IsNumeric(para.Range.Words(1).Characters(1)) = True Then
        para.Range.Font.Bold = True
    End If
    
    'Checks if first word is "Page" and therefore a page marker
    'Bolds and underlines line
    If para.Range.Words.First = "Page " Then
        para.Range.Font.Bold = True
        para.Range.Font.Underline = True
    End If
    
    'Uses function to highlight notes
    HighlightBetween para.Range, "[Note:", "]"
    'Uses function to highlight symbols, e.g. <heart> <music note>
    HighlightBetween para.Range, "<", ">"
    
Next para

End Sub

'Current functions:
'Formates ellipses
'Deletes escape markers
'That's it
Sub C_SevenSeasFinalFormatter()

'The guts of the code
'Iterates through each "paragraph" and formats accordingly
For Each para In ActiveDocument.Paragraphs
    
    'Deletes the line escaper
    If para.Range.Words(1).Characters(1) = "\" Then
        para.Range.Characters(1).Delete
    End If
    
Next para

End Sub


'Function to search for particular markers and highlight the section
Function HighlightBetween(ByVal paragraph As Range, startMarker As String, endMarker As String)
    Dim Searching
    Searching = True
    Set sourceParagraph = paragraph
    
    Do While Searching = True
    
        'Setup all variables
        Dim startIndex 'The position at which the startMarker starts
        Dim endIndex 'The position where the endMarker ends

        Dim tempRange 'Storage for the finding method
        
        '-1 is assumed to be the default, and if not altered, then the marker was not found
        startIndex = -1
        endIndex = -1
    
        'Setting paragraph range and setting the starting index.
        Set tempRange = sourceParagraph
        With tempRange.Find
            .Forward = True
            .Wrap = wdFindStop
            .Text = startMarker
            .Execute
            If .Found = True Then
                startIndex = tempRange.Start
            Else
                Searching = False
                Exit Function
            End If
        End With
        
        'Finds the endMarker, but only sets if it is positioned after the startMarker
        'Sets beginning of the search field to startIndex if startMarker already found.
        Set tempRange = sourceParagraph
        If startIndex <> -1 Then
            tempRange.Start = startIndex
            'I have no idea how this bit works. Only recognizes symbols up to
            '200 characters long.
            tempRange.End = sourceParagraph.End + 200
        End If
        
        With tempRange.Find
            .Forward = True
            .Wrap = wdFindStop
            .Text = endMarker
            .Execute
            If .Found = True Then
                'Checks the end position is valid
                If tempRange.End > startIndex Then
                    endIndex = tempRange.End
                End If
            Else
                Seaching = False
                Exit Function
            End If
        End With
        
        'Rechecks that start and end are valid and actuall found
        'Highlights position if so
        If endIndex > startIndex And startIndex <> -1 And endIndex <> -1 Then
            tempRange.Start = startIndex
            tempRange.End = endIndex
            tempRange.HighlightColorIndex = wdYellow
        Else
            Searching = False
            Exit Function
        End If
        
        Set sourceParagraph = Selection.Range
        sourceParagraph.SetRange Start:=endIndex, End:=sourceParagraph.End
                
    Loop
    

End Function

Sub A_SevenSeasElementNumberer()

Dim elementIndex
Dim isSpeech
Dim shortTable As New Collection

'The guts of the code

elementIndex = 1
isSpeech = False
pageNum = 0

'Iterates through each "paragraph" and formats accordingly
For Each para In Selection.Paragraphs
    'For each line, do the following:
       
    'If speech - Then do nothing, next line is not speech
    'If !speech - If starts with "\", then is actually speech, next line is not speech
    'If !speech but Page marker - Nothing, Next line is !speech, Reset elementIndex
    'If !speech - If numbered and second word is ".", then is a panel marker,
    '   next line is not speech
    'If !speech - If numbered and second word is non ".", then is an element marker,
    '   next line is speech
    'If !speech - If not numbered and first character is not "\", then is
    '   element marker, next line is speech and requires numbering
    
    'If line is blank, skip, next line is not speech
    If Asc(para.Range.Words(1).Characters(1)) = 13 Then
        isSpeech = False
        
    'If this line is speech, next line is not
    ElseIf isSpeech = True Then
        isSpeech = False
    
    'If this line is escaped, and next line is not speech
    'Do NOT delete line escaper - This is done in the formatter
    ElseIf para.Range.Words(1).Characters(1) = "\" Then
        isSpeech = False
    
    'If first word is page, is a page marker
    'Reset elementIndex
    'Next line is not speech
    ElseIf para.Range.Words.First = "Page " Then
        isSpeech = False
        elementIndex = 1
        pageNum = CInt(para.Range.Words(2))
    
    'If has number and period as second word, is panel marker
    'Next line is not speech
    ElseIf IsNumeric(para.Range.Words(1).Characters(1)) = True And _
        para.Range.Words(2).Characters(1) = "." Then
        pageNum = CInt(para.Range.Words(1))
        isSpeech = False
        
    'If starts with period and second word is number, then is an
    'unnumbered panel marker. Do number.
    ElseIf para.Range.Words(1).Characters(1) = "." And _
        IsNumeric(para.Range.Words(2)) = True Then
        para.Range.InsertBefore Text:=Str(pageNum)
        para.Range.Characters(1).Delete
        isSpeech = False
                
    
    'If starts with number and no period, is a numbered element marker
    'Next line is speech
    'Increment elementIndex (just in case)
    ElseIf IsNumeric(para.Range.Words(1)) = True And _
        para.Range.Words(2).Characters(1) <> "." Then

        elementIndex = CInt(para.Range.Words(1)) + 1
        
        isSpeech = True
    
    'If starts with word, then is an unnumbered element marker
    'Insert element number
    'Delete space (it inserts a space at beginning of line for some reason)
    'Increment elementIndex
    'Next line is speech
    ElseIf IsNumeric(para.Range.Words(1)) = False And _
        para.Range.Words(1).Characters(1) <> "\" Then

        If para.Range.Words(1) = "Nar" And _
            para.Range.Words(2).Characters(1) = "*" Then
            para.Range.Words(1).Delete
            para.Range.Words(1).Delete
            para.Range.InsertBefore Text:="Narration by "
        End If
        
        If para.Range.Words(1) = "Tho" And _
            para.Range.Words(2).Characters(1) = "*" Then
            para.Range.Words(1).Delete
            para.Range.Words(1).Delete
            para.Range.InsertBefore Text:="Thought by "
        End If

        para.Range.InsertBefore Text:=Str(elementIndex) + " "
        para.Range.Characters(1).Delete
        elementIndex = elementIndex + 1
        isSpeech = True

    End If

    'This part is messed up, but if isSpeech is True at this point,
    'it means that we're still at an element marker, because
    'we haven't moved to the next line yet.
    If isSpeech = True Then
        ReplaceShorthand para.Range, shortTable
    End If
    
Next para

End Sub

Sub Z_SevenSeasElementDeNumberer()

'Iterates through each "paragraph" and formats accordingly
For Each para In Selection.Paragraphs

    'For each line, do the following:
       
    'If blank - Then do nothing
    'If numbered and second word is non ".", then is an element marker,
    '   next line is speech
    
    'If line is blank, skip, next line is not speech
    If Asc(para.Range.Words(1).Characters(1)) = 13 Then

    'If starts with number and no period, is a numbered element marker
    ElseIf IsNumeric(para.Range.Words(1)) = True And _
        para.Range.Words(2).Characters(1) <> "." Then
        para.Range.Words(1).Delete
        
    'If starts with number but has period, is a panel marker
    ElseIf IsNumeric(para.Range.Words(1)) = True And _
        para.Range.Words(2).Characters(1) = "." Then
        
        para.Range.Words(1).Delete
    End If
    
Next para

End Sub




'On first speech by speaker, use the firstChar and
'lastChar to mark the shorthand.
'Using curly braces.
'e.g., KIRI {Kirishima} -> Kirishima
Function ReplaceShorthand(paragraph As Range, shortTable As Collection)

Dim searchString As String
Dim firstChar As String
Dim lastChar As String

Dim shortKey As String
Dim val As String

Dim FirstCharIndex As Integer
Dim LastCharIndex As Integer

Dim iter As Integer
Dim iterEnd As Integer

'Static shortTable As Collection
Dim getValue As String
Dim compareKey As String

On Error Resume Next

'shortTable.Add Item:="value", key:="key"

firstChar = "{"
lastChar = "}"

For Each para In paragraph.Paragraphs

    iter = 1
    iterEnd = para.Range.Words.Count
    
    Do While iter < iterEnd
        getValue = ""
        compareKey = Trim(para.Range.Words(iter).Text)
        getValue = shortTable(compareKey)

        If getValue <> "" Then
            para.Range.Words(iter).Delete
            para.Range.Words(iter - 1).InsertAfter (shortTable(compareKey) + " ")
            
        ElseIf Trim(para.Range.Words(iter).Text) = firstChar Then
            searchString = para.Range.Text
            FirstCharIndex = 0
            LastCharIndex = 0
            
            FirstCharIndex = InStr(searchString, firstChar) + 1
            
            If FirstCharIndex > 0 Then
                LastCharIndex = InStr(FirstCharIndex, searchString, lastChar)
            End If
            
            If LastCharIndex > 0 Then
                val = Mid(searchString, FirstCharIndex, (LastCharIndex - FirstCharIndex))
                shortKey = Trim(para.Range.Words(iter - 1).Text)
                shortTable.Add Item:=val, key:=shortKey
                
                Do While Trim(para.Range.Words(iter)) <> lastChar
                    para.Range.Words(iter).Delete
                Loop
                para.Range.Words(iter).Delete
                                
                para.Range.Words(iter - 1).Delete
                para.Range.Words(iter - 2).InsertAfter (shortTable(shortKey) + " ")

                Exit Do
            End If
        End If
        
        iter = iter + 1
    Loop
    
Next para

End Function


