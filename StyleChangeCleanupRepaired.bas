Attribute VB_Name = "StyleChangeCleanup"
' ***** BEGIN LICENSE BLOCK *****
'
'MIT License
'
'Copyright (c) 2017 A. Ilin-Tomich (Johannes Gutenberg University Mainz)
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.
'
' ***** END LICENSE BLOCK *****
Private Const ZoteroFieldIdentifier = " ADDIN ZOTERO_ITEM CSL_CITATION"
Private Const PunctuationPrecedingNoteReference = "[.,:;?!]"
'This macro places references before the punctuation marks and adds spaces before references
Sub CleanUpAfterChangingNotesToAuthorDate()
Dim uUndo As UndoRecord
Set uUndo = Application.UndoRecord
uUndo.StartCustomRecord ("Clean up citations after converting notes to author-date citations") 'Make the macro appear as a single operation on the Undo list
Dim fField As Field, rRange As Range, strPrevChar As String
Dim rPrevChar As Range
Dim rCurrentPosition As Range
Set rCurrentPosition = Selection.Range
For Each fField In ActiveDocument.Fields
    If Left(fField.Code, Len(ZoteroFieldIdentifier)) = ZoteroFieldIdentifier Then 'Check if this is a Zotero item field
        fField.Select
        Set rRange = Selection.Range
        If rRange.Start > ActiveDocument.Range.Start Then
            Set rPrevChar = ActiveDocument.Range(rRange.Start - 1, rRange.Start)
            If rPrevChar.Text Like PunctuationPrecedingNoteReference Then 'If the note reference was preceded by a punctuation sign, move it after the Author-Date citation
                Do While rPrevChar.Start > ActiveDocument.Range.Start
                    If ActiveDocument.Range(rPrevChar.Start - 1, rPrevChar.Start).Text Like PunctuationPrecedingNoteReference Then
                        rPrevChar.Start = rPrevChar.Start - 1 'if there is more than one punctuation sign preceding the reference, place them all after the reference
                    Else
                        Exit Do
                    End If
                Loop
                rRange.InsertAfter rPrevChar.Text
                rPrevChar.Delete
                If rRange.Start > ActiveDocument.Range.Start Then
                    Set rPrevChar = ActiveDocument.Range(rRange.Start - 1, rRange.Start)
                Else
                    GoTo Skip
                End If
            End If
            If Not (rPrevChar.Text = " " Or rPrevChar.Text = ChrW(160)) Then 'Insert a space before the Author-Date citation if there is no space or non-breaking space
                rRange.InsertBefore " "
            End If
        End If
    End If
Skip:
Next fField
rCurrentPosition.Select
uUndo.EndCustomRecord
End Sub

' ========================================================================
' Clean Up Citations After Converting Author-Date Style to Note Style
' ========================================================================
' Purpose: After changing Zotero citation style from author-date format 
'          (e.g., "Smith, 2020") to note format (e.g., "[1]"), this macro:
'          1. Removes spaces before citation numbers
'          2. Moves punctuation from after citation numbers to before them
'          
' Example: "some text [1]." becomes "some text.[1]"
'          "research [2], shows" becomes "research,[2] shows"
' ========================================================================

' Main subroutine: Process all Zotero citation fields in the document
Sub CleanUpCitationsAfterPunctuation()
    Dim uUndo As UndoRecord
    Set uUndo = Application.UndoRecord
    uUndo.StartCustomRecord ("Clean up citations after style change")
    
    Dim fField As Field
    Dim rCurrentPosition As Range
    Dim processedCount As Long
    Dim i As Long
    Dim totalFields As Long
    
    Set rCurrentPosition = Selection.Range
    Application.ScreenUpdating = False
    
    ' Store total field count to avoid accessing .Count property in loop
    totalFields = ActiveDocument.Fields.Count
    
    ' Process fields from end to beginning to avoid index shifting issues
    For i = totalFields To 1 Step -1
        ' Check if index is still valid (document modifications may change count)
        If i <= ActiveDocument.Fields.Count Then
            Set fField = ActiveDocument.Fields(i)
            
            ' Check if this is a Zotero citation field
            If InStr(fField.Code.Text, "ZOTERO_ITEM") > 0 Or _
               InStr(fField.Code.Text, "ZOTERO_CITATION") > 0 Then
                
                ' Select field for visual feedback
                fField.Select
                
                ' Process this field
                Call ProcessZoteroCitationField(fField)
                
                processedCount = processedCount + 1
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    rCurrentPosition.Select
    uUndo.EndCustomRecord
    
    MsgBox "Processing complete! " & processedCount & " citations processed.", vbInformation
End Sub

' Process a single Zotero citation field
' Parameters:
'   fField - The Zotero field object to process
Private Sub ProcessZoteroCitationField(fField As Field)
    On Error Resume Next
    
    Dim rBefore As Range, rAfter As Range
    Dim fieldStart As Long, fieldEnd As Long
    Dim punctText As String
    Dim docStart As Long, docEnd As Long
    Dim loopCounter As Integer
    Dim maxLoops As Integer
    
    maxLoops = 10  ' Maximum consecutive spaces to remove (prevents infinite loops)
    
    docStart = ActiveDocument.Content.Start
    docEnd = ActiveDocument.Content.End
    
    ' Get field position
    fField.Select
    fieldStart = Selection.Start
    fieldEnd = Selection.End
    
    ' ===== Step 1: Remove spaces before the citation =====
    loopCounter = 0
    Do While fieldStart > docStart And loopCounter < maxLoops
        ' Create a range containing the previous character
        Set rBefore = ActiveDocument.Range(fieldStart - 1, fieldStart)
        
        ' Check if it's a space (regular space or non-breaking space)
        If rBefore.Text = " " Or rBefore.Text = ChrW(160) Then
            ' Delete the space
            rBefore.Delete
            loopCounter = loopCounter + 1
            
            ' Re-acquire field position
            On Error Resume Next
            fField.Select
            If Err.Number <> 0 Then
                ' Field is no longer valid, exit safely
                Exit Sub
            End If
            On Error GoTo 0
            
            fieldStart = Selection.Start
            fieldEnd = Selection.End
        Else
            ' Not a space, exit loop
            Exit Do
        End If
    Loop
    
    ' ===== Step 2: Move punctuation from after citation to before =====
    ' Re-acquire field position
    fField.Select
    fieldStart = Selection.Start
    fieldEnd = Selection.End
    
    ' Check if there's punctuation after the citation
    If fieldEnd < docEnd Then
        Set rAfter = ActiveDocument.Range(fieldEnd, fieldEnd + 1)
        
        If Not rAfter Is Nothing Then
            If rAfter.Text Like PunctuationPrecedingNoteReference Then
                ' Collect all consecutive punctuation marks (max 5 to prevent anomalies)
                punctText = ""
                loopCounter = 0
                maxLoops = 5
                
                Do While rAfter.End <= docEnd And loopCounter < maxLoops
                    Dim currentChar As String
                    currentChar = ActiveDocument.Range(rAfter.Start + loopCounter, rAfter.Start + loopCounter + 1).Text
                    
                    If currentChar Like PunctuationPrecedingNoteReference Then
                        punctText = punctText & currentChar
                        loopCounter = loopCounter + 1
                    Else
                        Exit Do
                    End If
                Loop
                
                ' If punctuation was collected
                If Len(punctText) > 0 Then
                    ' Insert punctuation before the citation
                    Set rBefore = ActiveDocument.Range(fieldStart, fieldStart)
                    rBefore.InsertBefore punctText
                    
                    ' Delete original punctuation after the citation
                    fField.Select
                    fieldEnd = Selection.End
                    Set rAfter = ActiveDocument.Range(fieldEnd, fieldEnd + Len(punctText))
                    rAfter.Delete
                End If
            End If
        End If
    End If
    
    On Error GoTo 0
End Sub

