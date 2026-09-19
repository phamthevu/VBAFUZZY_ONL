Sub QFSFormat_Substring()
    Dim rng As Range
    Dim findRange As Range
    
    ' Làm vi?c trong do?n van b?n du?c ch?n
    Set rng = Selection.Range
   
    ' B2: Tìm "ej2p..."
    Set rng = Selection.Range
    With rng.Find
        .ClearFormatting
        .Text = "ej2p[0-9,\.]{1,}"
        .MatchWildcards = True
        .Wrap = wdFindStop
    End With
    
    Do While rng.Find.Execute
        Set findRange = rng.Duplicate
        ' B? ch? "e" d?u, ch? superscript t? "j"
        findRange.MoveStart wdCharacter, 1
        findRange.Font.Superscript = True
        rng.Collapse wdCollapseEnd
    Loop
    
End Sub
