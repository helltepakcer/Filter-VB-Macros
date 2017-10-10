Sub filterDomain()
' Macros will delete domains from list in "List1"
'
'
Dim domain As String
Dim i As Integer

        
        i = 0
        Do
        
        i = i + 1
        
        Sheets("Ëèñò1").Select
'NEXT DOMAIN
        domain = Worksheets("Ëèñò1").Cells(i, 1).Value
        
        Sheets("Ëèñò2").Select
'FINDER
        
        Do

        Set smvar = Cells.Find(What:=domain, After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False)

        If Not smvar Is Nothing Then smvar.Activate

        Selection.ClearContents
        
        Loop Until Not IsEmpty(smvar) Or Not Len(domain) <> 0
'END FINDER
        Loop Until Not Len(domain) <> 0

        MsgBox "Âûïîëíåíî"
        
End Sub
