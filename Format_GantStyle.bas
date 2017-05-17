Attribute VB_Name = "Format_GantStyle"
Sub FormatStyleBar()
'Version 0.0.1
'Initial Release of the macro
'Author : Florent Tainturier florent.tainturier@gmail.com
'
'This Macro permit to format the text (right text) for each Bar of Gant Diagram.
'It is based on the method GantBarEditEx and a dummy for/next loop instruction.
'The "RightText" parameters is arbitrary set to "Nom" (in French).
'Replace this parameter if needed.
'Since, the "Item" parameter is waiting a string, the Cstr() function is used to converse the n into string.

    Dim n As Integer
    n = 0
    
    For n = 1 To 50
        GanttBarEditEx Item:=CStr(n), RightText:="Nom"
        'Cstr is used to convert n integer into string type
        'Replace "Nom" by the text you want to place
    Next n

End Sub

