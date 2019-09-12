Sub CommandButton1_Click()
    
    Dim Filenames As Variant
    Filenames = Application.GetOpenFilename(MultiSelect:=True)
    
    If Not IsArray(Filenames) Then Exit Sub
    
    Dim Index As Integer
    For Index = LBound(Filenames) To UBound(Filenames)
        Dim Text As String
        
        Text = ""
        
        Open Filenames(Index) For Input As #1
        Do Until EOF(1)
            Dim Textline As String
           
            Line Input #1, Textline
            Text = Text & Textline
        Loop
        Close #1
       
        Dim Jsc As Integer, Voc As Integer, FF As Integer, PCE As Integer
        
        xrow = Range("E60000").End(xlUp).Offset(1).Row
       
        Jsc = InStr(Text, "Jsc")
        Voc = InStr(Text, "Voc")
        FF = InStr(Text, "Fill Factor")
        PCE = InStr(Text, "Power Conversion Efficiency")
        
        Range("$E$" & xrow).Value = Mid(Text, Jsc + 14, 11)
        Range("$F$" & xrow).Value = Mid(Text, Voc + 9, 8)
        Range("$G$" & xrow).Value = Mid(Text, FF + 13, 8)
        
    Next Index
End Sub
