Sub FormatForReport()

'Select row of variant to be transferred
Dim rng As Range
Set rng = Selection

'Set active workbook variables
Dim wb As Workbook
Dim ws As Worksheet
Set wb = ActiveWorkbook
Set ws = ActiveSheet

'Loop through columns in row 2. If column header matches the field we want, extract its corresponding value from selected row. Column 130 set as upper limit due to duplicate "Gene" header in later columns.
For i = 1 To 130
    If ws.Cells(2, i) = "Chr" Then
        Dim chrom As String
        chrom = ws.Cells(rng.Row, i).Text
    ElseIf ws.Cells(2, i) = "Start" Then
        Dim start As String
        start = ws.Cells(rng.Row, i).Text
    ElseIf ws.Cells(2, i) = "Ref" Then
        Dim ref As String
        ref = ws.Cells(rng.Row, i).Text
    ElseIf ws.Cells(2, i) = "Alt" Then
        Dim alt As String
        alt = ws.Cells(rng.Row, i).Text
    ElseIf ws.Cells(2, i) = "Gene" Then
        Dim gene As String
        gene = ws.Cells(rng.Row, i).Text
    ElseIf ws.Cells(2, i) = "DNA_Change|(SampleData|Source|SourceVer|Datetime|IsValid|Gene|GeneIDs|Transcript|TranscriptIDs|DNA Change|Type)" Then
        transcript = ws.Cells(rng.Row, i).Text
        transcript_arr = Split(transcript, "|")
        transcript = transcript_arr(7)
    ElseIf ws.Cells(2, i) = "Zygosity" Then
        Dim zygosity As String
        zygosity = ws.Cells(rng.Row, i).Text
        'Format zygosity
        If zygosity = "het" Then
        zygosity = "Het"
        ElseIf zygosity = "hom" Then
        zygosity = "Hom"
        ElseIf zygosity = "hem" Or zygosity = "hemi" Then
        zygosity = "Hem"
        End If
    ElseIf ws.Cells(2, i) = "DNA Change" Then
        Dim nuc_change As String
        nuc_change = ws.Cells(rng.Row, i).Text
    ElseIf ws.Cells(2, i) = "AA Change" Then
        Dim prot_change As String
        prot_change = ws.Cells(rng.Row, i).Text
        'If no protein change (intronic variant), write "N/A"
        If prot_change = vbNullString Then: prot_change = "N/A"
    ElseIf ws.Cells(2, i) = "Interpretation" Then
        Dim variant_interp As String
        variant_interp = ws.Cells(rng.Row, i).Text
    End If
Next

'Extract parental variant status using same loop as above
For i = 1 To 130
    If ws.Cells(2, i) = "in Mother" Then
        Dim mom As String
        mom = ws.Cells(rng.Row, i).Text
        'If column exists and variant not inherited from mom, then mom = "N". If column doesn't exist, mom = vbNullString
        If mom = vbNullString Then
            mom = "N"
        End If
    ElseIf ws.Cells(2, i) = "in Father" Then
        Dim dad As String
        dad = ws.Cells(rng.Row, i).Text
        'If column exists and variant not inherited from dad, then dad = "N". If column doesn't exist, dad = vbNullString
        If dad = vbNullString Then
            dad = "N"
        End If
    End If
Next

'Inheritance logic
If zygosity = "Hom" And mom = "Y" And dad = "Y" Then
    inheritance = "Mat/Pat"
ElseIf (mom = "Y" And dad = "Y") Or (mom = "N" And dad = vbNullString) Or (mom = vbNullString And dad = "N") Or (mom = vbNullString And dad = vbNullString) Then
    inheritance = "Unk"
ElseIf mom = "Y" Then
    inheritance = "Mat"
ElseIf dad = "Y" Then
    inheritance = "Pat"
ElseIf mom = "N" And dad = "N" Then
    inheritance = "De novo"
End If

'User input
Dim response As Integer
response = MsgBox("Gene: " & gene & vbCr & "Transcript ID: " & transcript & vbCr & "Genomic coordinates: " & "chr" & chrom & ":" & start & ref & ">" & alt & vbCr & "Nucleotide change: " & nuc_change & vbCr & "Zygosity: " & zygosity & vbCr & "Inheritance: " & inheritance & vbCr & "Protein change: " & prot_change & vbCr & "Variant Interpretation: " & variant_interp, Buttons:=vbYesNoCancel, Title:="Format for Report?")

'If yes is clicked
If response = 6 Then
    'User Input
    omim_disease = InputBox("Please enter associated OMIM disease." & vbCr & vbCr & "e.g. Fanconi anemia, complementation group C")
    omim_inheritance = InputBox("Please enter disease associated inheritance pattern." & vbCr & vbCr & "e.g. AD, AR, XLD, XLR")
    omim_id = InputBox("Please enter OMIM disease ID." & vbCr & vbCr & "e.g. 227645")
    'Calculate last row in active sheet
    Dim LastRow As Integer
    LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    'Write data to bottom of active sheet
    ws.Cells(LastRow + 2, 16).Value = gene & " (" & transcript & ")"
    ws.Cells(LastRow + 2, 16).Characters(1, Len(gene)).Font.Italic = True
    ws.Cells(LastRow + 2, 17).Value = "chr" & chrom & ":" & start & ref & ">" & alt
    ws.Cells(LastRow + 2, 18).Value = nuc_change
    ws.Cells(LastRow + 2, 19).Value = zygosity & "/" & inheritance
    ws.Cells(LastRow + 2, 20).Value = prot_change
    ws.Cells(LastRow + 2, 21).Value = "(" & omim_inheritance & ") " & omim_disease & " (OMIM: " & omim_id & ")"
    ws.Cells(LastRow + 2, 22).Value = variant_interp
    ws.Range(ws.Cells(LastRow + 2, 16), ws.Cells(LastRow + 2, 22)).Copy
    MsgBox "Variant details copied to clipboard. Please paste directly into report table."
End If

End Sub
