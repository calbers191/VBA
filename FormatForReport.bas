Sub FormatForReport()

'Select row of variant to be transferred
Dim rng As Range
Set rng = Selection

'Extract gene
Dim gene As String
gene = ActiveSheet.Cells(rng.Row, 10).Text

'Extract transcript
Dim transcript As String, transcript_arr() As String
transcript = ActiveSheet.Cells(rng.Row, 105).Text
transcript_arr = Split(transcript, "|")
transcript = transcript_arr(7)

'Extract genomic coordinates
Dim genomic_coords As String
genomic_coords = "chr" & ActiveSheet.Cells(rng.Row, 5).Text & ":" & ActiveSheet.Cells(rng.Row, 6).Value & ActiveSheet.Cells(rng.Row, 7).Value & ">" & ActiveSheet.Cells(rng.Row, 8).Value

'Extract nucleotide change
Dim nuc_change As String
nuc_change = ActiveSheet.Cells(rng.Row, 11).Text

'Extract zygosity
Dim zygosity As String
zygosity = ActiveSheet.Cells(rng.Row, 13).Text
If zygosity = "het" Then
    zygosity = "Het"
End If
If zygosity = "hom" Then
    zygosity = "Hom"
End If
If zygosity = "hem" Or zygosity = "hemi" Then
    zygosity = "Hem"
End If

'Extract inheritance
Dim inheritance As String
If ActiveSheet.Cells(2, 14).Text = "in Father" Then
    If ActiveSheet.Cells(rng.Row, 14).Text = "Y" Then
        inheritance = "Pat"
    ElseIf ActiveSheet.Cells(rng.Row, 15).Text = "Y" Then
        inheritance = "Mat"
    Else: inheritance = "De novo"
    End If
ElseIf ActiveSheet.Cells(2, 14).Text = "in Mother" Then
    If ActiveSheet.Cells(rng.Row, 14).Text = "Y" Then
        inheritance = "Mat"
    ElseIf ActiveSheet.Cells(rng.Row, 15).Text = "Y" Then
        inheritance = "Pat"
        Else: inheritance = "De novo"
    End If
End If

'Extract protein change
Dim prot_change As String
prot_change = ActiveSheet.Cells(rng.Row, 12).Text

'Extract variant interpretation
Dim variant_interp As String
variant_interp = ActiveSheet.Cells(rng.Row, 16).Text

'User input
Dim response As Integer
response = MsgBox("Gene: " & gene & vbCr & "Transcript ID: " & transcript & vbCr & "Genomic coordinates: " & genomic_coords & vbCr & "Nucleotide change: " & nuc_change & vbCr & "Zygosity: " & zygosity & vbCr & "Inheritance: " & inheritance & vbCr & "Protein change: " & prot_change & vbCr & "Variant Interpretation: " & variant_interp, Buttons:=vbYesNoCancel, Title:="Format for Report?")

'If yes is clicked
If response = 6 Then
    'User Input
    omim_disease = InputBox("Please enter associated OMIM disease." & vbCr & vbCr & "e.g. Fanconi anemia, complementation group C")
    omim_inheritance = InputBox("Please enter disease associated inheritance pattern." & vbCr & vbCr & "e.g. AD, AR, XLD, XLR")
    omim_id = InputBox("Please enter OMIM disease ID." & vbCr & vbCr & "e.g. 227645")
    'Set workbook and worksheet variables
    Dim wb As Workbook, ws As Worksheet, LastRow As Integer
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets(1)
    'Calculate last row in active sheet
    LastRow = ActiveWorkbook.Sheets(1).Cells(Rows.Count, 9).End(xlUp).Row
    'Write data to bottom of active sheet
    ws.Cells(LastRow + 2, 16).Value = gene & " (" & transcript & ")"
    ws.Cells(LastRow + 2, 16).Characters(1, Len(gene)).Font.Italic = True
    ws.Cells(LastRow + 2, 17).Value = genomic_coords
    ws.Cells(LastRow + 2, 18).Value = nuc_change
    ws.Cells(LastRow + 2, 19).Value = zygosity & "/" & inheritance
    ws.Cells(LastRow + 2, 20).Value = prot_change
    ws.Cells(LastRow + 2, 21).Value = "(" & omim_inheritance & ") " & omim_disease & " (OMIM: " & omim_id & ")"
    ws.Cells(LastRow + 2, 22).Value = variant_interp
    ws.Range(ws.Cells(LastRow + 2, 16), ws.Cells(LastRow + 2, 22)).Copy
    MsgBox "Variant details copied to clipboard. Please paste directly into report table."
End If
End Sub
