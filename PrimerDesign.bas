Attribute VB_Name = "PrimerDesign"
Sub PrimerDesign()

On Error Resume Next

Dim rng As Range
Set rng = selection

'Extract CoPath # from Workbook Title
Dim copath_no, copath_no_arr() As String
copath_no = ActiveWorkbook.Name
copath_no_arr = Split(copath_no, "_")
copath_no = copath_no_arr(1)

'Extract necessary identifiers
For I = 1 To Columns.Count
    If ActiveSheet.Cells(2, I) = "Gene" Then
        Dim gene As String
        gene = ActiveSheet.Cells(rng.Row, I).Text
    ElseIf ActiveSheet.Cells(2, I) = "Chr" Then
        Dim chrom As String
        chrom = ActiveSheet.Cells(rng.Row, I).Text
    ElseIf ActiveSheet.Cells(2, I) = "Start" Then
        Dim start As String
        start = ActiveSheet.Cells(rng.Row, I).Text
    ElseIf ActiveSheet.Cells(2, I) = "AAChange.ensGene" Then
        Dim exon, transcript_id, fields_arr() As String
        exon = ActiveSheet.Cells(rng.Row, I).Text
        fields_arr = Split(exon, ":")
        transcript_id = fields_arr(1)
        exon = fields_arr(2)
        For J = 1 To Len(exon)
            If IsNumeric(Mid(exon, J, 1)) Then
                exon_no = exon_no + Mid(exon, J, 1)
            End If
        Next
    End If
Next

verify = MsgBox("CoPath #: " & copath_no & vbCr & "Gene: " & gene & vbCr & "Chrom: " & chrom & vbCr & "Start: " & start & vbCr & "Exon: " & exon_no & vbCr & "Transcript: " & transcript_id, Buttons:=vbYesNoCancel, Title:="Get primer design sequences?")

If verify = 6 Then
    Shell ("python " & "U:\primer_design\primer_design.py " & chrom & " " & start & " " & gene & " " & exon_no & " " & transcript_id & " " & copath_no)
End If

End Sub
