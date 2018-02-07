Attribute VB_Name = "OpenVariantLinks"
Sub OpenVariantLinks()

Application.ScreenUpdating = False

Dim rng As Range
Set rng = selection

'Extract CoPath # from Workbook Title
Dim copath_no As String, copath_no_arr() As String
copath_no = ActiveWorkbook.Name
copath_no_arr = Split(copath_no, "_")
copath_no = copath_no_arr(1)

'Extract necessary identifiers
For I = 1 To Columns.Count
    If ActiveSheet.Cells(2, I) = "Gene.ensGene" Then
        Dim ensembl_id As String
        ensembl_id = ActiveSheet.Cells(rng.Row, I).Text
    ElseIf ActiveSheet.Cells(2, I) = "Gene" Then
        Dim gene As String
        gene = ActiveSheet.Cells(rng.Row, I).Text
    ElseIf ActiveSheet.Cells(2, I) = "OMIM" Then
        Dim omim_id As String
        omim_id = ActiveSheet.Cells(rng.Row, I).Text
    ElseIf ActiveSheet.Cells(2, I) = "Chr" Then
        Dim chrom As String
        chrom = ActiveSheet.Cells(rng.Row, I).Text
    ElseIf ActiveSheet.Cells(2, I) = "Start" Then
        Dim start As String
        start = ActiveSheet.Cells(rng.Row, I).Text
    ElseIf ActiveSheet.Cells(2, I) = "Ref" Then
        Dim ref As String
        ref = ActiveSheet.Cells(rng.Row, I).Text
    ElseIf ActiveSheet.Cells(2, I) = "Alt" Then
        Dim alt As String
        alt = ActiveSheet.Cells(rng.Row, I).Text
    ElseIf ActiveSheet.Cells(2, I) = "ClinVar-ID|(SampleData|Source|SourceVer|Datetime|IsValid|Gene|GeneIDs|ID)" Then
        Dim clinvar_id As String
        clinvar_id = ActiveSheet.Cells(rng.Row, I).Text
        If clinvar_id <> vbNullString Then
            clinvar_id_arr = Split(clinvar_id, "|||")
            clinvar_id = clinvar_id_arr(1)
        End If
    ElseIf ActiveSheet.Cells(2, I) = "DNA Change" Then
        Dim nuc_change As String
        nuc_change = ActiveSheet.Cells(rng.Row, I).Text
    ElseIf ActiveSheet.Cells(2, I) = "AA Change" Then
        Dim prot_change As String
        prot_change = ActiveSheet.Cells(rng.Row, I).Text
    ElseIf ActiveSheet.Cells(2, I) = "DNA_Change|(SampleData|Source|SourceVer|Datetime|IsValid|Gene|GeneIDs|Transcript|TranscriptIDs|DNA Change|Type)" Then
        transcript = ActiveSheet.Cells(rng.Row, I).Text
        transcript_arr = Split(transcript, "|")
        transcript = transcript_arr(7)
    End If
Next

'Extract nucleotide and amino acid positions
For I = 1 To Len(nuc_change)
    If IsNumeric(Mid(nuc_change, I, 1)) Then
        nuc_pos = nuc_pos + Mid(nuc_change, I, 1)
    End If
Next
For I = 1 To Len(prot_change)
    If IsNumeric(Mid(prot_change, I, 1)) Then
        aa_pos = aa_pos + Mid(prot_change, I, 1)
    End If
Next

'Get amino acid abbreviations from prot_change
ref_aa = Mid(prot_change, 3, 3)
alt_aa = Right(prot_change, 3)

'AA collection
Dim aa_coll As New Collection
aa_coll.Add "F", "Phe"
aa_coll.Add "L", "Leu"
aa_coll.Add "I", "Ile"
aa_coll.Add "M", "Met"
aa_coll.Add "V", "Val"
aa_coll.Add "S", "Ser"
aa_coll.Add "P", "Pro"
aa_coll.Add "T", "Thr"
aa_coll.Add "A", "Ala"
aa_coll.Add "Y", "Tyr"
aa_coll.Add "H", "His"
aa_coll.Add "Q", "Gln"
aa_coll.Add "N", "Asn"
aa_coll.Add "K", "Lys"
aa_coll.Add "D", "Asp"
aa_coll.Add "E", "Glu"
aa_coll.Add "C", "Cys"
aa_coll.Add "W", "Trp"
aa_coll.Add "R", "Arg"
aa_coll.Add "G", "Gly"

'Get one letter codes
On Error Resume Next
ref_aa_oneletter = aa_coll.Item(ref_aa)
alt_aa_oneletter = aa_coll.Item(alt_aa)

'Extract first letter of gene to determine file path
Dim first_letter As String
first_letter = Left(gene, 1)

'Extract Pubmed ID's from HGMD
Dim hgmd_book As Workbook, hgmd_sheet As Worksheet
Set hgmd_book = Workbooks.Open("\\columbuschildrens.net\apps\ngs\samples\Exome Production Files\HGMD View\Previous Versions and View Validations\HGMD_2017.2 View pubmed_ids.xlsx")
Set hgmd_sheet = hgmd_book.Sheets(1)

For I = 1 To Rows.Count
    If hgmd_sheet.Cells(I, 1) = gene And (hgmd_sheet.Cells(I, 2) = nuc_change Or hgmd_sheet.Cells(I, 3) = prot_change) Then
        pubmed_ids = hgmd_sheet.Cells(I, 4)
    End If
Next

hgmd_book.Close

'Set up links
Dim varsome, gnomad, exac, clinvar_gene, omim As String
varsome = "https://varsome.com/variant/hg19/" & chrom & "-" & start & "-" & ref & "-" & alt
gnomad = "http://gnomad.broadinstitute.org/variant/" & chrom & "-" & start & "-" & ref & "-" & alt
clinvar_variant = "https://www.ncbi.nlm.nih.gov/clinvar/variation/" & clinvar_id
omim = "https://www.omim.org/entry/" & omim_id
clinvar_gene = "https://www.ncbi.nlm.nih.gov/clinvar/?term=" & gene & "%5Bgene%5D"
exac = "http://exac.broadinstitute.org/gene/" & ensembl_id
mastermind = "https://mastermind.genomenon.com/detail?gene=" & gene & "&disease=all%20diseases&mutation=" & gene & ":" & ref_aa_oneletter & aa_pos & alt_aa_oneletter & "&mutation_source=fulltext"
'
''Open in Default Browser
'On Error Resume Next
'ActiveWorkbook.FollowHyperlink varsome
'ActiveWorkbook.FollowHyperlink gnomad
'If clinvar_variant <> vbNullString Then
'    ActiveWorkbook.FollowHyperlink clinvar_variant
'End If
'ActiveWorkbook.FollowHyperlink omim
'ActiveWorkbook.FollowHyperlink clinvar_gene
'ActiveWorkbook.FollowHyperlink exac
'If pubmed_ids <> vbNullString Then
'    ActiveWorkbook.FollowHyperlink "https://www.ncbi.nlm.nih.gov/pubmed/" & pubmed_ids
'End If
'ActiveWorkbook.FollowHyperlink "\\columbuschildrens.net\apps\ngs\samples\Exome Production Files\HGMD View\plots_DM_only\" & first_letter & "\" & gene & ".png"
'ActiveWorkbook.FollowHyperlink "https://www.google.com/search?q=" & gene & "+" & "(" & nuc_pos & "|" & aa_pos & ")"
'ActiveWorkbook.FollowHyperlink mastermind

'Open in Chrome
On Error Resume Next
Shell ("C:\Users\cja001\AppData\Local\Google\Chrome\Application\chrome.exe -url " & varsome)
Shell ("C:\Users\cja001\AppData\Local\Google\Chrome\Application\chrome.exe -url " & gnomad)
If clinvar_id <> vbNullString Then
    Shell ("C:\Users\cja001\AppData\Local\Google\Chrome\Application\chrome.exe -url " & clinvar_variant)
End If
Shell ("C:\Users\cja001\AppData\Local\Google\Chrome\Application\chrome.exe -url " & omim)
Shell ("C:\Users\cja001\AppData\Local\Google\Chrome\Application\chrome.exe -url " & clinvar_gene)
Shell ("C:\Users\cja001\AppData\Local\Google\Chrome\Application\chrome.exe -url " & exac)
Shell ("C:\Users\cja001\AppData\Local\Google\Chrome\Application\chrome.exe -url " & """" & "\\columbuschildrens.net\apps\ngs\samples\Exome Production Files\HGMD View\plots_DM_only\" & first_letter & "\" & gene & ".png" & """")
If pubmed_ids <> vbNullString Then
    pubmed_ids_arr = Split(pubmed_ids, ",")
    For Each ID In pubmed_ids_arr
        ID = Replace(ID, " ", "")
        If ID <> vbNullString Then
           Shell ("C:\Users\cja001\AppData\Local\Google\Chrome\Application\chrome.exe -url " & "https://www.ncbi.nlm.nih.gov/pubmed/" & ID)
        End If
    Next
End If
Shell ("C:\Users\cja001\AppData\Local\Google\Chrome\Application\chrome.exe -url " & "https://www.google.com/search?q=" & gene & "+" & "(" & nuc_pos & "|" & aa_pos & ")")
Shell ("C:\Users\cja001\AppData\Local\Google\Chrome\Application\chrome.exe -url " & mastermind)

'Open VAF for variant
If prot_change = vbNullString Then
    Workbooks.Open (ActiveWorkbook.Path & "\VAFs\" & copath_no & "_" & transcript & "(" & gene & ")" & "_" & Replace(nuc_change, ">", "^") & ".xlsx")
Else
    Workbooks.Open (ActiveWorkbook.Path & "\VAFs\" & copath_no & "_" & transcript & "(" & gene & ")" & "_" & Replace(nuc_change, ">", "^") & "_" & prot_change & ".xlsx")
End If

End Sub
