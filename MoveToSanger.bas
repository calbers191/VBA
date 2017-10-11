Sub MoveToSanger()

'Select row of variant to be transferred
Dim rng As Range
Set rng = Selection

Dim book As Workbook
Dim sheet As Worksheet
Set book = ActiveWorkbook
Set sheet = ActiveSheet

'Extract CoPath # from Workbook Title
Dim copath_no As String, copath_no_arr() As String
copath_no = ActiveWorkbook.Name
copath_no_arr = Split(copath_no, "_")
copath_no = copath_no_arr(1)

'Extract identifiers from selected row
For i = 1 To 130
    If sheet.Cells(2, i) = "Chr" Then
        Dim chrom As String
        chrom = sheet.Cells(rng.Row, i).Text
    ElseIf sheet.Cells(2, i) = "Start" Then
        Dim start As String
        start = sheet.Cells(rng.Row, i).Text
    ElseIf sheet.Cells(2, i) = "Ref" Then
        Dim ref As String
        ref = sheet.Cells(rng.Row, i).Text
    ElseIf sheet.Cells(2, i) = "Alt" Then
        Dim alt As String
        alt = sheet.Cells(rng.Row, i).Text
    ElseIf sheet.Cells(2, i) = "Gene" Then
        Dim gene As String
        gene = sheet.Cells(rng.Row, i).Text
    ElseIf sheet.Cells(2, i) = "HGVS" Then
        Dim variant_details As String
        variant_details = sheet.Cells(rng.Row, i).Text
    ElseIf sheet.Cells(2, i) = "Zygosity" Then
        Dim zygosity As String
        zygosity = sheet.Cells(rng.Row, i).Text
    End If
Next

'User Input
Dim response As Integer
response = MsgBox("CoPath #: " & copath_no & vbCr & "Gene: " & gene & vbCr & "Genomic coordinates: " & "chr" & chrom & ":" & start & ref & ">" & alt & vbCr & "Variant details: " & variant_details & vbCr & "Zygosity: " & zygosity, Buttons:=vbYesNoCancel, Title:="Transfer to Sanger Confirmation Log?")

'If yes is clicked
If response = 6 Then
    'User Input
    Dim initials As String
    initials = InputBox("Please enter your initials.")
    'Open Sanger Tracking Log and set variables
    Dim wb As Workbook, ws As Worksheet, LastRow As Integer
    Set wb = Workbooks.Open("Y:\Exome Production Files\Sanger Confirmation\Sanger Tracking.xlsm")
    Set ws = wb.Sheets(1)
    'Calculate last row in Sanger tracking log
    LastRow = ActiveWorkbook.Sheets(1).Cells(Rows.Count, "A").End(xlUp).Row
    'Write data to Sanger Tracking Log
    ws.Cells(LastRow + 1, 1).Value = copath_no
    ws.Cells(LastRow + 1, 2).Value = gene
    ws.Cells(LastRow + 1, 3).Value = genomic_coords
    ws.Cells(LastRow + 1, 4).Value = variant_details
    ws.Cells(LastRow + 1, 5).Value = zygosity
    ws.Cells(LastRow + 1, 6).Value = initials
    ws.Cells(LastRow + 1, 7).Value = Format(Date, "mm/dd/yyyy")
    'Save and close the file
    wb.Save
    wb.Close
    'Call sub to send notification email
    Email = MsgBox("Send notification email?", Buttons:=vbYesNo)
    If Email = 6 Then
        SendNotificationEmail
    End If
End If

End Sub
