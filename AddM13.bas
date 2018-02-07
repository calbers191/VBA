Attribute VB_Name = "AddM13"
Sub AddM13()

Dim sequence As Range
Set sequence = selection

Dim m13f, m13r As String
m13f = "GTAAAACGACGGCCAG"
m13r = "CAGGAAACAGCTATGAC"

Dim primer_f, primer_r As String
primer_f = sequence
primer_f = m13f & primer_f
primer_r = Cells(sequence.Row + 1, sequence.Column)
primer_r = m13r & primer_r

Cells(sequence.Row, sequence.Column + 4) = primer_f
Cells(sequence.Row + 1, sequence.Column + 4) = primer_r

End Sub
