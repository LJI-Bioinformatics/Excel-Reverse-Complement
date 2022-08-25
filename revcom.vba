Function reverse(input_str As String)
    ' reverse a string
    xLen = VBA.Len(input_str)
    rev_str = ""
    For i = 1 To xLen
        getChar = VBA.Right(input_str, 1)
        input_str = VBA.Left(input_str, xLen - i)
        rev_str = rev_str & getChar
    Next
    
    reverse = rev_str

End Function

Function revcom(input_str As String, Optional ByVal isRNA = 0)
    ' calculate the reverse complement of a DNA/RNA sequence
    revcom = complement(reverse(input_str), isRNA)
    
End Function

Function complement(input_str As String, Optional ByVal isRNA = 0)

    ' calculate the complement of a DNA/RNA sequence
    If isRNA = 1 Then
        input_str = swap_letters(input_str, "A", "U")
    Else
        input_str = swap_letters(input_str, "A", "T")
    End If
    input_str = swap_letters(input_str, "C", "G")
    
    'now deal with the ambiguous codes
    input_str = swap_letters(input_str, "R", "Y")
    input_str = swap_letters(input_str, "K", "M")
    input_str = swap_letters(input_str, "B", "V")
    input_str = swap_letters(input_str, "D", "H")

    complement = input_str

End Function

Function swap_letters(input_str As String, l1 As String, l2 As String)
    'swap all instances of L1 and L2 in the input string
    input_str = Replace(input_str, l1, "1")
    input_str = Replace(input_str, l2, l1)
    input_str = Replace(input_str, "1", l2)
    
    swap_letters = input_str

End Function
