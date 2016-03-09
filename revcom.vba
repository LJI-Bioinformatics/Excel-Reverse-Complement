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
    input_str = Replace(input_str, "A", "1")
    If isRNA = 1 Then
        input_str = Replace(input_str, "U", "A")
        input_str = Replace(input_str, "1", "U")
    Else
        input_str = Replace(input_str, "T", "A")
        input_str = Replace(input_str, "1", "T")
    End If
    input_str = Replace(input_str, "C", "1")
    input_str = Replace(input_str, "G", "C")
    input_str = Replace(input_str, "1", "G")

    ' now deal with lowercase letters
    ' this could be more elegant
    input_str = Replace(input_str, "a", "1")
    If isRNA = 1 Then
        input_str = Replace(input_str, "u", "a")
        input_str = Replace(input_str, "1", "u")
    Else
        input_str = Replace(input_str, "t", "a")
        input_str = Replace(input_str, "1", "t")
    End If
    input_str = Replace(input_str, "c", "1")
    input_str = Replace(input_str, "g", "c")
    input_str = Replace(input_str, "1", "g")

    
    complement = input_str

End Function
