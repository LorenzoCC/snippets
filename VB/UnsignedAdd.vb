Function UnsignedAdd(Addr As LongPtr, Incr As LongPtr) As LongPtr
' Add Incr to Addr, treating Addr as unsigned
' Incr is treated as signed.
' This function raises overflow errors when the unsigned LongPtr would overflow.
Dim SignBit As LongPtr

#If Win64 Then
    '64 bit version of Office
    SignBit = CLngPtr("-9,223,372,036,854,775,808")
#Else
    SignBit = &H80000000
#End If

If Incr >= 0 Then
    If Addr And SignBit Then
        ' Adr < 0, need to check whether unsigned would overflow
        If Addr >= -Incr Then
            Err.Raise 6, "UnsignedAdd", "Overflow"
        Else
            UnsignedAdd = Addr + Incr
        End If
    ElseIf (Addr Or SignBit) < -Incr Then
        'Adr >= 0, signed would not overflow
        UnsignedAdd = Addr + Incr
    Else
        'Adr >= 0, signed would overflow, need to wrap around
        UnsignedAdd = (Addr Or SignBit) + (Incr Or SignBit)
    End If
Else 'Inc < 0
    If Not Addr And SignBit Then
        ' Adr >= 0, need to check whether unsigned would neg. overflow
        If Addr < -Incr Then
            Err.Raise 6, "UnsignedAdd", "Overflow"
        Else
            UnsignedAdd = Addr + Incr
        End If
    ElseIf (Addr Or SignBit) > -Incr Then
        'Adr < 0, signed would not overflow
        UnsignedAdd = Addr + Incr
    Else
        'Adr < 0, signed would overflow, need to wrap around
        UnsignedAdd = (Addr And Not SignBit) + (Incr And Not SignBit)
    End If
End If
End Function