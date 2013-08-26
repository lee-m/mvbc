Module Foo

    Public Sub Main()

    'Hex literal
    Dim hexLiteral1 = &H1234567890ABCDEF
    Dim hexLiteral2 = &H123456789abcdef
    Dim hexLiteral3 = &h123456789abcdef
    Dim hexLiteral4 = -&h123456789abcdef
    
    'Octal literal
    Dim octLiteral1 = &O1234567
    Dim octLiteral2 = &o1234567
    Dim octLiteral3 = -&o1234567
    
    'Short literal
    Dim shortLiteral1 = 32767S
    Dim shortLiteral2 = 32767s
    Dim shortLiteral3 = -32767s
    
    'Unsigned short
    Dim unsignedShort1 = 65535US
    Dim unsignedShort2 = 65535us
    Dim unsignedShort3 = -65535us
    
    'Integer literal
    Dim intLiteral1 = 1234567890
    Dim intLiteral2 = 1234567890%
    Dim intLiteral3 = 1234567890I
    Dim intLiteral4 = 1234567890i
    Dim intLiteral5 = -1234567890
    
    'Unsigned integer
    Dim unsignedInt1 = 4294967295
    Dim unsignedInt2 = 4294967295UI
    Dim unsignedInt3 = 4294967295ui
    Dim unsignedInt4 = -4294967295
    
    'Long literal
    Dim longLiteral1 = 8223372036854775807
    Dim longLiteral2 = 8223372036854775807&
    Dim longLiteral3 = 8223372036854775807L
    Dim longLiteral4 = 8223372036854775807l
    Dim longLiteral5 = 999&
    Dim longLiteral6 = 1234L
    Dim longLiteral7 = -1234L
    
    'Unsigned long
    Dim unsignedLongLiteral1 = 16446744073709551615UL
    Dim unsignedLongLiteral2 = 16446744073709551615ul
    Dim unsignedLongLiteral3 = -16446744073709551615ul
    
    'Decimal literal
    Dim decimalLiteral1 = 9223372036854775808@
    Dim decimalLiteral2 = 9223372036854775808D
    Dim decimalLiteral3 = 9223372036854775808d
    Dim decimalLiteral4 = 9223372036854775808d
    
    End Sub

End Module
