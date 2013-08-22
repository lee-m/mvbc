'#compiler-message:30036-1.vb(25:36) : error BC30036: Overflow
'#compiler-message:30036-1.vb(26:36) : error BC30036: Overflow
'#compiler-message:30036-1.vb(29:36) : error BC30036: Overflow
'#compiler-message:30036-1.vb(30:36) : error BC30036: Overflow
'#compiler-message:30036-1.vb(33:34) : error BC30036: Overflow
'#compiler-message:30036-1.vb(34:34) : error BC30036: Overflow
'#compiler-message:30036-1.vb(35:34) : error BC30036: Overflow
'#compiler-message:30036-1.vb(38:35) : error BC30036: Overflow
'#compiler-message:30036-1.vb(39:35) : error BC30036: Overflow
'#compiler-message:30036-1.vb(42:35) : error BC30036: Overflow
'#compiler-message:30036-1.vb(43:35) : error BC30036: Overflow
'#compiler-message:30036-1.vb(44:35) : error BC30036: Overflow
'#compiler-message:30036-1.vb(45:35) : error BC30036: Overflow
'#compiler-message:30036-1.vb(48:43) : error BC30036: Overflow
'#compiler-message:30036-1.vb(49:43) : error BC30036: Overflow
'#compiler-message:30036-1.vb(50:43) : error BC30036: Overflow
'#compiler-message:30036-1.vb(53:31) : error BC30036: Overflow
'#compiler-message:30036-1.vb(55:31) : error BC30036: Overflow

Module Foo

    Public Sub Main()

        'Short literal overflow
        Dim shortLiteralOverflow1 = 42767S
        Dim shortLiteralOverflow2 = 42767s

        'Unsigned short overflow
        Dim unsignedShortOverflow = 75535US
        Dim unsignedShortOverflow = 75535us

        'Integer literal overflow
        Dim intLiteralOverflow1 = 99999999999999999999I
        Dim intLiteralOverflow2 = 99999999999999999999%
        Dim intLiteralOverflow3 = 99999999999999999999i

        'Unsigned integer overflow
        Dim unsignedIntOverflow1 = 5294967295UI
        Dim unsignedIntOverflow2 = 5294967295ui

        'Long literal overflow
        Dim longLiteralOverflow1 = 10223372036854775807
        Dim longLiteralOverflow2 = 10223372036854775807&
        Dim longLiteralOverflow3 = 10223372036854775807L
        Dim longLiteralOverflow4 = 10223372036854775807l

        'Unsigned long overflow
        Dim unsignedLongLiteralOverflow1 = 19446744073709551615
        Dim unsignedLongLiteralOverflow2 = 19446744073709551615UL
        Dim unsignedLongLiteralOverflow3 = 19446744073709551615ul

        'Decimal overflow
        Dim decimalOverflow1 = 9223372036854775808
        Dim decimalOverflow2 = 9223372036854775807
        Dim decimalOverflow3 = 79228162514264337593543950335

    End Sub
    
End Module
