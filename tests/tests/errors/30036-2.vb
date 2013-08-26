'#compiler-message:30036-2.vb(9:31) : error BC30036: Overflow
'#compiler-message:30036-2.vb(10:30) : error BC30036: Overflow
'#exit-code:1

Module Test

    Public Sub Main()

        Dim singleOverflow =  3.402823e40f
        Dim doubleOverflow = 1.7976931348623157E+400
        
    End Sub
    
End Module