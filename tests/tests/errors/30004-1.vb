'#compiler-message:30004-1.vb(9:23) : error BC30004: Character constant must contain exactly one character.
'#compiler-message:30004-1.vb(10:23) : error BC30004: Character constant must contain exactly one character.
'#exit-code:1

Module Test

    Public Sub Main()

        Dim a = "abc"c  
        Dim b = "abc"C
        
    End Sub
    
End Module
