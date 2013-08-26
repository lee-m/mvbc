Module Test

	Public Sub Main()
		
		'Various types of floating point literals with no type characters
		Dim literal1 = .123
		Dim literal2 = .123E+123
		Dim literal3 = .123E-123
		Dim literal4 = 1.234
		Dim literal5 = 1.234E+123
		Dim literal6 = 1.234E-123
		Dim literal7 = 1234E+12
		Dim literal8 = 1234E-12
		Dim literal9 = 1234E12
		
		'Decimal floating point literals
		Dim decimalLiteral1 = .123@
		Dim decimalLiteral2 = .123E+12@
		Dim decimalLiteral3 = 1.234@
		Dim decimalLiteral4 = 1.234E+12@
		Dim decimalLiteral5 = 1234E-12@
		Dim decimalLiteral6 = .123D
		Dim decimalLiteral7 = .123E+12D
		Dim decimalLiteral8 = 1.234d
		Dim decimalLiteral9 = 1.234E+12d
		Dim decimalLiteral10 = 1234E-12d
		
		'Single floating point literals
		Dim singleLiteral1 = .123!
		Dim singleLiteral2 = .123E+12!
		Dim singleLiteral3 = 1.234!
		Dim singleLiteral4 = 1.234E+12!
		Dim singleLiteral5 = 1234E-123!
		Dim singleLiteral6 = .123F
		Dim singleLiteral7 = .123E+12F
		Dim singleLiteral8 = 1.234f
		Dim singleLiteral9 = 1.234E+12f
		Dim singleLiteral10 = 1234E-12f
		
		'Double floating point literals
		Dim doubleLiteral1 = .123#
		Dim doubleLiteral2 = .123E+123#
		Dim doubleLiteral3 = 1.234#
		Dim doubleLiteral4 = 1.234E+123#
		Dim doubleLiteral5 = 1234E-1234#
		Dim doubleLiteral6 = .123R
		Dim doubleLiteral7 = .123E+123R
		Dim doubleLiteral8 = 1.234r
		Dim doubleLiteral9 = 1.234E+123r
		Dim doubleLiteral10 = 1234E-1234r
		
	End Sub
	
End Module