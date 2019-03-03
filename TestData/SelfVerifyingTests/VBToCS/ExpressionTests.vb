Imports System
Imports System.Linq
Imports Xunit

Module Program

    Public Class Tests

        <Fact>
        Public Sub TestFloatingPointDivisionOfIntegers()
            Dim x = 7 / 2
            Assert.Equal(x, 3.5)
        End Sub

        <Fact>
        Public Sub TestIntegerDivisionOfIntegers()
            Dim x = 7 \ 2
            Assert.Equal(x, 3)
        End Sub

        <Fact>
        Public Sub TestDecimalDivisionOfDecimals()
            Dim x = 7D / 2D
            Assert.Equal(x, 3.5D)
        End Sub

        <Fact>
        Public Sub TestIntFunctionFloorsDecimal() 'https://github.com/icsharpcode/CodeConverter/issues/238
            Dim a = 30.4
            Dim b = 20.5
            Dim c = 10.6
            Dim d = -10.4
            Dim e = -20.5
            Dim f = -30.6
            Assert.Equal(30, Int(a))
            Assert.Equal(20, Int(b))
            Assert.Equal(10, Int(c))
            Assert.Equal(-11, Int(d))
            Assert.Equal(-21, Int(e))
            Assert.Equal(-31, Int(f))
        End Sub

        <Fact> 'https://github.com/icsharpcode/CodeConverter/issues/105
        Public Sub VisualBasicEqualityOfEmptyStringAndNothingIsPreserved()
            Dim record = ""

            Dim nullObject As Object = Nothing
            Dim nullString As String = Nothing
            Dim emptyStringObject As Object = ""
            Dim emptyString = ""
            Dim nonEmptyString = "a"
            Dim emptyCharArray = New Char(){}
            Dim nullCharArray As Char() = Nothing

            If nullObject = nullObject Then record &= "1" Else record &= "0"
            If nullObject = nullString Then record &= "1" Else record &= "0"
            If nullObject = emptyStringObject Then record &= "1" Else record &= "0"
            If nullObject = emptyString Then record &= "1" Else record &= "0"
            If nullObject = nonEmptyString Then record &= "1" Else record &= "0"
            If nullObject = emptyCharArray Then record &= "1" Else record &= "0"
            If nullObject = nullCharArray Then record &= "1" Else record &= "0"
            record &= " "
            If nullString = nullObject Then record &= "1" Else record &= "0"
            If nullString = nullString Then record &= "1" Else record &= "0"
            If nullString = emptyStringObject Then record &= "1" Else record &= "0"
            If nullString = emptyString Then record &= "1" Else record &= "0"
            If nullString = nonEmptyString Then record &= "1" Else record &= "0"
            If nullString = emptyCharArray Then record &= "1" Else record &= "0"
            If nullString = nullCharArray Then record &= "1" Else record &= "0"
            record &= " "
            If emptyStringObject = nullObject Then record &= "1" Else record &= "0"
            If emptyStringObject = nullString Then record &= "1" Else record &= "0"
            If emptyStringObject = emptyStringObject Then record &= "1" Else record &= "0"
            If emptyStringObject = emptyString Then record &= "1" Else record &= "0"
            If emptyStringObject = nonEmptyString Then record &= "1" Else record &= "0"
            If emptyStringObject = emptyCharArray Then record &= "1" Else record &= "0"
            If emptyStringObject = nullCharArray Then record &= "1" Else record &= "0"
            record &= " "
            If emptyString = nullObject Then record &= "1" Else record &= "0"
            If emptyString = nullString Then record &= "1" Else record &= "0"
            If emptyString = emptyStringObject Then record &= "1" Else record &= "0"
            If emptyString = emptyString Then record &= "1" Else record &= "0"
            If emptyString = nonEmptyString Then record &= "1" Else record &= "0"
            If emptyString = emptyCharArray Then record &= "1" Else record &= "0"
            If emptyString = nullCharArray Then record &= "1" Else record &= "0"
            record &= " "
            If nonEmptyString = nullObject Then record &= "1" Else record &= "0"
            If nonEmptyString = nullString Then record &= "1" Else record &= "0"
            If nonEmptyString = emptyStringObject Then record &= "1" Else record &= "0"
            If nonEmptyString = emptyString Then record &= "1" Else record &= "0"
            If nonEmptyString = nonEmptyString Then record &= "1" Else record &= "0"
            If nonEmptyString = emptyCharArray Then record &= "1" Else record &= "0"
            If nonEmptyString = nullCharArray Then record &= "1" Else record &= "0"
            record &= " "
            If emptyCharArray = nullObject Then record &= "1" Else record &= "0"
            If emptyCharArray = nullString Then record &= "1" Else record &= "0"
            If emptyCharArray = emptyStringObject Then record &= "1" Else record &= "0"
            If emptyCharArray = emptyString Then record &= "1" Else record &= "0"
            If emptyCharArray = nonEmptyString Then record &= "1" Else record &= "0"
            If emptyCharArray = emptyCharArray Then record &= "1" Else record &= "0"
            If emptyCharArray = nullCharArray Then record &= "1" Else record &= "0"
            record &= " "
            If nullCharArray = nullObject Then record &= "1" Else record &= "0"
            If nullCharArray = nullString Then record &= "1" Else record &= "0"
            If nullCharArray = emptyStringObject Then record &= "1" Else record &= "0"
            If nullCharArray = emptyString Then record &= "1" Else record &= "0"
            If nullCharArray = nonEmptyString Then record &= "1" Else record &= "0"
            If nullCharArray = emptyCharArray Then record &= "1" Else record &= "0"
            If nullCharArray = nullCharArray Then record &= "1" Else record &= "0"
            
            Assert.Equal("1111011 1111011 1111011 1111011 0000100 1111011 1111011", record)

            Assert.True(emptyCharArray = New Char(){}, "Char arrays should be compared as strings because that's what happens in VB")

            Dim a1 As Object = 3
            Dim a2 As Object = 3
            Dim b As Object = 4
            Assert.True(a1 = a2, "Identical values stored in objects should be equal")
            Assert.False(a1 = b, "Different values stored in objects should not be equal")
        End Sub
    End Class

End Module
