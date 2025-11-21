

Module Validate

    Public Function IsFormat(ByVal pstrArg As String, ByVal pstrFormat As String) As Boolean
        Dim i As Integer ' In my pgms, i is always an integer used in a for-next loop

        ' Check length of string
        If pstrArg.Length <> pstrFormat.Length Then Return False
        ' Check to see that each of the characters matches up, where 'X' is an integer
        For i = 0 To pstrArg.Length - 1
            If pstrFormat.Substring(0, 1) = "X" Then
                If Not IsIntegerInRange(pstrArg.Substring(0, 1), 0, 9) Then Return False
            Else
                If Not pstrFormat.Substring(0, 1) = pstrArg.Substring(0, 1) Then Return False
            End If
        Next i
        Return True
    End Function
    Public Function IsIntegerGreaterThan(ByVal pstrArg As String, ByVal pintMin As Integer) As Boolean
        If Not IsInteger(pstrArg) Then Return False
        If System.Convert.ToInt64(pstrArg) <= pintMin Then Return False
        Return True
    End Function
    Public Function IsIntegerLessThan(ByVal pstrArg As String, ByVal pintMax As Integer) As Boolean
        If Not IsInteger(pstrArg) Then Return False
        If System.Convert.ToInt64(pstrArg) >= pintMax Then Return False
        Return True
    End Function
    Public Function IsSingle(ByVal pstrArg As String) As Boolean
        Dim psngArg As Single
        If IsNumeric(pstrArg) Then
            psngArg = System.Convert.ToSingle(pstrArg)
            If psngArg > 0 Then
                Return True
            End If
        End If
        Return False
    End Function
    Public Function IsStringInRange(ByVal pstrArg As String, ByVal pintMin As Integer, ByVal pintMax As Integer) As Boolean
        If (pstrArg.Length >= pintMin) And (pstrArg.Length <= pintMax) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub ValidateTextBoxDisplay(ByRef ptxt As TextBox, ByVal pbln As Boolean)
        If pbln Then
            ptxt.ForeColor = Color.Black
        Else
            ptxt.ForeColor = Color.Red
        End If
    End Sub
    Public Sub ValidateTextBoxFormat(ByRef ptxt As TextBox, ByVal pstrFormat As String)
        ValidateTextBoxDisplay(ptxt, IsFormat(ptxt.Text, pstrFormat))
    End Sub
    Public Sub ValidateTextBoxLength(ByRef ptxt As TextBox, ByVal pintMin As Integer, ByVal pintMax As Integer)
        ValidateTextBoxDisplay(ptxt, IsStringInRange(ptxt.Text, pintMin, pintMax))
    End Sub

    ' The next 2 functions came from the book. Why reinvent the wheel?
    Public Function IsIntegerInRange(ByVal pstrArg As String, ByVal pintMin As Integer, ByVal pintMax As Integer) As Boolean
        ' Determine wheter a string argument is contains an Integer value between
        ' a specified range. The first argument contains the string to validate. The
        ' second argument contains the minimum valid value, and the third argument
        ' contains the maximum valid value.
        Dim pintArg As Integer
        ' Call in IsInteger function to first determine whether the argument
        ' is an integer.
        If IsInteger(pstrArg) Then
            pintArg = System.Convert.ToInt32(pstrArg)
            ' Second determine whether the argument is between the minimum and
            ' maximum allowable values.
            If pintArg >= pintMin And pintArg <= pintMax Then
                Return True
            End If
        End If
    End Function

    Public Function IsInteger(ByVal pstrArg As String) As Boolean
        ' Determine whether an argument contains an Integer value. First test that
        ' the argument is numeric. If it is, convert the argument to both an 
        ' Integer and a Single. If the values are the same, then the argument
        ' is an Integer, and the function returns True. Otherwise the argument
        ' is not an Integer, and the function returns False.
        Dim pdblArg As Double
        If IsNumeric(pstrArg) Then
            pdblArg = System.Convert.ToDouble(pstrArg)
            If System.Convert.ToInt32(pdblArg) = pdblArg Then
                Return True
            End If
        End If
        Return False
    End Function

End Module
