' Program name: Midterm Exam - Calculator
' Author:       Crystal Pendergrass
' Date:         9 March 2015
' Purpose:      The Calculator works similar to the Windows calculator. The user 
'               types the first number using the keypad buttons of the calculator 
'               (no keyboard input required), presses an operation, then the user types 
'               the second number using the keypad buttons of the calculator (no keyboard 
'               input required). If it clicks the ‘=’ it performs the operation. If it 
'               clicks an operation button, it performs the previous calculation, shows 
'               the result of the first calculation and waits for the second operand of 
'               the new operation. If the user clicks Clear, the calculator shows the 
'               value 0 and resets its memory.

Public Class frmCalculator

    Dim decLeftOperand As Decimal
    Dim decRightOperand As Decimal
    Dim decResult As Decimal
    Dim strOperator As String
    Dim strLastOperator As String
    Dim blnIsFirstDigit As Boolean

    '****************************************************************************** 
    'When the form is loaded the display is set to 0 and the flag to indicate that
    'the next number entered is the first digit of some number.  
    '******************************************************************************
    Private Sub frmCalculator_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtDisplay.Text = "0"
        blnIsFirstDigit = True
    End Sub

    '****************************************************************************** 
    'When the "Clear" button is pressed the display is reset to "0" and any values 
    'stored for the operands and operators are erased.  The user can now begin
    'entering digits for a new number
    '******************************************************************************
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtDisplay.Text = "0"
        blnIsFirstDigit = True
        decLeftOperand = Nothing
        decRightOperand = Nothing
        decResult = Nothing
        strOperator = Nothing

        btnDecimal.Enabled = True
    End Sub

    '****************************************************************************** 
    'The following subroutines control how the number buttons behave.  Because they 
    'essentially behave the same, this comment will describe all of them.  
    'If either operand is empty and the user is entering the first digit of a 
    'number, then the display is set to the value of that button.  Otherwise 
    'any pressed number will be added to the exisiting numbers in the display.
    '******************************************************************************
    Private Sub btnOne_Click(sender As Object, e As EventArgs) Handles btnOne.Click
        If ((decLeftOperand = Nothing Or decRightOperand = Nothing) And decResult = Nothing) Then
            If (blnIsFirstDigit) Then
                txtDisplay.Text = btnOne.Text
                blnIsFirstDigit = False
            Else
                txtDisplay.Text = txtDisplay.Text & btnOne.Text
            End If
        End If
    End Sub

    Private Sub btnTwo_Click(sender As Object, e As EventArgs) Handles btnTwo.Click
        If ((decLeftOperand = Nothing Or decRightOperand = Nothing) And decResult = Nothing) Then
            If (blnIsFirstDigit) Then
                txtDisplay.Text = btnTwo.Text
                blnIsFirstDigit = False
            Else
                txtDisplay.Text = txtDisplay.Text & btnTwo.Text
            End If
        End If
    End Sub

    Private Sub btnThree_Click(sender As Object, e As EventArgs) Handles btnThree.Click
        If ((decLeftOperand = Nothing Or decRightOperand = Nothing) And decResult = Nothing) Then
            If (blnIsFirstDigit) Then
                txtDisplay.Text = btnThree.Text
                blnIsFirstDigit = False
            Else
                txtDisplay.Text = txtDisplay.Text & btnThree.Text
            End If
        End If
    End Sub

    Private Sub btnFour_Click(sender As Object, e As EventArgs) Handles btnFour.Click
        If ((decLeftOperand = Nothing Or decRightOperand = Nothing) And decResult = Nothing) Then
            If (blnIsFirstDigit) Then
                txtDisplay.Text = btnFour.Text
                blnIsFirstDigit = False
            Else
                txtDisplay.Text = txtDisplay.Text & btnFour.Text
            End If
        End If
    End Sub

    Private Sub btnFive_Click(sender As Object, e As EventArgs) Handles btnFive.Click
        If ((decLeftOperand = Nothing Or decRightOperand = Nothing) And decResult = Nothing) Then
            If (blnIsFirstDigit) Then
                txtDisplay.Text = btnFive.Text
                blnIsFirstDigit = False
            Else
                txtDisplay.Text = txtDisplay.Text & btnFive.Text
            End If
        End If
    End Sub

    Private Sub btnSix_Click(sender As Object, e As EventArgs) Handles btnSix.Click
        If ((decLeftOperand = Nothing Or decRightOperand = Nothing) And decResult = Nothing) Then
            If (blnIsFirstDigit) Then
                txtDisplay.Text = btnSix.Text
                blnIsFirstDigit = False
            Else
                txtDisplay.Text = txtDisplay.Text & btnSix.Text
            End If
        End If
    End Sub

    Private Sub btnSeven_Click(sender As Object, e As EventArgs) Handles btnSeven.Click
        If ((decLeftOperand = Nothing Or decRightOperand = Nothing) And decResult = Nothing) Then
            If (blnIsFirstDigit) Then
                txtDisplay.Text = btnSeven.Text
                blnIsFirstDigit = False
            Else
                txtDisplay.Text = txtDisplay.Text & btnSeven.Text
            End If
        End If
    End Sub

    Private Sub btnEight_Click(sender As Object, e As EventArgs) Handles btnEight.Click
        If ((decLeftOperand = Nothing Or decRightOperand = Nothing) And decResult = Nothing) Then
            If (blnIsFirstDigit) Then
                txtDisplay.Text = btnEight.Text
                blnIsFirstDigit = False
            Else
                txtDisplay.Text = txtDisplay.Text & btnEight.Text
            End If
        End If
    End Sub

    Private Sub btnNine_Click(sender As Object, e As EventArgs) Handles btnNine.Click
        If ((decLeftOperand = Nothing Or decRightOperand = Nothing) And decResult = Nothing) Then
            If (blnIsFirstDigit) Then
                txtDisplay.Text = btnNine.Text
                blnIsFirstDigit = False
            Else
                txtDisplay.Text = txtDisplay.Text & btnNine.Text
            End If
        End If
    End Sub

    Private Sub btnZero_Click(sender As Object, e As EventArgs) Handles btnZero.Click
        If (blnIsFirstDigit) Then
            txtDisplay.Text = btnZero.Text
            blnIsFirstDigit = False
        Else
            If (txtDisplay.Text = "0") Then
                txtDisplay.Text = btnZero.Text
            Else
                txtDisplay.Text = txtDisplay.Text & btnZero.Text
            End If
        End If
    End Sub

    '****************************************************************************** 
    'When the decimal button is pressed it is added to the display and then the 
    'decimal button is disabled. 
    '******************************************************************************
    Private Sub btnDecimal_Click(sender As Object, e As EventArgs) Handles btnDecimal.Click
        If (txtDisplay.Text <> "0") Then
            txtDisplay.Text = txtDisplay.Text & btnDecimal.Text
        Else
            txtDisplay.Text = "0" & btnDecimal.Text
        End If
        btnDecimal.Enabled = False
    End Sub

    '****************************************************************************** 
    'This comment covers all operator buttons because they behave essentially the
    'same.  When an operator button is pressed the user can enter a new number and
    'the decimal button is enabled. If the user presses an operator in series, the
    'first calculation is performed and its results are shown in the display.  If
    'the user attempts to divide by zero, "ERROR" is displayed and the user must 
    'start a new calculation (ie all values stored are erased).
    '******************************************************************************
    Private Sub btnPlus_Click(sender As Object, e As EventArgs) Handles btnPlus.Click
        blnIsFirstDigit = True
        btnDecimal.Enabled = True

        If (decLeftOperand = Nothing And decRightOperand = Nothing) Then
            decLeftOperand = Convert.ToDecimal(txtDisplay.Text)
            strOperator = btnPlus.Text
            txtDisplay.Text = Convert.ToString(decLeftOperand)
        Else
            strLastOperator = strOperator
            decRightOperand = Convert.ToDecimal(txtDisplay.Text)

            If (decRightOperand = 0 And strLastOperator = "/") Then
                txtDisplay.Text = "ERROR"
                decRightOperand = Nothing
                decLeftOperand = Nothing
                strOperator = Nothing
            Else
                decLeftOperand = PerformCalculation(decLeftOperand, decRightOperand, strLastOperator)
                decRightOperand = Nothing
                strOperator = btnPlus.Text
                txtDisplay.Text = Convert.ToString(decLeftOperand)
            End If
        End If
    End Sub

    Private Sub btnMinus_Click(sender As Object, e As EventArgs) Handles btnMinus.Click
        blnIsFirstDigit = True
        btnDecimal.Enabled = True

        If (decLeftOperand = Nothing And decRightOperand = Nothing) Then
            decLeftOperand = Convert.ToDecimal(txtDisplay.Text)
            strOperator = btnMinus.Text
            txtDisplay.Text = Convert.ToString(decLeftOperand)
        Else
            strLastOperator = strOperator
            decRightOperand = Convert.ToDecimal(txtDisplay.Text)

            If (decRightOperand = 0 And strLastOperator = "/") Then
                txtDisplay.Text = "ERROR"
                decRightOperand = Nothing
                decLeftOperand = Nothing
                strOperator = Nothing
            Else
                decLeftOperand = PerformCalculation(decLeftOperand, decRightOperand, strLastOperator)
                decRightOperand = Nothing
                strOperator = btnMinus.Text
                txtDisplay.Text = Convert.ToString(decLeftOperand)
            End If
        End If
    End Sub

    Private Sub btnMultiply_Click(sender As Object, e As EventArgs) Handles btnMultiply.Click
        blnIsFirstDigit = True
        btnDecimal.Enabled = True

        If (decLeftOperand = Nothing And decRightOperand = Nothing) Then
            decLeftOperand = Convert.ToDecimal(txtDisplay.Text)
            strOperator = btnMultiply.Text
            txtDisplay.Text = Convert.ToString(decLeftOperand)
        Else
            strLastOperator = strOperator
            decRightOperand = Convert.ToDecimal(txtDisplay.Text)

            If (decRightOperand = 0 And strLastOperator = "/") Then
                txtDisplay.Text = "ERROR"
                decRightOperand = Nothing
                decLeftOperand = Nothing
                strOperator = Nothing
            Else
                decLeftOperand = PerformCalculation(decLeftOperand, decRightOperand, strLastOperator)
                decRightOperand = Nothing
                strOperator = btnMultiply.Text
                txtDisplay.Text = Convert.ToString(decLeftOperand)
            End If
        End If
    End Sub

    Private Sub btnDivide_Click(sender As Object, e As EventArgs) Handles btnDivide.Click
        blnIsFirstDigit = True
        btnDecimal.Enabled = True

        If (decLeftOperand = Nothing And decRightOperand = Nothing) Then
            decLeftOperand = Convert.ToDecimal(txtDisplay.Text)
            strOperator = btnDivide.Text
            txtDisplay.Text = Convert.ToString(decLeftOperand)
        Else
            strLastOperator = strOperator
            decRightOperand = Convert.ToDecimal(txtDisplay.Text)

            If (decRightOperand = 0 And strLastOperator = "/") Then
                txtDisplay.Text = "ERROR"
                decRightOperand = Nothing
                decLeftOperand = Nothing
                strOperator = Nothing
            Else
                decLeftOperand = PerformCalculation(decLeftOperand, decRightOperand, strLastOperator)
                decRightOperand = Nothing
                strOperator = btnDivide.Text
                txtDisplay.Text = Convert.ToString(decLeftOperand)
            End If
        End If
    End Sub

    '****************************************************************************** 
    'When the equal button is pressed the final calculation is performed and the
    'result is displayed.  All values stored in memory are erased.  
    '******************************************************************************
    Private Sub btnEqual_Click(sender As Object, e As EventArgs) Handles btnEqual.Click
        decRightOperand = Convert.ToDecimal(txtDisplay.Text)
        If (decRightOperand = 0 And strOperator = "/") Then
            txtDisplay.Text = "ERROR"
        ElseIf (decLeftOperand = Nothing) Then
            decResult = Convert.ToDecimal(txtDisplay.Text)
        Else
            decResult = PerformCalculation(decLeftOperand, decRightOperand, strOperator)
            txtDisplay.Text = Convert.ToString(decResult)
        End If

        decRightOperand = Nothing
        decLeftOperand = Nothing
        decResult = Nothing
        strOperator = Nothing
        blnIsFirstDigit = True

        btnDecimal.Enabled = True
    End Sub

    '****************************************************************************** 
    'Function which accepts two values and an operator and performs a function on 
    'the two values and returns a value.
    '******************************************************************************
    Function PerformCalculation(ByVal decOperand1 As Decimal, ByVal decOperand2 As Decimal, ByVal strOperator As String) As Decimal
        Dim decResult As Decimal

        Select Case strOperator
            Case "+"
                decResult = decOperand1 + decOperand2
            Case "-"
                decResult = decOperand1 - decOperand2
            Case "*"
                decResult = decOperand1 * decOperand2
            Case "/"
                decResult = decOperand1 / decOperand2
        End Select

        Return decResult
    End Function

End Class
