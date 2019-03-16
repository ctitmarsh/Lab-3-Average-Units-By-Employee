Option Strict On
Public Class AverageUnitsShippedByEmployee
    ' Net Development - Lab 3
    ' By: Christian Titmarsh
    ' ID: 100527601
    ' This program is designed to use 2D arrays to store 7 values for 3 employees. 
    ' It will display each employees average plus the total average.
#Region "Variable and constant Declaration"


    Const MIN_VALUE As Integer = 0
    Const MAX_VALUE As Integer = 1000
    Const ERROR_MESSAGE As String = "Invalid input, must be between 0 and 1000, please try again"

    Dim units(2, 6) As Integer

    Dim day As Integer = 1
    Dim employeeNumber As Integer = 0

    Dim sum As Integer
    Dim total As Integer
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
#End Region

#Region "Functions and subs"

    ''' <summary>
    '''     Resets form to initial state
    ''' </summary>
    Sub resetForm()
        day = 1
        Array.Clear(units, 0, units.Length)

        txtUnits.Text = ""
        lblCurrentDay.Text = "Day 1"
        txtDataDisplay1.Text = ""
        txtDataDisplay2.Text = ""
        txtDataDisplay3.Text = ""
        lblAverage1.Text = ""
        lblAverage2.Text = ""
        lblAverage3.Text = ""
        lblDailyAverage.Text = ""

        btnEnter.Enabled = True
        txtUnits.ReadOnly = False

        txtUnits.Focus()
    End Sub

    ''' <summary>
    '''     Validates inputed user data
    ''' </summary>
    ''' <param name="input">User input to be validated</param>
    ''' <returns>Whether the input was valid or invalid as boolean</returns>
    Function validateInput(ByVal input As String) As Boolean
        Dim inputNumber As Integer
        Dim isValidInput As Boolean = False

        Try
            inputNumber = CInt(input)
            If (input.Equals(inputNumber.ToString)) Then
                If (inputNumber >= MIN_VALUE AndAlso inputNumber <= MAX_VALUE) Then
                    isValidInput = True
                End If
            End If
        Catch ex As Exception

        End Try

        Return isValidInput
    End Function

#End Region

#Region "Event Handlers"
    ''' <summary>
    '''     Handles enter button
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>


    Private Sub btnEnter_Click(sender As Object, e As EventArgs) Handles btnEnter.Click
        Dim userInput As String = txtUnits.Text
        userInput = txtUnits.Text
        If (employeeNumber < 7) Then
            If (day < 8) Then
                'validates the user input
                If (validateInput(userInput)) Then
                    'adds value to array
                    units(employeeNumber, day - 1) = Convert.ToInt32(userInput)
                    'adds value to sum of current employee
                    sum += units(employeeNumber, day - 1)
                    'adds value to the total sum used to calculate total average
                    total += units(employeeNumber, day - 1)
                Else
                    'Prompts user to enter new input
                    day = day - 1
                    lblCurrentDay.Text = "Day " + day.ToString
                    lblDailyAverage.Text = ERROR_MESSAGE

                End If
            End If

        End If

        If (employeeNumber = 0) Then
            'displays input in the first employee's data display textbox
            txtDataDisplay1.Text += userInput + vbCrLf
            'resets the textbox
            txtUnits.Text = ""
            'calculates the average and resets the sum
            If (day = 7) Then
                lblAverage1.Text = "Average: " + (Math.Round(sum / 7, 2)).ToString()
                sum = 0
                lblCurrentDay.Text = "Day 1"
            Else
                txtUnits.Text = ""
            End If

        ElseIf (employeeNumber = 1) Then
            'displays input in the second employee's data display textbox
            txtDataDisplay2.Text += userInput + vbCrLf
            'resets the textbox
            txtUnits.Text = ""
            'performs the average calculation and resets the sum 
            If (day = 7) Then
                lblAverage2.Text = "Average: " + (Math.Round(sum / 7, 2)).ToString()
                sum = 0
                lblCurrentDay.Text = "Day 1"
            End If

        ElseIf (employeeNumber = 2) Then
            'displays input in the third employee's data display textbox
            txtDataDisplay3.Text += userInput + vbCrLf
            'resets the textbox
            txtUnits.Text = ""
            'performs the average calculation and resets the sum
            If (day = 7) Then
                lblAverage3.Text = "Average: " + (Math.Round(sum / 7, 2)).ToString()
                sum = 0
                btnEnter.Enabled = False
                txtUnits.ReadOnly = True
                lblDailyAverage.Text = "Total Average: " + (Math.Round(total / 21, 2).ToString())
            End If
        End If
        'increment day by 1
        day += 1
        'If all values have been entered for current employee, continue to the next employee and reset day to 1
        If day = 8 Then
            employeeNumber += 1
            day = 1
        ElseIf day <= 7 Then
            lblCurrentDay.Text = "Day " + day.ToString()

        End If



        txtUnits.Focus()
        txtUnits.Text = ""

    End Sub
    ''' <summary>
    '''     Handles exit button
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Application.Exit()
    End Sub
    ''' <summary>
    '''     Handles reset button
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Call resetForm()
    End Sub


#End Region
End Class
