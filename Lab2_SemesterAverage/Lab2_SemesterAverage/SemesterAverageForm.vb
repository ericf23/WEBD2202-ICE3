'Name: Eric
'Date: June 13th, 2020
'Program: Semester Average Form
'Program Description: This program assigns letter grades to percentages entered for courses, and then calculates an overall average and letter grade.

'Option Strict On

Public Class SemesterAverageForm

#Region "Constants and Variables"
    ''' <summary>
    ''' This region stores the variables And constants For the program.
    ''' </summary>
    Private Const MIN_GRADE As Integer = 0
    Private Const MAX_GRADE As Integer = 100

    Dim inputTextboxList As TextBox()
    Dim outputLabelList As Label()
    Dim gradeTotal As Double = 0


#End Region

#Region "Functions and Subroutines"

    ''' <summary>
    ''' This clears the text property of all controls in the array of controls that is passed in
    ''' </summary>
    ''' <param name="controlArray">An array of controls with a text property to clear</param>
    Sub ClearControls(controlArray As Control())

        ' For every control in the list that is passed in, empty its Text property
        For Each controlToClear As Control In controlArray
            controlToClear.Text = String.Empty
        Next

    End Sub

    ''' <summary>
    ''' This enables or disables all textboxes in the array that is passed in
    ''' </summary>
    ''' <param name="textboxArray">An array of textboxes to disable</param>
    ''' <param name="enabledStatus">Boolean: set textboxes to enabled?</param>
    Sub SetTextboxesEnabled(textboxArray As TextBox(), enabledStatus As Boolean)

        ' For every textbox in the list that is passed in, set its Enabled property
        For Each textboxToSet As TextBox In textboxArray
            textboxToSet.Enabled = enabledStatus
        Next

    End Sub

    ''' <summary>
    ''' The instructions to set the form back to default
    ''' </summary>
    Sub SetDefaults()

        'Clear input and output fields
        ClearControls(inputTextboxList) 'Clear input textboxes
        ClearControls(outputLabelList) 'Clear output labels
        lblSemesterAverage.Text = String.Empty 'Clear overall average box
        lblSemesterLetterGrade.Text = String.Empty 'Clear overall average letter grade label
        lblResponseOutput.Text = String.Empty 'Clear error output
        'Re-enable
        SetTextboxesEnabled(inputTextboxList, True)
        btnCalculate.Enabled = True

        'Set focus on first textbox
        txtCourseOneGrade.Focus()
    End Sub

    ''' <summary>
    ''' This function validates individual textboxes
    ''' </summary>
    ''' <param name="txtInput"></param>
    Function ValidateTextbox(txtInput As TextBox)
        'Declare variable
        Dim inputGrade As Double

        'if entry can be parsed to a double, check for validation range, if in range, grade is added to total. If either validation is false, error
        If Double.TryParse(txtInput.Text, inputGrade) Then
            If inputGrade >= MIN_GRADE AndAlso inputGrade <= MAX_GRADE Then
                gradeTotal += inputGrade
                Return True
            Else
                lblResponseOutput.Text &= "Please ensure that your input a number between 0 and 100." & vbCrLf
                txtInput.SelectAll()
                txtInput.Focus()
                Return False
            End If

        Else
            lblResponseOutput.Text &= "Please ensure that your input a number between 0 and 100." & vbCrLf
            txtInput.SelectAll()
            txtInput.Focus()
            Return False
        End If

    End Function

    ''' <summary>
    ''' Checks validity of all textboxes in array.
    ''' </summary>
    ''' <param name="textboxArray"></param>
    ''' <returns></returns>
    Function ValidateTextboxes(textboxArray As TextBox())

        'Declare variable
        Dim isValid As Boolean = True

        'For each item in array check it's validation
        For Each textboxToCheck As TextBox In textboxArray

            isValid = isValid And ValidateTextbox(textboxToCheck)
        Next

        Return isValid

    End Function

    ''' <summary>
    ''' This function assigns a letter grade to the input textboxes using arrays
    ''' </summary>
    ''' <returns></returns>
    Function GetLetterGrade(grade As Double) As String
        'variable declarations 
        Dim gradeThresholds As Double() = {0, 50, 55, 60, 65, 70, 75, 80, 85, 90}
        Dim letterGrades As String() = {"F", "D", "D+", "C", "B-", "B", "B+", "A-", "A", "A+"}
        Dim letterGradeOutput As String = String.Empty

        'Count through array of grade thresholds
        For assignedGrade As Double = 0 To gradeThresholds.Length - 1
            If grade >= gradeThresholds(assignedGrade) Then
                letterGradeOutput = letterGrades(assignedGrade)
            End If
        Next
        Return letterGradeOutput
    End Function
#End Region

#Region "Event Handlers"
    ''' <summary>
    ''' When the form loads, values are assigned to arrays. One array is for the user inputs, and another for output labels
    ''' </summary>
    Private Sub SemesterAverageForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        inputTextboxList = {txtCourseOneGrade, txtCourseTwoGrade, txtCourseThreeGrade, txtCourseFourGrade, txtCourseFiveGrade, txtCourseSixGrade}
        outputLabelList = {lblCourseOneLetterGrade, lblCourseTwoLetterGrade, lblCourseThreeLetterGrade, lblCourseThreeLetterGrade, lblCourseFourLetterGrade, lblCourseFiveLetterGrade, lblCourseSixLetterGrade}
    End Sub

    ''' <summary>
    ''' When the calculate button is clicked, calculation is performed. Average grade and it's assigned letter grade appear
    ''' </summary>
    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        'Declare variable
        Dim averageGrade As Double = 0
        'Check if the textboxes are valid
        If ValidateTextboxes(inputTextboxList) Then

            'Clear error message
            lblResponseOutput.Text = String.Empty

            'Calculate average grade for semester, rounded to two decimals
            averageGrade = Math.Round(gradeTotal / inputTextboxList.Length, 2)

            'Display results
            lblSemesterAverage.Text = averageGrade


            'Uses if statements to assign letter grade to overall average
            If averageGrade >= 90 AndAlso averageGrade <= 100 Then
                lblSemesterLetterGrade.Text = "A+"
            ElseIf averageGrade >= 85 AndAlso averageGrade < 90 Then
                lblSemesterLetterGrade.Text = "A"
            ElseIf averageGrade >= 80 AndAlso averageGrade < 85 Then
                lblSemesterLetterGrade.Text = "A-"
            ElseIf averageGrade >= 75 AndAlso averageGrade < 80 Then
                lblSemesterLetterGrade.Text = "B+"
            ElseIf averageGrade >= 70 AndAlso averageGrade < 75 Then
                lblSemesterLetterGrade.Text = "B"
            ElseIf averageGrade >= 65 AndAlso averageGrade < 70 Then
                lblSemesterLetterGrade.Text = "B-"
            ElseIf averageGrade >= 60 AndAlso averageGrade < 65 Then
                lblSemesterLetterGrade.Text = "C"
            ElseIf averageGrade >= 55 AndAlso averageGrade < 60 Then
                lblSemesterLetterGrade.Text = "D+"
            ElseIf averageGrade >= 50 AndAlso averageGrade < 55 Then
                lblSemesterLetterGrade.Text = "D"
            Else
                averageGrade = "F"


            End If

            'Disable calculate button
            btnCalculate.Enabled = False
        End If

    End Sub

    ''' <summary>
    ''' When reset button is activated, SetDefaults() is called
    ''' </summary>
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        SetDefaults()
    End Sub

    ''' <summary>
    ''' When exit button is hit, form closes
    ''' </summary>
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    ''' <summary>
    ''' When focus is lost on textbox, letter grade is assigned
    ''' </summary>
    Private Sub TextBoxLostFocus(sender As Object, e As EventArgs) Handles txtCourseOneGrade.LostFocus, txtCourseTwoGrade.LostFocus, txtCourseThreeGrade.LostFocus, txtCourseFourGrade.LostFocus, txtCourseFiveGrade.LostFocus, txtCourseSixGrade.LostFocus
        For controlIndex As Integer = 0 To inputTextboxList.Length - 1
            Dim inputGrade As Double = 0
            If Double.TryParse(inputTextboxList(controlIndex).Text, inputGrade) Then
                outputLabelList(controlIndex).Text = GetLetterGrade(inputGrade)
            End If
        Next
    End Sub

#End Region

End Class
