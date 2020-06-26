<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class SemesterAverageForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnReset = New System.Windows.Forms.Button()
        Me.btnCalculate = New System.Windows.Forms.Button()
        Me.lblResponseOutput = New System.Windows.Forms.Label()
        Me.lblForCourseOneGrade = New System.Windows.Forms.Label()
        Me.lblForCourseTwoGrade = New System.Windows.Forms.Label()
        Me.lblForCourseThreeGrade = New System.Windows.Forms.Label()
        Me.lblForCourseFourGrade = New System.Windows.Forms.Label()
        Me.lblForCourseFiveGrade = New System.Windows.Forms.Label()
        Me.lblForCourseSixGrade = New System.Windows.Forms.Label()
        Me.lblForSemester = New System.Windows.Forms.Label()
        Me.txtCourseOneGrade = New System.Windows.Forms.TextBox()
        Me.txtCourseTwoGrade = New System.Windows.Forms.TextBox()
        Me.txtCourseThreeGrade = New System.Windows.Forms.TextBox()
        Me.txtCourseFourGrade = New System.Windows.Forms.TextBox()
        Me.txtCourseFiveGrade = New System.Windows.Forms.TextBox()
        Me.txtCourseSixGrade = New System.Windows.Forms.TextBox()
        Me.lblSemesterAverage = New System.Windows.Forms.Label()
        Me.lblCourseOneLetterGrade = New System.Windows.Forms.Label()
        Me.lblCourseTwoLetterGrade = New System.Windows.Forms.Label()
        Me.lblCourseThreeLetterGrade = New System.Windows.Forms.Label()
        Me.lblCourseFourLetterGrade = New System.Windows.Forms.Label()
        Me.lblCourseFiveLetterGrade = New System.Windows.Forms.Label()
        Me.lblCourseSixLetterGrade = New System.Windows.Forms.Label()
        Me.lblSemesterLetterGrade = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(206, 340)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 24
        Me.btnExit.Text = "E&xit"
        Me.ToolTip1.SetToolTip(Me.btnExit, "Click this button, click the 'x' button in the upper righthand corner, or press t" &
        "he 'x' button on the keyboard to exit.")
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnReset
        '
        Me.btnReset.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnReset.Location = New System.Drawing.Point(125, 340)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(75, 23)
        Me.btnReset.TabIndex = 23
        Me.btnReset.Text = "&Reset"
        Me.ToolTip1.SetToolTip(Me.btnReset, "Click this button, press the 'esc' button, or press the 'r' button to reset the f" &
        "orm.")
        Me.btnReset.UseVisualStyleBackColor = True
        '
        'btnCalculate
        '
        Me.btnCalculate.Location = New System.Drawing.Point(44, 340)
        Me.btnCalculate.Name = "btnCalculate"
        Me.btnCalculate.Size = New System.Drawing.Size(75, 23)
        Me.btnCalculate.TabIndex = 22
        Me.btnCalculate.Text = "&Calculate"
        Me.ToolTip1.SetToolTip(Me.btnCalculate, "Click this button, press the 'enter' button, or press the 'c' button on your keyb" &
        "oard to calculate an average grade for the semester.")
        Me.btnCalculate.UseVisualStyleBackColor = True
        '
        'lblResponseOutput
        '
        Me.lblResponseOutput.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblResponseOutput.Location = New System.Drawing.Point(2, 229)
        Me.lblResponseOutput.Name = "lblResponseOutput"
        Me.lblResponseOutput.Size = New System.Drawing.Size(279, 108)
        Me.lblResponseOutput.TabIndex = 21
        '
        'lblForCourseOneGrade
        '
        Me.lblForCourseOneGrade.AutoSize = True
        Me.lblForCourseOneGrade.Location = New System.Drawing.Point(12, 19)
        Me.lblForCourseOneGrade.Name = "lblForCourseOneGrade"
        Me.lblForCourseOneGrade.Size = New System.Drawing.Size(52, 13)
        Me.lblForCourseOneGrade.TabIndex = 0
        Me.lblForCourseOneGrade.Text = "Course 1:"
        Me.lblForCourseOneGrade.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblForCourseTwoGrade
        '
        Me.lblForCourseTwoGrade.AutoSize = True
        Me.lblForCourseTwoGrade.Location = New System.Drawing.Point(12, 48)
        Me.lblForCourseTwoGrade.Name = "lblForCourseTwoGrade"
        Me.lblForCourseTwoGrade.Size = New System.Drawing.Size(52, 13)
        Me.lblForCourseTwoGrade.TabIndex = 3
        Me.lblForCourseTwoGrade.Text = "Course 2:"
        Me.lblForCourseTwoGrade.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblForCourseThreeGrade
        '
        Me.lblForCourseThreeGrade.AutoSize = True
        Me.lblForCourseThreeGrade.Location = New System.Drawing.Point(12, 77)
        Me.lblForCourseThreeGrade.Name = "lblForCourseThreeGrade"
        Me.lblForCourseThreeGrade.Size = New System.Drawing.Size(52, 13)
        Me.lblForCourseThreeGrade.TabIndex = 6
        Me.lblForCourseThreeGrade.Text = "Course 3:"
        Me.lblForCourseThreeGrade.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblForCourseFourGrade
        '
        Me.lblForCourseFourGrade.AutoSize = True
        Me.lblForCourseFourGrade.Location = New System.Drawing.Point(12, 106)
        Me.lblForCourseFourGrade.Name = "lblForCourseFourGrade"
        Me.lblForCourseFourGrade.Size = New System.Drawing.Size(52, 13)
        Me.lblForCourseFourGrade.TabIndex = 9
        Me.lblForCourseFourGrade.Text = "Course 4:"
        Me.lblForCourseFourGrade.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblForCourseFiveGrade
        '
        Me.lblForCourseFiveGrade.AutoSize = True
        Me.lblForCourseFiveGrade.Location = New System.Drawing.Point(12, 135)
        Me.lblForCourseFiveGrade.Name = "lblForCourseFiveGrade"
        Me.lblForCourseFiveGrade.Size = New System.Drawing.Size(52, 13)
        Me.lblForCourseFiveGrade.TabIndex = 12
        Me.lblForCourseFiveGrade.Text = "Course 5:"
        Me.lblForCourseFiveGrade.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblForCourseSixGrade
        '
        Me.lblForCourseSixGrade.AutoSize = True
        Me.lblForCourseSixGrade.Location = New System.Drawing.Point(12, 164)
        Me.lblForCourseSixGrade.Name = "lblForCourseSixGrade"
        Me.lblForCourseSixGrade.Size = New System.Drawing.Size(52, 13)
        Me.lblForCourseSixGrade.TabIndex = 15
        Me.lblForCourseSixGrade.Text = "Course 6:"
        Me.lblForCourseSixGrade.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblForSemester
        '
        Me.lblForSemester.AutoSize = True
        Me.lblForSemester.Location = New System.Drawing.Point(12, 193)
        Me.lblForSemester.Name = "lblForSemester"
        Me.lblForSemester.Size = New System.Drawing.Size(54, 13)
        Me.lblForSemester.TabIndex = 18
        Me.lblForSemester.Text = "Semester:"
        Me.lblForSemester.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtCourseOneGrade
        '
        Me.txtCourseOneGrade.Location = New System.Drawing.Point(71, 16)
        Me.txtCourseOneGrade.Name = "txtCourseOneGrade"
        Me.txtCourseOneGrade.Size = New System.Drawing.Size(100, 20)
        Me.txtCourseOneGrade.TabIndex = 1
        Me.txtCourseOneGrade.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCourseOneGrade, "Please enter a grade between 0 and 100.")
        '
        'txtCourseTwoGrade
        '
        Me.txtCourseTwoGrade.Location = New System.Drawing.Point(71, 45)
        Me.txtCourseTwoGrade.Name = "txtCourseTwoGrade"
        Me.txtCourseTwoGrade.Size = New System.Drawing.Size(100, 20)
        Me.txtCourseTwoGrade.TabIndex = 4
        Me.txtCourseTwoGrade.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCourseTwoGrade, "Please enter a grade between 0 and 100.")
        '
        'txtCourseThreeGrade
        '
        Me.txtCourseThreeGrade.Location = New System.Drawing.Point(71, 74)
        Me.txtCourseThreeGrade.Name = "txtCourseThreeGrade"
        Me.txtCourseThreeGrade.Size = New System.Drawing.Size(100, 20)
        Me.txtCourseThreeGrade.TabIndex = 7
        Me.txtCourseThreeGrade.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCourseThreeGrade, "Please enter a grade between 0 and 100.")
        '
        'txtCourseFourGrade
        '
        Me.txtCourseFourGrade.Location = New System.Drawing.Point(71, 103)
        Me.txtCourseFourGrade.Name = "txtCourseFourGrade"
        Me.txtCourseFourGrade.Size = New System.Drawing.Size(100, 20)
        Me.txtCourseFourGrade.TabIndex = 10
        Me.txtCourseFourGrade.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCourseFourGrade, "Please enter a grade between 0 and 100.")
        '
        'txtCourseFiveGrade
        '
        Me.txtCourseFiveGrade.Location = New System.Drawing.Point(71, 132)
        Me.txtCourseFiveGrade.Name = "txtCourseFiveGrade"
        Me.txtCourseFiveGrade.Size = New System.Drawing.Size(100, 20)
        Me.txtCourseFiveGrade.TabIndex = 13
        Me.txtCourseFiveGrade.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCourseFiveGrade, "Please enter a grade between 0 and 100.")
        '
        'txtCourseSixGrade
        '
        Me.txtCourseSixGrade.Location = New System.Drawing.Point(71, 161)
        Me.txtCourseSixGrade.Name = "txtCourseSixGrade"
        Me.txtCourseSixGrade.Size = New System.Drawing.Size(100, 20)
        Me.txtCourseSixGrade.TabIndex = 16
        Me.txtCourseSixGrade.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtCourseSixGrade, "Please enter a grade between 0 and 100.")
        '
        'lblSemesterAverage
        '
        Me.lblSemesterAverage.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSemesterAverage.Location = New System.Drawing.Point(72, 192)
        Me.lblSemesterAverage.Name = "lblSemesterAverage"
        Me.lblSemesterAverage.Size = New System.Drawing.Size(98, 20)
        Me.lblSemesterAverage.TabIndex = 19
        Me.lblSemesterAverage.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.lblSemesterAverage, "The percentage grade for the overall average is displayed here.")
        '
        'lblCourseOneLetterGrade
        '
        Me.lblCourseOneLetterGrade.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCourseOneLetterGrade.Location = New System.Drawing.Point(177, 16)
        Me.lblCourseOneLetterGrade.Name = "lblCourseOneLetterGrade"
        Me.lblCourseOneLetterGrade.Size = New System.Drawing.Size(98, 20)
        Me.lblCourseOneLetterGrade.TabIndex = 2
        Me.lblCourseOneLetterGrade.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ToolTip1.SetToolTip(Me.lblCourseOneLetterGrade, "The letter grade for course 1 is displayed here.")
        '
        'lblCourseTwoLetterGrade
        '
        Me.lblCourseTwoLetterGrade.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCourseTwoLetterGrade.Location = New System.Drawing.Point(177, 45)
        Me.lblCourseTwoLetterGrade.Name = "lblCourseTwoLetterGrade"
        Me.lblCourseTwoLetterGrade.Size = New System.Drawing.Size(98, 20)
        Me.lblCourseTwoLetterGrade.TabIndex = 5
        Me.lblCourseTwoLetterGrade.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ToolTip1.SetToolTip(Me.lblCourseTwoLetterGrade, "The letter grade for course 2 is displayed here.")
        '
        'lblCourseThreeLetterGrade
        '
        Me.lblCourseThreeLetterGrade.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCourseThreeLetterGrade.Location = New System.Drawing.Point(177, 74)
        Me.lblCourseThreeLetterGrade.Name = "lblCourseThreeLetterGrade"
        Me.lblCourseThreeLetterGrade.Size = New System.Drawing.Size(98, 20)
        Me.lblCourseThreeLetterGrade.TabIndex = 8
        Me.lblCourseThreeLetterGrade.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ToolTip1.SetToolTip(Me.lblCourseThreeLetterGrade, "The letter grade for course 3 is displayed here.")
        '
        'lblCourseFourLetterGrade
        '
        Me.lblCourseFourLetterGrade.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCourseFourLetterGrade.Location = New System.Drawing.Point(177, 103)
        Me.lblCourseFourLetterGrade.Name = "lblCourseFourLetterGrade"
        Me.lblCourseFourLetterGrade.Size = New System.Drawing.Size(98, 20)
        Me.lblCourseFourLetterGrade.TabIndex = 11
        Me.lblCourseFourLetterGrade.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ToolTip1.SetToolTip(Me.lblCourseFourLetterGrade, "The letter grade for course 4 is displayed here.")
        '
        'lblCourseFiveLetterGrade
        '
        Me.lblCourseFiveLetterGrade.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCourseFiveLetterGrade.Location = New System.Drawing.Point(177, 132)
        Me.lblCourseFiveLetterGrade.Name = "lblCourseFiveLetterGrade"
        Me.lblCourseFiveLetterGrade.Size = New System.Drawing.Size(98, 20)
        Me.lblCourseFiveLetterGrade.TabIndex = 14
        Me.lblCourseFiveLetterGrade.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ToolTip1.SetToolTip(Me.lblCourseFiveLetterGrade, "The letter grade for course 5 is displayed here.")
        '
        'lblCourseSixLetterGrade
        '
        Me.lblCourseSixLetterGrade.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCourseSixLetterGrade.Location = New System.Drawing.Point(177, 161)
        Me.lblCourseSixLetterGrade.Name = "lblCourseSixLetterGrade"
        Me.lblCourseSixLetterGrade.Size = New System.Drawing.Size(98, 20)
        Me.lblCourseSixLetterGrade.TabIndex = 17
        Me.lblCourseSixLetterGrade.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ToolTip1.SetToolTip(Me.lblCourseSixLetterGrade, "The letter grade for course 6 is displayed here.")
        '
        'lblSemesterLetterGrade
        '
        Me.lblSemesterLetterGrade.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSemesterLetterGrade.Location = New System.Drawing.Point(177, 192)
        Me.lblSemesterLetterGrade.Name = "lblSemesterLetterGrade"
        Me.lblSemesterLetterGrade.Size = New System.Drawing.Size(98, 20)
        Me.lblSemesterLetterGrade.TabIndex = 20
        Me.lblSemesterLetterGrade.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ToolTip1.SetToolTip(Me.lblSemesterLetterGrade, "The letter grade for the overall average is displayed here.")
        '
        'SemesterAverageForm
        '
        Me.AcceptButton = Me.btnCalculate
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnReset
        Me.ClientSize = New System.Drawing.Size(284, 365)
        Me.Controls.Add(Me.lblSemesterLetterGrade)
        Me.Controls.Add(Me.lblCourseSixLetterGrade)
        Me.Controls.Add(Me.lblCourseFiveLetterGrade)
        Me.Controls.Add(Me.lblCourseFourLetterGrade)
        Me.Controls.Add(Me.lblCourseThreeLetterGrade)
        Me.Controls.Add(Me.lblCourseTwoLetterGrade)
        Me.Controls.Add(Me.lblCourseOneLetterGrade)
        Me.Controls.Add(Me.lblSemesterAverage)
        Me.Controls.Add(Me.txtCourseSixGrade)
        Me.Controls.Add(Me.txtCourseFiveGrade)
        Me.Controls.Add(Me.txtCourseFourGrade)
        Me.Controls.Add(Me.txtCourseThreeGrade)
        Me.Controls.Add(Me.txtCourseTwoGrade)
        Me.Controls.Add(Me.txtCourseOneGrade)
        Me.Controls.Add(Me.lblForSemester)
        Me.Controls.Add(Me.lblForCourseSixGrade)
        Me.Controls.Add(Me.lblForCourseFiveGrade)
        Me.Controls.Add(Me.lblForCourseFourGrade)
        Me.Controls.Add(Me.lblForCourseThreeGrade)
        Me.Controls.Add(Me.lblForCourseTwoGrade)
        Me.Controls.Add(Me.lblForCourseOneGrade)
        Me.Controls.Add(Me.lblResponseOutput)
        Me.Controls.Add(Me.btnCalculate)
        Me.Controls.Add(Me.btnReset)
        Me.Controls.Add(Me.btnExit)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(300, 404)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(300, 404)
        Me.Name = "SemesterAverageForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Semester Average"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnExit As Button
    Friend WithEvents btnReset As Button
    Friend WithEvents btnCalculate As Button
    Friend WithEvents lblResponseOutput As Label
    Friend WithEvents lblForCourseOneGrade As Label
    Friend WithEvents lblForCourseTwoGrade As Label
    Friend WithEvents lblForCourseThreeGrade As Label
    Friend WithEvents lblForCourseFourGrade As Label
    Friend WithEvents lblForCourseFiveGrade As Label
    Friend WithEvents lblForCourseSixGrade As Label
    Friend WithEvents lblForSemester As Label
    Friend WithEvents txtCourseOneGrade As TextBox
    Friend WithEvents txtCourseTwoGrade As TextBox
    Friend WithEvents txtCourseThreeGrade As TextBox
    Friend WithEvents txtCourseFourGrade As TextBox
    Friend WithEvents txtCourseFiveGrade As TextBox
    Friend WithEvents txtCourseSixGrade As TextBox
    Friend WithEvents lblSemesterAverage As Label
    Friend WithEvents lblCourseOneLetterGrade As Label
    Friend WithEvents lblCourseTwoLetterGrade As Label
    Friend WithEvents lblCourseThreeLetterGrade As Label
    Friend WithEvents lblCourseFourLetterGrade As Label
    Friend WithEvents lblCourseFiveLetterGrade As Label
    Friend WithEvents lblCourseSixLetterGrade As Label
    Friend WithEvents lblSemesterLetterGrade As Label
    Friend WithEvents ToolTip1 As ToolTip
End Class
