Public Class clsStudent
    '---------------------------------------------------------------------------------------
    '--- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES ---
    '--- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES ---
    '--- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES ---
    '---------------------------------------------------------------------------------------

    Public Property strInitials As String       'Holds the student's initials
    Public Property strLastName As String       'Holds the student's last name
    Public Property sngHomework1 As Single      'Holds the student's HW 1 score
    Public Property sngHomework2 As Single      'Holds the student's HW 2 score
    Public Property sngHomework3 As Single      'Holds the student's HW 3 score
    Public Property sngHomework4 As Single      'Holds the student's HW 4 score
    Public Property sngExamScore As Single      'Holds the student's exam score
    Public Property sngNumericGrade As Single   'Holds the student's numeric course score
    Public Property strLetterGrade As String    'Holds the student's course letter grade

    Public Sub New(ByVal Initials As String, ByVal LastName As String,
    ByVal Homework1 As Single, ByVal Homework2 As Single, ByVal Homework3 As Single,
    ByVal Homework4 As Single, ByVal ExamScore As Single, ByVal NumericGrade As Single,
    ByVal LetterGrade As String)
        '------------------------------------------------------------
        '-            Subprogram Name: New                          -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: February 11, 2021             -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- Constructor                                              -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- strInitials:     Holds the student's initials            -
        '- strLastName:     Holds the student's last name           -
        '- sngHomework1:    Holds the student's HW 1 score          -
        '- sngHomework2:    Holds the student's HW 1 score          -
        '- sngHomework3:    Holds the student's HW 1 score          -
        '- sngHomework4:    Holds the student's HW 1 score          -
        '- sngExamScore:    Holds the student's exam score          -
        '- sngNumericGrade: Holds the student's numeric course score-
        '- strLetterGrade:  Holds the student's course letter grade -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        'Assign parameter
        Me.strInitials = Initials
        Me.strLastName = LastName
        Me.sngHomework1 = Homework1
        Me.sngHomework2 = Homework2
        Me.sngHomework3 = Homework3
        Me.sngHomework4 = Homework4
        Me.sngExamScore = ExamScore
        Me.sngNumericGrade = NumericGrade
        Me.strLetterGrade = LetterGrade
    End Sub

    Public Overrides Function ToString() As String
        '------------------------------------------------------------
        '-            Subprogram Name: ToString                     -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: February 11, 2021             -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- Returns a nicely formatted String variable               -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        'Format and return string
        Return String.Format(" {0, 4} {1, -11} {2, 5}   {3, 5}   {4, 5}   {5, 5}  {6, 6}   {7, 6}     {8, -2}",
        strInitials, strLastName, Format(sngHomework1, "0.00"), Format(sngHomework2, "0.00"), Format(sngHomework3, "0.00"),
        Format(sngHomework4, "0.00"), Format(sngExamScore, "0.00"), Format(sngNumericGrade, "0.00"), strLetterGrade)
    End Function
End Class