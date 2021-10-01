Imports System.IO
Module Module1
    '------------------------------------------------------------
    '-                File Name : Module1.vb                    - 
    '-                Part of Project: Assignment5              -
    '------------------------------------------------------------
    '-                Written By: Benjamin Neeb                 -
    '-                Written On: February 15, 2021              -
    '------------------------------------------------------------
    '- File Purpose:                                            -
    '-                                                          -
    '- This file contains the main Sub for the console          -
    '- application. The user will input all their data in this  - 
    '- file. The error handling, file I/O, calculations, and    -
    '- output are performed in this file.                       -
    '------------------------------------------------------------
    '- Program Purpose:                                         -
    '-                                                          -
    '- This program gathers all text from a text file of the    -
    '- user's choice. It then performs statistical analysis on  -
    '- the student grades and enformation contained in the file.-
    '- The program uses LINQ to then gather relevant            -
    '- information about the student's and how they perfromed.  -
    '- All output is presented on the console window.           -
    '------------------------------------------------------------
    '- Global Variable Dictionary (alphabetically):             -
    '- (None)                                                   –
    '------------------------------------------------------------

    '---------------------------------------------------------------------------------------
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '---------------------------------------------------------------------------------------

    Const intDUPNUMBER As Integer = 50              'Integer for repeat count
    Const intHWPOSSIBLE As Integer = 25             'Integer for possible points on a homework assignment
    Const intEXAMPOSSIBLE As Integer = 100          'Integer for possible points on an exam

    '-----------------------------------------------------------------------------------
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '-----------------------------------------------------------------------------------

    Sub Main()
        '------------------------------------------------------------
        '-            Subprogram Name: Main                         -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: February 11, 2021             -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine sets the console's attributesm, gathers  -
        '- all input from the user, and terminates the program. No  -
        '- calculations, file I/O, or error handling is performed   -
        '- in this Sub.                                             -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- strAnalysis:     String variable which holds the entire  -
        '-                  statistical analysis that was performed -
        '- strFileContents: String variable which holds the entire  -
        '-                  contents of the user's file chosen to   -
        '-                  be statistically analyzed.              -
        '- lstStudents:     List (of clsStudents) that contains the -
        '-                  list of students that were read in from -
        '-                  the text file.                          -
        '------------------------------------------------------------

        'Ask user for path of file
        Console.WriteLine("Please enter the path and name of the file to process:")

        'Get the contents of the file and store in a string
        Dim strFileContents As String = GetFile(Console.ReadLine())

        'If the read did not fail (returned a string), continue program
        If Not (strFileContents.Equals("")) Then
            'Create arraylist for lines in file
            Dim lstStudents As New List(Of clsStudent)

            'Populate the list
            GetLines(strFileContents, lstStudents)

            'Write the report section
            WriteReport(lstStudents)

            'Get the distribution statistics
            GetDistStats(lstStudents)

            'Get the range statistics
            GetRangeStats(lstStudents)

            'Get the overall statistics
            GetOveralStats(lstStudents)

            'Simply adding a line of space
            Console.WriteLine()
            Console.WriteLine("Application has completed. Press any key to end.")
        End If

        'Ending program here
        Console.ReadKey()
    End Sub

    Private Function GetFile(strPath As String) As String
        '------------------------------------------------------------
        '-                Function Name: GetFile                    -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: January 28, 2021              -
        '------------------------------------------------------------
        '- Function Purpose:                                        -
        '-                                                          -
        '- This function accepts the file path that the user wants  -
        '- to read from. The function then ensures that the file    -
        '- be successfully read from without errors. If not, the    -
        '- function takes care of error handling and will return an -
        '- empty String. If the file read is successful, the        -
        '- function reads the file's contents into a String         -
        '- variable and returns the String.                         -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- strPath:         String variable that contains the path  -
        '-                  of the file the user wants to read from -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- strContents:     String variable that contains the       -
        '-                  contents of the read file.              -
        '------------------------------------------------------------
        '- Returns:                                                 -
        '- String:          Contains the file contents              -
        '------------------------------------------------------------

        'String for file contents
        Dim strContents As String = ""

        'Must get data from an existing text file as outlined in assignment constraints
        If Not (strPath.EndsWith(".txt") Or File.Exists(strPath)) Then
            Console.WriteLine()
            Console.WriteLine(StrDup(intDUPNUMBER, "*"))
            Console.WriteLine("*** Must input data from a text (.txt) file! ***")
            Console.WriteLine(StrDup(intDUPNUMBER, "*"))
            Console.WriteLine()
            Console.WriteLine("*** Application will exit -- press any key... ***")
        Else
            Try
                'Read the contents of the file into strContents
                strContents = File.ReadAllText(strPath)
            Catch ex As Exception
                'If this failed, alert user
                Console.WriteLine()
                Console.WriteLine(StrDup(intDUPNUMBER, "*"))
                Console.WriteLine("*** Could not open input file for processing! ***")
                Console.WriteLine(StrDup(intDUPNUMBER, "*"))
                Console.WriteLine()
                Console.WriteLine("*** Application will exit -- press any key... ***")
            End Try
        End If

        'Return strContents and cast to uppercase. It will either contain the contents of the file or will equal ""
        Return strContents.ToUpper
    End Function

    Private Sub GetLines(strContents As String, ByRef lstStudents As List(Of clsStudent))
        '------------------------------------------------------------
        '-            Subprogram Name: GetLines                     -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: February 11, 2021             -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine reads the text from the String variable  -
        '- and populates the list with the students' information.   -
        '- This information includes their initials, lastname,      -
        '- homework scores, exam score, overall numeric score, and  -
        '- calculated letter score.                                 -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- strContents:     String type variable that contains the  -
        '-                  file's contents                         -
        '- lstStudents:     List (of clsStudent) that will be       -
        '-                  populated with student information      -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- sngExam:     Single type variable that holds exam score  -
        '- sngGrade:    Single type variable that holds the         -
        '-              calculated score for entire course          -
        '- sngHW1:      Single type variable that holds HW1's score -
        '- sngHW2:      Single type variable that holds HW2's score -
        '- sngHW3:      Single type variable that holds HW3's score -
        '- sngHW4:      Single type variable that holds HW4's score -
        '- strLetter:   String type variable that holds the         -
        '-              calculated letter grade for each student    -
        '- sr:          String reader used to read text contents    -
        '- arrString:   String array used to hold each row of the   -
        '-              text file.                                  -
        '------------------------------------------------------------

        'Create a StringReader to read the lines from the path text file
        Dim sr As StringReader = New StringReader(strContents)

        'Read until the end of the file
        Do While sr.Peek() >= 0
            'String split the file into a String array
            Dim arrString() As String = sr.ReadLine.Split(" ")

            Dim sngHW1 As Single = CSng(arrString(2))

            Dim sngHW2 As Single = CSng(arrString(3))

            Dim sngHW3 As Single = CSng(arrString(4))

            Dim sngHW4 As Single = CSng(arrString(5))

            Dim sngExam As Single = CSng(arrString(6))

            Dim sngGrade As Single = CSng(Math.Round(((sngHW1 + sngHW2 + sngHW3 + sngHW4) * 0.5) + (sngExam * 0.5), 2))

            Dim strLetter As String
            Select Case sngGrade
                Case >= 95  'A
                    strLetter = "A"
                Case >= 90  'A-
                    strLetter = "A-"
                Case >= 87  'B+
                    strLetter = "B+"
                Case >= 84  'B
                    strLetter = "B"
                Case >= 80  'B-
                    strLetter = "B-"
                Case >= 77  'C+
                    strLetter = "C+"
                Case >= 74  'C
                    strLetter = "C"
                Case >= 70  'C-
                    strLetter = "C-"
                Case >= 67  'D+
                    strLetter = "D+"
                Case >= 64  'D
                    strLetter = "D"
                Case >= 60  'D-
                    strLetter = "D-"
                Case Else
                    strLetter = "F"
            End Select

            'Add the student to arraylist
            lstStudents.Add(New clsStudent(arrString(0), arrString(1), sngHW1, sngHW2, sngHW3, sngHW4, sngExam, sngGrade, strLetter))
        Loop
    End Sub

    Private Sub WriteReport(lstStudents As List(Of clsStudent))
        '------------------------------------------------------------
        '-            Subprogram Name: WriteReport                  -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: February 11, 2021             -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine writes each student's report nicely      -
        '- formatted to the console. The report includes the        -
        '- students' performance information such as scores and     -
        '- letter grades.                                           -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- lstStudents:     List (of clsStudent) that is populated  -
        '-                  with student information                -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- objStudents:     Object type variable used with LINQ to  -
        '-                  perform analysis on lstStudent's        -
        '-                  contents.                               -
        '------------------------------------------------------------

        'Write header
        Console.WriteLine(vbCrLf & String.Format("{0, 49}", "Ye Old Country School"))
        Console.WriteLine(String.Format("{0, 53}", "*** Semester Grade Report ***"))
        Console.WriteLine(String.Format("{0, 53}", StrDup(29, "-")) & vbCrLf)
        Console.WriteLine(String.Format("{0, 40}{1, 14}{2, 10}{3, 8}", "Homework Scores", "Exam", "Numeric", "Letter"))
        Console.WriteLine(String.Format("{0, 9}{1, 12}{2, 8}{3, 8}{4, 8}{5, 10}{6, 8}{7, 9}", "Name", "1", "2", "3", "4", "Score", "Grade", "Grade"))
        Console.WriteLine(String.Format("{0, 14}{1, 9}{2, 8}{3, 8}{4, 8}{5, 8}{6, 9}{7, 8}",
        StrDup(14, "-"), StrDup(5, "-"), StrDup(5, "-"), StrDup(5, "-"), StrDup(5, "-"), StrDup(6, "-"), StrDup(7, "-"), StrDup(6, "-")))

        'Sort report
        Dim objStudents As Object
        objStudents = From students In lstStudents
                      Order By students.strLastName 'Ordering result set by students' last name
                      Select students

        'Write report body
        For Each student In objStudents
            Console.WriteLine(student.ToString)
        Next
    End Sub

    Private Sub GetDistStats(lstStudents As List(Of clsStudent))
        '------------------------------------------------------------
        '-            Subprogram Name: GetDistStats                 -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: February 11, 2021             -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine calls other Subs to get the students'    -
        '- grade distribution statistics, which then nicely write   -
        '- the results out to the console.                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- lstStudents:     List (of clsStudent) that is populated  -
        '-                  with student information                -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        'Write Grade Distribution stats header
        Console.WriteLine(vbCrLf & StrDup(73, "-"))
        Console.WriteLine(String.Format("{0, 51}", "Grade Distribution Statistics"))
        Console.WriteLine(StrDup(73, "-"))

        'Print students with letter grade A
        WriteDistStats(lstStudents, "A")

        'Simply adding a line of space
        Console.WriteLine()

        'Print students with letter grade B
        WriteDistStats(lstStudents, "B")

        'Simply adding a line of space
        Console.WriteLine()

        'Print students with letter grade C
        WriteDistStats(lstStudents, "C")

        'Simply adding a line of space
        Console.WriteLine()

        'Print students with letter grade D
        WriteDistStats(lstStudents, "D")

        'Simply adding a line of space
        Console.WriteLine()

        'Print students with letter grade F
        WriteDistStats(lstStudents, "F")

        'Simply adding a line of space
        Console.WriteLine()
    End Sub

    Private Sub WriteDistStats(lstStudents As List(Of clsStudent), strLetter As String)
        '------------------------------------------------------------
        '-            Subprogram Name: WriteDistStats               -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: February 11, 2021             -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine generates each report for individual     -
        '- letter grades using LINQ. After analysis is performed,   -
        '- the results are formatted and written to console.        -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- lstStudents:     List (of clsStudent) that is populated  -
        '-                  with student information                -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- blnFlag:         Boolean type variable used to flag if   -
        '-                  there are no students with the current  -
        '-                  letter grade.                           -
        '- objStudents:     Object type variable used with LINQ to  -
        '-                  perform analysis on lstStudent's        -
        '-                  contents.                               -
        '------------------------------------------------------------

        Console.WriteLine("Those students earning a " & strLetter & " grade are:")

        'Get all students with matching letter grade
        Dim objStudents As Object
        objStudents = From students In lstStudents
                      Where students.strLetterGrade Like strLetter & "*"
                      Order By students.strLastName 'Ordering result set by students' last name
                      Select students.strInitials, students.strLastName, students.strLetterGrade

        'Boolean flag for students having letter grade
        Dim blnFlag As Boolean = True

        'Print A students
        For Each aStudents In objStudents
            Console.WriteLine(String.Format("{0, 7} {1, -10} {2, 3} {3, -2}", aStudents.strInitials, aStudents.strLastName, "-->", aStudents.strLetterGrade))
            blnFlag = False
        Next

        'If no students had this letter grade, write that
        If blnFlag Then
            Console.WriteLine(String.Format("{0, 7}", "None"))
        End If
    End Sub

    Private Sub GetRangeStats(lstStudents As List(Of clsStudent))
        '------------------------------------------------------------
        '-            Subprogram Name: GetRangeStats                -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: February 11, 2021             -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine generates the range statistics for each  -
        '- graded item within the course. The data is then send to  -
        '- another Sub for formatting and printing.                 -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- lstStudents:     List (of clsStudent) that is populated  -
        '-                  with student information                -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- intCtr:          Integer variable used to keep count of  -
        '-                  loop cycles.                            -
        '- sngHW1:          Single type Array used to hold the      -
        '-                  grades from each student for homework 1 -
        '- sngHW2:          Single type Array used to hold the      -
        '-                  grades from each student for homework 2 -
        '- sngHW3:          Single type Array used to hold the      -
        '-                  grades from each student for homework 3 -
        '- sngHW4:          Single type Array used to hold the      -
        '-                  grades from each student for homework 4 -
        '- sngExam:         Single type Array used to hold the      -
        '-                  grades from each student for the exam   -
        '- objStudents:     Object type variable used with LINQ to  -
        '-                  perform analysis on lstStudent's        -
        '-                  contents.                               -
        '------------------------------------------------------------

        'Write Homework/Exam Grade Range stats header
        Console.WriteLine(vbCrLf & StrDup(73, "-"))
        Console.WriteLine(String.Format("{0, 53}", "Homework/Exam Grade Range Statistics"))
        Console.WriteLine(StrDup(73, "-"))
        Console.WriteLine(String.Format("{0, 18}{1, 25}{2, 27}", "Low", "Ave", "High"))

        Dim objStudents As Object

        'Get all students' homework and exam score
        objStudents = From students In lstStudents
                      Select students.sngHomework1, students.sngHomework2, students.sngHomework3, students.sngHomework4, students.sngExamScore

        Dim sngHW1(lstStudents.Count - 1) As Single
        Dim sngHW2(lstStudents.Count - 1) As Single
        Dim sngHW3(lstStudents.Count - 1) As Single
        Dim sngHW4(lstStudents.Count - 1) As Single
        Dim sngExam(lstStudents.Count - 1) As Single

        'Counter for For Loop
        Dim intCtr As Integer = 0

        'Populate Single arrays for each homework and exam
        For Each student In objStudents
            sngHW1(intCtr) = CSng(student.sngHomework1)
            sngHW2(intCtr) = CSng(student.sngHomework2)
            sngHW3(intCtr) = CSng(student.sngHomework3)
            sngHW4(intCtr) = CSng(student.sngHomework4)
            sngExam(intCtr) = CSng(student.sngExamScore)
            intCtr += 1
        Next

        'Write HW1 stats
        WriteRangeStats(sngHW1, "Homework 1", intHWPOSSIBLE)

        'Write HW2 stats
        WriteRangeStats(sngHW2, "Homework 2", intHWPOSSIBLE)

        'Write HW3 stats
        WriteRangeStats(sngHW3, "Homework 3", intHWPOSSIBLE)

        'Write HW4 stats
        WriteRangeStats(sngHW4, "Homework 4", intHWPOSSIBLE)

        'Write Exam stats
        WriteRangeStats(sngExam, "Exam", intEXAMPOSSIBLE)
    End Sub

    Private Sub WriteRangeStats(sngArray() As Single, strGrading As String, intPossible As Integer)
        '------------------------------------------------------------
        '-            Subprogram Name: WriteRangeStats              -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: February 11, 2021             -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine prints the range statistics for the      -
        '- current graded item within the course.                   -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- sngArray:        Single type Array that holds the scores -
        '-                  for the current graded item             -
        '- strGrading:      String variable that holds the name of  -
        '-                  current graded assignment.              -
        '- intPossible:     Integer variable that holds the maximum -
        '-                  amount of points available to earn on   -
        '-                  the currend graded assignment.          -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- sngAve:          Single variable used to hold the        -
        '-                  average of the current assignment's     -
        '-                  score.                                  -
        '- sngMax:          Single variable used to hold the max    -
        '-                  score on the current assignment.        -
        '- sngMin:          Single variable used to hold the min    -
        '-                  score on the current assignment.        -
        '------------------------------------------------------------

        'Dim and populate variables for min, ave, and max as percentages
        Dim sngMin As Single = (Aggregate nums In sngArray Into Min()) * 100 / intPossible
        Dim sngAve As Single = (Aggregate nums In sngArray Into Average()) * 100 / intPossible
        Dim sngMax As Single = (Aggregate nums In sngArray Into Max()) * 100 / intPossible

        'Write stats
        Console.WriteLine(String.Format("{0, -10} {1, 1}  {2, 7}{3, 18}{4, 7}{3, 18}{5, 8}", strGrading, ":",
        Format(sngMin, "0.00") & " %", StrDup(18, "."), Format(sngAve, "0.00") & " %", Format(sngMax, "0.00") & " %"))
    End Sub

    Private Sub GetOveralStats(lstStudents As List(Of clsStudent))
        '------------------------------------------------------------
        '-            Subprogram Name: GetOveralStats               -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: February 15, 2021             -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine generates the overall statistics for the -
        '- course as a whole. The overall statistics are both       -
        '- generated and printed in this Sub.                       -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- lstStudents:     List (of clsStudent) that is populated  -
        '-                  with student information                -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- sngAve:          Single variable used to hold the        -
        '-                  average of the current course's         -
        '-                  score.                                  -
        '- intCtr:          Integer variable used to keep count of  -
        '-                  loop cycles.                            -
        '- sngGrade:        Single type Array used to hold the      -
        '-                  grades of each student within the       -
        '-                  course.                                 -
        '- sngMax:          Single variable used to hold the max    -
        '-                  score on the current course.            -
        '- sngMin:          Single variable used to hold the min    -
        '-                  score on the current course.            -
        '- objStudents:     Object type variable used with LINQ to  -
        '-                  perform analysis on lstStudent's        -
        '-                  contents.                               -
        '------------------------------------------------------------

        'Write Overal Grade stats header
        Console.WriteLine(vbCrLf & StrDup(73, "-"))
        Console.WriteLine(String.Format("{0, 52}", "Overall Course Grade Statistics"))
        Console.WriteLine(StrDup(73, "-"))

        'Get all students' numeric course grade
        Dim objStudents As Object = From students In lstStudents
                                    Select students.sngNumericGrade, students.strInitials, students.strLastName

        'Create single array
        Dim sngGrade(lstStudents.Count - 1) As Single

        'Counter for loop
        Dim intCtr As Integer = 0

        'Translate object into Single array
        For Each student In objStudents
            sngGrade(intCtr) = CSng(student.sngNumericGrade)
            intCtr += 1
        Next

        'Create overal grade vars
        Dim sngMax As Single = Aggregate nums In sngGrade Into Max()
        Dim sngMin As Single = Aggregate nums In sngGrade Into Min()
        Dim sngAve As Single = Aggregate nums In sngGrade Into Average()

        'Write the highest overal grades
        Console.WriteLine("The highest course grade of " & sngMax & " was earned by:")

        'Write any students that had the highest grade
        For Each student In objStudents
            If student.sngNumericGrade = sngMax Then
                Console.WriteLine(String.Format("   {0, -5}{1, -11}{2, -4}{3, -3}", student.strInitials, student.strLastName, "-->", student.sngNumericGrade))
            End If
        Next

        'Write the lowest overal grades
        Console.WriteLine(vbCrLf & "The lowest course grade of " & sngMin & " was earned by:")

        'Write any students that had the lowest grade
        For Each student In objStudents
            If student.sngNumericGrade = sngMin Then
                Console.WriteLine(String.Format("   {0, -5}{1, -11}{2, -4}{3, -3}", student.strInitials, student.strLastName, "-->", student.sngNumericGrade))
            End If
        Next

        'Write the average grades
        Console.WriteLine(vbCrLf & "The average course grade was " & sngAve)
    End Sub
End Module