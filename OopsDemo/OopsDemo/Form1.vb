Imports System.Runtime.CompilerServices
Imports OopsDemoClassLibrary

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim t As Test
        t = New Test
        t.SayHello()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim p As New Person
        p.FirstName = "Simran"
        p.LastName = "Multani"
        p.DateOfBirth = #10/03/2021#

        MsgBox(p.FirstName & " " & p.LastName & " " & p.DateOfBirth)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim p As New Person

        p.FirstName = "Simran"
        p.LastName = "Multani"

        If p.SaveDetails() Then
            MsgBox("sucess!")
        Else
            MsgBox("error!")
        End If

        If p.SaveDetails("C:\Users\Public\Documents\simran1.txt") Then
            MsgBox("sucess! 1 paramet method")
        Else
            MsgBox("error!  1 paramet method")
        End If


        If p.SaveDetails("C:\Users\Public\Documents\simran2.txt", "sm") Then
            MsgBox("sucess! 2 paramet method")
        Else
            MsgBox("error!  2 paramet method")
        End If


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim employeeDetail As New EmployeeDetail
        Dim employee As New Employee

        employee.SaveDetails()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim polomorphism As New POlymorphism
        Dim mul = polomorphism.Calculation(5, 6)
        MsgBox("Multiplication of values =" & mul)

        Dim p1 As New Person
        Dim Sum = p1.Calculation(5, 6)
        MsgBox("Sum of values =" & Sum)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim school As New School

        school.Salary()
        school.PeimaryDetails()
        school.totalSubjects()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim example As String = "Example string"
        example.Print()

        example = "Hello"
        example.PrintAndPunctuate(".")
        example.PrintAndPunctuate("!!!!")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim test As New Test

        Const highRate As Double = 12.5
        Dim lowRate = highRate * 0.6

        Dim initialDebt = 4999.99

        Dim debtWithInterest = initialDebt

        test.Calculate(highRate, debtWithInterest)

        Dim debtString = Format(debtWithInterest, "C")
        MsgBox("high interest: " & debtString)

        debtWithInterest = initialDebt
        test.Calculate(lowRate, debtWithInterest)
        debtString = Format(debtWithInterest, "C")
        MsgBox("low INtrest" & debtString)
    End Sub

    Dim i As Integer
    Dim i2 As Integer

    Dim thread As System.Threading.Thread
    Dim thread2 As System.Threading.Thread

    Private Sub ContUp()
        Do Until i = 1000
            i = i + 1
            Label1.Text = i
            Me.Refresh()
        Loop
    End Sub

    Private Sub ContUp2()
        Do Until i = 1000
            i2 = i + 1
            Label2.Text = i2
            Me.Refresh()
        Loop
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        thread = New System.Threading.Thread(AddressOf ContUp)
        thread.Start()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        thread2 = New System.Threading.Thread(AddressOf ContUp2)
        thread2.Start()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CheckForIllegalCrossThreadCalls = False
    End Sub
End Class



Public Class Test
    Public Sub Calculate(ByVal rate As Double, ByRef debt As Double)
        debt = debt + (debt * rate / 100)
    End Sub

    Public Function SayHello()
        MsgBox("Hello !")
    End Function

End Class

Public Class School
    Implements ITeacher
    Implements IStudent

    Public Sub Salary() Implements ITeacher.Salary
        MsgBox("Salary called from Teacher")
    End Sub

    Public Sub PeimaryDetails() Implements IStudent.PrimaryDetails
        MsgBox("Primary Details of students")
    End Sub

    Public Sub totalSubjects() Implements IStudent.totalSubjects
        MsgBox("total Subjects from Subjects")
    End Sub
End Class


Public Interface ISubjects
    Sub totalSubjects()

End Interface

Public Interface ITeacher
    Inherits ISubjects
    Sub Salary()

End Interface
Public Interface IStudent
    Inherits ISubjects
    Sub PrimaryDetails()

End Interface

Public NotInheritable Class SealedClass
    Private Sub New()

    End Sub

End Class

Module ExtensionTest
    <Extension()>
    Public Sub Print(aString As String)
        MsgBox(aString)
    End Sub

    <Extension()>
    Public Sub PrintAndPunctuate(aString As String, punc As String)
        MsgBox(aString & punc)
    End Sub

    Public Delegate Sub Sum(ByVal a As Integer)
    Sub Main(ByVal args As String())

        Dim y As Integer = 10
        Dim Total As Sum = Sub(ByVal x As Integer)
                               MsgBox("Total = " & x + y)
                           End Sub


    End Sub
End Module

