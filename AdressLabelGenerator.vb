Option Strict On
Option Explicit On

'Taylor Herndon
'RCET0265
'Spring 2021
'Adress Label Program
'https://github.com/TaylorHerndon/AdressLabel

Public Class AdressLabelGenerator

    Sub SubmitButtonPress() Handles SubmitButton.Click

        Dim hasProblem(6) As Boolean

        'Check to see if any of the folling text boxes contain a number and store the results in an array and message
        hasProblem(6) = CheckString(FirstNameTextBox.Text, "First Name")
        hasProblem(5) = CheckString(MITextBox.Text, "Middle Inital")
        hasProblem(4) = CheckString(LastNameTextBox.Text, "Last Name")
        hasProblem(2) = CheckString(CityTextBox.Text, "City")
        hasProblem(1) = CheckString(StateTextBox.Text, "State")

        'Check to see if the street adress text box is empty and store the results in an array and message
        If StreetAdressTextBox.Text = "" Then

            StoreMessage("Street Adress", False)
            hasProblem(3) = True

        End If

        'Check to see if the zipcode text box contains only numbers and store the results in an array and message
        Try
            'Try to convert to integer
            Convert.ToInt32(ZipCodeTextBox.Text)

        Catch ex As Exception

            'If ZipCode cannot be converted to an integer there is a problem.
            hasProblem(0) = True

            'Check to see if Zip Code is empty and store the correct response.
            If ZipCodeTextBox.Text = "" Then

                StoreMessage("Zip Code is empty", False)

            Else

                StoreMessage("Zip Code", False)

            End If

        End Try

        'If there are any stored messages write them out in a msg box
        If StoreMessage("", False) <> "" Then

            MsgBox("An error occoured in the following fields." & vbNewLine & StoreMessage("", False))
            StoreMessage("", True)

            'Set the focus to the top most field that has a problem
            For i = 0 To 6

                If hasProblem(i) Then

                    Select Case i
                        Case 0
                            ZipCodeTextBox.Focus()
                        Case 1
                            StateTextBox.Focus()
                        Case 2
                            CityTextBox.Focus()
                        Case 3
                            StreetAdressTextBox.Focus()
                        Case 4
                            LastNameTextBox.Focus()
                        Case 5
                            MITextBox.Focus()
                        Case 6
                            FirstNameTextBox.Focus()
                    End Select

                End If

            Next

            Exit Sub

        End If

        'If all tests are passed write out the adress label
        AdressLabel.Text = FirstNameTextBox.Text & " " & MITextBox.Text & " " & LastNameTextBox.Text & vbNewLine &
                           StreetAdressTextBox.Text & vbNewLine &
                           CityTextBox.Text & ", " & StateTextBox.Text & " " & ZipCodeTextBox.Text

    End Sub

    Sub ClearButtonClick() Handles ClearButton.MouseDown

        'Return adress label to default state
        AdressLabel.Text = "Name " & vbNewLine & "Street Adress" & vbNewLine & "City State, Zip"

        'Clear all text boxes
        FirstNameTextBox.Text = ""
        MITextBox.Text = ""
        LastNameTextBox.Text = ""
        StreetAdressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""

    End Sub

    Sub ExitButtonClick() Handles ExitButton.Click

        End

    End Sub

    Function CheckString(checkThisString As String, testedField As String) As Boolean

        Dim stringLength As Integer
        Dim hasNumber As Boolean = False

        'Gets the length of the string in question
        stringLength = Len(checkThisString)

        'If the string is empty then return "Empty"
        If checkThisString = "" Then

            StoreMessage(testedField & " is empty.", False)
            Return True

        Else

            'If the string is not empty test if each character is a number
            For i = 0 To stringLength - 1

                Try

                    Convert.ToInt32(checkThisString.Substring(i, 1)) 'Test the character
                    hasNumber = True 'If the code continues then the tested character is a number
                    StoreMessage(testedField, False) 'Store what field has a problem
                    Return True

                Catch ex As Exception

                End Try

            Next

        End If

        'Return whether or not the tested string has a number in it
        Return hasNumber

    End Function

    Function StoreMessage(Message As String, Clear As Boolean) As String

        Static storedMessages As String

        'If clear is true reset stored messages
        If Clear Then

            storedMessages = ""
            Return storedMessages

        End If

        'If message is empty then return the stored messages and continue
        If Message = "" Then

            Return storedMessages

        End If

        'Add the new message to the StoredMessages String
        storedMessages = storedMessages & vbNewLine & Message

        Return storedMessages

    End Function

End Class
