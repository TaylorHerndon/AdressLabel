Option Strict On
Option Explicit On

'Taylor Herndon
'RCET0265
'Spring 2021
'Adress Label Program
'https://github.com/TaylorHerndon/AdressLabel

Public Class AdressLabelGenerator

    Sub SubmitButtonPress() Handles SubmitButton.MouseDown

        'Check to see if any of the folling text boxes contain a number or are empty
        CheckString(FirstNameTextBox.Text, "First Name")
        CheckString(MITextBox.Text, "Middle Inital")
        CheckString(LastNameTextBox.Text, "Last Name")
        CheckString(CityTextBox.Text, "City")
        CheckString(StateTextBox.Text, "State")

        'Check to see if the street adress text box is empty
        If StreetAdressTextBox.Text = "" Then

            StoreMessage("Street Adress", False)

        End If

        'Check to see if the zipcode text box contains only numbers
        Try
            'Try to conver to integer
            Convert.ToInt32(ZipCodeTextBox.Text)

        Catch ex As Exception

            'If ZipCode cannot be converted to an integer there is a problem.

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

    Sub ExitButtonClick() Handles ExitButton.MouseDown

        End

    End Sub

    Function CheckString(CheckThisString As String, TestedField As String) As String

        Dim StringLength As Integer
        Dim HasNumber As String = "False"

        'Gets the length of the string in question
        StringLength = Len(CheckThisString)

        'If the string is empty then return "Empty"
        If CheckThisString = "" Then

            StoreMessage(TestedField & " is empty.", False)
            Return "Empty"

        Else

            'If the string is not empty test if each character is a number
            For i = 0 To StringLength - 1

                Try

                    Convert.ToInt32(CheckThisString.Substring(i, 1)) 'Test the character
                    HasNumber = "True" 'If the code continues then the tested character is a number
                    StoreMessage(TestedField, False) 'Store what field has a problem

                Catch ex As Exception

                End Try

            Next

        End If

        'Return whether or not the tested string has a number in it
        Return HasNumber

    End Function

    Function StoreMessage(Message As String, Clear As Boolean) As String

        Static StoredMessages As String

        'If clear is true reset stored messages
        If Clear Then

            StoredMessages = ""
            Return StoredMessages

        End If

        'If message is empty then return the stored messages and continue
        If Message = "" Then

            Return StoredMessages

        End If

        'Add the new message to the StoredMessages String
        StoredMessages = StoredMessages & vbNewLine & Message

        Return StoredMessages

    End Function

End Class
