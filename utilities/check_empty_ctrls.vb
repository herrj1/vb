Function ValidateInput() As Boolean
    
    Dim EmptyInputsFound As Boolean

    Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        'check text box
        If TypeName(Ctrl) = "TextBox" Then
            If Ctrl.Text = vbNullString Then
                EmptyInputsFound = True
                Exit For
            End If
        End If
        'check combo box
        If TypeName(Ctrl) = "ComboBox" Then
            If Ctrl.Text = vbNullString Then
                EmptyInputsFound = True
                Exit For
            End If
        End If
    Next Ctrl

    If EmptyInputsFound Then
        ValidateInput = True
    Else
        ValidateInput = False
    End If
End Function
