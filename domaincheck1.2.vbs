Dim domains
Dim test_domain
Dim isValid
Dim result

test_domain = InputBox("Enter Your Desired Domain")
domains = Array(".com", ".net", ".live", ".edu", ".gov")
isValid = False

' Check if the domain ends with .gov and not with any other domain extension
If Right(test_domain, 4) = ".com" Then
    isValid = True
    MsgBox "The domain you have selected " & test_domain & " is an acceptable domain because it ends with .com"
Else
    MsgBox "The domain you have selected " & test_domain & " is not an acceptable domain because it ends with .edu"
End If

'Make sure this applies to the email adress of the main receiver not the CC or BCC'