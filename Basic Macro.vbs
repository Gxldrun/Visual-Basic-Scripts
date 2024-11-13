Dim message, frequency, duration, delay, startTime, elapsedTime

message = InputBox("Enter the spam message:", "Spam Message")

frequency = InputBox("Enter the number of times per second to send the message:", "Frequency")
If Not IsNumeric(frequency) Or CInt(frequency) <= 0 Then
    MsgBox "Please enter a valid positive number for frequency.", vbExclamation, "Invalid Input"
    WScript.Quit
End If

delay = 1000 / CInt(frequency)

duration = InputBox("Enter the spam time in seconds. (type 'Inf' for infinite):", "Duration")
If duration <> "Inf" And (Not IsNumeric(duration) Or CInt(duration) <= 0) Then
    MsgBox "Please enter a valid positive number for duration or type 'Inf'.", vbExclamation, "Invalid Input"
    WScript.Quit
End If

startTime = Timer

Do

    Set WshShell = CreateObject("WScript.Shell")
    WshShell.SendKeys message

    WScript.Sleep delay

    If duration <> "Inf" Then
        elapsedTime = Timer - startTime
        If elapsedTime >= CInt(duration) Then Exit Do
    End If
Loop
