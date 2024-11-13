'Code made by https://github.com/Gxldrun
Dim morseCodeDict, reverseMorseCodeDict
Set morseCodeDict = CreateObject("Scripting.Dictionary")
Set reverseMorseCodeDict = CreateObject("Scripting.Dictionary")

morseCodeDict.Add ".-", "A"
morseCodeDict.Add "-...", "B"
morseCodeDict.Add "-.-.", "C"
morseCodeDict.Add "-..", "D"
morseCodeDict.Add ".", "E"
morseCodeDict.Add "..-.", "F"
morseCodeDict.Add "--.", "G"
morseCodeDict.Add "....", "H"
morseCodeDict.Add "..", "I"
morseCodeDict.Add ".---", "J"
morseCodeDict.Add "-.-", "K"
morseCodeDict.Add ".-..", "L"
morseCodeDict.Add "--", "M"
morseCodeDict.Add "-.", "N"
morseCodeDict.Add "---", "O"
morseCodeDict.Add ".--.", "P"
morseCodeDict.Add "--.-", "Q"
morseCodeDict.Add ".-.", "R"
morseCodeDict.Add "...", "S"
morseCodeDict.Add "-", "T"
morseCodeDict.Add "..-", "U"
morseCodeDict.Add "...-", "V"
morseCodeDict.Add ".--", "W"
morseCodeDict.Add "-..-", "X"
morseCodeDict.Add "-.--", "Y"
morseCodeDict.Add "--..", "Z"
morseCodeDict.Add "-----", "0"
morseCodeDict.Add ".----", "1"
morseCodeDict.Add "..---", "2"
morseCodeDict.Add "...--", "3"
morseCodeDict.Add "....-", "4"
morseCodeDict.Add ".....", "5"
morseCodeDict.Add "-....", "6"
morseCodeDict.Add "--...", "7"
morseCodeDict.Add "---..", "8"
morseCodeDict.Add "----.", "9"
morseCodeDict.Add "-.-.--", "!"
morseCodeDict.Add "..--..", "?"
morseCodeDict.Add "--..--", ","
morseCodeDict.Add "-.-.-.", ";"
morseCodeDict.Add ".-.-.-", "."
morseCodeDict.Add "---...", ":"
morseCodeDict.Add "-...-", "="
morseCodeDict.Add ".-.-.", "+"

For Each key In morseCodeDict.Keys
    reverseMorseCodeDict.Add morseCodeDict(key), key
Next

Do
    inputText = InputBox("Enter text to encode/decode. Use spaces to separate Morse code letters and '/' for word breaks." & vbCrLf & vbCrLf & "Click Cancel or leave blank and press OK to exit.", "Morse Code Converter")
    If inputText = "" Then Exit Do
    If InStr(inputText, ".") > 0 Or InStr(inputText, "-") > 0 Then
        words = Split(inputText, "/")
        decodedText = ""
        For Each word In words
            letters = Split(Trim(word), " ")
            For Each letter In letters
                If morseCodeDict.Exists(letter) Then
                    decodedText = decodedText & LCase(morseCodeDict(letter))
                Else
                    decodedText = decodedText & "*"
                End If
            Next
            decodedText = decodedText & " "
        Next
        InputBox "Decoded Text:", "Morse Code Converter", decodedText
    Else
        encodedText = ""
        inputText = UCase(inputText)
        For i = 1 To Len(inputText)
            char = Mid(inputText, i, 1)
            If reverseMorseCodeDict.Exists(char) Then
                encodedText = encodedText & reverseMorseCodeDict(char) & " "
            ElseIf char = " " Then
                encodedText = encodedText & "/ "
            Else
                encodedText = encodedText & "* "
            End If
        Next
        InputBox "Encoded Morse Code:", "Morse Code Converter", Trim(encodedText)
    End If
Loop