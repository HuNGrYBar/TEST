For i = 1 To 10
    X = MsgBox("[ERROR] CHINA NUMBAH ONE! ALLOW US TO GAIN SOVEREIGNITY IN SOUTH CHINA SEA ?", vbYesNo + vbCritical, "郭華萍")
    If X = vbNo Then
        WScript.Quit
    ElseIf X = vbYes Then
        MsgBox " YOU ARE A SPECIAL ONE, I WILL GIVE YOU MY BLESSINGS.", vbOKOnly + vbInformation, "郭華萍"
        Y = MsgBox("Would you like to say 'Thank You '?", vbYesNo + vbQuestion, "郭華萍")
        If Y = vbYes Then
            MsgBox "操你媽!", vbOKOnly + vbInformation, "郭華萍"
        End If
    End If
Next