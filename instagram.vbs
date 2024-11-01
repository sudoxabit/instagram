' Display message for support
MsgBox("SUPPORT XABIT FOR MORE TRICKS")

' Collect multi-line message by combining multiple input prompts
Message = InputBox("Enter the first line of your message:", "XABIT")
Do
    line = InputBox("Enter the next line of your message (leave blank to finish):", "XABIT")
    If line = "" Then Exit Do
    Message = Message & vbCrLf & line
Loop

' Number of times to send the message
T = InputBox("How many times should the message be sent?", "XABIT")

' Validate that T is a positive integer
If Not IsNumeric(T) Or T <= 0 Then
    MsgBox "Please enter a valid number greater than 0", vbExclamation, "Invalid Input"
    WScript.Quit
End If

' Confirmation message
If MsgBox("You've filled everything in correctly", 1024 + vbSystemModal, "XABIT") = vbOk Then
    Set WshShell = WScript.CreateObject("WScript.Shell")
    WshShell.Run("https://www.instagram.com/direct/inbox/")
    MsgBox("Instagram DM will open shortly. Select a friend to send messages.")

    ' Short delay for the Instagram page to load
    WScript.Sleep 3000

    ' Loop to send messages
    For i = 1 To T
        WshShell.SendKeys Message
        WScript.Sleep 50 ' Short delay before pressing enter
        WshShell.SendKeys "{ENTER}"
        WScript.Sleep 200 ' Short delay between messages for smoothness
    Next

    ' Notification after all messages are sent
    MsgBox "All messages have been sent!", vbInformation, "XABIT"
Else
    MsgBox "Process Has Been Cancelled", vbSystemModal, "XABIT"
End If
