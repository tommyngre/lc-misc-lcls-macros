Attribute VB_Name = "setPddUnits"
Sub setPddUnits()

    'notes:
    'As is, PDD needs to exist before this will work
    'This uses option 3-Change on the Optional Procedures screen
    '
    'resources:
    'http://analystcave.com/vba-read-file-vba/
    'https://msdn.microsoft.com/en-us/library/6x627e5f(v=vs.90).aspx
    
    Const NEVER_TIME_OUT = 0
    Dim HT As String    ' Chr(rcHT) = Chr(9) = Control-I
    HT = Chr(9)
    
    'make sure in SYS9 by grabbing text from console
    With Session
    MyString1 = .GetText(23, 16, 23, 20)
    x = .CursorRow
    y = .CursorColumn
        
    If MyString1 = "CPU:9" And x = 0 And y = 1 Or MyString1 = "CPU:9" And x = 22 And y = 34 Then
    Else
        MsgBox "Aborting... This only runs on SYS9 from the Main Menu."
        Exit Sub
    End If

    'get lab code
    Dim LabCode As String
    LabCode = InputBox("Enter lab code for optional", "Lab Code")
    If (LabCode = "") Or (Not LabCode Like "[A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9]") Then
        MsgBox "Aborting... That doesn't look like a valid lab code."
        Exit Sub
    End If
    
    Dim InputFileName As String
    Dim line As String
    Dim lineArry As Variant
    Dim textRow As String
    Dim testCode As String
    Dim units As String
        
    Close
    
    InputFileName = "\\bnwcfs02\shared\frontend\share\EDI Testing\Macros\Tommy\testing\pddUnitsTest.txt"
    InputFileNum = FreeFile
       
    Open InputFileName For Input As #1
        
    Do While Not EOF(1)
        
        Line Input #1, textRow
        'MsgBox textRow
        lineArry = split(textRow, Chr(9))
        testCode = lineArry(0)
        'MsgBox testCode
        units = lineArry(1)
        'MsgBox units
            
        'set units
        'make sure test code looks like a test code
        If Not testCode Like "[A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9]" Then
            MsgBox "Aborting... " & testCode & " doesn't look like a valid test code."
            MsgBox "Check your string of test codes... " & vbNewLine & vbNewLine & TestCodes
        End If
                    
        'navigate to Option screen
        .Transmit "26"
        .TransmitTerminalKey rcVtPF4Key
        .Transmit "26"
        .TransmitTerminalKey rcVtPF4Key
        .Transmit "13"
        .TransmitTerminalKey rcVtPF4Key
        .Transmit testCode
        'behavior: automatically moves you to next field after 6 chars
        .Transmit LabCode
        .Transmit "2" 'Shift
        .TransmitTerminalKey rcVtDownKey
        .Transmit "3" 'Option: Change
        .TransmitTerminalKey rcVtPF4Key
        
        'now you're in the sunken zone (Optional screen)
        '6 tabs, then enter units
        .Transmit HT
        .Transmit HT
        .Transmit HT
        .Transmit HT
        .Transmit HT
        .Transmit HT
        .Transmit units
        .TransmitTerminalKey rcVtRemoveKey
        'MsgBox LabCode & "-" & testCode & "-" & units
        
        
        'transmit 3x to approve change
        .TransmitTerminalKey rcVtPF4Key
        .TransmitTerminalKey rcVtPF4Key
        'MsgBox "<" & Format(Date, "MMDDYYYY") & ">"
        .Transmit Format(Date, "MMDDYYYY")
        .TransmitTerminalKey rcVtPF4Key
    
        'escape 3x back to main screen
        .TransmitTerminalKey rcVtF14Key
        .TransmitTerminalKey rcVtF14Key
        .TransmitTerminalKey rcVtF14Key
    Loop
    
    Close InputFileNum
    
    End With 'end programmatic stuff in reflection
        
    Exit Sub

End Sub
