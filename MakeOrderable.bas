Attribute VB_Name = "MakeOrderable"
Sub setPddAutoPRC()

    'automatically-declared during macro recording
    On Error GoTo ErrorHandler
    Const NEVER_TIME_OUT = 0
    Dim BS As String    ' Chr(rcBS) = Chr(8) = Control-H
    Dim ESC As String   ' Chr(rcESC) = Chr(27) = Control-[
    BS = Chr(Reflection2.ControlCodes.rcBS)
    ESC = Chr(Reflection2.ControlCodes.rcESC)
    
    Dim HT As String    ' Chr(rcHT) = Chr(9) = Control-I
    HT = Chr(Reflection2.ControlCodes.rcHT)
    
    
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
    End With

    'get lab code
    Dim LabCode As String
    LabCode = InputBox("Enter lab code for optional", "Lab Code")
    If (LabCode = "") Or (Not LabCode Like "[A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9]") Then
        MsgBox "Aborting... That doesn't look like a valid lab code."
        Exit Sub
    End If
    
    'get test codes
    Dim TestCodes As String
    TestCodes = InputBox("Paste string of comma-delimited test codes" & vbNewLine & vbNewLine & "E.g. 100000,100001,100002", "Test Codes")
    
        Dim TCodeDelimiter As String
        TCodeDelimiter = ","
    
    'build an array from the string
    Dim TestCodeAry() As String
    TestCodeAry = split(TestCodes, TCodeDelimiter)
    Dim TestCodeAryLen As Integer
    TestCodeAryLen = (UBound(TestCodeAry) - LBound(TestCodeAry) + 1)
    
    'individual-test-code and count vars for loop
    Dim testCode As Variant
    Dim i As Integer
    i = 1

    'functional stuff in reflection begins here
    With Session
    For Each testCode In TestCodeAry 'loop through test codes in array
        
        'make sure test code looks like a test code
        If Not testCode Like "[A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9]" Then
            MsgBox "Aborting... " & testCode & " doesn't look like a valid test code."
            MsgBox "Check your string of test codes... " & vbNewLine & vbNewLine & TestCodes
            Exit For
        End If
        
        'MsgBox TestCode & " (" & i & " of " & TestCodeAryLen & ")" 'debug
        
        'navigate to Option screen
        .Transmit "26"
        .TransmitTerminalKey rcVtPF4Key
        .Transmit "26"
        .TransmitTerminalKey rcVtPF4Key
        
        'ask user to specify test (15), panel (9), or super (10)
        Dim TestType As String
        TestType = InputBox("Enter" + vbNewLine + "Superpanel = 10" + vbNewLine + "Panel      =9" + vbNewLine + "Test       =15", "TestType")
        If (TestType = "") Or (Not TestType Like "9" Or "10" Or "15") Then
        MsgBox "Aborting... Not a valid option."
        Exit Sub
        End If
        'SET UP CASE BY CASE
        
        Select Case TestType
        
            Case TestType = 10
                   
                .Transmit "10"
                .Transmit testCode
                .Transmit "3" 'change super panel
                .TransmitTerminalKey rcVtPF4Key
                .TransmitTerminalKey rcVtDownKey
                .TransmitTerminalKey rcVtDownKey
                .Transmit HT
                .Transmit HT
                .Transmit HT
                .Transmit HT
                .Transmit HT
                .Transmit HT
                .Transmit "Y"
                .TransmitTerminalKey rcVtPF4Key
                .Transmit Format(Date, "MMDDYYYY")
                .TransmitTerminalKey rcVtPF4Key
                .Transmit "Y"
                .TransmitTerminalKey rcVtPF4Key
                
            Case TestType = 9
            
            Case TestType = 10
            
            
        End Select
        
        .Transmit "13"
        .TransmitTerminalKey rcVtPF4Key

        .Transmit testCode
        'behavior: automatically moves you to next field after 6 chars
        .Transmit LabCode
        .Transmit "2" 'Shift
        'behavior: must 'tab' or 'down' to Option field
        .TransmitTerminalKey rcVtDownKey
        .Transmit "3" 'Option: Change
        'TODO? add option to enter 2 (Add) or 3 (Change)?
        .TransmitTerminalKey rcVtPF4Key
        
        'now you're in the sunken zone (Optional screen)
        .TransmitTerminalKey rcVtDownKey 'Key down to RSLT TYPE field
        .Transmit "N" 'Set to auto-PRC
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
    
        i = i + 1
    Next 'end For Each
    
    End With 'end programmatic stuff in reflection
    
    MsgBox (i - 1) & " test codes set to PRC"
    
    Exit Sub

'automatically generated during macro recording
ErrorHandler:
    Session.MsgBox Err.Description, vbExclamation + vbOKOnly

End Sub

