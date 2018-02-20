Attribute VB_Name = "setPddRouting"
Sub setPddRouting()
    
    On Error GoTo ErrorHandler
    Const NEVER_TIME_OUT = 0
    Dim BS As String    ' Chr(rcBS) = Chr(8) = Control-H
    Dim ESC As String   ' Chr(rcESC) = Chr(27) = Control-[
    Dim HT As String   ' Chr(rcHT) = Chr(9)= Control-I

    BS = Chr(Reflection2.ControlCodes.rcBS)
    ESC = Chr(Reflection2.ControlCodes.rcESC)
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
    If ((LabCode = "") Or ((Not Len(LabCode) = 2) And (Not Len(LabCode) = 5))) Then
        MsgBox "Aborting... That doesn't look like a valid lab code."
        Exit Sub
    End If

    'get shift
    Dim ShiftNum As String
    ShiftNum = InputBox("Enter shift number", "Shift")
    If (ShiftNum = "") Or (Not ShiftNum Like "[0-9]") Then
        MsgBox "That doesn't look like a valid shift... defaulting to shift 2"
        ShiftNum = 2
    End If
    'these are the only shifts i've seen
    If ((Not ShiftNum Like "2") And (Not ShiftNum Like "3")) Then
        ShiftNum = 2
    End If
  
    'get option
    Dim AddOrChange As String
    AddOrChange = InputBox("Enter 2 to Add, 3 to Change", "Option")
    If ((Not AddOrChange Like "2") And (Not AddOrChange Like "3")) Then
        MsgBox "Only valid options are 2 (Add) and 3 (Change)... defaulting to 2 (Add)"
        AddOrChange = 2
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

    'stuff in session begins here
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
        .Transmit "13"
        .TransmitTerminalKey rcVtPF4Key

        .Transmit testCode
        'Note: behavior automatically moves you to next field after 6 chars
        .Transmit LabCode
        'Note: behavior automatically moves you to next field after 5 chars
        '      the following rcVtLeftKey HT combo helps handle 2-5 digit LabCodes
        .TransmitTerminalKey rcVtLeftKey
        .Transmit HT
        .Transmit ShiftNum 'Shift
        .TransmitTerminalKey rcVtDownKey 'jumps to Option field
        .Transmit AddOrChange
        .TransmitTerminalKey rcVtPF4Key
        
        'now you're in the sunken zone (Optional screen)
            
        'delete attached abbrev
        .Transmit HT 'tab
        .TransmitTerminalKey rcVtRemoveKey 'delete
            
        'delete reflex code
        .Transmit HT
        .TransmitTerminalKey rcVtRemoveKey
            
        'delete reflex test code
        .Transmit HT
        .TransmitTerminalKey rcVtRemoveKey
            
        'delete abrev
        .Transmit HT
        .TransmitTerminalKey rcVtRemoveKey
            
        'delete RPT OPT
        .Transmit HT
        .TransmitTerminalKey rcVtRemoveKey
            
        'norm type N
        .Transmit HT
        .Transmit HT
        .Transmit HT
        .Transmit "N"
            
        'change worksheet to 0007
        .TransmitTerminalKey rcVtDownKey
        '.Transmit ESC & "[Z" 'back tab
        .Transmit "0007" '0007
        'set regional "Y"; continues into next field
        .Transmit "Y"
        .TransmitTerminalKey rcVtRemoveKey 'delete override lab/shift
        .Transmit HT 'tab
        .TransmitTerminalKey rcVtRemoveKey
            
        'transmit 3x to approve change
        .TransmitTerminalKey rcVtPF4Key
        .TransmitTerminalKey rcVtPF4Key
        'MsgBox "<" & Format(Date, "MMDDYYYY") & ">" 'debug
        .Transmit Format(Date, "MMDDYYYY")
        .TransmitTerminalKey rcVtPF4Key
    
        'escape 3x back to main screen
        .TransmitTerminalKey rcVtF14Key
        .TransmitTerminalKey rcVtF14Key
        .TransmitTerminalKey rcVtF14Key
    
        i = i + 1
    Next 'end For Each
    
    End With 'end programmatic stuff in reflection
    
    MsgBox (i - 1) & " test codes route changes"
    
    Exit Sub

'automatically generated during macro recording
ErrorHandler:
    Session.MsgBox Err.Description, vbExclamation + vbOKOnly

End Sub

