Attribute VB_Name = "BillingScenarios"
Sub billingScenarioWrapper()
'enhancement idea: reduce redundant code by using this subroutine to...
'  validate SYS9
'  collect specID
'  collect scenario

End Sub
Sub inputMedicareInfo()
' accepts a spec ID and inputs medicare info
'
    On Error GoTo ErrorHandler

    Const NEVER_TIME_OUT = 0

    Dim BS As String    ' Chr(rcBS) = Chr(8) = Control-H
    Dim HT As String    ' Chr(rcHT) = Chr(9) = Control-I
    Dim LF As String    ' Chr(rcLF) = Chr(10) = Control-J
    Dim ESC As String   ' Chr(rcESC) = Chr(27) = Control-[

    BS = Chr(Reflection2.ControlCodes.rcBS)
    HT = Chr(Reflection2.ControlCodes.rcHT)
    LF = Chr(Reflection2.ControlCodes.rcLF)
    ESC = Chr(Reflection2.ControlCodes.rcESC)

    With Session
    
        'validate SYS9, accept specID, goto host maintenance
        Dim MyString1 As String, x1 As String, y1 As String
        MyString1 = .GetText(23, 16, 23, 20)
        x1 = .CursorRow
        y1 = .CursorColumn
        
        If (MyString1 = "CPU:9" And x1 = 0 And y1 = 1) Or (MyString1 = "CPU:9" And x1 = 22 And y1 = 34) Then
        Else
            MsgBox "ERROR.  This only runs on SYS 9 from the Main Menu.  The macro has stopped."
            Exit Sub
        End If

        Dim specID As String
        specID = InputBox("Enter spec ID", "specID")
        If (specID = "") Or (Not specID Like "[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]") Then
            MsgBox "Aborting... That doesn't look like a valid spec ID."
        Exit Sub
        End If
        
        .Transmit "11"
        .TransmitTerminalKey rcVtPF4Key
        .Transmit specID
        'option: edit
        .Transmit "2"
        .TransmitTerminalKey rcVtPF4Key
        'end spec entry - hereafter, modifying spec

        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        'tab
        .Transmit HT
        .Transmit HT
        .Transmit HT
        'enter bill code
        .Transmit "05"
        'continue typing into clinical info field (optional)
        .Transmit "MEDICARE BILL TEST"
        .TransmitTerminalKey rcVtPF4Key

        'enter diagnosis code
        .Transmit "331.0"
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .Transmit HT
        .Transmit HT
        'enter medicare ID
        .Transmit "123456789A"
        .Transmit HT
        .Transmit HT
        'set ABN to Y? Y not?
        .Transmit "Y"
        .Transmit HT
        .Transmit "MEDICARE INS NAME"
        .Transmit HT
        .Transmit HT
        .Transmit HT
        .Transmit HT
        .Transmit "MEDICARE ADDR LN 1"
        .Transmit HT
        .Transmit "MEDICARE ADDR LN 2"
        .Transmit HT
        .Transmit "BURLINGTON"
        .Transmit HT
        .Transmit "NC"
        .Transmit "27215"
        .TransmitTerminalKey rcVtPF4Key
    
    End With
    Exit Sub

ErrorHandler:
    Session.MsgBox Err.Description, vbExclamation + vbOKOnly

End Sub
Sub inputMedicaidInfo()
' accepts spec ID and inputs Medicaid info
'
    On Error GoTo ErrorHandler

    Const NEVER_TIME_OUT = 0

    Dim BS As String    ' Chr(rcBS) = Chr(8) = Control-H
    Dim HT As String    ' Chr(rcHT) = Chr(9) = Control-I
    Dim LF As String    ' Chr(rcLF) = Chr(10) = Control-J
    Dim ESC As String   ' Chr(rcESC) = Chr(27) = Control-[

    BS = Chr(Reflection2.ControlCodes.rcBS)
    HT = Chr(Reflection2.ControlCodes.rcHT)
    LF = Chr(Reflection2.ControlCodes.rcLF)
    ESC = Chr(Reflection2.ControlCodes.rcESC)

    With Session
    
        'validate SYS9, accept specID, goto host maintenance
        Dim MyString1 As String, x1 As String, y1 As String
        MyString1 = .GetText(23, 16, 23, 20)
        x1 = .CursorRow
        y1 = .CursorColumn
        
        If (MyString1 = "CPU:9" And x1 = 0 And y1 = 1) Or (MyString1 = "CPU:9" And x1 = 22 And y1 = 34) Then
        Else
            MsgBox "ERROR.  This only runs on SYS 9 from the Main Menu.  The macro has stopped."
            Exit Sub
        End If

        Dim specID As String
        specID = InputBox("Enter spec ID", "specID")
        If (specID = "") Or (Not specID Like "[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]") Then
            MsgBox "Aborting... That doesn't look like a valid spec ID."
        Exit Sub
        End If
        
        .Transmit "11"
        .TransmitTerminalKey rcVtPF4Key
        .Transmit specID
        'option: edit
        .Transmit "2"
        .TransmitTerminalKey rcVtPF4Key
        'end spec entry - hereafter, modifying spec

        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        'tab
        .Transmit HT
        .Transmit HT
        .Transmit HT
        'enter bill code
        .Transmit "NC"
        'continue typing into clinical info field (optional)
        .Transmit "MEDICAID BILL TEST"
        .TransmitTerminalKey rcVtPF4Key

        'enter diagnosis code
        .Transmit "331.0"
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .Transmit ESC & "[Z"
        'enter medicaid ID
        .Transmit "123456789M"
        .Transmit HT
        .Transmit HT
        .Transmit HT
        .Transmit "MEDICAID NAME"
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .Transmit "MEDICAID ADDR LN 1"
        .Transmit HT
        .Transmit "MEDICAID ADDR LN 2"
        .Transmit HT
        .Transmit "BURLINGTON"
        .Transmit HT
        .Transmit "NC"
        .Transmit "27215"
        .TransmitTerminalKey rcVtPF4Key

    End With
    Exit Sub

ErrorHandler:
    Session.MsgBox Err.Description, vbExclamation + vbOKOnly
    
End Sub
Sub inputThirdPartyBillInfo()
' accepts spec ID and inputs 3rd party billing info
'
    On Error GoTo ErrorHandler

    Const NEVER_TIME_OUT = 0

    Dim BS As String    ' Chr(rcBS) = Chr(8) = Control-H
    Dim HT As String    ' Chr(rcHT) = Chr(9) = Control-I
    Dim LF As String    ' Chr(rcLF) = Chr(10) = Control-J
    Dim ESC As String   ' Chr(rcESC) = Chr(27) = Control-[

    BS = Chr(Reflection2.ControlCodes.rcBS)
    HT = Chr(Reflection2.ControlCodes.rcHT)
    LF = Chr(Reflection2.ControlCodes.rcLF)
    ESC = Chr(Reflection2.ControlCodes.rcESC)

    With Session
        
        'validate SYS9, accept specID, goto host maintenance
        Dim MyString1 As String, x1 As String, y1 As String
        MyString1 = .GetText(23, 16, 23, 20)
        x1 = .CursorRow
        y1 = .CursorColumn
        
        If (MyString1 = "CPU:9" And x1 = 0 And y1 = 1) Or (MyString1 = "CPU:9" And x1 = 22 And y1 = 34) Then
        Else
            MsgBox "ERROR.  This only runs on SYS 9 from the Main Menu.  The macro has stopped."
            Exit Sub
        End If

        Dim specID As String
        specID = InputBox("Enter spec ID", "specID")
        If (specID = "") Or (Not specID Like "[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]") Then
            MsgBox "Aborting... That doesn't look like a valid spec ID."
        Exit Sub
        End If
        
        .Transmit "11"
        .TransmitTerminalKey rcVtPF4Key
        .Transmit specID
        'option: edit
        .Transmit "2"
        .TransmitTerminalKey rcVtPF4Key
        'end spec entry - hereafter, modifying spec

        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        'tab
        .Transmit HT
        .Transmit HT
        .Transmit HT
        'enter bill code
        .Transmit "XI"
        'continue typing into clinical info field (optional)
        .Transmit "THIRD PARTY BILL TEST"
        .TransmitTerminalKey rcVtPF4Key

        'enter diagnosis code
        .Transmit "331.0"
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .TransmitTerminalKey rcVtDownKey
        .Transmit HT
        .Transmit HT
        .Transmit "AETNA INSURANCE"
        .Transmit HT
        .Transmit "AETNA"
        'no tab necessary here
        .Transmit "123456789I"
        .Transmit HT
        .Transmit "123456"
        .Transmit HT
        .StatusBar = "Waiting for Prompt: Insurance Address 1    <"
        .WaitForString ESC & "[16;25H", NEVER_TIME_OUT, rcAllowKeystrokes
        .StatusBar = ""
        .Transmit "INSURANCE ADDR LN 1"
        .Transmit HT
        .Transmit "INSURANCE ADDR LN 2"
        .Transmit HT
        .Transmit "BURLINGTON"
        .Transmit HT
        .Transmit "NC"
        .Transmit "27215"
        .Transmit HT
        .Transmit "22223330"
        .Transmit HT
        .Transmit "LABCORP"
        .TransmitTerminalKey rcVtPF4Key
    
    End With
    Exit Sub

ErrorHandler:
    Session.MsgBox Err.Description, vbExclamation + vbOKOnly

End Sub
