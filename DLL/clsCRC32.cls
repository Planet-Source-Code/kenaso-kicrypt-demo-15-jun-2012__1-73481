VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCRC32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Project:       Cyclic Redundancy Check 32-bit (CRC32)
'
' Module:        clsCRC32.cls
'
' Description:   The CRC is a very powerful but easily implemented technique
'                to obtain data reliability. The CRC technique is used to
'                protect blocks of data called Frames. Using this technique,
'                the transmitter appends an extra n- bit sequence to every
'                frame called Frame Check Sequence (FCS). The FCS holds
'                redundant information about the frame that helps the
'                transmitter detect errors in the frame. The CRC is one of
'                the most used techniques for error detection in data
'                communications. The technique gained its popularity because
'                it combines three advantages:
'
'                      - Extreme error detection capabilities
'                      - Little overhead
'                      - Ease of implementation
'
' References:    This code was originally written in Visual Basic by David
'                Ireland.  DI Management Services Pty Ltd, Sydney Australia
'                http://www.di-mgt.com.au/
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 26-JUL-2004  Kenneth Ives  kenaso@tx.rr.com
'              Documented, modified, updated variable names
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 08-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Moved CRC_BlockSize() routine to common module.
' 01-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced Err.Raise with InfoMsg so programs will continue to
'              flow smoothly.
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug in table sizing.  Only appeared if last element
'              of table were accessed.
'              Removed obsolete variables.
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Added ability to see progress display while calculating the
'                Cyclic Redundancy Check of the data (CalcCRC32() routine).
'              - Removed obsolete code.
'              - Updated documentation.
' 18-Feb-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Fixed a bug.  I was reloading table and resetting total
'                accumulator with every pass.
'              - Renamed LoadTable() routine to Initialize().
'              - Created constant MINUS_ONE to replace &HFFFFFFFF.
'              - Moved initial accumulator setting from CalcCRC32() to
'                Initialize() routine.
' 19-Feb-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug.  I was emptying a module level array in the
'              CalcCRC32() routine.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "clsCRC32"
  Private Const ZEROES_8    As String = "00000000"
  Private Const KB_4        As Long = &H1000&     ' 4096
  Private Const MAX_BYTE    As Long = 256
  Private Const MAX_LONG    As Long = &H7FFFFFFF  '  2147483647
  Private Const MIN_LONG    As Long = &H80000000  ' -2147483648
  Private Const MINUS_ONE   As Long = -1
  
' ***************************************************************************
' Module Variables
'
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (Long)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a lng Table
' Variable name:     malngTable
'
' ***************************************************************************
  Private mcurMaxSize  As Currency
  Private mcurAccrued  As Currency
  Private mlngCRC      As Long
  Private malngTable() As Long


' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************
    
' Update progress bar
Public Event CRCProgress(ByVal lngProgress As Long)


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

' Comment out Version Property if not used in this application
Public Property Get Version() As String
    ' Version of the DLL this class is associated with.
    Version = DLL_NAME & ".dll (tm) v" & App.Major & "." & App.Minor & "." & App.Revision & _
              " Copyright (c) 2004-2012 - Kenneth Ives  kenaso@tx.rr.com"
End Property

Public Property Let StopProcessing(ByVal blnData As Boolean)
Attribute StopProcessing.VB_Description = "Switch that designates if the user opts to stop processing."
    gblnStopProcessing = blnData
    DoEvents
End Property

Public Property Get StopProcessing() As Boolean
Attribute StopProcessing.VB_Description = "Switch that designates if the user opts to stop processing."
    StopProcessing = gblnStopProcessing
    DoEvents
End Property


' ***************************************************************************
' ****                         Methods                                   ****
' ***************************************************************************

' ***************************************************************************
' Routine:       CRC32_File
'
' Description:   Returns a hex value representing a file.
' 
' Syntax:        ' convert string data to byte array 
'                abytFile() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)
' 
' Parameters:    abytFile() - Path/filename in byte array
'                blnCheckSum - [Optional] - TRUE - A checksum must return
'                             a full 8 characters.
'                             FALSE - A non-checksum removes leading zeroes
'
' Returns:       TRUE - Successful completion
'                FALSE - An error occurred
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 26-JUL-2004  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 08-MAY-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added call for file verification.
' ***************************************************************************
Public Function CRC32_File(ByRef abytFile() As Byte, _
                  Optional ByVal blnCheckSum As Boolean = False) As String

    Dim hFile        As Long
    Dim lngBlockSize As Long
    Dim curAmtLeft   As Currency
    Dim curFilePos   As Currency
    Dim strCRC       As String
    Dim strSource    As String
    Dim abytData()   As Byte
    Dim objBigFiles  As cBigFiles

    Const ROUTINE_NAME As String = "CRC32_File"

    On Error GoTo CRC32_File_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
         
    ' Test incoming data array
    If Not CBool(IsArrayInitialized(abytFile())) Then
        InfoMsg "Incoming array has not been properly initialized." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If

    Set objBigFiles = New cBigFiles                     ' Instantiate class object
    strSource = ByteArrayToString(abytFile())           ' convert byte array to string
    SetFileAttributes strSource, FILE_ATTRIBUTE_NORMAL  ' Reset file attributes to normal
    
    mcurAccrued = 0@   ' Init accrued progress
    curFilePos = 0@    ' set to first position in file
    Initialize         ' load table array
    strCRC = vbNullString
    
    With objBigFiles
        .CalcFileSize strSource, curAmtLeft  ' Get size of file
        mcurMaxSize = curAmtLeft             ' Save file size for progress calcs
                 
        ' If zero byte file then leave
        If curAmtLeft < 1 Then
            gblnStopProcessing = True
            GoTo CRC32_File_CleanUp
        End If
            
        ' Open source file
        If Not .OpenReadOnly(strSource, hFile) Then
            gblnStopProcessing = True
            GoTo CRC32_File_CleanUp
        End If
        
        Do
            Erase abytData()                              ' empty data array
            lngBlockSize = CRC_BlockSize(curAmtLeft)      ' Process source file in chunks
            ReDim abytData(lngBlockSize - 1)              ' Size receiving array
            curAmtLeft = curAmtLeft - CCur(lngBlockSize)  ' Adjust amount of file left to process
            
            ' Read source file
            If Not .API_ReadFile(hFile, curFilePos, abytData()) Then
                gblnStopProcessing = True
                Exit Do  ' exit Do..Loop
            End If
                       
            ' Calculate the CRC from the byte array
            CalcCRC32 abytData()
        
            ' Adjust pointers accordingly
            curFilePos = curFilePos + CCur(UBound(abytData) + 1)
            
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do   ' exit Do..Loop
            End If
    
            ' If nothing else to read then leave
            DoEvents
            If curAmtLeft < 1 Then
                Exit Do  ' exit Do..Loop
            End If
            
        Loop
        
        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            GoTo CRC32_File_CleanUp
        End If
    
        .API_CloseFile hFile    ' Close all files opened by this routine
    End With
    
    ' convert to hex
    If blnCheckSum Then
        ' A checksum must be a full 8 characters
        strCRC = Right$(ZEROES_8 & Hex$(mlngCRC Xor MINUS_ONE), 8)
    Else
        ' A non-checksum drops any leading zeroes
        strCRC = Hex$(mlngCRC Xor MINUS_ONE)
    End If
    
    
CRC32_File_CleanUp:
    objBigFiles.API_CloseFile hFile  ' Close all files opened by this routine
    Erase abytData()                 ' Always empty arrays when not needed
    Erase malngTable()
    Set objBigFiles = Nothing        ' Always free objects when not needed
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strCRC = vbNullString
    End If

    CRC32_File = strCRC   ' Return hex calculation
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

CRC32_File_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strCRC = vbNullString
    Resume CRC32_File_CleanUp

End Function

' ***************************************************************************
' Routine:       CRC32_String
'
' Description:   Returns a hex value representing a file.
' 
' Parameters:    abytData() - String data in byte array
'                blnCheckSum - [Optional] - TRUE - A checksum must return
'                             a full 8 characters.
'                             FALSE - A non-checksum removes leading zeroes
'
' Returns:       TRUE - Successful completion
'                FALSE - An error occurred
'
' Example:       CRC32_String('hello world') =  D4A1185  checksum = FALSE
'                CRC32_String('hello world') = 0D4A1185  checksum = TRUE
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 26-JUL-2004  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array for security purposes
' ***************************************************************************
Public Function CRC32_String(ByRef abytData() As Byte, _
                    Optional ByVal blnCheckSum As Boolean = False) As String

    Dim strCRC As String

    Const ROUTINE_NAME As String = "CRC32_String"

    On Error GoTo CRC32_String_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Is there any data to process?
    If Not CBool(IsArrayInitialized(abytData())) Then
        gblnStopProcessing = True   ' Set flag to stop processing
        Exit Function
    End If

    mcurMaxSize = CCur(UBound(abytData))  ' Get string length for progress display
    mcurAccrued = 0@                      ' Init accrued progress
    Initialize                            ' load table array
    strCRC = vbNullString
    
    ' Calculate the CRC from the byte array
    If Not CalcCRC32(abytData()) Then
        InfoMsg "Failed to calculate a CRC value." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        GoTo CRC32_String_CleanUp
    End If

    ' convert to hex
    If blnCheckSum Then
        ' A checksum must be a full 8 characters
        strCRC = Right$(ZEROES_8 & Hex$(mlngCRC Xor MINUS_ONE), 8)
    Else
        ' A non-checksum drops the leading zeroes
        strCRC = Hex$(mlngCRC Xor MINUS_ONE)
    End If
    
CRC32_String_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strCRC = vbNullString
    End If

    CRC32_String = strCRC   ' Return hex calculation
    
    Erase abytData()    ' Always empty arrays when not needed
    Erase malngTable()
    On Error GoTo 0     ' Nullify this error trap
    Exit Function

CRC32_String_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strCRC = vbNullString
    Resume CRC32_String_CleanUp

End Function


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Function CalcCRC32(ByRef abytData() As Byte) As Boolean

    Dim bytIdx      As Byte
    Dim bytChar     As Byte
    Dim lngIndex    As Long
    Dim lngTemp     As Long
    Dim lngProgress As Long

    On Error GoTo CalcCRC32_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo CalcCRC32_CleanUp
    End If

    For lngIndex = 0 To UBound(abytData)
        
        bytChar = abytData(lngIndex)
        bytIdx = (mlngCRC And &HFF) Xor bytChar
        lngTemp = ShiftRight8(mlngCRC)
        mlngCRC = lngTemp Xor malngTable(bytIdx)
    
        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If

    Next lngIndex
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo CalcCRC32_CleanUp
    End If

    ' 19-Feb-2012 Update progress bar. Separate
    '             steps for easier debugging
    DoEvents
    mcurAccrued = mcurAccrued + UBound(abytData)
    lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
    RaiseEvent CRCProgress(lngProgress)
    
    CalcCRC32 = True

CalcCRC32_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        CalcCRC32 = False
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

CalcCRC32_Error:
    CalcCRC32 = False
    Resume CalcCRC32_CleanUp

End Function

Private Function ShiftRight8(ByVal lngValue As Long) As Long
    
    ' Shift 32-bit long value to right by 8 bits
    ' Avoiding problem with sign bit
    
    Dim lngNewValue As Long
    
    lngNewValue = (lngValue And MAX_LONG) \ MAX_BYTE
    
    If (lngValue And MIN_LONG) <> 0 Then
        lngNewValue = lngNewValue Or &H800000
    End If
    
    ShiftRight8 = lngNewValue
    
End Function

' **************************************************************************
' Routine:       CRC_BlockSize
'
' Description:   Determines the size of the data to be processed.
'
' Parameters:    curAmtLeft - Amount of data left
'
' Returns:       New record size as a long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Function CRC_BlockSize(ByVal curAmtLeft As Currency) As Long

    ' Determine amount of data to process
    Select Case curAmtLeft
           Case Is >= KB_4: CRC_BlockSize = KB_4
           Case Else:       CRC_BlockSize = CLng(curAmtLeft)
    End Select
    
End Function

Private Sub Initialize()

    Dim strData  As String
    Dim lngIndex As Long
    Dim avntData As Variant

    mlngCRC = MINUS_ONE  ' Init accumulator to -1
    
    strData = vbNullString   ' Empty variables
    avntData = Empty         ' Always start with empty variants
    Erase malngTable()       ' Always start with empty arrays

    ReDim malngTable(MAX_BYTE)  ' Size table array
    
    strData = strData & "00000000 77073096 ee0e612c 990951ba 076dc419 706af48f e963a535 9e6495a3 "
    strData = strData & "0edb8832 79dcb8a4 e0d5e91e 97d2d988 09b64c2b 7eb17cbd e7b82d07 90bf1d91 "
    strData = strData & "1db71064 6ab020f2 f3b97148 84be41de 1adad47d 6ddde4eb f4d4b551 83d385c7 "
    strData = strData & "136c9856 646ba8c0 fd62f97a 8a65c9ec 14015c4f 63066cd9 fa0f3d63 8d080df5 "
    strData = strData & "3b6e20c8 4c69105e d56041e4 a2677172 3c03e4d1 4b04d447 d20d85fd a50ab56b "
    strData = strData & "35b5a8fa 42b2986c dbbbc9d6 acbcf940 32d86ce3 45df5c75 dcd60dcf abd13d59 "
    strData = strData & "26d930ac 51de003a c8d75180 bfd06116 21b4f4b5 56b3c423 cfba9599 b8bda50f "
    strData = strData & "2802b89e 5f058808 c60cd9b2 b10be924 2f6f7c87 58684c11 c1611dab b6662d3d "
    strData = strData & "76dc4190 01db7106 98d220bc efd5102a 71b18589 06b6b51f 9fbfe4a5 e8b8d433 "
    strData = strData & "7807c9a2 0f00f934 9609a88e e10e9818 7f6a0dbb 086d3d2d 91646c97 e6635c01 "
    strData = strData & "6b6b51f4 1c6c6162 856530d8 f262004e 6c0695ed 1b01a57b 8208f4c1 f50fc457 "
    strData = strData & "65b0d9c6 12b7e950 8bbeb8ea fcb9887c 62dd1ddf 15da2d49 8cd37cf3 fbd44c65 "
    strData = strData & "4db26158 3ab551ce a3bc0074 d4bb30e2 4adfa541 3dd895d7 a4d1c46d d3d6f4fb "
    strData = strData & "4369e96a 346ed9fc ad678846 da60b8d0 44042d73 33031de5 aa0a4c5f dd0d7cc9 "
    strData = strData & "5005713c 270241aa be0b1010 c90c2086 5768b525 206f85b3 b966d409 ce61e49f "
    strData = strData & "5edef90e 29d9c998 b0d09822 c7d7a8b4 59b33d17 2eb40d81 b7bd5c3b c0ba6cad "
    strData = strData & "edb88320 9abfb3b6 03b6e20c 74b1d29a ead54739 9dd277af 04db2615 73dc1683 "
    strData = strData & "e3630b12 94643b84 0d6d6a3e 7a6a5aa8 e40ecf0b 9309ff9d 0a00ae27 7d079eb1 "
    strData = strData & "f00f9344 8708a3d2 1e01f268 6906c2fe f762575d 806567cb 196c3671 6e6b06e7 "
    strData = strData & "fed41b76 89d32be0 10da7a5a 67dd4acc f9b9df6f 8ebeeff9 17b7be43 60b08ed5 "
    strData = strData & "d6d6a3e8 a1d1937e 38d8c2c4 4fdff252 d1bb67f1 a6bc5767 3fb506dd 48b2364b "
    strData = strData & "d80d2bda af0a1b4c 36034af6 41047a60 df60efc3 a867df55 316e8eef 4669be79 "
    strData = strData & "cb61b38c bc66831a 256fd2a0 5268e236 cc0c7795 bb0b4703 220216b9 5505262f "
    strData = strData & "c5ba3bbe b2bd0b28 2bb45a92 5cb36a04 c2d7ffa7 b5d0cf31 2cd99e8b 5bdeae1d "
    strData = strData & "9b64c2b0 ec63f226 756aa39c 026d930a 9c0906a9 eb0e363f 72076785 05005713 "
    strData = strData & "95bf4a82 e2b87a14 7bb12bae 0cb61b38 92d28e9b e5d5be0d 7cdcefb7 0bdbdf21 "
    strData = strData & "86d3d2d4 f1d4e242 68ddb3f8 1fda836e 81be16cd f6b9265b 6fb077e1 18b74777 "
    strData = strData & "88085ae6 ff0f6a70 66063bca 11010b5c 8f659eff f862ae69 616bffd3 166ccf45 "
    strData = strData & "a00ae278 d70dd2ee 4e048354 3903b3c2 a7672661 d06016f7 4969474d 3e6e77db "
    strData = strData & "aed16a4a d9d65adc 40df0b66 37d83bf0 a9bcae53 debb9ec5 47b2cf7f 30b5ffe9 "
    strData = strData & "bdbdf21c cabac28a 53b39330 24b4a3a6 bad03605 cdd70693 54de5729 23d967bf "
    strData = strData & "b3667a2e c4614ab8 5d681b02 2a6f2b94 b40bbe37 c30c8ea1 5a05df1b 2d02ef8d "
    
    avntData = Split(strData, Chr$(32))
    
    For lngIndex = 0 To UBound(avntData) - 1
        malngTable(lngIndex) = Val("&H" & avntData(lngIndex))
    Next lngIndex
    
    strData = vbNullString   ' Empty variables
    avntData = Empty         ' Always empty variants when not needed
    
End Sub

Private Sub Class_Initialize()
    StopProcessing = False
End Sub

Private Sub Class_Terminate()
    Erase malngTable()   ' Always empty arrays when not needed
End Sub
