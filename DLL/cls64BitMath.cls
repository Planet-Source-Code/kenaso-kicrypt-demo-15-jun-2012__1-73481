VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMath64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        cls64BitMath.cls
'
' Description:   Multiple 64-bit manipulating routines.  Some are not used.
'                I left the unused ones so I would not have to search for
'                them if I use this module in another application.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-JAN-2000  Kenneth Ives  kenaso@tx.rr.com
' 15-Mar-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated IsHexData() and IsBinaryData() routines.
' 01-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced Err.Raise with InfoMsg so programs will continue to
'              flow smoothly.
'              Added parameter testing in HexRepresentation() routine.
' 10-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Updated GetPrimeNumbers(), IsHexData() routines.
' 11-Jan-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated ConvertBase() and its supporting routines
' 08-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modifed LoadCubeRoots() and LoadSquareRoots() routines to
'              display output details as shown in their flowerboxes if user
'              hard codes the exact prime number indicated in the routine.
' 02-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modifed ConvertBase() routine.
' 10-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Updated logic testing in IsHexData() and IsBinaryData()
'              routines.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME   As String = "cls64BitMath"
  Private Const ZEROES_8      As String = "00000000"
  Private Const ZEROES_16     As String = ZEROES_8 & ZEROES_8
  Private Const ZEROES_32     As String = ZEROES_16 & ZEROES_16
  Private Const ZEROES_64     As String = ZEROES_32 & ZEROES_32
  Private Const HEX_LENGTH    As Long = 16
  Private Const MAX_LONG      As Long = &H7FFFFFFF  '  2147483647

' ***************************************************************************
' Public Enumerations
' ***************************************************************************
  Public Enum enumHEX_REPRESENTATION
      eCubeRoots    ' 0
      eSquareRoots  ' 1
  End Enum

  Public Enum enumHEX_RETURN_FORMAT
      e16Chars    ' 0
      eLeft8      ' 1
      eRight8     ' 2
  End Enum

' ***************************************************************************
' Private Enumerations
' ***************************************************************************
  Private Enum enum64_BaseValues
      e64_Binary = 2
      e64_Numeric = 10
      e64_Hex = 16
      e64_Max = 36
  End Enum

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function.
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
          (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

' ***************************************************************************
' Module Variables
'
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (String)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a str Digits
' Variable name:     mastrDigits
'
' ***************************************************************************
  Private mastrDigits()   As String   ' Bases digits
  Private malngValueIn()  As Long     ' Input number
  Private malngValueOut() As Long     ' Output number
  Private mlngFromBase    As Long     ' Input base
  Private mlngToBase      As Long     ' Output base


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let StopProcessing(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.StopProcessing = True
    gblnStopProcessing = blnData
    DoEvents
End Property

Public Property Get StopProcessing() As Boolean
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.StopProcessing
    StopProcessing = gblnStopProcessing
    DoEvents
End Property


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       w64HexToNumber
'
' Description:   Convert 16 character hex string (64-bit) to a big number in
'                string format.
'
' Parameters:    strHex - hex string
'
' Returns:       Big number in string format
'
' Example:       Input:   428a2f98d728ae22
'                Output:  4794697086780616226
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function w64HexToNumber(ByVal strHex As String) As String

    ' Called by w64HexMultiply()
    '           w64HexDivide()

    w64HexToNumber = ConvertBase(strHex, e64_Hex, e64_Numeric)

End Function

' ***************************************************************************
' Routine:       w64HexToBinary
'
' Description:   Convert 16 character hex string (64-bit) to binary string.
'
' Parameters:    strHex - hex string
'
' Returns:       Binary string
'
' Example:       Input:   428a2f98d728ae22
'                Output:  0100001010001010001011111001100011010111001010001010111000100010
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function w64HexToBinary(ByVal strHex As String) As String

    ' Called by w64Shift()
    '           w64Rotate()

    w64HexToBinary = ConvertBase(strHex, e64_Hex, e64_Binary)

End Function

' ***************************************************************************
' Routine:       w64BinaryToNumber
'
' Description:   Convert a binary string to a big number in string format.
'
' Parameters:    strBinary - Binary string
'
' Returns:       Big number in string format
'
' Example:       Input:   0100001010001010001011111001100011010111001010001010111000100010
'                Output:  4794697086780616226
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function w64BinaryToNumber(ByVal strBinary As String) As String

    w64BinaryToNumber = ConvertBase(strBinary, e64_Binary, e64_Numeric)

End Function

' ***************************************************************************
' Routine:       w64BinaryToHex
'
' Description:   Convert a binary string to a hex string.
'
' Parameters:    strBinary - Binary string
'
' Returns:       16 character hex string (64-bit)
'
' Example:       Input:   0100001010001010001011111001100011010111001010001010111000100010
'                Output:  428a2f98d728ae22
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function w64BinaryToHex(ByVal strBinary As String) As String

    w64BinaryToHex = ConvertBase(strBinary, e64_Binary, e64_Hex)

End Function

' ***************************************************************************
' Routine:       w64NumberToHex
'
' Description:   Convert a big (whole) number in string format to a sixteen
'                character hex string.
'
' Parameters:    strBinary - Binary string
'
' Returns:       16 character hex string (64-bit)
'
' Example:       Input:   4794697086780616226
'                Output:  428a2f98d728ae22
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function w64NumberToHex(ByVal strNumber As String) As String

    ' Called by w64HexMultiply()
    '           w64HexDivide()

    w64NumberToHex = ConvertBase(strNumber, e64_Numeric, e64_Hex)

End Function

' ***************************************************************************
' Routine:       w64NumberToBinary
'
' Description:   Convert a big (whole) number in string format to a binary
'                string.
'
' Parameters:    strBinary - Binary string
'
' Returns:       Binary string
'
' Example:       Input:   4794697086780616226
'                Output:  0100001010001010001011111001100011010111001010001010111000100010
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function w64NumberToBinary(ByVal strNumber As String) As String

    w64NumberToBinary = ConvertBase(strNumber, e64_Numeric, e64_Binary)

End Function

' ***************************************************************************
' Routine:       w64HexAdd
'
' Description:   Perform Hex addition on two 16 char hex values (64-bit).
'
' Special note:  Visit http://mathforum.org/library/drmath/view/55725.html
'
' Parameters:    strHex1 - First hex string (64-bit)
'                strHex2 - Second hex string (64-bit)
'
' Returns:       16 char hex string
'
'     Input:     b5c0fbcfec4d3b2f
'                428a2f98d728ae22
'                ----------------
'     Output:    f84b2b68c375e951
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine.  I kept this original routine because I needed
'              the speed.
' ***************************************************************************
Public Function w64HexAdd(ByVal strHex1 As String, _
                          ByVal strHex2 As String) As String

    Dim lngResult As Long
    Dim intIndex  As Integer
    Dim strOutput As String

    Const ROUTINE_NAME As String = "w64HexAdd"

    On Error GoTo w64HexAdd_Error

    strOutput = vbNullString
    lngResult = 0

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo w64HexAdd_CleanUp
    End If

    ' Make sure we have some valid data
    If Len(strHex1) = 0 Or StrComp(ZEROES_16, strHex1, vbTextCompare) = 0 Then
        If Len(strHex2) = 0 Or StrComp(ZEROES_16, strHex2, vbTextCompare) = 0 Then
            w64HexAdd = ZEROES_16
        Else
            w64HexAdd = strHex2
        End If

        Exit Function
    Else
        If Len(strHex2) = 0 Or StrComp(ZEROES_16, strHex2, vbTextCompare) = 0 Then
            w64HexAdd = strHex1
            Exit Function
        End If
    End If

    ' Step thru one value at a time and perform the hex addition
    For intIndex = Len(strHex1) To 1 Step -1
        lngResult = lngResult + _
                    Val("&H" & Mid$(strHex1, intIndex, 1)) + _
                    Val("&H" & Mid$(strHex2, intIndex, 1))

        strOutput = Hex$(lngResult Mod HEX_LENGTH) & strOutput
        lngResult = Int(lngResult / HEX_LENGTH)

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If

    Next intIndex

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strOutput = vbNullString
    Else
        ' If the final calculated value is greater than
        ' zero then prefix it to the output string
        If Fix(lngResult) > 0 Then
            strOutput = Hex$(lngResult) & strOutput
        End If

        strOutput = Right$(ZEROES_16 & strOutput, HEX_LENGTH)  ' Verify 16 character length
        strOutput = LCase$(strOutput)                            ' Convert hex string to lowercase
    End If

w64HexAdd_CleanUp:
    w64HexAdd = strOutput

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

w64HexAdd_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strOutput = vbNullString
    gblnStopProcessing = True
    Resume w64HexAdd_CleanUp

End Function
' ***************************************************************************
' Routine:       w64HexSubtract
'
' Description:   Perform Hex subtraction on two 16 char hex values (64-bit).
'
' Special note:  Visit http://mathforum.org/library/drmath/view/55725.html
'
' Parameters:    strHex1 - First hex string (64-bit)
'                strHex2 - Second hex string (64-bit)
'
' Returns:       16 char hex string
'
'     Input:     b5c0fbcfec4d3b2f
'                428a2f98d728ae22
'                ----------------
'     Output:    7336cc3715248d0d
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine.  I kept this original routine because I needed
'              the speed.
' ***************************************************************************
Public Function w64HexSubtract(ByVal strHex1 As String, _
                               ByVal strHex2 As String) As String

    Dim lngResult  As Long
    Dim lngValue1  As Long
    Dim lngValue2  As Long
    Dim lngCarry   As Long
    Dim intIndex   As Integer
    Dim strOutput  As String

    Const ROUTINE_NAME As String = "w64HexSubtract"

    On Error GoTo w64HexSubtract_Error

    strOutput = vbNullString

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Make sure we have some valid data
    If Len(strHex1) = 0 Or StrComp(ZEROES_16, strHex1, vbTextCompare) = 0 Then
        If Len(strHex2) = 0 Or StrComp(ZEROES_16, strHex2, vbTextCompare) = 0 Then
            w64HexSubtract = ZEROES_16
        Else
            w64HexSubtract = strHex2
        End If

        Exit Function
    Else
        If Len(strHex2) = 0 Or StrComp(ZEROES_16, strHex2, vbTextCompare) = 0 Then
            w64HexSubtract = strHex1
            Exit Function
        End If
    End If

    ' Step thru one value at a time and perform the hex addition
    For intIndex = Len(strHex1) To 1 Step -1

        lngValue1 = Val("&H" & Mid$(strHex1, intIndex, 1))
        lngValue2 = Val("&H" & Mid$(strHex2, intIndex, 1))

        If lngCarry > 0 Then
            lngValue1 = lngValue1 - 1
        End If

        If lngValue1 < lngValue2 Then
            lngCarry = HEX_LENGTH
        Else
            lngCarry = 0
        End If

        lngResult = (lngValue1 + lngCarry) - lngValue2
        strOutput = Hex$(lngResult Mod HEX_LENGTH) & strOutput

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If

    Next intIndex

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strOutput = vbNullString
    Else
        strOutput = Right$(ZEROES_16 & strOutput, HEX_LENGTH)  ' Verify 16 character length
        strOutput = LCase$(strOutput)                            ' Convert hex string to lowercase
    End If

w64HexSubtract_CleanUp:
    w64HexSubtract = strOutput

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

w64HexSubtract_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strOutput = vbNullString
    gblnStopProcessing = True
    Resume w64HexSubtract_CleanUp

End Function

' ***************************************************************************
' Routine:       w64HexMultiply
'
' Description:   Perform Hex multiplication on two 16 char hex values (64-bit).
'
' Parameters:    strHex1 - First hex string (64-bit)
'                strHex2 - Second hex string (64-bit)
'
' Returns:       Left-most 16 chars of the newly calculated hex string
'
'     Input:     428a2f98d728ae22
'                b5c0fbcfec4d3b2f
'                ----------------
'     Output:    2f3ddcbe4a93effec7c951e798d3ce3e  Complete output
'     Return:    2f3ddcbe4a93effe                  Return first 16 chars only
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function w64HexMultiply(ByVal strHex1 As String, _
                               ByVal strHex2 As String) As String

    Dim strNumber1 As String
    Dim strNumber2 As String
    Dim strOutput  As String

    Const ROUTINE_NAME As String = "w64HexSubtract"

    On Error GoTo w64HexMultiply_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Make sure we have some valid data
    If Len(strHex1) = 0 Or StrComp(ZEROES_16, strHex1, vbTextCompare) = 0 Or _
       Len(strHex2) = 0 Or StrComp(ZEROES_16, strHex2, vbTextCompare) = 0 Then

        ' cannot multiply with either value being all zeroes
        w64HexMultiply = ZEROES_16
        Exit Function
    End If

    strNumber1 = w64HexToNumber(strHex1)              ' convert hex values to large number
    strNumber2 = w64HexToNumber(strHex2)
    strOutput = BigMultiply(strNumber1, strNumber2)   ' perform multiplication

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strOutput = vbNullString
    Else
        strOutput = w64NumberToHex(strOutput)                    ' convert computation to hex
        strOutput = Left$(strOutput, HEX_LENGTH)               ' Save first 16 characters
        strOutput = Right$(ZEROES_16 & strOutput, HEX_LENGTH)  ' Verify 16 character length
        strOutput = LCase$(strOutput)                            ' Convert hex string to lowercase
    End If

w64HexMultiply_CleanUp:
    w64HexMultiply = strOutput

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

w64HexMultiply_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strOutput = vbNullString
    gblnStopProcessing = True
    Resume w64HexMultiply_CleanUp

End Function

' ***************************************************************************
' Routine:       w64HexDivide
'
' Description:   Perform Hex division on two 16 char hex values (64-bit).
'
' Parameters:    strHex1 - First hex string
'                strHex2 - Second hex string
'
' Returns:       16 char hex string
'
'   Example:     00000005db89fc59 = 428a2f98d728ae22 \ 1a2b
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function w64HexDivide(ByVal strHex1 As String, _
                             ByVal strHex2 As String) As String

    Dim strDividend As String
    Dim strDivisor  As String
    Dim strOutput   As String

    Const ROUTINE_NAME As String = "w64HexDivide"

    On Error GoTo w64HexDivide_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Make sure we have some valid data
    If Len(strHex1) = 0 Or StrComp(ZEROES_16, strHex1, vbTextCompare) = 0 Or _
       Len(strHex2) = 0 Or StrComp(ZEROES_16, strHex2, vbTextCompare) = 0 Then

        InfoMsg "Cannot divide by zero." & vbNewLine & vbNewLine & _
                "Dividend:   " & strHex1 & vbNewLine & _
                "Divisor:    " & strHex2 & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        GoTo w64HexDivide_CleanUp
    End If

    strDividend = w64HexToNumber(strHex1)   ' convert hex values to large number
    strDivisor = w64HexToNumber(strHex2)

    ' Test the numbers
    If StrComp(strDividend, strDivisor, vbBinaryCompare) = -1 Then

        InfoMsg "The Divisor cannot exceed the Dividend." & vbNewLine & vbNewLine & _
                "Dividend:   " & strHex1 & " (" & strDividend & ")" & vbNewLine & _
                "Divisor:    " & strHex2 & " (" & strDivisor & ")" & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        GoTo w64HexDivide_CleanUp
    End If

    strOutput = BigDivide(strDividend, strDivisor)  ' perform the calculation

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strOutput = vbNullString
    Else
        strOutput = w64NumberToHex(strOutput)                    ' convert computation to hex
        strOutput = Left$(strOutput, HEX_LENGTH)               ' Save first 16 characters
        strOutput = Right$(ZEROES_16 & strOutput, HEX_LENGTH)  ' Verify 16 character length
        strOutput = LCase$(strOutput)                            ' Convert hex string to lowercase
    End If

w64HexDivide_CleanUp:
    w64HexDivide = strOutput

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

w64HexDivide_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strOutput = vbNullString
    gblnStopProcessing = True
    Resume w64HexDivide_CleanUp

End Function

' ***************************************************************************
' Routine:       w64Shift
'
'' Description:  Shifts the bits to the right/Left the specified number of
'                positions and returns the new value.  Bits "falling off"
'                the direction of the shift do not wrap around.  Fill bits
'                coming in from the opposite side are zeros.  Some common
'                languages like C/C++ or Java have an operator for this
'                job: ">>" or "<<".
'
' Parameters:    strHex - hex value to be manipulated
'                lngBitShift - Number of shift positions
'                              Positive value = left shift
'                              Negative value = right shift
'                blnReturnAsHex - [Optional] how the return data is formatted
'                                TRUE - Data returned in hex format (Default)
'                                FALSE - Data returned in binary format
'
' Returns:       New manipulated hex string
'
'                  Hex                          Binary
' Original:  1234567890abcdef  0001001000110100010101100111100010010000101010111100110111101111
'   Left 5:  468acf121579bde0  0100011010001010110011110001001000010101011110011011110111100000
'  Right 5:  0091a2b3c4855e6f  0000000010010001101000101011001111000100100001010101111001101111
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-May-2006  Dermot Hogan  dermot@bitwisemag.com
'              http://www.bitwisemag.com/2/Bit-Shifting-in-Visual-Basic-6
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function w64Shift(ByVal strHex As String, _
                         ByVal lngBitShift As Long, _
                Optional ByVal blnReturnAsHex As Boolean = True) As String

    Dim lngIndex    As Long
    Dim lngTmpLeft  As Long
    Dim lngHexLeft  As Long
    Dim lngHexRight As Long
    Dim lngTmpRight As Long

    Const ROUTINE_NAME As String = "w64Shift"

    On Error GoTo w64Shift_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Test the amount of bit shifting
    Select Case lngBitShift

           Case 0:        w64Shift = strHex      ' Return original value if no shift positions
           Case Is < -63: w64Shift = ZEROES_16   ' Return zeroes if too many negative shift positions
           Case Is > 63:  w64Shift = ZEROES_16   ' Return zeroes if too many positive shift positions

           Case Is > 0
                ' A positive bit shift value means shift left
                '
                ' Prepare data prior to bit movement.
                ' Split hex value into two long integers.
                strHex = Right$(ZEROES_16 & strHex, HEX_LENGTH)
                lngHexLeft = "&H" & Mid$(strHex, 1, 8)
                lngHexRight = "&H" & Mid$(strHex, 9, 8)

                ' We need to shift left by one bit but
                ' take account of overflow error from
                ' VB.  So mask accordingly.
                For lngIndex = 1 To lngBitShift

                    ' Remember the 2 most significant bits
                    lngTmpLeft = lngHexRight And &HC0000000
                    lngTmpRight = lngHexLeft And &HC0000000

                    ' To prevent an overflow from the transition from
                    ' a positive to a negative quantity when you shift
                    ' to the left.  Mask the top two bits of a 32 bit
                    ' word before multiplying (the code shifts a single
                    ' bit to the left).
                    lngHexRight = lngHexRight And &H3FFFFFFF
                    lngHexRight = lngHexRight * 2

                    lngHexLeft = lngHexLeft And &H3FFFFFFF
                    lngHexLeft = lngHexLeft * 2

                    ' Set the Most Significant Bit (MSB)
                    If lngTmpRight And &H40000000 Then
                        lngHexLeft = lngHexLeft Or &H80000000
                    Else
                        lngHexLeft = lngHexLeft And &H7FFFFFFF
                    End If

                    ' If the number was negative, then add 1
                    ' if bit 30 was set, then set the sign bit
                    If lngTmpLeft And &H80000000 Then
                        lngHexLeft = lngHexLeft Or 1
                    End If

                    ' Reset the Most Significant Bit (MSB)
                    If lngTmpLeft And &H40000000 Then
                        lngHexRight = lngHexRight Or &H80000000
                    Else
                        lngHexRight = lngHexRight And &H7FFFFFFF
                    End If

                Next lngIndex

                ' rejoin the long integers and convert back to hex
                strHex = Right$(ZEROES_8 & Hex$(lngHexLeft), 8) & _
                         Right$(ZEROES_8 & Hex$(lngHexRight), 8)

                ' Return manipulated value
                If blnReturnAsHex Then
                    w64Shift = LCase$(strHex)          ' return hex string
                Else
                    w64Shift = w64HexToBinary(strHex)  ' return as binary string
                End If

           Case Is < 0
                ' A negative bit shift value means shift right
                '
                ' Prepare data prior to bit movement.
                ' Split hex value into two long integers.
                strHex = Right$(ZEROES_16 & strHex, HEX_LENGTH)
                lngHexLeft = "&H" & Mid$(strHex, 1, 8)
                lngHexRight = "&H" & Mid$(strHex, 9, 8)

                ' We need to shift right by one bit but
                ' take account of overflow error from
                ' VB.  So mask accordingly.
                For lngIndex = lngBitShift To -1

                    ' Save the sign bit
                    lngTmpLeft = lngHexRight And &H80000001
                    lngTmpRight = lngHexLeft And &H80000001

                    ' To prevent an overflow from the transition from
                    ' a positive to a negative quantity when you shift
                    ' to the left.  Mask the top two bits of a 32 bit
                    ' word before dividing (the code shifts a single
                    ' bit to the right).
                    lngHexRight = lngHexRight And &H7FFFFFFF
                    lngHexRight = lngHexRight \ 2

                    lngHexLeft = lngHexLeft And &H7FFFFFFF
                    lngHexLeft = lngHexLeft \ 2

                    If lngTmpRight And 1 Then
                        lngHexRight = lngHexRight Or &H80000000
                    End If

                    ' Set or clear the old sign bit
                    If lngTmpLeft And &H80000000 Then
                        lngHexRight = lngHexRight Or &H40000000
                    Else
                        lngHexRight = lngHexRight And &HBFFFFFFF
                    End If

                    ' if the number was negative, then re-insert
                    ' the bit if bit 0 was set, then set the sign bit
                    If lngTmpRight And &H80000000 Then
                        lngHexLeft = lngHexLeft Or &H40000000
                    Else
                        lngHexLeft = lngHexLeft And &HBFFFFFFF
                    End If

                Next lngIndex

                ' rejoin the long integers and convert back to hex
                strHex = Right$(ZEROES_8 & Hex$(lngHexLeft), 8) & _
                         Right$(ZEROES_8 & Hex$(lngHexRight), 8)

                ' Return manipulated value
                If blnReturnAsHex Then
                    w64Shift = LCase$(strHex)          ' return hex string
                Else
                    w64Shift = w64HexToBinary(strHex)  ' return as binary string
                End If
    End Select

w64Shift_CleanUp:
    On Error GoTo 0  ' Nullify this error trap
    Exit Function

w64Shift_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    w64Shift = vbNullString
    gblnStopProcessing = True
    Resume w64Shift_CleanUp

End Function

' ***************************************************************************
' Routine:       w64Rotate
'
' Description:   Shifts the bits to the right/Left the specified number of
'                positions and returns the new value.  Bits "falling off"
'                the direction of the shift wrap around to the opposite end.
'                Some common languages like C/C++ or Java have an operator
'                for this job: ">>>" or "<<<".  Sometimes referred to as
'                circular bit shifting or bit rotation.
'
' Parameters:    strHex - hex value to be manipulated
'                lngBitShift - Number of shift positions
'                              Positive value = left rotation
'                              Negative value = right rotation
'                blnReturnAsHex - [Optional] how the return data is formatted
'                                TRUE - Data returned in hex format (Default)
'                                FALSE - Data returned in binary format
'
' Returns:       New manipulated hex string
'
'                  Hex                          Binary
' Original:  1234567890abcdef  0001001000110100010101100111100010010000101010111100110111101111
'   Left 5:  468acf121579bde2  0100011010001010110011110001001000010101011110011011110111100010
'  Right 5:  7891a2b3c4855e6f  0111100010010001101000101011001111000100100001010101111001101111
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-May-2006  Dermot Hogan  dermot@bitwisemag.com
'              http://www.bitwisemag.com/2/Bit-Shifting-in-Visual-Basic-6
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function w64Rotate(ByVal strHex As String, _
                          ByVal lngBitShift As Long, _
                 Optional ByVal blnReturnAsHex As Boolean = True) As String

    Dim lngIndex    As Long
    Dim lngTmpLeft  As Long
    Dim lngHexLeft  As Long
    Dim lngHexRight As Long
    Dim lngTmpRight As Long

    Const ROUTINE_NAME As String = "w64Rotate"

    On Error GoTo w64Rotate_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Test the amount of bit shifting
    Select Case lngBitShift

           Case 0:        w64Rotate = strHex      ' Return original value if no shift positions
           Case Is < -63: w64Rotate = ZEROES_16   ' Return zeroes if too many negative shift positions
           Case Is > 63:  w64Rotate = ZEROES_16   ' Return zeroes if too many positive shift positions

           Case Is > 0
                ' A positive bit shift value means rotate left
                '
                ' Prepare data prior to bit movement.
                ' Split hex value into two long integers.
                strHex = Right$(ZEROES_16 & strHex, HEX_LENGTH)
                lngHexLeft = "&H" & Mid$(strHex, 1, 8)
                lngHexRight = "&H" & Mid$(strHex, 9, 8)

                ' We need to shift left by one bit but
                ' take account of overflow error from
                ' VB.  So mask accordingly.
                For lngIndex = 1 To lngBitShift

                    ' Save the sign bit
                    lngTmpLeft = lngHexRight And &HC0000000
                    lngTmpRight = lngHexLeft And &HC0000000

                    ' To prevent an overflow from the transition from
                    ' a positive to a negative quantity when you shift
                    ' to the left.  Mask the top two bits of a 32 bit
                    ' word before multiplying (the code shifts a single
                    ' bit to the left).
                    lngHexRight = lngHexRight And &H3FFFFFFF
                    lngHexRight = lngHexRight * 2

                    lngHexLeft = lngHexLeft And &H3FFFFFFF
                    lngHexLeft = lngHexLeft * 2

                    ' Set the Most Significant Bit (MSB)
                    If lngTmpRight And &H40000000 Then
                        lngHexLeft = lngHexLeft Or &H80000000
                    Else
                        lngHexLeft = lngHexLeft And &H7FFFFFFF
                    End If

                    ' Rotate Most Significant Bit (MSB) into
                    ' Least Significant Bit (LSB)
                    If lngTmpRight And &H80000000 Then
                        lngHexRight = lngHexRight Or 1
                    End If

                    If lngTmpLeft And &H80000000 Then
                        lngHexLeft = lngHexLeft Or 1
                    End If

                    ' Reset the Most Significant Bit (MSB)
                    If lngTmpLeft And &H40000000 Then
                        lngHexRight = lngHexRight Or &H80000000
                    Else
                        lngHexRight = lngHexRight And &H7FFFFFFF
                    End If

                Next lngIndex

                ' rejoin the long integers and convert back to hex
                strHex = Right$(ZEROES_8 & Hex$(lngHexLeft), 8) & _
                         Right$(ZEROES_8 & Hex$(lngHexRight), 8)

                ' Return manipulated value
                If blnReturnAsHex Then
                    w64Rotate = LCase$(strHex)           ' return hex string
                Else
                    w64Rotate = w64HexToBinary(strHex)   ' return as binary string
                End If

           Case Is < 0
                ' A negative bit shift value means rotate right
                '
                ' Prepare data prior to bit movement.
                ' Split hex value into two long integers.
                strHex = Right$(ZEROES_16 & strHex, HEX_LENGTH)
                lngHexLeft = "&H" & Mid$(strHex, 1, 8)
                lngHexRight = "&H" & Mid$(strHex, 9, 8)

                ' We need to shift right by one bit but
                ' take account of overflow error from
                ' VB.  So mask accordingly.
                For lngIndex = lngBitShift To -1

                    ' Save the sign bit
                    lngTmpLeft = lngHexRight And &H80000001
                    lngTmpRight = lngHexLeft And &H80000001

                    ' To prevent an overflow from the transition from
                    ' a positive to a negative quantity when you shift
                    ' to the left.  Mask the top two bits of a 32 bit
                    ' word before dividing (the code shifts a single
                    ' bit to the right).
                    lngHexRight = lngHexRight And &H7FFFFFFF
                    lngHexRight = lngHexRight \ 2

                    lngHexLeft = lngHexLeft And &H7FFFFFFF
                    lngHexLeft = lngHexLeft \ 2

                    ' Rotate Most Significant Bit (MSB) into
                    ' Least Significant Bit (LSB)
                    If lngTmpLeft And 1 Then
                        lngHexLeft = lngHexLeft Or &H80000000
                    End If

                    If lngTmpRight And 1 Then
                        lngHexRight = lngHexRight Or &H80000000
                    End If

                    ' Set or clear the old sign bit
                    If lngTmpLeft And &H80000000 Then
                        lngHexRight = lngHexRight Or &H40000000
                    Else
                        lngHexRight = lngHexRight And &HBFFFFFFF
                    End If

                    ' if the number was negative, then re-insert
                    ' the bit if bit 0 was set, then set the sign bit
                    If lngTmpRight And &H80000000 Then
                        lngHexLeft = lngHexLeft Or &H40000000
                    Else
                        lngHexLeft = lngHexLeft And &HBFFFFFFF
                    End If

                Next lngIndex

                ' rejoin the long integers and convert back to hex
                strHex = Right$(ZEROES_8 & Hex$(lngHexLeft), 8) & _
                         Right$(ZEROES_8 & Hex$(lngHexRight), 8)

                ' Return manipulated value
                If blnReturnAsHex Then
                    w64Rotate = LCase$(strHex)          ' return hex string
                Else
                    w64Rotate = w64HexToBinary(strHex)  ' return as binary string
                End If
    End Select

w64Rotate_CleanUp:
    On Error GoTo 0  ' Nullify this error trap
    Exit Function

w64Rotate_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    w64Rotate = vbNullString
    gblnStopProcessing = True
    Resume w64Rotate_CleanUp

End Function

' ***************************************************************************
' Routine:       w64Hex_NOT
'
' Description:   NOT expression used for bit comparison does not exist
'                in Visual Basic.
'
'                This function looks at the binary representation of
'                the values of the expression and does a bitwise negation
'                operation on it. The result of this operation behaves
'                as shown in this example:
'
'                     X = (NOT 1011)
'
'                         1011
'                        ------
'                     X = 0100    The bits are reversed
'
' Special Note:  This is the same as the Visual Basic bit function NOT except
'                that it works with 64-bit binary strings.
'
' Input:    428a2f98d728ae22  0100001010001010001011111001100011010111001010001010111000100010
' Output:   bd75d06728d751dd  1011110101110101110100000110011100101000110101110101000111011101
'
' Parameters:    strHex - Hex string to be evaluated
'
' Returns:       New hex data string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function w64Hex_NOT(ByVal strHex As String) As String

    Dim strOutput As String
    Dim lngValue  As Long
    Dim intIndex  As Integer

    Const ROUTINE_NAME As String = "w64Hex_NOT"

    On Error GoTo w64Hex_NOT_Error

    DoEvents
    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Initialize variables
    strHex = Right$(ZEROES_16 & strHex, HEX_LENGTH)
    strOutput = vbNullString

    For intIndex = 1 To HEX_LENGTH Step 2

        lngValue = (Not Val("&H" & Mid$(strHex, intIndex, 2)))
        strOutput = strOutput & Right$("00" & Hex$(lngValue), 2)

    Next intIndex

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strOutput = vbNullString
    Else
        strOutput = LCase$(strOutput)   ' return hex string
    End If

w64Hex_NOT_CleanUp:
    w64Hex_NOT = strOutput

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

w64Hex_NOT_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strOutput = vbNullString
    gblnStopProcessing = True
    Resume w64Hex_NOT_CleanUp

End Function

' ***************************************************************************
' Routine:       w64Hex_AND
'
' Description:   The And operator also performs a bitwise comparison of
'                identically positioned bits in two numeric expressions
'                and sets the corresponding bit in result according to
'                the following example:
'
'                    X = 1011 AND 1001
'
'                         1011
'                         1001
'                        ------
'                     X = 1001    Different bits become "0"
'
' Special Note:  This is the same as the Visual Basic bit function AND except
'                that it works with 64-bit binary strings.
'
' Input:    428a2f98d728ae22  0100001010001010001011111001100011010111001010001010111000100010
'           27b70a8546d22ffc  0010011110110111000010101000010101000110110100100010111111111100
'          -----------------  ----------------------------------------------------------------
' Output:   02820a8046002e20  0000001010000010000010101000000001000110000000000010111000100000
'
' Parameters:    strHex1 - Hex string to be evaluated
'                strHex2 - Hex string to be evaluated
'
' Returns:       New hex data string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function w64Hex_AND(ByVal strHex1 As String, _
                           ByVal strHex2 As String) As String

    Dim strOutput As String
    Dim lngValue  As Long
    Dim intIndex  As Integer

    Const ROUTINE_NAME As String = "w64Hex_AND"

    On Error GoTo w64Hex_AND_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Initialize variables
    strHex1 = Right$(ZEROES_16 & strHex1, HEX_LENGTH)
    strHex2 = Right$(ZEROES_16 & strHex2, HEX_LENGTH)
    strOutput = vbNullString

    ' Perform the "AND" binary comparisons
    For intIndex = 1 To HEX_LENGTH Step 2

        lngValue = Val("&H" & Mid$(strHex1, intIndex, 2) & "&") And _
                   Val("&H" & Mid$(strHex2, intIndex, 2) & "&")
        strOutput = strOutput & Right$("00" & Hex$(lngValue), 2)

    Next intIndex

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strOutput = vbNullString
    Else
        strOutput = LCase$(strOutput)   ' return hex string
    End If

w64Hex_AND_CleanUp:
    w64Hex_AND = strOutput

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

w64Hex_AND_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strOutput = vbNullString
    gblnStopProcessing = True
    Resume w64Hex_AND_CleanUp

End Function

' ***************************************************************************
' Routine:       w64Hex_OR
'
' Description:   The OR operator also performs a bitwise comparison of
'                identically positioned bits in two numeric expressions
'                and sets the corresponding bit in result according to
'                the following example:
'
'                     X = 1011 OR 1001
'
'                         1011
'                         1001
'                        ------
'                     X = 1011    The higher bit value wins
'
' Special Note:  This is the same as the Visual Basic bit function OR except
'                that it works with 64-bit binary strings.
'
' Input:    428a2f98d728ae22  0100001010001010001011111001100011010111001010001010111000100010
'           27b70a8546d22ffc  0010011110110111000010101000010101000110110100100010111111111100
'          -----------------  ----------------------------------------------------------------
' Output:   67bf2f9dd7faaffe  0110011110111111001011111001110111010111111110101010111111111110
'
' Parameters:    strHex1 - Hex string to be evaluated
'                strHex2 - Hex string to be evaluated
'
' Returns:       New hex data string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function w64Hex_OR(ByVal strHex1 As String, _
                          ByVal strHex2 As String) As String

    Dim strOutput As String
    Dim lngValue  As Long
    Dim intIndex  As Integer

    Const ROUTINE_NAME As String = "w64Hex_OR"

    On Error GoTo w64Hex_OR_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Initialize variables
    strHex1 = Right$(ZEROES_16 & strHex1, HEX_LENGTH)
    strHex2 = Right$(ZEROES_16 & strHex2, HEX_LENGTH)
    strOutput = vbNullString

    ' Perform the "Or" binary comparisons
    For intIndex = 1 To HEX_LENGTH Step 2

        lngValue = Val("&H" & Mid$(strHex1, intIndex, 2) & "&") Or _
                   Val("&H" & Mid$(strHex2, intIndex, 2) & "&")
        strOutput = strOutput & Right$("00" & Hex$(lngValue), 2)

    Next intIndex

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strOutput = vbNullString
    Else
        strOutput = LCase$(strOutput)   ' return hex string
    End If

w64Hex_OR_CleanUp:
    w64Hex_OR = strOutput

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

w64Hex_OR_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strOutput = vbNullString
    gblnStopProcessing = True
    Resume w64Hex_OR_CleanUp

End Function

' ***************************************************************************
' Routine:       w64Hex_XoR
'
' Description:   The Xor operator performs as both a logical and bitwise
'                operator. A bit-wise comparison of two expressions using
'                exclusive-or logic to form the result, as shown in the
'                following example:
'
'                         X = 1001 XoR 1101
'
'                         1001
'                         1101
'                        ------
'                     X = 0100   Same values become "0"
'
' Special Note:  This is the same as the Visual Basic bit function XoR except
'                that it works with 64-bit binary strings.
'
' Input:    428a2f98d728ae22  0100001010001010001011111001100011010111001010001010111000100010
'           27b70a8546d22ffc  0010011110110111000010101000010101000110110100100010111111111100
'          -----------------  ----------------------------------------------------------------
' Output:   653d251d91fa81de  0110010100111101001001010001110110010001111110101000000111011110
'
' Parameters:    strHex1 - Hex string to be evaluated
'                strHex2 - Hex string to be evaluated
'
' Returns:       New hex data string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function w64Hex_XoR(ByVal strHex1 As String, _
                           ByVal strHex2 As String) As String

    Dim strOutput As String
    Dim lngValue  As Long
    Dim intIndex  As Integer

    Const ROUTINE_NAME As String = "w64Hex_XoR"

    On Error GoTo w64Hex_XoR_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Initialize variables
    strHex1 = Right$(ZEROES_16 & strHex1, HEX_LENGTH)
    strHex2 = Right$(ZEROES_16 & strHex2, HEX_LENGTH)
    strOutput = vbNullString

    ' Perform the "XoR" binary comparisons
    For intIndex = 1 To HEX_LENGTH Step 2

        lngValue = Val("&H" & Mid$(strHex1, intIndex, 2) & "&") Xor _
                   Val("&H" & Mid$(strHex2, intIndex, 2) & "&")
        strOutput = strOutput & Right$("00" & Hex$(lngValue), 2)

    Next intIndex

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strOutput = vbNullString
    Else
        strOutput = LCase$(strOutput)   ' return hex string
    End If

w64Hex_XoR_CleanUp:
    w64Hex_XoR = strOutput

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

w64Hex_XoR_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strOutput = vbNullString
    gblnStopProcessing = True
    Resume w64Hex_XoR_CleanUp

End Function

' ***************************************************************************
' Routine:       BigAdd
'
' Description:   A solution for handling large numbers (up to 32768 digits)
'                in Visual Basic 6.  Adds strNumber2 to strNumber1 and
'                returns the result in a string.
'
' Example:       13096744586834688815
'                 4794697086780616226
'                --------------------
'                17891441673615305041
'
' Parameters:    strNumber1 - Numeric string
'                strNumber1 - Numeric string to be added to strNumber1
'
' Returns:       New numeric string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 23-Nov-2002  Philipp Emanuel Weidmann   http://www.philippsworld.net/
'              Author of three of the "Big" routines
' 29-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified variable names and renamed routine for easier
'              maintenance.  Added error trapping.
' ***************************************************************************
Public Function BigAdd(ByVal strNumber1 As String, _
                       ByVal strNumber2 As String) As String

    Dim intIndex       As Integer
    Dim intAddBuffer   As Integer
    Dim intTempDigit1  As Integer
    Dim intTempDigit2  As Integer
    Dim strCalcResult  As String

    Const ROUTINE_NAME As String = "BigAdd"

    On Error GoTo BigAdd_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    If Not IsNumeric(strNumber1) Or Not IsNumeric(strNumber2) Then
        InfoMsg "Values to be added must be numeric." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        Exit Function
    End If

    ' Fill up the shorter number with zeroes at the beginning to make them equal in length
    If Len(strNumber1) > Len(strNumber2) Then
        strNumber2 = String$(Len(strNumber1) - Len(strNumber2), "0") & strNumber2

    ElseIf Len(strNumber1) < Len(strNumber2) Then
        strNumber1 = String$(Len(strNumber2) - Len(strNumber1), "0") & strNumber1
    End If

    ' Prefix one zero at the beginning of each number
    ' to make sure no digit gets lost when the sum of
    ' two digits is greater than 10 (see below)
    strNumber1 = "0" & strNumber1
    strNumber2 = "0" & strNumber2

    For intIndex = Len(strNumber1) To 1 Step -1
        ' Add the numbers digit by digit to one another
        intTempDigit1 = CInt(Mid$(strNumber1, intIndex, 1))
        intTempDigit2 = CInt(Mid$(strNumber2, intIndex, 1))

        If intTempDigit1 + intTempDigit2 + intAddBuffer >= 10 Then

            strCalcResult = CStr(intTempDigit1 + intTempDigit2 + intAddBuffer - 10) & strCalcResult

            ' intAddBuffer contains 1 if the sum of two digits is greater than 10
            intAddBuffer = 1
        Else
            strCalcResult = CStr(intTempDigit1 + intTempDigit2 + intAddBuffer) & strCalcResult
            intAddBuffer = 0
        End If

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If

    Next intIndex

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strCalcResult = "0"
    Else
        strCalcResult = DataTrim(strCalcResult)
    End If

BigAdd_CleanUp:
    BigAdd = strCalcResult

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

BigAdd_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strCalcResult = "0"
    gblnStopProcessing = True
    Resume BigAdd_CleanUp

End Function

' ***************************************************************************
' Routine:       BigSubtract
'
' Description:   A solution for handling large numbers (up to 32768 digits)
'                in Visual Basic 6.  Subtracts strNumber2 from strNumber1
'                and returns the result in a string.
'
' Example:       13096744586834688815
'                 4794697086780616226
'                --------------------
'                 8302047500054072589
'
' Parameters:    strNumber1 - Numeric string
'                strNumber1 - Numeric string to be Subtracted
'
' Returns:       New numeric string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 23-Nov-2002  Philipp Emanuel Weidmann   http://www.philippsworld.net/
'              Author of three of the "Big" routines
' 29-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified variable names and renamed routine for easier
'              maintenance.  Added error trapping.
' ***************************************************************************
Public Function BigSubtract(ByVal strNumber1 As String, _
                            ByVal strNumber2 As String) As String

    Dim intIndex          As Integer
    Dim intTempDigit1     As Integer
    Dim intTempDigit2     As Integer
    Dim intSubtractBuffer As Integer
    Dim strCalcResult     As String
    Dim strTempBuffer     As String

    Const ROUTINE_NAME As String = "BigSubtract"

    On Error GoTo BigSubtract_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    If Not IsNumeric(strNumber1) Or Not IsNumeric(strNumber2) Then
        InfoMsg "Values to be subtracted must be numeric." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        Exit Function
    End If

    ' If the numbers are equal, the result is zero
    If BigCompare(strNumber1, strNumber2) = 0 Then
        BigSubtract = "0"
        Exit Function
    End If

    ' Fill up the shorter number with zeroes at the beginning to make them equal in length
    If Len(strNumber1) > Len(strNumber2) Then
        strNumber2 = String$(Len(strNumber1) - Len(strNumber2), "0") & strNumber2

    ElseIf Len(strNumber1) < Len(strNumber2) Then
        strNumber1 = String$(Len(strNumber2) - Len(strNumber1), "0") & strNumber1
    End If

    ' Put the larger number in strNumber1 (otherwise digit-by-digit subtraction won't work)
    If BigCompare(strNumber1, strNumber2) = 2 Then
        SwapData strNumber1, strNumber2
    End If

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    For intIndex = Len(strNumber1) To 1 Step -1

        ' Subtract the numbers digit by digit from one another
        intTempDigit1 = CInt(Mid$(strNumber1, intIndex, 1))
        intTempDigit2 = CInt(Mid$(strNumber2, intIndex, 1))

        If intTempDigit1 - intTempDigit2 - intSubtractBuffer < 0 Then

            strCalcResult = CStr(intTempDigit1 - intTempDigit2 - intSubtractBuffer + 10) & strCalcResult

            ' if Digit1 - Digit2 < 0 then
            ' intSubtractBuffer contains 1
            intSubtractBuffer = 1
        Else
            strCalcResult = CStr(intTempDigit1 - intTempDigit2 - intSubtractBuffer) & strCalcResult
            intSubtractBuffer = 0
        End If

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If

    Next intIndex

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strCalcResult = "0"
    Else
        strCalcResult = DataTrim(strCalcResult)

        If Len(Trim$(strTempBuffer)) > 0 Then
            strCalcResult = "-" & strCalcResult
        End If
    End If

BigSubtract_CleanUp:
    BigSubtract = strCalcResult

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

BigSubtract_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strCalcResult = "0"
    gblnStopProcessing = True
    Resume BigSubtract_CleanUp

End Function

' ***************************************************************************
' Routine:       BigMultiply
'
' Description:   A solution for handling large numbers (up to 32768 digits)
'                in Visual Basic 6.  Multiplies strNumber1 with strNumber2
'                and returns the result in a string.
'
' Example:                       13096744586834688815
'                                 4794697086780616226
'              --------------------------------------
'              62794923116806087757258117158649712190
'
' Parameters:    strNumber1 - Numeric string
'                strNumber1 - Numeric string as multiplier
'
' Returns:       New numeric string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 23-Nov-2002  Philipp Emanuel Weidmann   http://www.philippsworld.net/
'              Author of three of the "Big" routines
' 29-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified variable names and renamed routine for easier
'              maintenance.  Added error trapping.
' ***************************************************************************
Public Function BigMultiply(ByVal strNumber1 As String, _
                            ByVal strNumber2 As String) As String

    Dim intIndex1     As Integer
    Dim intindex2     As Integer
    Dim intTempDigit1 As Integer
    Dim intTempDigit2 As Integer
    Dim strCalcResult As String

    Const ROUTINE_NAME As String = "BigMultiply"

    On Error GoTo BigMultiply_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    If Not IsNumeric(strNumber1) Or Not IsNumeric(strNumber2) Then
        InfoMsg "Values to be multiplied must be numeric." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        Exit Function
    End If

    strCalcResult = "0"

    For intIndex1 = Len(strNumber1) To 1 Step -1

        For intindex2 = Len(strNumber2) To 1 Step -1

            intTempDigit1 = CInt(Mid$(strNumber1, intIndex1, 1))
            intTempDigit2 = CInt(Mid$(strNumber2, intindex2, 1))

            ' Split the multiplication into additions:
            '
            ' abc * def = 10^0 * cf + 10^1 * ce + 10^2 * cd +
            '             10^1 * bf + 10^2 * be + 10^3 * bd +
            '             10^2 * af + 10^3 * ae + 10^4 * ad
            strCalcResult = BigAdd(strCalcResult, _
                                   CStr(intTempDigit1 * intTempDigit2) & _
                                   String$((Len(strNumber1) - intIndex1) + _
                                           (Len(strNumber2) - intindex2), "0"))

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If

        Next intindex2

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If

    Next intIndex1

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strCalcResult = "0"
    Else
        strCalcResult = DataTrim(strCalcResult)
    End If

BigMultiply_CleanUp:
    BigMultiply = strCalcResult

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

BigMultiply_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strCalcResult = "0"
    gblnStopProcessing = True
    Resume BigMultiply_CleanUp

End Function

' ***************************************************************************
' Routine:       BigDivide
'
' Description:   A solution for handling large numbers (up to 32768 digits)
'                in Visual Basic 6.  Divides strNumber1 by strNumber2 and
'                returns the truncated result (no decimals) in a string.
'
' Example:       Full return with decimals:
'                1060894660739950.4912920210611584 = 13096744586834688815 / 12345
'
'                Only the truncated value is returned
'                1060894660739950 = 13096744586834688815 \ 12345
'
' Parameters:    strNumber1 - Numeric string as dividend
'                strNumber1 - Numeric string as divisor
'
' Returns:       New numeric string as quotient
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-Oct-2005  HAANDI
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=62738&lngWId=1
' 29-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified variable names and renamed routine for easier
'              maintenance.  Added error trapping.
' ***************************************************************************
Public Function BigDivide(ByVal strNumber1 As String, _
                          ByVal strNumber2 As String) As String

    Dim strCurrent    As String
    Dim strTemp       As String
    Dim strCount      As String
    Dim strCalcResult As String
    Dim intLength     As Integer

    Const ROUTINE_NAME As String = "BigDivide"

    On Error GoTo BigDivide_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    If Not IsNumeric(strNumber1) Or Not IsNumeric(strNumber2) Then
        InfoMsg "Values to be divided must be numeric." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        Exit Function
    End If

    If Val(strNumber1) <= 0 Or Val(strNumber2) <= 0 Then
        InfoMsg "Both values must be greater than zero." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        Exit Function
    End If

    If bGreaterThan(strNumber2, strNumber1) Then
        InfoMsg "The divisor must be less than or equal to the dividend." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        Exit Function
    End If

    strCurrent = strNumber1
    strCalcResult = "0"

    Do While Not (bGreaterThan(strNumber2, strCurrent))

        strTemp = strNumber2
        strCount = "1"
        intLength = Len(strCurrent) - Len(strTemp) + _
                    bGreaterThan(strTemp, Left$(strCurrent, Len(strTemp)))

        strTemp = strTemp & String$(intLength, "0")
        strCount = strCount & String$(intLength, "0")

        Do While bGreaterThan(strCurrent, strTemp) Or strCurrent = strTemp

            strCurrent = BigSubtract(strCurrent, strTemp)

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do   ' exit Do..Loop
            End If

            strCalcResult = BigAdd(strCalcResult, strCount)

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do   ' exit Do..Loop
            End If

        Loop

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            Exit Do   ' exit Do..Loop
        End If

        strCurrent = DataTrim(strCurrent)

    Loop

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strCalcResult = "0"
    End If

BigDivide_CleanUp:
    BigDivide = strCalcResult

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

BigDivide_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strCalcResult = "0"
    gblnStopProcessing = True
    Resume BigDivide_CleanUp

End Function

' ***************************************************************************
' Routine:       SwapData
'
' Description:   Swap data with each other.  I wrote this function since
'                BASIC stopped having its own SWAP function.  I use this
'                for swapping strings, type structures, numbers with
'                decimal values, etc.
'
' Parameters:    vntValue1 - Incoming data to be swapped with Value2
'                vntValue2 - Incoming data to be swapped with Value1
'
' Returns:       Swapped data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 09-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Sub SwapData(ByRef vntData1 As Variant, _
                    ByRef vntData2 As Variant)

    Dim vntHold As Variant

    vntHold = Empty   ' Start with empty variants (prevents crashing)

    vntHold = vntData1
    vntData1 = vntData2
    vntData2 = vntHold

    vntHold = Empty   ' Always empty variants (prevents crashing)

End Sub

' ***************************************************************************
' Routine:       SwapEndianHex
'
' Description:   Converts hex data from Big-Endian to Little-Endian or
'                Little-Endian to Big_Endian format. Used with certain
'                hash algorithms.
'
' Parameters:    strHex - Incoming hex string to be manipulated
'                lngRetLength - Optional - Desired length of the return
'                         string.  Default = 16
'
' Returns:       Hex string in new format
'
' Example:         Big-Endian     to    Little-Endian  to    Big-Endian
'                0123456789ABCDEF --> EFCDAB8967452301 --> 0123456789ABCDEF
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function SwapEndianHex(ByVal strHex As String, _
                     Optional ByVal lngRetLength As Long = 16) As String

    Dim strOutput   As String
    Dim intPosition As Integer
    Dim intPointer  As Integer

    ' No negative or zero lengths allowed
    If lngRetLength < 1 Then
        lngRetLength = 2
    End If

    ' Return length must be in multiples of 2
    If lngRetLength Mod 2 <> 0 Then
        lngRetLength = lngRetLength + 1
    End If

    ' Verify the incoming hex string is the
    ' same as the return length including
    ' any leading zeroes.
    strHex = Right$(String$(lngRetLength, "0") & strHex, lngRetLength)
    strOutput = Space$(lngRetLength * 2)  ' Preload output string
    intPointer = lngRetLength - 1         ' Initialize pointer

    ' Parse hex string capturing two characters
    ' at a time and inserting them into the
    ' output string
    For intPosition = 1 To lngRetLength Step 2

        Mid$(strOutput, intPosition, 2) = Mid$(strHex, intPointer, 2)
        intPointer = intPointer - 2

    Next intPosition

    ' Remove leading and trailing spaces
    ' before returning hex string
    SwapEndianHex = Trim$(strOutput)

End Function

' ***************************************************************************
' Routine:       IsHexData
'
' Description:   Parses a string of data to determine if it is in hex format.
'
' Parameters:    strData - String of data to be evaluated
'
' Returns:       TRUE  - Data string is in hex format
'                FALSE - Not in hex format
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' 10-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Updated logic testing
' ***************************************************************************
Public Function IsHexData(ByVal strData As String) As Boolean

    Dim lngIndex  As Long
    Dim lngLength As Long

    Const ROUTINE_NAME As String = "IsHexData"
    Const HEX_DATA     As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"

    IsHexData = True             ' Preset to TRUE
    strData = TrimStr(strData)   ' Remove unwanted characters
    lngLength = Len(strData)     ' Capture length of data string

    If lngLength > 0 Then

        ' Parse data string to verify
        ' each character is valid
        For lngIndex = 1 To lngLength

            If InStr(1, HEX_DATA, Mid$(strData, lngIndex, 1)) = 0 Then
                IsHexData = False
                Exit For   ' Found invalid character
            End If

        Next lngIndex

    Else

        InfoMsg "Incoming data string is empty." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        IsHexData = False

    End If

End Function

' ***************************************************************************
' Routine:       IsBinaryData
'
' Description:   Parses a string of data to determine if it is in
'                binary format.  Must be zero or one.
'
' Parameters:    strData - String of data to be evaluated
'
' Returns:       TRUE  - Data string is in binary format
'                FALSE - Not in binary format
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' 10-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Updated logic testing
' ***************************************************************************
Public Function IsBinaryData(ByVal strData As String) As Boolean

    Dim lngIndex  As Long
    Dim lngLength As Long

    Const ROUTINE_NAME As String = "IsBinaryData"
    Const BINARY_DATA  As String = "01"

    IsBinaryData = True          ' Preset to TRUE
    strData = TrimStr(strData)   ' Remove unwanted characters
    lngLength = Len(strData)     ' Capture length of data string

    If lngLength > 0 Then

        ' Parse data string to verify
        ' each character is valid
        For lngIndex = 1 To lngLength

            If InStr(1, BINARY_DATA, Mid$(strData, lngIndex, 1)) = 0 Then
                IsBinaryData = False
                Exit For   ' Found invalid character
            End If

        Next lngIndex

    Else

        InfoMsg "Incoming data string is empty." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        IsBinaryData = False

    End If

End Function

' ***************************************************************************
' Routine:       HexRepresentation
'
' Description:   This is where the hex representation of an array of numbers
'                is calculated and returned in a string array.  If the
'                parameter lngNumber is greater than zero then only a single
'                value is processed.  If the parameter lngNumber equal zero
'                then an array of prime numbers are determined in sequence
'                and their hex representation is calculated accordingly.
'
'                These types of hex values are generally used in the work
'                and constant arrays for the SHA2 family of hash algorithms.
'
' Parameters:    lngQtyReq - Number of items to be returned
'                lngRetFmt - [Optional] - Return format.  Return all 16 hex
'                            values, just the first 8 or the last 8.
'                            Default = e16Chars (All 16 characters)
'                lngHexRep - [Optional] - Type of hex representation to
'                            perform.  Either Cube Roots or Square Roots.
'                            Default = eCubeRoots
'                lngNumber - [Optional] - This value is greater than zero if
'                            a single value is to be processed.
'                            Default = 0
'
' Returns:       A string array of hex values
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Nov-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 01-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added parameter testing
' ***************************************************************************
Public Function HexRepresentation(ByVal lngQtyReq As Long, _
                         Optional ByVal lngRetFmt As enumHEX_RETURN_FORMAT = e16Chars, _
                         Optional ByVal lngHexRep As enumHEX_REPRESENTATION = eCubeRoots, _
                         Optional ByVal lngNumber As Long = 0) As String()

    Dim astrHex()     As String
    Dim alngNumbers() As Long

    Const ROUTINE_NAME As String = "HexRepresentation"

    On Error GoTo HexRepresentation_Error

    Erase astrHex()     ' Always start with empty arrays
    Erase alngNumbers()

    ' If a negative value is passed
    ' then present a message
    If lngQtyReq < 1 Then
        InfoMsg "Quantity requested must be a positive value greater than zero." & _
                vbNewLine & vbNewLine & "Quantity:   " & CStr(lngQtyReq) & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        GoTo HexRepresentation_CleanUp
    End If

    Select Case lngRetFmt
           Case e16Chars, eLeft8, eRight8
                ' Valid values
           Case Else
                InfoMsg "Cannot identify number of characters to be returned." & _
                        vbNewLine & vbNewLine & "Valid values:" & vbNewLine & _
                        Space$(5) & "e16Chars or 16" & vbNewLine & _
                        Space$(5) & "eLeft8 or 8" & vbNewLine & _
                        Space$(5) & "eRight8 or 8" & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                gblnStopProcessing = True
                GoTo HexRepresentation_CleanUp
    End Select

    Select Case lngHexRep
           Case eCubeRoots, eSquareRoots
                ' Valid values
           Case Else
                InfoMsg "Hex representation must be either CUBE ROOT or SQUARE ROOT." & _
                        vbNewLine & vbNewLine & "Valid values:" & vbNewLine & _
                        Space$(5) & "eCubeRoots or 0" & vbNewLine & _
                        Space$(5) & "eSquareRoots or 1" & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                gblnStopProcessing = True
                GoTo HexRepresentation_CleanUp
    End Select

    ' Evaluate if one or more numbers are to be processed
    Select Case lngNumber

           Case Is > 0
                ' lngNumber is greater than zero therefore
                ' a single value is to be processed
                lngQtyReq = 1
                ReDim alngNumbers(lngQtyReq)
                CopyMemory alngNumbers(0), lngNumber, 4&

           Case 0
                ' Use an array of prime numbers
                '
                ' load prime number array
                alngNumbers() = GetPrimeNumbers(lngQtyReq)

                ' Test to see if the array has any data
                If Not CBool(IsArrayInitialized(alngNumbers())) Then
                    InfoMsg "Failed to load prime number array." & _
                            vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                    GoTo HexRepresentation_CleanUp
                End If

           Case Is < 0
                InfoMsg "Number to be evaluated must be a positive value of zero or greater." & _
                        vbNewLine & vbNewLine & "Number:   " & CStr(lngNumber) & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                        gblnStopProcessing = True
                        GoTo HexRepresentation_CleanUp
    End Select

    ' Fill the hex array with computed values
    Select Case lngHexRep
           Case eCubeRoots:   astrHex() = LoadCubeRoots(lngQtyReq, lngRetFmt, alngNumbers())
           Case eSquareRoots: astrHex() = LoadSquareRoots(lngQtyReq, lngRetFmt, alngNumbers())
    End Select

    HexRepresentation = astrHex()

HexRepresentation_CleanUp:
    Erase astrHex()     ' Always empty arrays when not needed
    Erase alngNumbers()

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

HexRepresentation_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume HexRepresentation_CleanUp

End Function

' ***************************************************************************
' Routine:       IsArrayInitialized
'
' Description:   This is an ArrPtr function that determines if the passed
'                array is initialized, and if so will return the pointer
'                to the safearray header. If the array is not initialized,
'                it will return zero. Normally you need to declare a VarPtr
'                alias into msvbvm50.dll or msvbvm60.dll depending on the
'                VB version, but this function will work with vb5 or vb6.
'                It is handy to test if the array is initialized as the
'                return value is non-zero.  Use CBool to convert the return
'                value into a boolean value.
'
'                This function returns a pointer to the SAFEARRAY header of
'                any Visual Basic array, including a Visual Basic string
'                array. Substitutes both ArrPtr and StrArrPtr. This function
'                will work with vb5 or vb6 without modification.
'
'                ex:  If CBool(IsArrayInitialized(array_being_tested)) Then ...
'
' Parameters:    vntData - Data to be evaluated
'
' Returns:       Zero     - Bad data (FALSE)
'                Non-zero - Good data (TRUE)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 30-Mar-2008  RD Edwards
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=69970
' ***************************************************************************
Public Function IsArrayInitialized(ByVal avntData As Variant) As Long

    Dim intDataType As Integer   ' Must be an integer

    On Error GoTo IsArrayInitialized_Exit

    IsArrayInitialized = 0  ' preset to FALSE

    ' Get the real VarType of the argument, this is similar
    ' to VarType(), but returns also the VT_BYREF bit
    CopyMemory intDataType, avntData, 2&

    ' if a valid array was passed
    If (intDataType And vbArray) = vbArray Then

        ' get the address of the SAFEARRAY descriptor
        ' stored in the second half of the Variant
        ' parameter that has received the array.
        ' Thanks to Francesco Balena and Monte Hansen.
        CopyMemory IsArrayInitialized, ByVal VarPtr(avntData) + 8&, 4&

    End If

IsArrayInitialized_Exit:
    On Error GoTo 0   ' Nullify this error trap

End Function

' ***************************************************************************
' Routine:       GetPrimeNumbers
'
' Description:   This is where an array is filled with LONG INTEGERS only.
'
'                Prime Number - A prime number is a positive integer that
'                has exactly two positive integer factors, 1 and itself.
'                Two is the only even prime number, since any bigger even
'                number is divided by two.  For example, if we list the
'                factors of 28, we have 1, 2, 4, 7, 14, and 28.  That's
'                six factors.  If we list the factors of 29, we only have
'                1 and 29.  Prime numbers have exactly two factors.  So we
'                say that 29 is a prime number and 28 is a composite number
'                and not a prime.
'
'                Eratosthenes worked on prime numbers.  He is remembered for
'                his prime number sieve, the "Sieve of Eratosthenes" which,
'                in Modified form, is still an important tool in number
'                theory research.  He was also one of the first ancient Greek
'                scientists to calculate the circumference of the earth, and
'                with an astonishing accuracy.  He lived 276 BC to 194 BC.
'
' Reference:     Ask Dr. Math
'                http://mathforum.org/dr.math/faq/faq.prime.num.html
'
'                The Prime Pages
'                http://primes.utm.edu/
'
'                Primality Testing Is Easy
'                http://mathworld.wolfram.com/news/2002-08-07/primetest/
'
'                10,000 Smallest primes by Peter Alfeld
'                Department of Mathematics, University of Utah
'                http://www.math.utah.edu/~pa/math/p10000.html
'
' Parameters:    lngQtyNeeded - [Optional] - Number of prime numbers to return
'                                            Default = 1
'
' Returns:       An array of prime numbers
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Jul-2009  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 10-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Fixed logic bug in determining starting value
' ***************************************************************************
Public Function GetPrimeNumbers(ByVal lngQtyNeeded As Long, _
                       Optional ByVal lngStartValue As Long = 2, _
                       Optional ByVal lngMaxValue As Long = MAX_LONG) As Long()

    ' Called by HexRepresentation()

    Dim lngNumber    As Long
    Dim lngPrimeCnt  As Long
    Dim alngPrimes() As Long

    Const ROUTINE_NAME As String = "GetPrimeNumbers"

    On Error GoTo GetPrimeNumbers_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo GetPrimeNumbers_CleanUp
    End If

    Erase alngPrimes()  ' Always start with empty arrays
    lngPrimeCnt = 0     ' Init array index

    '-------------------------------------------------
    ' Evaluate input parameters
    If lngQtyNeeded < 1 Then
        InfoMsg "Return quantity must be a positive number greater than zero." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        GoTo GetPrimeNumbers_CleanUp
    End If

    If (lngStartValue < 2) Or (lngMaxValue < 2) Then
        InfoMsg "Starting and maximum values must be a positive number greater than one." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        GoTo GetPrimeNumbers_CleanUp
    End If

    If lngStartValue > lngMaxValue Then
        InfoMsg "Starting value must be less than or equal maximum value." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        GoTo GetPrimeNumbers_CleanUp
    End If
    '-------------------------------------------------

    ' Size return array to number of
    ' primes needed with a buffer.
    ReDim alngPrimes(lngQtyNeeded + 2)

    ' Determine starting value
    If lngStartValue = 2 Then
        alngPrimes(0) = 2  ' Insert first two primes into array
        alngPrimes(1) = 3
        lngPrimeCnt = 2    ' Number of primes
        lngStartValue = 5  ' Next odd value

    ElseIf lngStartValue = 3 Then
        lngStartValue = 5  ' Next odd value

    Else
        ' Start value must be odd
        If lngStartValue Mod 2 = 0 Then
            lngStartValue = lngStartValue + 1
        End If
    End If

    If lngPrimeCnt >= lngQtyNeeded Then
        ReDim Preserve alngPrimes(lngQtyNeeded)  ' Resize prime number array
        GetPrimeNumbers = alngPrimes()           ' Return prime number array
        GoTo GetPrimeNumbers_CleanUp             ' Time to leave
    End If

    ' Second test to verify values
    ' are still within range
    If lngStartValue > lngMaxValue Then
        InfoMsg "Starting value must be less than or equal maximum value." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        GoTo GetPrimeNumbers_CleanUp
    End If

    ' Load rest of requested prime numbers
    ' into an array.  Use step 2 because
    ' only odd value numbers are processed.
    For lngNumber = lngStartValue To lngMaxValue Step 2

        ' Test to see if this is a prime number
        If IsPrime(lngNumber) Then
            alngPrimes(lngPrimeCnt) = lngNumber   ' Add number to array
            lngPrimeCnt = lngPrimeCnt + 1         ' Increment array index
        End If

        ' If enough data has been collected
        ' then exit this loop
        DoEvents
        If lngPrimeCnt >= lngQtyNeeded Then
            Exit For   ' exit For..Next loop
        End If

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            Exit For   ' exit For..Next loop
        End If

    Next lngNumber                                        '

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo GetPrimeNumbers_CleanUp
    End If

    ReDim Preserve alngPrimes(lngQtyNeeded)  ' Resize prime number array
    GetPrimeNumbers = alngPrimes()           ' Return prime number array

GetPrimeNumbers_CleanUp:
    Erase alngPrimes()  ' Always empty arrays when not needed
    On Error GoTo 0     ' Nullify this error trap
    Exit Function

GetPrimeNumbers_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume GetPrimeNumbers_CleanUp

End Function

' ***************************************************************************
' Routine:       IsPrime
'
' Description:   Determines whether a positive whole number is a prime.
'
'                Prime Number - A prime number is a positive integer that
'                has exactly two positive integer factors, 1 and itself.
'                Two is the only even prime number, since any bigger even
'                number is divided by two.  For example, if we list the
'                factors of 28, we have 1, 2, 4, 7, 14, and 28.  That's
'                six factors.  If we list the factors of 29, we only have
'                1 and 29.  Prime numbers have exactly two factors.  So we
'                say that 29 is a prime number and 28 is a composite number.
'
'                Eratosthenes worked on prime numbers.  He is remembered for
'                his prime number sieve, the "Sieve of Eratosthenes" which,
'                in Modified form, is still an important tool in number
'                theory research.  He was also one of the first ancient Greek
'                scientists to calculate the circumference of the earth, and
'                with an astonishing accuracy.  He lived 276 BC to 194 BC.
'
' Reference:     Ask Dr. Math
'                http://mathforum.org/dr.math/faq/faq.prime.num.html
'
'                The Prime Pages
'                http://primes.utm.edu/
'
'                Primality Testing Is Easy
'                http://mathworld.wolfram.com/news/2002-08-07/primetest/
'
' Parameters:    dblNumber - Number to be evaluated
'
' Returns:       TRUE - Number is a prime value
'                FALSE - Number is a composite (not a prime)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 27-May-2000  Francesco Balena
'              http://www.devx.com/vb2themax/Tip/19051
' 12-Jul-2009  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function IsPrime(ByVal dblNumber As Double) As Boolean

    ' Called by GetPrimeNumbers()

    Dim lngStep    As Long
    Dim dblDivisor As Double
    Dim dblSqrRoot As Double

    Const ROUTINE_NAME As String = "IsPrime"

    IsPrime = False             ' Preset to FALSE for input testing
    dblNumber = Fix(dblNumber)  ' Remove any decimal positions

    If dblNumber < 1# Then
        InfoMsg "Number to be evaluated must be greater than zero." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If

    ' Pretest with two smallest primes
    If dblNumber Mod 2# = 0 Or _
       dblNumber Mod 3# = 0 Then

        Exit Function
    End If

    IsPrime = True    ' Assume input value is a prime
    dblDivisor = 5#   ' Starting divisor
    lngStep = 2       ' Incremental steps

    ' Calc square root of incoming number
    dblSqrRoot = Int(Sqr(dblNumber)) + 1

    ' No need to go higher than the
    ' square root of the number
    Do While dblSqrRoot > dblDivisor

        If (dblNumber Mod dblDivisor) = 0 Then
            IsPrime = False   ' Not a prime number
            Exit Do           ' exit Do..Loop
        End If

        dblDivisor = dblDivisor + lngStep  ' Update divisor
        lngStep = IIf(lngStep = 2, 4, 2)   ' Toggle between 2 and 4

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            IsPrime = False   ' Set flag to FALSE
            Exit Do           ' exit Do..Loop
        End If

    Loop

End Function



' ***************************************************************************
' ****               Internal Procedures and Functions                   ****
' ***************************************************************************

' ***************************************************************************
' Description:   This will handle values up to 9999999999999999999999999999
'                which is 204FCE5E3E2502610FFFFFFF in hex.  That is
'                twenty-eight digits.  When converting from hex and the value
'                is negative, numeric values will not show a minus sign.
'
'                 Base
'                    2 (Binary)  -> 10 (Numeric) or 16 (Hex)
'                   10 (Numeric) ->  2 (Binary)  or 16 (Hex)
'                   16 (Hex)     ->  2 (Binary)  or 10 (Numeric)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-Jan-2002  Guillaume Giffard   Guiland@mail.com
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=30364&lngWId=1
' 11-Jan-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function ConvertBase(ByVal strDataToConvert As String, _
                             ByVal lngFromBase As enum64_BaseValues, _
                             ByVal lngToBase As enum64_BaseValues, _
                    Optional ByVal strSeparator As String = ".") As String

    ' Called by w64NumberToHex()
    '           w64NumberToBinary()
    '           w64HexToNumber()
    '           w64HexToBinary()
    '           w64BinaryToHex()
    '           w64BinaryToNumber()

    Dim blnNegativeNumber As Boolean

    Const ROUTINE_NAME As String = "ConvertBase"

    On Error GoTo ConvertBase_Error

    blnNegativeNumber = False
    ConvertBase = vbNullString
    Erase mastrDigits()   ' Always start with empty arrays
    Erase malngValueIn()
    Erase malngValueOut()

    ' Evaluate incoming data
    If Len(Trim$(strDataToConvert)) = 0 Or _
       (lngFromBase < 2 Or lngFromBase > 36) Or _
       (lngToBase < 2 Or lngToBase > 36) Then

        Exit Function
    End If

    Select Case lngFromBase

           Case e64_Binary   ' 2
                ' Verify a full 64 character string
                strDataToConvert = Right$(ZEROES_64 & strDataToConvert, 64)

                ' If first digit in binary string, 16-bit or bigger, is one (1),
                ' value will be negative.  If zero(0), value will be positive.
                blnNegativeNumber = IIf(Left$(strDataToConvert, 1) = "1", True, False)

           Case e64_Numeric   ' 10
                ' Remove minus sign, if any
                ' "-" in first position designates a negative value
                If Mid$(strDataToConvert, 1, 1) = "-" Then
                    strDataToConvert = Mid$(strDataToConvert, 2)
                    blnNegativeNumber = True
                End If

           Case e64_Hex   ' 16
                ' Verify a full 16 character string
                strDataToConvert = Right$(ZEROES_16 & strDataToConvert, HEX_LENGTH)
    End Select

    LoadDigitsArray
    StringToArray strDataToConvert, lngFromBase, strSeparator
    Convert lngToBase
    ConvertBase = DeleteZeros(ArrayToString(strSeparator), strSeparator)

    ' Prefix leading zeroes for
    ' binary and hex return values
    Select Case lngToBase

           Case e64_Binary   ' 2
                ConvertBase = Right$(ZEROES_64 & ConvertBase, 64)

                ' If From Base was negative then
                ' replace first digit with a one (1)
                If blnNegativeNumber Then
                    Mid$(ConvertBase, 1, 1) = "1"
                End If

           Case e64_Numeric  ' 10
                ' If From Base was negative then
                ' prefix numeric with minus (-)
                If blnNegativeNumber Then
                    ConvertBase = "-" & ConvertBase
                End If

           Case e64_Hex      ' 16
                ConvertBase = Right$(ZEROES_16 & ConvertBase, HEX_LENGTH)
                ConvertBase = LCase$(ConvertBase)
    End Select

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        ConvertBase = vbNullString
        Exit Function
    End If

ConvertBase_CleanUp:
    Erase mastrDigits()   ' Always empty arrays when not needed
    Erase malngValueIn()
    Erase malngValueOut()

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

ConvertBase_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    ConvertBase = vbNullString
    gblnStopProcessing = True
    Resume ConvertBase_CleanUp

End Function

' Saves base digits in an array
Private Sub LoadDigitsArray()

    ' Called by ConvertBase()

    Dim lngIndex As Long

    Erase mastrDigits()
    ReDim mastrDigits(0 To 35)

    For lngIndex = 0 To 9
        mastrDigits(lngIndex) = lngIndex
    Next lngIndex

    For lngIndex = 10 To 35
        mastrDigits(lngIndex) = Chr$(lngIndex + 55)
    Next lngIndex

End Sub

' Saves a number as string in an array of Longs.
' Each array element is one digit.
Private Sub StringToArray(ByVal strData As String, _
                          ByVal lngBase As Long, _
                          ByVal strSeparator As String)

    ' Called by ConvertBase()

    Dim lngMin       As Long
    Dim lngMax       As Long
    Dim lngIndex     As Long
    Dim lngPointer   As Long
    Dim lngNoPointer As Long

    mlngFromBase = lngBase
    lngPointer = InStr(1, strData, strSeparator, vbTextCompare)

    If lngPointer = 0 Then
        lngMax = Len(strData) - 1
        lngMin = 0
    Else
        lngMax = lngPointer - 2
        lngMin = lngPointer - Len(strData) + Len(strSeparator) - 1
    End If

    ReDim malngValueIn(lngMin To lngMax)

    For lngIndex = 0 To Len(strData) - 1

        If lngIndex <= (Len(strData) - lngPointer) And _
           lngIndex >= (Len(strData) - lngPointer - Len(strSeparator) + 1) Then

            lngNoPointer = lngNoPointer - 1
        Else
            malngValueIn(lngMin + lngIndex + lngNoPointer) = Numeric(Left$(Right$(strData, lngIndex + 1), 1), lngBase)
        End If

    Next lngIndex

End Sub

' Returns number corresponding to a digit
' as string if the digit is allowed
' by the base. e.g. : C is allowed in
' hex but not in decimal or in octal
Private Function Numeric(ByVal strData As String, _
                         ByVal lngBase As Long) As Long

    ' Called by StringToArray()

    Dim lngIndex As Long

    For lngIndex = 0 To 35

        If lngIndex = lngBase Then
            Exit Function
        End If

        If StrComp(strData, mastrDigits(lngIndex), vbTextCompare) = 0 Then
            Numeric = lngIndex
        End If

    Next lngIndex

End Function

' sub that converts malngValueIn to malngValueOut
' with mlngFromBase and mlngToBase
Private Sub Convert(ByVal lngToBase As Long)

    ' Called by ConvertBase()

    Dim lngIdx        As Long
    Dim lngMin        As Long
    Dim lngMax        As Long
    Dim lngIndex      As Long
    Dim lngFromBase   As Long
    Dim alngTemp()    As Long
    Dim alngValue()   As Long
    Dim alngDataIn()  As Long
    Dim alngDataOut() As Long

    Erase alngTemp()   ' Always start with empty arrays
    Erase alngValue()
    Erase alngDataIn()
    Erase alngDataOut()

    ReDim alngDataIn(UBound(malngValueIn) + 1)
    CopyMemory alngDataIn(0), malngValueIn(0), UBound(alngDataIn) * 4&

    mlngToBase = lngToBase
    lngFromBase = mlngFromBase

    lngMax = RoundOverInt(Int((UBound(alngDataIn, 1) + 1) * Log(lngFromBase) / Log(lngToBase)))
    lngMin = Int(LBound(alngDataIn, 1) * Log(lngFromBase) / Log(lngToBase))

    ReDim alngDataOut(lngMin To lngMax)
    lngIndex = 0   ' LBound(malngValueOut, 1)

    DivideVeryLong alngDataIn(), lngToBase, alngTemp(), alngDataOut(lngIndex), lngFromBase
    lngIndex = lngIndex + 1

    Do Until lngIndex > UBound(alngDataOut, 1)

        DivideVeryLong alngTemp(), lngToBase, alngValue(), alngDataOut(lngIndex), lngFromBase
        ReDim alngTemp(LBound(alngValue, 1) To UBound(alngValue, 1))

        For lngIdx = LBound(alngValue, 1) To UBound(alngValue, 1)
            alngTemp(lngIdx) = alngValue(lngIdx)
        Next lngIdx

        lngIndex = lngIndex + 1

    Loop

    ReDim malngValueOut(lngMin To lngMax + 1)
    CopyMemory malngValueOut(0), alngDataOut(0), UBound(alngDataOut) * 4&
    ReDim Preserve malngValueOut(lngMin To lngMax)

    Erase alngTemp()   ' Always empty arrays when not needed
    Erase alngValue()
    Erase alngDataIn()
    Erase alngDataOut()

End Sub

' round numbers to the closest higher Long
' e.g. : 3.9 gives 4 ; 3.4 gives 4 ; 3 gives 3
Private Function RoundOverInt(ByVal dblValue As Double) As Integer

    ' Called by Convert()

    If dblValue = Int(dblValue) Then
        RoundOverInt = CInt(dblValue)
    Else
        RoundOverInt = CInt(Int(dblValue) + 1)
    End If

End Function

' Divides a huge number by a Long and
' returns huge quotient and remainder
Private Sub DivideVeryLong(ByRef alngNumerator() As Long, _
                           ByVal lngDenominator As Long, _
                           ByRef alngQuotient() As Long, _
                           ByRef lngRemainder As Long, _
                           ByVal lngBase As Long)

    ' Called by Convert()

    Dim lngTemp  As Long
    Dim lngDecal As Long
    Dim lngIndex As Long

    lngTemp = 0
    lngDecal = 0
    ReDim alngQuotient(LBound(alngNumerator, 1) To UBound(alngNumerator, 1))

    For lngIndex = UBound(alngNumerator, 1) To 0 Step -1

        lngTemp = lngTemp * lngBase + alngNumerator(lngIndex)
        alngQuotient(lngIndex - lngDecal) = lngTemp \ lngDenominator
        lngTemp = lngTemp - alngQuotient(lngIndex - lngDecal) * lngDenominator

    Next lngIndex

    lngRemainder = lngTemp

End Sub

' Saves an array to a string
Private Function ArrayToString(ByVal strSeparator As String) As String

    ' Called by ConvertBase()

    Dim lngIndex As Long

    ArrayToString = vbNullString

    For lngIndex = UBound(malngValueOut, 1) To LBound(malngValueOut, 1) Step -1

        If lngIndex = -1 Then
            ArrayToString = ArrayToString & strSeparator
        End If

        ArrayToString = ArrayToString & mastrDigits(malngValueOut(lngIndex))

    Next lngIndex

End Function

' Deletes zeros before and after the data string
' and, if possible, deletes the separator
Private Function DeleteZeros(ByVal strData As String, _
                             ByVal strSeparator As String) As String

    ' Called by ConvertBase()

    Dim lngPointer As Long
    Dim strTemp    As String

    strTemp = strData

    Do
        lngPointer = InStr(1, strTemp, "0", vbTextCompare)

        If lngPointer = 1 Then
            strTemp = Right$(strTemp, Len(strTemp) - 1)
        Else
            Exit Do
        End If

    Loop

    If InStr(1, strTemp, strSeparator, vbTextCompare) <> 0 Then

        Do
            lngPointer = InStr(Len(strTemp) - 1, strTemp, "0", vbTextCompare)

            If lngPointer = Len(strTemp) - 1 Then
                strTemp = Left$(strTemp, Len(strTemp) - 1)
            Else
                Exit Do
            End If

        Loop

        Do
            lngPointer = InStr(Len(strTemp) - 1, strTemp, "0", vbTextCompare)

            If lngPointer = Len(strTemp) Then
                strTemp = Left$(strTemp, Len(strTemp) - 1)
            Else
                Exit Do
            End If

        Loop
    End If

    If Len(Trim$(strTemp)) = 0 Then
        strTemp = "0"
    End If

    If InStr(1, strTemp, strSeparator, vbTextCompare) = Len(strTemp) - Len(strSeparator) + 1 Then
        strTemp = Left$(strTemp, Len(strTemp) - Len(strSeparator))
    End If

    DeleteZeros = strTemp

End Function

' ***************************************************************************
'
' Description:   A solution for handling large numbers (up to 32768 digits)
'                in Visual Basic 6.  Compares strNumber1 with strNumber2:
'
'                     Returns  1 if strNumber1 is greater than strNumber2
'                     Returns  2 if strNumber2 is greater than strNumber1
'                     Returns  0 if the numbers are equal
'                     Returns -1 if an error occurs
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 23-Nov-2002  Philipp Emanuel Weidmann   http://www.philippsworld.net/
'              Author of the "Big" routines
' 29-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified variable names and renamed routine for easier
'              maintenance.
' ***************************************************************************
Private Function BigCompare(ByVal strNumber1 As String, _
                            ByVal strNumber2 As String) As Integer

    ' Called by BigSubtract()

    Dim intIndex As Integer

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    If Not IsNumeric(strNumber1) Or Not IsNumeric(strNumber2) Then
        BigCompare = -1
        Exit Function
    End If

    ' Test if one of the numbers is shorter than the other one
    If Len(strNumber1) > Len(strNumber2) Then
        BigCompare = 1
        Exit Function

    ElseIf Len(strNumber1) < Len(strNumber2) Then
        BigCompare = 2
        Exit Function
    End If

    ' The numbers are equal in length => compare them digit by digit
    For intIndex = 1 To Len(strNumber1)

        If CInt(Mid$(strNumber1, intIndex, 1)) > CInt(Mid$(strNumber2, intIndex, 1)) Then
            BigCompare = 1
            Exit Function

        ElseIf CInt(Mid$(strNumber1, intIndex, 1)) < CInt(Mid$(strNumber2, intIndex, 1)) Then
            BigCompare = 2
            Exit Function
        End If

    Next intIndex

    BigCompare = 0

End Function

Private Function bGreaterThan(ByVal strNumber1 As String, _
                              ByVal strNumber2 As String) As Boolean

    ' Called by BigDivide()

    Dim intIndex   As Integer
    Dim intDigit1  As Integer
    Dim intDigit2  As Integer

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    strNumber1 = DataTrim(strNumber1)
    strNumber2 = DataTrim(strNumber2)

    If Len(strNumber1) > Len(strNumber2) Then
        bGreaterThan = True
    ElseIf Len(strNumber1) < Len(strNumber2) Then
        bGreaterThan = False
    Else
        ' same length
        ' loop thru data until 1 digit does not equal the other
        For intIndex = 1 To Len(strNumber1)

            intDigit1 = CInt(Mid$(strNumber1, intIndex, 1))
            intDigit2 = CInt(Mid$(strNumber2, intIndex, 1))

            If intDigit1 > intDigit2 Then
                bGreaterThan = True
                Exit For    ' exit For..Next loop
            End If

            If intDigit2 > intDigit1 Then
                bGreaterThan = False
                Exit For    ' exit For..Next loop
            End If

        Next intIndex
    End If

End Function

Private Function DataTrim(ByVal strNumber As String) As String

    ' Drop all leading "0"s
    '
    ' Called by BigAdd()
    '           BigSubtract()
    '           BigMultiply()
    '           BigDivide()
    '           bGreaterThan()

    If Len(strNumber) = 0 Then
        DataTrim = "0"
    Else
        Do While Left$(strNumber, 1) = "0" And Len(strNumber) > 1

            strNumber = Right$(strNumber, Len(strNumber) - 1)

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                strNumber = "0"
                Exit Do   ' exit Do..Loop
            End If

        Loop

        DataTrim = strNumber

    End If

End Function

' ***************************************************************************
' Routine:       LoadSquareRoots
'
' Description:   Got this explanation from Ask Dr. Math web site.
'                http://mathforum.org/dr.math/ 
'                 
'                There is a simple process for converting a base 10 decimal 
'                fraction to a "decimal" fraction in another base.  In this
'                case, hex will be our new base (16).  We repeatedly multiply
'                our decimal fraction by the new base, picking off the whole
'                number part each time as the next digit of the final output. 
'                 
' Example:   1st Prime number = 2
'            SQUARE ROOT of 2 = 1.4142135623730950488016887242
'
'            0.4142135623730950488016887242 * 16 =  6.6274169979695207808270195872
'            0.6274169979695207808270195872 * 16 = 10.038671967512332493232313395
'            0.038671967512332493232313395  * 16 =  0.61875148019731989171701432
'            0.61875148019731989171701432   * 16 =  9.90002368315711826747222912
'            0.90002368315711826747222912   * 16 = 14.40037893051389227955566592
'            0.40037893051389227955566592   * 16 =  6.40606288822227647289065472
'            0.40606288822227647289065472   * 16 =  6.49700621155642356625047552
'            0.49700621155642356625047552   * 16 =  7.95209938490277706000760832
'            0.95209938490277706000760832   * 16 = 15.23359015844443296012173312
'            0.23359015844443296012173312   * 16 =  3.73744253511092736194772992
'            0.73744253511092736194772992   * 16 = 11.79908056177483779116367872
'            0.79908056177483779116367872   * 16 = 12.78528898839740465861885952
'            0.78528898839740465861885952   * 16 = 12.56462381435847453790175232
'            0.56462381435847453790175232   * 16 =  9.03398102973559260642803712
'            0.03398102973559260642803712   * 16 =  0.54369647576948170284859392
'            0.54369647576948170284859392   * 16 =  8.69914361231170724557750272
'                                                   |
'               Whole numbers are converted to hex__|
'
'            The hex representation of the fractional parts of the
'            SQUARE ROOTS of 2 is:   6a09e667f3bcc908
'
' Parameters:  lngQtyReq    - Number of items to return
'              lngRetFmt    - Return format.
'                                e16Chars - All 16 characters
'                                eLeft8   - First 8 characters
'                                eRight8  - Last 8 characters
'
'              alngNumbers() - An array of numbrs to be processed
'
' Returns:     An array of hex strings
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Nov-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 08-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated documentation
' 08-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modifed to display output details as shown in flowerbox above
'              if user hard codes the exact prime number below.
' ***************************************************************************
Private Function LoadSquareRoots(ByVal lngQtyReq As Long, _
                                 ByVal lngRetFmt As enumHEX_RETURN_FORMAT, _
                                 ByRef alngNumbers() As Long) As String()

    ' Called by HexRepresentation()

    Dim blnMyNumber As Boolean   ' Flag for flowerbox output
    Dim lngLoop     As Long      ' loop counter
    Dim lngIndex    As Long      ' loop counter
    Dim lngNumber   As Long      ' Temp number
    Dim lngWholeNum As Long      ' Whole number only
    Dim vntBase     As Variant   ' fractional part of number
    Dim vntTemp     As Variant   ' new computed value
    Dim avntHex     As Variant   ' base 16 hex values
    Dim strHex      As String    ' hex character
    Dim strSuffix   As String    ' Used for flowerbox output
    Dim astrHex()   As String    ' hex data

    Const HEX_BASE     As Double = 16#
    Const ROUTINE_NAME As String = "LoadSquareRoots"

    On Error GoTo LoadSquareRoots_Error

    avntHex = Empty  ' Always start with empty variants and arrays
    vntBase = Empty
    vntTemp = Empty
    Erase astrHex()
    blnMyNumber = False

    ReDim astrHex(lngQtyReq)

    ' load hex array
    avntHex = Array("0", "1", "2", "3", "4", "5", "6", "7", _
                    "8", "9", "a", "b", "c", "d", "e", "f")

    ' loop thru the array of prime numbers
    For lngIndex = 0 To lngQtyReq - 1

        lngNumber = CDec(alngNumbers(lngIndex))   ' convert to decimal base
        vntBase = SquareRoot(lngNumber)           ' get the square root
        strHex = vbNullString

        '------------------------------------------------------
        ' To print out detailed data as shown in the above
        ' flowerbox, replace 999 with the prime number desired
        '------------------------------------------------------
        DoEvents
        If lngNumber = 999 Then
            blnMyNumber = True
        End If

        If blnMyNumber Then
            strSuffix = "th"  ' Change accordingly when finished
            Debug.Print CStr(lngIndex + 1) & strSuffix & " Prime number = " & CStr(lngNumber)
            Debug.Print "SQUARE ROOT of " & CStr(lngNumber) & " = " & CStr(vntBase)
            Debug.Print " "
        End If

        For lngLoop = 1 To HEX_LENGTH

            vntBase = CDec(CDec(vntBase) - CDec(Fix(vntBase)))  ' drop the whole number
            vntTemp = CDec(CDec(vntBase) * CDec(HEX_BASE))      ' multiply by hex base

            If blnMyNumber Then
                Debug.Print Format$(vntBase, String$(30, "@")) & " * 16 = " & CStr(vntTemp)
            End If

            lngWholeNum = CDec(Fix(vntTemp))        ' capture whole number
            strHex = strHex & avntHex(lngWholeNum)  ' append hex character to output string

            vntTemp = CDec(vntTemp)   ' Convert to decimal base
            vntBase = CDec(vntTemp)   ' Transfer to base value

        Next lngLoop

        If blnMyNumber Then
            Debug.Print " "
            Debug.Print "The hex representation of the fractional parts of the"
            Debug.Print "SQUARE ROOTS of " & CStr(lngNumber) & " is:   " & strHex
            Debug.Print String$(60, "*")
            blnMyNumber = False    ' Reset flag
        End If

        Select Case lngRetFmt
               Case e16Chars: strHex = Left$(strHex, HEX_LENGTH)   ' save all 16 characters
               Case eLeft8:   strHex = Left$(strHex, 8)              ' save first 8 characters
               Case eRight8:  strHex = Right$(strHex, 8)             ' save last 8 characters
        End Select

        astrHex(lngIndex) = LCase$(strHex)   ' verify data is lowercase

    Next lngIndex

    LoadSquareRoots = astrHex()

LoadSquareRoots_CleanUp:
    avntHex = Empty   ' Always empty variants when not needed
    vntBase = Empty
    vntTemp = Empty
    Erase astrHex()   ' Always empty arrays when not needed

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

LoadSquareRoots_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume LoadSquareRoots_CleanUp

End Function

' ***************************************************************************
' Routine:       LoadCubeRoots
'
' Description:   Got this explanation from Ask Dr. Math web site.
'                http://mathforum.org/dr.math/ 
'                 
'                There is a simple process for converting a base 10 decimal 
'                fraction to a "decimal" fraction in another base.  In this
'                case, hex will be our new base (16).  We repeatedly multiply
'                our decimal fraction by the new base, picking off the whole
'                number part each time as the next digit of the final output. 
'                 
' Example:   1st Prime number = 2
'            CUBE ROOT of 2 = 1.2599210498948731647672106073 
' 
'            0.2599210498948731647672106073 * 16 =  4.1587367983179706362753697168 
'            0.1587367983179706362753697168 * 16 =  2.5397887730875301804059154688 
'            0.5397887730875301804059154688 * 16 =  8.636620369400482886494647501 
'            0.636620369400482886494647501  * 16 = 10.185925910407726183914360016 
'            0.185925910407726183914360016  * 16 =  2.974814566523618942629760256 
'            0.974814566523618942629760256  * 16 = 15.597033064377903082076164096 
'            0.597033064377903082076164096  * 16 =  9.552529030046449313218625536 
'            0.552529030046449313218625536  * 16 =  8.840464480743189011498008576 
'            0.840464480743189011498008576  * 16 = 13.447431691891024183968137216 
'            0.447431691891024183968137216  * 16 =  7.158907070256386943490195456 
'            0.158907070256386943490195456  * 16 =  2.542513124102191095843127296 
'            0.542513124102191095843127296  * 16 =  8.680209985635057533490036736 
'            0.680209985635057533490036736  * 16 = 10.883359770160920535840587776 
'            0.883359770160920535840587776  * 16 = 14.133756322574728573449404416 
'            0.133756322574728573449404416  * 16 =  2.140101161195657175190470656 
'            0.140101161195657175190470656  * 16 =  2.241618579130514803047530496 
'                                                   |
'               Whole numbers are converted to hex__|
' 
'            The hex representation of the fractional parts of the
'            CUBE ROOTS of 2 is:   428a2f98d728ae22 
'
' Parameters:  lngQtyReq    - Number of items to return
'              lngRetFmt    - Return format.
'                                e16Chars - All 16 characters
'                                eLeft8   - First 8 characters
'                                eRight8  - Last 8 characters
'
'              alngNumbers() - An array of numbrs to be processed
'
' Returns:     An array of hex strings
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Nov-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 08-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated documentation
' 08-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modifed to display output details as shown in flowerbox above
'              if user hard codes the exact prime number below.
' ***************************************************************************
Private Function LoadCubeRoots(ByVal lngQtyReq As Long, _
                               ByVal lngRetFmt As enumHEX_RETURN_FORMAT, _
                               ByRef alngNumbers() As Long) As String()

    ' Called by HexRepresentation()

    Dim blnMyNumber As Boolean   ' Flag for flowerbox output
    Dim lngLoop     As Long      ' loop counter
    Dim lngIndex    As Long      ' loop counter
    Dim lngNumber   As Long      ' Temp number
    Dim lngWholeNum As Long      ' Whole number only
    Dim vntBase     As Variant   ' fractional part of number
    Dim vntTemp     As Variant   ' new computed value
    Dim avntHex     As Variant   ' base 16 hex values
    Dim strHex      As String    ' hex character
    Dim strSuffix   As String    ' Used for flowerbox output
    Dim astrHex()   As String    ' hex data

    Const HEX_BASE     As Double = 16#
    Const ROUTINE_NAME As String = "LoadCubeRoots"

    On Error GoTo LoadCubeRoots_Error

    avntHex = Empty  ' Always start with empty variants and arrays
    vntBase = Empty
    vntTemp = Empty
    Erase astrHex()
    blnMyNumber = False

    ReDim astrHex(lngQtyReq)

    ' load hex array
    avntHex = Array("0", "1", "2", "3", "4", "5", "6", "7", _
                    "8", "9", "a", "b", "c", "d", "e", "f")

    ' loop thru the array of prime numbers
    For lngIndex = 0 To lngQtyReq - 1

        lngNumber = CDec(alngNumbers(lngIndex))   ' convert to decimal base
        vntBase = CubeRoot(lngNumber)             ' get the cube root
        strHex = vbNullString

        '------------------------------------------------------
        ' To print out detailed data as shown in the above
        ' flowerbox, replace 999 with the prime number desired
        '------------------------------------------------------
        DoEvents
        If lngNumber = 999 Then
            blnMyNumber = True
        End If

        If blnMyNumber Then
            strSuffix = "th"  ' Change accordingly when finished
            Debug.Print CStr(lngIndex + 1) & strSuffix & " Prime number = " & CStr(lngNumber)
            Debug.Print "CUBE ROOT of " & CStr(lngNumber) & " = " & CStr(vntBase)
            Debug.Print " "
        End If

        For lngLoop = 1 To HEX_LENGTH

            vntBase = CDec(CDec(vntBase) - CDec(Fix(vntBase)))  ' drop the whole number
            vntTemp = CDec(CDec(vntBase) * CDec(HEX_BASE))      ' multiply by hex base

            If blnMyNumber Then
                Debug.Print Format$(vntBase, String$(30, "@")) & " * 16 = " & CStr(vntTemp)
            End If

            lngWholeNum = CDec(Fix(vntTemp))        ' capture whole number
            strHex = strHex & avntHex(lngWholeNum)  ' append hex character to output string

            vntTemp = CDec(vntTemp)   ' Convert to decimal base
            vntBase = CDec(vntTemp)   ' Transfer to base value

        Next lngLoop

        If blnMyNumber Then
            Debug.Print " "
            Debug.Print "The hex representation of the fractional parts of the"
            Debug.Print "CUBE ROOTS of " & CStr(lngNumber) & " is:   " & strHex
            Debug.Print String$(60, "*")
            blnMyNumber = False    ' Reset flag
        End If

        Select Case lngRetFmt
               Case e16Chars: strHex = Left$(strHex, HEX_LENGTH)   ' save all 16 characters
               Case eLeft8:   strHex = Left$(strHex, 8)              ' save first 8 characters
               Case eRight8:  strHex = Right$(strHex, 8)             ' save last 8 characters
        End Select

        astrHex(lngIndex) = LCase$(strHex)   ' verify data is lowercase

    Next lngIndex

    LoadCubeRoots = astrHex()

LoadCubeRoots_CleanUp:
    avntHex = Empty   ' Always empty variants when not needed
    vntBase = Empty
    vntTemp = Empty
    Erase astrHex()   ' Always empty arrays when not needed

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

LoadCubeRoots_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume LoadCubeRoots_CleanUp

End Function

' ***************************************************************************
' Routine:       CubeRoot
'
' Description:   Compute cube root of (vntInput), by iteration, to up to 29
'                digits.  The iteration continues until the limit of
'                precision is reached.
'
' Parameters:    vntInput - Value to be evaluated
'
' Returns:       Result of computation
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Aug-2002  Jay Tanner
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=21048&lngWId=1
' 17-Nov-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function CubeRoot(ByVal vntInput As Variant) As Variant

    ' Called by LoadCubeRoots()

    Dim vntTemp    As Variant   ' Argument - May be positive or negative value
    Dim vntWork1   As Variant   ' Any general approximation to cube root
    Dim vntWork2   As Variant   ' Next successive approximation to cube root
    Dim intCounter As Integer   ' Cycle loop control counter
    Dim strSign    As String    ' Sign of argument - Attached to result

    Const ROUTINE_NAME As String = "CubeRoot"

    On Error GoTo CubeRoot_Error

    vntWork1 = CDec(0)          ' Initialize (vntWork1) as decimal data type
    vntWork2 = CDec(0)          ' Initialize (vntWork2) as decimal data type
    vntTemp = Trim$(vntInput)   ' Check for invalid numeric argument

    If Not IsNumeric(vntTemp) Then
        GoTo CubeRoot_Error
    End If

    vntTemp = CDec(vntTemp)  ' Convert argument into decimal data type
    strSign = vbNullString             ' Consider the sign of the argument

    If vntTemp < 0 Then
        vntTemp = -vntTemp
        strSign = "-"
    End If

    If vntTemp = 0 Then
        vntTemp = 1
    Else
        vntWork1 = vntTemp ^ (1 / 3)  ' Use VB cube root as 1st approximation
        intCounter = 50               ' Set limit of iterations to 50 max

        ' A loop to grind out the cube root using a series of
        ' successive approximations, starting with (vntWork1).
        Do
            ' Compute next approx (vntWork2) from (vntWork1)
            vntWork2 = ((2 * vntWork1) + vntTemp / (vntWork1 * vntWork1)) / 3

            ' Check if finished
            If (vntWork2 = vntWork1) Or intCounter <= 0 Then
                Exit Do   ' exit Do..Loop
            End If

            vntWork1 = vntWork2           ' Update approx to current value
            intCounter = intCounter - 1   ' Update limit counter
        Loop
    End If

CubeRoot_CleanUp:
    CubeRoot = Trim$(strSign & vntWork2)

    vntWork1 = Empty  ' Always empty variants when not needed
    vntWork2 = Empty
    vntTemp = Empty

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

CubeRoot_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    GoTo CubeRoot_CleanUp

End Function

' ***************************************************************************
' Routine:       SquareRoot
'
' Description:   Compute square root of (vntInput), by iteration, to up to 29
'                digits.  The iteration continues until the limit of
'                precision is reached.
'
'                Generally, a square root refers to only positive arguments,
'                but this function will accept negative arguments and produce
'                an (imaginary) square root by returning a value with " i"
'                attached to the end.
'
' Parameters:    vntInput - Value to be evaluated
'
' Returns:       Result of computation
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Aug-2002  Jay Tanner
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=21048&lngWId=1
' 17-Nov-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function SquareRoot(ByVal vntInput As Variant) As Variant

    ' Called by LoadSquareRoots()

    Dim vntTemp       As Variant   ' Argument - May be positive or negative value
    Dim vntWork1      As Variant   ' Any general approximation to cube root
    Dim vntWork2      As Variant   ' Next successive approximation to cube root
    Dim intCounter    As Integer   ' Cycle loop control counter
    Dim strSqrMinus1  As String    ' Represents the square root of minus 1

    Const ROUTINE_NAME As String = "SquareRoot"

    On Error GoTo SquareRoot_Error

    strSqrMinus1 = vbNullString
    vntWork1 = CDec(0)          ' Initialize (vntWork1) as decimal data type
    vntWork2 = CDec(0)          ' Initialize (vntWork2) as decimal data type
    vntTemp = Trim$(vntInput)   ' Check for invalid numeric argument

    If Not IsNumeric(vntTemp) Then
        GoTo SquareRoot_Error
    End If

    vntTemp = CDec(vntTemp)  ' Convert argument into decimal data type

    ' Account for a negative argument
    If vntTemp < 0 Then
        vntTemp = -vntTemp
        strSqrMinus1 = " i"
    End If

    ' Check for zero argument
    If vntTemp = 0 Then
        SquareRoot = 0
        Exit Function
    End If

    vntWork1 = Sqr(vntTemp)   ' Use VB square root as 1st approximation
    intCounter = 50           ' Set limit of iterations to 50 max

    ' A loop to grind out the square root using a series of
    ' successive approximations, starting with (vntWork1).
    Do
        ' Compute next approx (vntWork2) from (vntWork1)
        vntWork2 = (vntWork1 + vntTemp / vntWork1) / 2

        ' Check if finished
        If (vntWork2 = vntWork1) Or intCounter <= 0 Then
            Exit Do   ' exit Do..Loop
        End If

        vntWork1 = vntWork2           ' Update approx to current value
        intCounter = intCounter - 1   ' Update limit counter
    Loop

SquareRoot_CleanUp:
    SquareRoot = Trim$(vntWork2 & strSqrMinus1)

    vntWork1 = Empty  ' Always empty variants when not needed
    vntWork2 = Empty
    vntTemp = Empty

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

SquareRoot_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    GoTo SquareRoot_CleanUp

End Function

Private Sub Class_Initialize()
    StopProcessing = False
End Sub



