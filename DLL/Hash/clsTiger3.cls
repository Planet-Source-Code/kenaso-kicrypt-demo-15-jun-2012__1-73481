VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTiger3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:       Tiger3 Hash Algorithm
'
' This module is my experimental version for Tiger-128 thru Tiger-512 bit
' output.  I call it Tiger3.  I have not had any problems with it thus far.
' If you should encounter any problems, please email me at:
'
'           Kenneth Ives  kenaso@tx.rr.com
'
' DO NOT CONTACT   Ross Anderson, http://www.cl.cam.ac.uk/users/rja14/
'                  Eli Biham, http://www.cs.technion.ac.il/~biham
'
' because they did not write this module nor are they responsible in any
' manner as to its content.
'
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
' *** WARNING *** WARNING *** WARNING *** WARNING *** WARNING *** WARNING ***
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
'
'    You acknowledge that this software is subject to the export control
'    laws and regulations of the United States ("U.S.") and agree to abide
'    by those laws and regulations. Under U.S. law, this software may not
'    be downloaded or otherwise exported, reexported, or transferred to
'    restricted countries, restricted end-users, or for restricted
'    end-uses. The U.S. currently has embargo restrictions against Cuba,
'    Iran, Iraq, Libya, North Korea, Sudan, and Syria. The lists of
'    restricted end-users are maintained on the U.S. Commerce Department's
'    Denied Persons List, the Commerce Department's Entity List, the
'    Commerce Department's List of Unverified Persons, and the U.S.
'    Treasury Department's List of Specially Designated Nationals and
'    Blocked Persons. In addition, this software may not be downloaded or
'    otherwise exported, reexported, or transferred to an end-user engaged
'    in activities related to weapons of mass destruction.
'
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
'
' Description:   Tiger3 is a fast New cHash function, designed to be very fast
'                on modern computers, and in particular on the state-of-the-art
'                64-bit computers (like DEC-Alpha), while it is still not
'                slower than other suggested hash functions on 32-bit machines.
'
'                Tiger hash has no usage restrictions nor patents. It can be
'                used freely, with the reference implementation, with other
'                implementations or with a modification to the reference
'                implementation (as long as it still implements Tiger2). We
'                only ask you to let us know about your implementation and to
'                cite the origin of Tiger and of the reference implementation.
'                http://www.cs.technion.ac.il/~biham/Reports/Tiger/
'
' My opinion:    Tiger2 is a very strong hash.  Unfortunately, the 32-bit
'                version seems to be just an off shoot and is not the primary
'                concern of the authors.  That is understandable because
'                everyone appears to be headed for 64-bit hash scenarios.
'                The test vectors, I believe, were created prior to the new
'                MD5 padding of Tiger2 (2007) and may be obsolete.  The time
'                stamp on the 32-bit version of the Tiger code is dated
'                30-Mar-1996.  When I contacted one of the authors, the response
'                was less than enthusiastic concerning the 32-bit version of
'                Tiger2.  Later, I learned the authors are deeply involved in
'                developing a hash for the SHA3 contest by NIST.  Good for them.
'                I wish their team the very best.
'
' Special Note:  I decided to make changes to the initial work data array
'                [malngHash()].  I believe that this is a cleaner and more
'                secure method as to the calculation of the Tiger hashes.
'                See Initialize() routine for more details.
'
' Reference:     A Fast New cHash function
'                http://www.cs.technion.ac.il/~biham/Reports/Tiger/
'
'                Original authors of the Tiger hash:
'                Ross Anderson, http://www.cl.cam.ac.uk/users/rja14/
'                Eli Biham, http://www.cs.technion.ac.il/~biham
'
'                Tiger 32-bit C source code:
'                http://www.cs.technion.ac.il/~biham/Reports/Tiger/tiger-src32.zip
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-SEP-2003  Kenneth Ives  kenaso@tx.rr.com
' 26-MAR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Removed my experimental versions until a later date.
' 02-MAY-2006  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote this module to use only long integers.  Much faster.
'              Found C source code at:
'              http://www.cs.technion.ac.il/~biham/Reports/Tiger/tiger-src32.zip
' 22-DEC-2006  Kenneth Ives  kenaso@tx.rr.com
'              ** MD5 padding - The data is stored in Little-Endian format
'                 with the Least Significant Bit (LSB) first.
'              ** SHA padding - The data is stored in Big-Endian format with
'                 the Most Significant Bit (MSB) first.
'              I elected to pad the data as per MD5.
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 07-MAY-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified output to 224, 256, 384, 512 bits.
' 27-Apr-2009  Kenneth Ives  kenaso@tx.rr.com
'              Modified Transform() routine for a better mix of data.
' 31-Dec-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Added property HashRounds() to set number of rounds for hashing.
'              - Added property HashMethod() to determine hash algorithm.
'              - Add hashed results to original hash values as last
'                calculation in Transform() routine for a better mix of data.
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Added ability to see progress display while hashing data.
'                See Transform() routine.
'              - Removed obsolete code.  Updated documentation.
' 23-Mar-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated Transform() routine for speed
' 31-Mar-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Modified the way to produce Tiger-224, Tiger-256, Tiger-384,
'                and Tiger-512 hashes by creating unique initialization values
'                for each.  See Initialize() routine for more details.
'              - Updated documentation.
' 19-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Updated reference documentation.
'              - Fixed logic bug concerning flow of events within the number
'                of passes being performed.  See Tranform() routine.
' 07-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug in KeySchedule() routine. One of the shift values
'              was 19 now is 9.
' 10-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Combined all Tiger hashes into a single module.
'              - Created unique initialization values for each of the Tiger
'                hashes except Tiger-192.
'              - Kept original values for constant arrays.
' 23-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Fixed a bug in KeySchedule() routine.  Corrected two of the
'                index pointers.
'              - Added some more comments.
' 04-May-2012  Kenneth Ives  kenaso@tx.rr.com
'              Added new property ReturnLowercse().  Designates if returned
'              hashed data string should be in upper or lowercase format.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "clsTiger3"
  Private Const KB_32       As Long = &H8000&      ' 32768
  Private Const MAX_BYTE    As Long = 255          ' Last byte value
  Private Const MSG_LEN_32  As Long = 128          ' 128 bytes long
  Private Const MSG_LEN_64  As Long = 256          ' 256 bytes long

' ***************************************************************************
' Enumerations
' ***************************************************************************
  Private Enum enumTIGER3_ALGORITHM
      eTIGER128  ' 0
      eTIGER160  ' 1
      eTIGER192  ' 2
      eTIGER224  ' 3
      eTIGER256  ' 4
      eTIGER384  ' 5
      eTIGER512  ' 6
  End Enum

' ***************************************************************************
' Module Variables
'
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (Long)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a lng Hash
' Variable name:     malngHash
'
' ***************************************************************************
  Private mstrBitsInHex    As String
  Private mblnRetLowercase As Boolean
  Private mcurMaxSize      As Currency
  Private mcurAccrued      As Currency
  Private mcurBitCount     As Currency
  Private mcurCurrBitCnt   As Currency
  Private mlngHashRounds   As Long
  Private mlngMsgLength    As Long
  Private malngHash()      As Long
  Private T1()             As Long
  Private T2()             As Long
  Private T3()             As Long
  Private T4()             As Long
  Private mobjMath32       As cMath32
  Private mlngHashMethod   As enumTIGER3_ALGORITHM


' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************

' Update progress bar
Public Event HashProgress(ByVal lngProgress As Long)


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let HashRounds(ByVal lngData As Long)

    ' Three (3) passes are minimum for hashing the data.
    ' Use four or more passes when you need extra security.
    ' I prefer six (6) passes.
    Select Case lngData
           Case 3 To 25
                mlngHashRounds = lngData   ' Good choice
           Case Else
                mlngHashRounds = 6         ' Use default value
    End Select

End Property

Public Property Let HashMethod(ByVal lngData As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashMethod = eTIGER192
    Select Case lngData

           Case 0 To 2    ' eTIGER128, eTIGER160, eTIGER192
                mlngHashMethod = lngData
                mlngMsgLength = MSG_LEN_32

           Case 3 To 6    ' eTIGER224, eTIGER256, eTIGER384, eTIGER512
                mlngHashMethod = lngData
                mlngMsgLength = MSG_LEN_64

           Case Else  ' Use default value
                mlngHashMethod = eTIGER192
                mlngMsgLength = MSG_LEN_32
    End Select

End Property

Public Property Let ReturnLowercase(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.ReturnLowercase = False
    mblnRetLowercase = blnData
End Property


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       HashFile
'
' Description:   Function to quickly convert a file into a hex string
'
' Syntax:
'    With mobjTiger3
'        abytData() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)  ' convert file location to byte array 
'        mabytHashed() = .HashFile(abytData())                        ' hash data and return as Byte array
'        txtData.Text = StrConv(mabytHashed(), vbUnicode)             ' convert byte array to string data
'    End With
' 
' Parameters:    abytData() - name of file to be hashed in byte array
'
' Returns:       Hashed data string in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array for security purposes
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 10-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              Adjusted mcurMaxSize so progress bar display would be close
'              to same hash finish time
' ***************************************************************************
Public Function HashFile(ByRef abytFile() As Byte) As Byte()

    Dim hFile        As Long
    Dim lngBlockSize As Long
    Dim curAmtLeft   As Currency
    Dim curFilePos   As Currency
    Dim strSource    As String
    Dim abytData()   As Byte
    Dim objBigFiles  As cBigFiles

    On Error GoTo HashFile_CleanUp

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    Set objBigFiles = New cBigFiles  ' Instantiate class object
    ResetVariables                   ' Reset class variables
    Initialize                       ' load the work and constant arrays
    curFilePos = 0@                  ' set to first position in the file
    mcurAccrued = 0@                 ' Init accrued progress

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo HashFile_CleanUp
    End If

    With objBigFiles
        strSource = ByteArrayToString(abytFile())           ' Convert byte array to string
        .CalcFileSize strSource, curAmtLeft, mstrBitsInHex  ' Get the size of the file
        mcurBitCount = curAmtLeft * 8@                                     ' Get total bit count for comparison
        mstrBitsInHex = Right$(String$(16, "0") & Hex$(mcurBitCount), 16)  ' Convert bit count into hex
        mstrBitsInHex = mobjMath32.SwapEndianHex(mstrBitsInHex)            ' Reverse hex values in string

        ' Calculate amount with number
        ' of hashing rounds requested
        mcurMaxSize = curAmtLeft * mlngHashRounds

        ' If zero byte file then leave
        If curAmtLeft < 1 Then
            gblnStopProcessing = True
            GoTo HashFile_CleanUp
        End If

        ' Open source file
        If Not .OpenReadOnly(strSource, hFile) Then
            gblnStopProcessing = True
            GoTo HashFile_CleanUp
        End If

        Do
            lngBlockSize = GetBlockSize(curAmtLeft)       ' Process the source file in chunks
            ReDim abytData(lngBlockSize - 1)              ' Size receiving array
            curAmtLeft = curAmtLeft - CCur(lngBlockSize)  ' Adjust amount of file left to process

            ' Read source file
            If Not .API_ReadFile(hFile, curFilePos, abytData()) Then
                gblnStopProcessing = True
                Exit Do  ' exit Do..Loop
            End If

            ' Adjust pointers accordingly
            curFilePos = curFilePos + CCur(UBound(abytData) + 1)

            ' Process the message data into the appropriate
            ' string length and perform the bit manipulations.
            If Not LoadWordArray(abytData()) Then
                gblnStopProcessing = True
            End If

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

        .API_CloseFile hFile    ' Close all files opened by this routine
    End With

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo HashFile_CleanUp
    End If

    ' Concatenate hashed data into a string
    HashFile = FormatOutput()

HashFile_CleanUp:
    objBigFiles.API_CloseFile hFile  ' Close all files opened by this routine
    ResetVariables                   ' Reset class variables
    Set objBigFiles = Nothing        ' Always free objects from memory

End Function

' ***************************************************************************
' Routine:       HashString
'
' Description:   Function to digest a text string and output the result as a
'                string of hexadecimal characters.
'
' Syntax:
'    With mobjTiger3
'        abytData() = StrConv("abc", vbFromUnicode)      ' convert string data to byte array 
'        abytHashed() = .HashFile(abytData())            ' hash data and return as Byte array
'        txtData.Text = StrConv(abytHashed(), vbUnicode) ' convert byte array to string data
'    End With
' 
' Parameters:    abytData() - data to be hashed in byte array
'
' Returns:       Hashed data string in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array for security purposes
' 10-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              Adjusted mcurMaxSize so progress bar display would be close
'              to same hash finish time
' ***************************************************************************
Public Function HashString(ByRef abytData() As Byte) As Byte()

    On Error GoTo HashString_CleanUp

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

    ResetVariables   ' Reset class variables
    Initialize       ' load work and constant arrays

    mcurAccrued = 0@                                                   ' Init accrued progress
    mcurMaxSize = CCur(UBound(abytData) + 1)                           ' Capture data size
    mcurBitCount = CCur((UBound(abytData) + 1) * 8)                    ' Get bit count for for comparison
    mstrBitsInHex = Right$(String$(16, "0") & Hex$(mcurBitCount), 16)  ' Convert bit count into hex
    mstrBitsInHex = mobjMath32.SwapEndianHex(mstrBitsInHex)            ' Reverse hex values in string

    ' Calculate amount with number
    ' of hashing rounds requested
    mcurMaxSize = mcurMaxSize * mlngHashRounds

    ' Process the message data into the appropriate
    ' string length and perform the bit manipulations.
    If LoadWordArray(abytData()) Then

        ' Concatenate hashed data into a string
        HashString = FormatOutput()
    Else
        gblnStopProcessing = True
    End If

HashString_CleanUp:
    ResetVariables     ' Reset class variables

End Function


' ***************************************************************************
' ****                  Internal functions and Procedures                ****
' ***************************************************************************

' ***************************************************************************
' Routine:       FormatOutput
'
' Description:   Concatenate hashed values into one string.
'
'                Output length in hex format:
'                    TIGER-128 =  32 characters
'                    TIGER-160 =  40 characters
'                    TIGER-192 =  48 characters
'                    TIGER-224 =  56 characters
'                    TIGER-256 =  64 characters
'                    TIGER-384 =  96 characters
'                    TIGER-512 = 128 characters
'
' Returns:       Concatenated hashed data string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-DEC-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 10-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              Combined all Tiger hashes
' ***************************************************************************
Private Function FormatOutput() As Byte()

    Dim strHex      As String
    Dim intIndex    As Integer
    Dim intPointer  As Integer
    Dim lngTemp1    As Long
    Dim lngTemp2    As Long

    strHex = Space$(150)  ' preload output buffer
    intPointer = 1        ' Start position in output buffer

    With mobjMath32
        For intIndex = 0 To (UBound(malngHash) - 1) Step 2

            ' Convert data to Big-Endian format
            lngTemp1 = .SwapEndianLong(malngHash(intIndex))
            lngTemp2 = .SwapEndianLong(malngHash(intIndex + 1))

            ' Concatenate Hash values into one
            ' string using alternating positions
            Mid$(strHex, intPointer, 16) = .LongToHex(lngTemp2) & _
                                           .LongToHex(lngTemp1)
            intPointer = intPointer + 16   ' increment output pointer

        Next intIndex
    End With

    If mblnRetLowercase Then
        strHex = LCase$(Trim$(strHex))   ' Convert string to lowercase
    Else
        strHex = UCase$(Trim$(strHex))   ' Convert string to uppercase
    End If

    ' Adjust according to desired length in bytes (characters)
    Select Case mlngHashMethod
           Case eTIGER128: strHex = Left$(strHex, 32)
           Case eTIGER160: strHex = Left$(strHex, 40)
           Case eTIGER192: strHex = Left$(strHex, 48)
           Case eTIGER224: strHex = Left$(strHex, 56)
           Case eTIGER256: strHex = Left$(strHex, 64)
           Case eTIGER384: strHex = Left$(strHex, 96)
           Case eTIGER512: strHex = Left$(strHex, 128)
    End Select

    FormatOutput = StringToByteArray(strHex)   ' convert string data to byte array

    RaiseEvent HashProgress(100)   ' Update progress bar

End Function

' ***************************************************************************
' Routine:       LoadWordArray
'
' Description:   All data is stored in Big_Endian format with the Most 
'                Significant Bit (MSB) first. 
' 
'                The message 'M' shall be padded before hash computation
'                begins. The purpose of this padding is to ensure that the
'                padded message is a multiple of 512 or 1024 bits, depending
'                on the algorithm.  Data will be converted to Little_Endian
'                format, with the Least Significant Bit (MSB) first, in the
'                Transform() routine prior to being manipulated.
'
'         32-Bit Format
'                Suppose the length of the message 'M', in bits, is 'l'
'                bits. Append the bit '1' to the end of the message,
'                followed by 'k' zero bits, where 'k' is the smallest
'                non-negative solution to the equation l+1+k=448 mod 512.
'                Then append the 64-bit block that is equal to the number
'                'l' expressed using a binary representation. The length of
'                the padded message should now be a multiple of 512 bits.
' 
'         64-Bit Format
'                Suppose the length of the message 'M', in bits, is 'l'
'                bits. Append the bit '1' to the end of the message,
'                followed by 'k' zero bits, where 'k' is the smallest
'                non-negative solution to the equation l+1+k=896 mod 1024.
'                Then append the 128-bit block that is equal to the number
'                'l' expressed using a binary representation. The length of
'                the padded message should now be a multiple of 1024 bits.
' 
' Parameters:    abytData() - message to be hashed in a byte array 
' 
' Returns:       TRUE - process completed successfully.
'                FALSE - user reqested to stop processing or an error
'                      occurred.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Function LoadWordArray(ByRef abytData() As Byte) As Boolean

    Dim lngIndex       As Long
    Dim lngPadding     As Long
    Dim lngMaxBytes    As Long
    Dim lngPosition    As Long
    Dim lngSpaceAlloc  As Long
    Dim lngMsgLenInHex As Long
    Dim strMsg         As String
    Dim strHex         As String
    Dim strTemp        As String
    Dim abytMsg()      As Byte

    On Error GoTo LoadWordArray_Error

    LoadWordArray = True  ' Preset flag to TRUE

    Select Case mlngHashMethod

           Case eTIGER128, eTIGER160, eTIGER192
                lngSpaceAlloc = 512
                lngMaxBytes = 64

           Case eTIGER224, eTIGER256, eTIGER384, eTIGER512
                lngSpaceAlloc = 1024
                lngMaxBytes = 128
    End Select

    lngMsgLenInHex = Len(mstrBitsInHex)
    strMsg = ByteArrayToString(abytData())                   ' convert to string data
    mcurCurrBitCnt = mcurCurrBitCnt + CCur(Len(strMsg) * 8)  ' calc the ongoing bit count

    With mobjMath32
        ' Incoming message is in equal multiples
        If Len(strMsg) Mod lngMaxBytes = 0 Then

            ' Process message in equal lengths. There
            ' is no need for any padding at this time.
            Do
                Erase abytMsg()          ' Always start with an empty array
                strTemp = vbNullString   ' Empty string variables
                strHex = vbNullString

                strTemp = Mid$(strMsg, 1, lngMaxBytes)   ' Capture chunk of data message
                strMsg = Mid$(strMsg, lngMaxBytes + 1)   ' Resize data message to what was not used

                ' See if we have some data to work with
                If Len(strTemp) > 0 Then

                    ' the string of data is the exact length needed
                    If Len(strTemp) = lngMaxBytes Then

                        strHex = Space$(lngSpaceAlloc)
                        lngPosition = 1

                        ' convert data to hex string
                        For lngIndex = 1 To Len(strTemp)
                            Mid$(strHex, lngPosition, 2) = Right$("00" & Hex$(Asc(Mid$(strTemp, lngIndex, 1))), 2)
                            lngPosition = lngPosition + 2
                        Next lngIndex

                        abytMsg() = .HexStringToByteArray(strHex)  ' Convert to byte array
                        Transform abytMsg()                        ' Hash data

                    End If
                Else
                    Exit Do  ' exit Do..Loop
                End If

            Loop While Len(strMsg) > 0
        End If

        strTemp = vbNullString   ' Empty string variables
        strHex = vbNullString

        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            GoTo LoadWordArray_CleanUp
        End If

        ' if the length of the message is not
        ' in equal multiples
        If Len(strMsg) > 0 Then

            If Len(strMsg) >= lngMaxBytes Then

                Do
                    Erase abytMsg()          ' Always start with an empty array
                    strTemp = vbNullString   ' Empty string variables
                    strHex = vbNullString

                    strTemp = Mid$(strMsg, 1, lngMaxBytes)   ' Capture chunk of data message
                    strMsg = Mid$(strMsg, lngMaxBytes + 1)   ' Resize data message to what was not used

                    If Len(strTemp) = lngMaxBytes Then

                        strHex = Space$(lngSpaceAlloc)
                        lngPosition = 1

                        ' convert data to hex string
                        For lngIndex = 1 To Len(strTemp)
                            Mid$(strHex, lngPosition, 2) = Right$("00" & Hex$(Asc(Mid$(strTemp, lngIndex, 1))), 2)
                            lngPosition = lngPosition + 2
                        Next lngIndex

                        abytMsg() = .HexStringToByteArray(strHex)  ' Convert to byte array
                        Transform abytMsg()                        ' Hash data

                    Else
                        strHex = vbNullString
                        Exit Do  ' exit Do..Loop
                    End If

                Loop

                ' An error occurred or user opted to STOP processing
                If gblnStopProcessing Then
                    GoTo LoadWordArray_CleanUp
                End If

            Else
                strTemp = Mid$(strMsg, 1, lngMaxBytes)   ' Capture chunk of data message
                strMsg = Mid$(strMsg, lngMaxBytes + 1)   ' Resize data message to what was not used
            End If

            ' the string of data is less than the length needed
            If (Len(strTemp) > 0) And (Len(strMsg) = 0) Then

                strTemp = strTemp & Chr$(128)
                strHex = Space$(lngSpaceAlloc)
                lngPosition = 1

                ' convert data to hex string
                For lngIndex = 1 To Len(strTemp)
                    Mid$(strHex, lngPosition, 2) = Right$("00" & Hex$(Asc(Mid$(strTemp, lngIndex, 1))), 2)
                    lngPosition = lngPosition + 2
                Next lngIndex

                strHex = Trim$(strHex)  ' remove all leading and trailing blanks
                strTemp = vbNullString

                ' if the len of the hex string plus the element length is less
                ' than or equal to the required record length then append the
                ' appropriate number of zeroes and the hex representation of
                ' the message length
                If (Len(strHex) < mlngMsgLength) Then

                    ' Calculate amount of padding needed
                    lngPadding = mlngMsgLength - Len(strHex)

                    Select Case lngPadding

                           Case lngMsgLenInHex
                                ' If the amount of padding needed is the
                                ' exact length of the bit string then
                                ' append, process and leave
                                strHex = strHex & mstrBitsInHex

                           Case Is < lngMsgLenInHex
                                ' *** String 1 ***
                                ' Hex string is exact length needed.
                                ' Will need to create an extra string.
                                strHex = strHex & String$(lngPadding, "0")
                                abytMsg() = .HexStringToByteArray(strHex)  ' Convert to byte array
                                Transform abytMsg()                        ' Hash data

                                ' *** String 2 ***
                                ' Create a string of zeroes and the
                                ' bit count of the message length.
                                strHex = vbNullString
                                lngPadding = mlngMsgLength - lngMsgLenInHex
                                strHex = String$(lngPadding, "0") & mstrBitsInHex

                           Case Is > lngMsgLenInHex
                                ' Adjust amount of padding to fill in
                                ' the gap between data and number of bits
                                lngPadding = mlngMsgLength - (Len(strHex) + lngMsgLenInHex)
                                strHex = strHex & String$(lngPadding, "0") & mstrBitsInHex
                    End Select

                    abytMsg() = .HexStringToByteArray(strHex)  ' Convert to byte array
                    Transform abytMsg()                        ' Hash data
                    GoTo LoadWordArray_CleanUp

                Else
                    ' *** String 1 ***
                    ' Hex string is exact length needed.
                    ' Will need to create an extra string.
                    abytMsg() = .HexStringToByteArray(strHex)  ' Convert to byte array
                    Transform abytMsg()                        ' Hash data

                    ' *** String 2 ***
                    ' Create a string of zeroes and the
                    ' bit count of the message length.
                    strHex = vbNullString
                    lngPadding = mlngMsgLength - lngMsgLenInHex
                    strHex = String$(lngPadding, "0") & mstrBitsInHex

                    abytMsg() = .HexStringToByteArray(strHex)  ' Convert to byte array
                    Transform abytMsg()                        ' Hash data
                    GoTo LoadWordArray_CleanUp

                End If
            End If
        End If

        ' No data found in current work string
        If Len(strMsg) = 0 Then

            ' See if we have reached the end
            ' of the message file or string
            If mcurCurrBitCnt = mcurBitCount Then

                ' End of message indicator is the starting character followed by a
                ' string of zeroes and the hex representation of the message length
                lngPadding = mlngMsgLength - (lngMsgLenInHex + 2)
                strHex = Hex$(128) & String$(lngPadding, "0") & mstrBitsInHex
                abytMsg() = .HexStringToByteArray(strHex)  ' Convert to byte array
                Transform abytMsg()                        ' Hash data
            End If
        End If
    End With

LoadWordArray_CleanUp:
    Erase abytMsg()         ' Always empty arrays when not needed
    strMsg = vbNullString   ' Verify string variables are empty
    strTemp = vbNullString
    strHex = vbNullString

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        LoadWordArray = False
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

LoadWordArray_Error:
    ErrorMsg MODULE_NAME, "LoadWordArray", Err.Description
    gblnStopProcessing = True
    Resume LoadWordArray_CleanUp

End Function

' ***************************************************************************
' Routine:       Transform
'
' Description:   Perform the TIGER main hashing functions here.
'
' Parameters:    abytMsg() - Message chunk in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Feb-2006  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote this routine
' 19-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Changed the input parameter to byte array
' 27-Apr-2009  Kenneth Ives  kenaso@tx.rr.com
'              Modified the first looping action to more thoroughly mix
'              the data
' 02-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Add hashed results to original hash values as last calculation
'              for a better mix of data.
' 28-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added RaiseEvents to track hash progress
' 19-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed logic bug concerning flow of events within the number
'              of passes being performed.
' ***************************************************************************
Private Sub Transform(ByRef abytMsg() As Byte)

    Dim lngIdx      As Long
    Dim lngLoop     As Long
    Dim lngTemp     As Long
    Dim lngStep     As Long
    Dim lngLoops    As Long
    Dim lngIndex    As Long
    Dim lngRound    As Long
    Dim lngProgress As Long
    Dim HH()        As Long
    Dim alngWord()  As Long

    lngIndex = 0

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If

    ' Determine work array sizes and loop iterations
    Select Case mlngHashMethod

           Case eTIGER128, eTIGER160, eTIGER192
                ReDim HH(6)
                ReDim alngWord(16)
                lngLoops = 1

           Case eTIGER224, eTIGER256
                ReDim HH(12)
                ReDim alngWord(32)
                lngLoops = 2

           Case eTIGER384, eTIGER512
                ReDim HH(18)
                ReDim alngWord(32)
                lngLoops = 3
    End Select

    ' ******** Message preparation **************************
    ' Convert data to Little-Endian format
    For lngIdx = 0 To UBound(alngWord) - 1

        alngWord(lngIdx) = UnsignedJoin(abytMsg(lngIndex), _
                                        abytMsg(lngIndex + 1), _
                                        abytMsg(lngIndex + 2), _
                                        abytMsg(lngIndex + 3))
        lngIndex = lngIndex + 4

    Next lngIdx
    ' ******** End of message preparation ********************

    ' Set inital values for the hash operators
    '
    ' save_abc \
    CopyMemory HH(0), malngHash(0), UBound(malngHash) * 4&

    lngStep = 0   ' Initialize index pointer

    ' Loop 1 - will process elements 0-5
    ' Loop 2 - will process elements 6-11
    ' Loop 3 - will process elements 12-17
    For lngLoop = 1 To lngLoops

        ' pass(a0,a1,b0,b1,c0,c1,5); \
        Pass HH(0 + lngStep), HH(1 + lngStep), HH(2 + lngStep), _
             HH(3 + lngStep), HH(4 + lngStep), HH(5 + lngStep), _
             alngWord(), 5

        ' key_schedule; \
        KeySchedule alngWord()

        ' pass(c0,c1,a0,a1,b0,b1,7); \
        Pass HH(4 + lngStep), HH(5 + lngStep), HH(0 + lngStep), _
             HH(1 + lngStep), HH(2 + lngStep), HH(3 + lngStep), _
             alngWord(), 7

        ' key_schedule; \
        KeySchedule alngWord()

        ' pass(b0,b1,c0,c1,a0,a1,9); \
        Pass HH(2 + lngStep), HH(3 + lngStep), HH(4 + lngStep), _
             HH(5 + lngStep), HH(0 + lngStep), HH(1 + lngStep), _
             alngWord(), 9

        lngStep = lngStep + 6

    Next lngLoop

    ' Three passes (rounds) are minimum for hashing the data.
    ' Use four or more passes when you need extra security.
    ' I prefer six passes.
    '
    ' for(pass_no=3; pass_no<PASSES; pass_no++) { \
    '   key_schedule \
    '   pass(a0,a1,b0,b1,c0,c1,9); \
    '   tmpa=a0; a0=c0; c0=b0; b0=tmpa; \
    '   tmpa=a1; a1=c1; c1=b1; b1=tmpa;} \
    For lngRound = 1 To mlngHashRounds

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If

        lngStep = 0   ' Initialize index pointer

        For lngLoop = 1 To lngLoops

            ' key_schedule \
            KeySchedule alngWord()

            ' pass(a0,a1,b0,b1,c0,c1,9); \
            Pass HH(0 + lngStep), HH(1 + lngStep), HH(2 + lngStep), _
                 HH(3 + lngStep), HH(4 + lngStep), HH(5 + lngStep), _
                 alngWord(), 9

            ' tmpa=a0; a0=c0; c0=b0; b0=tmpa; \
            lngTemp = HH(0 + lngStep)
            HH(0 + lngStep) = HH(4 + lngStep)
            HH(4 + lngStep) = HH(2 + lngStep)
            HH(2 + lngStep) = lngTemp

            ' tmpa=a1; a1=c1; c1=b1; b1=tmpa;} \
            lngTemp = HH(1 + lngStep)
            HH(1 + lngStep) = HH(5 + lngStep)
            HH(5 + lngStep) = HH(3 + lngStep)
            HH(3 + lngStep) = lngTemp

            lngStep = lngStep + 6

        Next lngLoop
    Next lngRound

    lngStep = 0   ' Initialize index pointer

    ' xor64(a0, a1, aa0, aa1); \
    ' sub64(b0, b1, bb0, bb1); \
    ' add64(c0, c1, cc0, cc1); \
    For lngLoop = 1 To lngLoops

        Xor64 HH(0 + lngStep), HH(1 + lngStep), malngHash(0 + lngStep), malngHash(1 + lngStep)
        Sub64 HH(2 + lngStep), HH(3 + lngStep), malngHash(2 + lngStep), malngHash(3 + lngStep)
        Add64 HH(4 + lngStep), HH(5 + lngStep), malngHash(4 + lngStep), malngHash(5 + lngStep)
        lngStep = lngStep + 6

    Next lngLoop

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        Exit Sub
    End If

    ' 19-Feb-2012 Update progress bar. Separate
    '             steps for easier debugging
    mcurAccrued = mcurAccrued + (UBound(abytMsg) * mlngHashRounds)
    lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
    RaiseEvent HashProgress(lngProgress)

    ' 02-Jan-2010 Add mixed hash operators to original operators
    With mobjMath32
        For lngIdx = 0 To UBound(HH) - 1
            malngHash(lngIdx) = .UnsignedAdd(malngHash(lngIdx), HH(lngIdx))
        Next lngIdx
    End With

End Sub

Private Function UnsignedJoin(ByVal a As Byte, _
                              ByVal b As Byte, _
                              ByVal c As Byte, _
                              ByVal d As Byte) As Long
    Dim abytData(4) As Byte

    abytData(0) = a
    abytData(1) = b
    abytData(2) = c
    abytData(3) = d

    ' Join 4 x 8-bit bytes into one 32-bit word
    With mobjMath32
        UnsignedJoin = .SwapEndianLong(.ByteArrayToLong(abytData()))
    End With
    
End Function

' ***************************************************************************
' Routine:       Bit manipulation routines
'
' Description:   these routines are considered part of a single pass.
'
' ***************************************************************************

' #define pass(a0,a1,b0,b1,c0,c1,mul) \
' a0,a1,b0,b1,c0,c1 values are manipulated and returned
Private Sub Pass(ByRef a0 As Long, _
                 ByRef a1 As Long, _
                 ByRef b0 As Long, _
                 ByRef b1 As Long, _
                 ByRef c0 As Long, _
                 ByRef c1 As Long, _
                 ByRef alngWord() As Long, _
                 ByVal lngMul As Long)

    Dim lngStep  As Long
    Dim lngLoop  As Long
    Dim lngLoops As Long

    lngStep = 0   ' Initialize index pointer

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If

    Select Case mlngHashMethod

           Case eTIGER128, eTIGER160, eTIGER192
                lngLoops = 1

           Case eTIGER224, eTIGER256, eTIGER384, eTIGER512
                lngLoops = 2
    End Select

    ' round(a0,a1,b0,b1,c0,c1,x00,x01,mul); \
    ' round(b0,b1,c0,c1,a0,a1,x10,x11,mul); \
    ' round(c0,c1,a0,a1,b0,b1,x20,x21,mul); \
    ' round(a0,a1,b0,b1,c0,c1,x30,x31,mul); \
    ' round(b0,b1,c0,c1,a0,a1,x40,x41,mul); \
    ' round(c0,c1,a0,a1,b0,b1,x50,x51,mul); \
    ' round(a0,a1,b0,b1,c0,c1,x60,x61,mul); \
    ' round(b0,b1,c0,c1,a0,a1,x70,x71,mul);
    '
    ' First loop will process word array elements 0-15
    ' Second loop will process word array elements 16-31
    For lngLoop = 1 To lngLoops

        CalcRound a0, a1, b0, b1, c0, c1, alngWord(0 + lngStep), alngWord(1 + lngStep), lngMul
        CalcRound b0, b1, c0, c1, a0, a1, alngWord(2 + lngStep), alngWord(3 + lngStep), lngMul
        CalcRound c0, c1, a0, a1, b0, b1, alngWord(4 + lngStep), alngWord(5 + lngStep), lngMul
        CalcRound a0, a1, b0, b1, c0, c1, alngWord(6 + lngStep), alngWord(7 + lngStep), lngMul
        CalcRound b0, b1, c0, c1, a0, a1, alngWord(8 + lngStep), alngWord(9 + lngStep), lngMul
        CalcRound c0, c1, a0, a1, b0, b1, alngWord(10 + lngStep), alngWord(11 + lngStep), lngMul
        CalcRound a0, a1, b0, b1, c0, c1, alngWord(12 + lngStep), alngWord(13 + lngStep), lngMul
        CalcRound b0, b1, c0, c1, a0, a1, alngWord(14 + lngStep), alngWord(15 + lngStep), lngMul

        ' This allows the processing
        ' of all parts of the message
        lngStep = lngStep + 16

    Next lngLoop

End Sub

' #define round(a0,a1,b0,b1,c0,c1,x0,x1,mul) \
' a0,a1,b0,b1,c0,c1 values are manipulated and returned
Private Sub CalcRound(ByRef a0 As Long, _
                      ByRef a1 As Long, _
                      ByRef b0 As Long, _
                      ByRef b1 As Long, _
                      ByRef c0 As Long, _
                      ByRef c1 As Long, _
                      ByVal lngWord1 As Long, _
                      ByVal lngWord2 As Long, _
                      ByVal lngMul As Long)

    Dim lngTemp0 As Long
    Dim lngTemp1 As Long

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If

    ' xor64(c0, c1, x0, x1); \
    Xor64 c0, c1, lngWord1, lngWord2

    With mobjMath32
        ' temp0  = t1[((c0)>>(0*8))&0xFF][0] ; \
        ' temp1  = t1[((c0)>>(0*8))&0xFF][1] ; \
        ' temp0 ^= T2[((c0)>>(2*8))&0xFF][0] ; \
        ' temp1 ^= T2[((c0)>>(2*8))&0xFF][1] ; \
        ' temp0 ^= t3[((c1)>>(0*8))&0xFF][0] ; \
        ' temp1 ^= t3[((c1)>>(0*8))&0xFF][1] ; \
        ' temp0 ^= t4[((c1)>>(2*8))&0xFF][0] ; \
        ' temp1 ^= t4[((c1)>>(2*8))&0xFF][1] ; \
        lngTemp0 = T1(0, c0 And &HFF&)
        lngTemp1 = T1(1, c0 And &HFF&)
        lngTemp0 = lngTemp0 Xor T2(0, .w32Shift(c0, -16) And &HFF&)
        lngTemp1 = lngTemp1 Xor T2(1, .w32Shift(c0, -16) And &HFF&)
        lngTemp0 = lngTemp0 Xor T3(0, c1 And &HFF&)
        lngTemp1 = lngTemp1 Xor T3(1, c1 And &HFF&)
        lngTemp0 = lngTemp0 Xor T4(0, .w32Shift(c1, -16) And &HFF&)
        lngTemp1 = lngTemp1 Xor T4(1, .w32Shift(c1, -16) And &HFF&)

        ' sub64(a0, a1, temp0, temp1); \
        Sub64 a0, a1, lngTemp0, lngTemp1

        ' temp0  = t4[((c0)>>(1*8))&0xFF][0] ; \
        ' temp1  = t4[((c0)>>(1*8))&0xFF][1] ; \
        ' temp0 ^= t3[((c0)>>(3*8))&0xFF][0] ; \
        ' temp1 ^= t3[((c0)>>(3*8))&0xFF][1] ; \
        ' temp0 ^= T2[((c1)>>(1*8))&0xFF][0] ; \
        ' temp1 ^= T2[((c1)>>(1*8))&0xFF][1] ; \
        ' temp0 ^= t1[((c1)>>(3*8))&0xFF][0] ; \
        ' temp1 ^= t1[((c1)>>(3*8))&0xFF][1] ; \
        lngTemp0 = T4(0, .w32Shift(c0, -8) And &HFF&)
        lngTemp1 = T4(1, .w32Shift(c0, -8) And &HFF&)
        lngTemp0 = lngTemp0 Xor T3(0, .w32Shift(c0, -24) And &HFF&)
        lngTemp1 = lngTemp1 Xor T3(1, .w32Shift(c0, -24) And &HFF&)
        lngTemp0 = lngTemp0 Xor T2(0, .w32Shift(c1, -8) And &HFF&)
        lngTemp1 = lngTemp1 Xor T2(1, .w32Shift(c1, -8) And &HFF&)
        lngTemp0 = lngTemp0 Xor T1(0, .w32Shift(c1, -24) And &HFF&)
        lngTemp1 = lngTemp1 Xor T1(1, .w32Shift(c1, -24) And &HFF&)

        ' add64(b0, b1, temp0, temp1); \
        Add64 b0, b1, lngTemp0, lngTemp1

    End With

    ' if((mul)==5) \
    '    {mul5(b0, b1);} \
    ' else \
    '    if((mul)==7) \
    '       {mul7(b0, b1);} \
    '    else \
    '       {mul9(b0, b1)};
    Select Case lngMul
           Case 5: Mul5 b0, b1
           Case 7: Mul7 b0, b1
           Case 9: Mul9 b0, b1
    End Select

End Sub

' #define key_schedule \
Private Sub KeySchedule(ByRef alngWord() As Long)

    ' key assignments (my reference)
    '  0 = X00    1 = X01
    '  2 = X10    3 = X11
    '  4 = X20    5 = X21
    '  6 = X30    7 = X31
    '  8 = X40    9 = X41
    ' 10 = X50   11 = X51
    ' 12 = X60   13 = X61
    ' 14 = X70   15 = X71

    Dim lngStep  As Long
    Dim lngLoop  As Long
    Dim lngLoops As Long

    Const K_12 As Long = &H1234567
    Const K_89 As Long = &H89ABCDEF
    Const K_A5 As Long = &HA5A5A5A5

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If

    lngStep = 0   ' Initialize index pointer

    Select Case mlngHashMethod

           Case eTIGER128, eTIGER160, eTIGER192
                lngLoops = 1

           Case eTIGER224, eTIGER256, eTIGER384, eTIGER512
                lngLoops = 2
    End Select

    With mobjMath32

        ' First loop will process word array elements 0-15
        ' Second loop will process word array elements 16-31
        For lngLoop = 1 To lngLoops

            ' sub64 x00, x01, x70^0xA5A5A5A5, x71^0xA5A5A5A5); \
            Sub64 alngWord(0 + lngStep), _
                  alngWord(1 + lngStep), _
                  (alngWord(14 + lngStep) Xor K_A5), _
                  (alngWord(15 + lngStep) Xor K_A5)

            ' xor64 x10, x11, x00, x01); \
            Xor64 alngWord(2 + lngStep), _
                  alngWord(3 + lngStep), _
                  alngWord(0 + lngStep), _
                  alngWord(1 + lngStep)

            ' add64 x20, x21, x10, x11); \
            Add64 alngWord(4 + lngStep), _
                  alngWord(5 + lngStep), _
                  alngWord(2 + lngStep), _
                  alngWord(3 + lngStep)

            ' sub64 x30, x31, x20^((~x10)<<19), ~x21^(((x11)<<19)|((x10)>>13))); \
            Sub64 alngWord(6 + lngStep), _
                  alngWord(7 + lngStep), _
                  (alngWord(4 + lngStep) Xor (.w32Shift((Not alngWord(2 + lngStep)), 19))), _
                  ((Not alngWord(5 + lngStep)) Xor (.w32Shift(alngWord(3 + lngStep), 19) Or .w32Shift(alngWord(2 + lngStep), -13)))

            ' xor64 x40, x41, x30, x31); \
            Xor64 alngWord(8 + lngStep), _
                  alngWord(9 + lngStep), _
                  alngWord(6 + lngStep), _
                  alngWord(7 + lngStep)

            ' add64 x50, x51, x40, x41); \
            Add64 alngWord(10 + lngStep), _
                  alngWord(11 + lngStep), _
                  alngWord(8 + lngStep), _
                  alngWord(9 + lngStep)

            ' sub64 x60, x61, ~x50^(((x40)>>23)|((x41)<<9)), x51^((~x41)>>23)); \
            Sub64 alngWord(12 + lngStep), _
                  alngWord(13 + lngStep), _
                  ((Not alngWord(10 + lngStep)) Xor (.w32Shift(alngWord(8 + lngStep), -23)) Or .w32Shift(alngWord(9 + lngStep), 9)), _
                  (alngWord(11 + lngStep)) Xor (.w32Shift((Not alngWord(9 + lngStep)), -23))

            ' xor64 x70, x71, x60, x61); \
            Xor64 alngWord(14 + lngStep), _
                  alngWord(15 + lngStep), _
                  alngWord(12 + lngStep), _
                  alngWord(13 + lngStep)

            ' add64 x00, x01, x70, x71); \
            Add64 alngWord(0 + lngStep), _
                  alngWord(1 + lngStep), _
                  alngWord(14 + lngStep), _
                  alngWord(15 + lngStep)

            ' sub64 x10, x11, x00^((~x70)<<19), ~x01^(((x71)<<19)|((x70)>>13))); \
            Sub64 alngWord(2 + lngStep), _
                  alngWord(3 + lngStep), _
                  (alngWord(0 + lngStep)) Xor (.w32Shift((Not alngWord(14 + lngStep)), 19)), _
                  ((Not alngWord(1 + lngStep)) Xor (.w32Shift(alngWord(15 + lngStep), 19)) Or .w32Shift(alngWord(14 + lngStep), -13))

            ' xor64 x20, x21, x10, x11); \
            Xor64 alngWord(4 + lngStep), _
                  alngWord(5 + lngStep), _
                  alngWord(2 + lngStep), _
                  alngWord(3 + lngStep)

            ' 23-Jun-2011 Corrected index values
            ' add64 x30, x31, x20, x21); \
            Add64 alngWord(6 + lngStep), _
                  alngWord(7 + lngStep), _
                  alngWord(4 + lngStep), _
                  alngWord(5 + lngStep)

            ' 07-May-2011 Corrected one of the shift values
            ' sub64 x40, x41, ~x30^(((x20)>>23)|((x21)<<9)), x31^((~x21)>>23)); \
            Sub64 alngWord(8 + lngStep), _
                  alngWord(9 + lngStep), _
                  ((Not alngWord(6 + lngStep)) Xor (.w32Shift(alngWord(4 + lngStep), -23)) Or .w32Shift(alngWord(5 + lngStep), 9)), _
                  (alngWord(7 + lngStep)) Xor (.w32Shift((Not alngWord(5 + lngStep)), -23))

            ' xor64 x50, x51, x40, x41); \
            Xor64 alngWord(10 + lngStep), _
                  alngWord(11 + lngStep), _
                  alngWord(8 + lngStep), _
                  alngWord(9 + lngStep)

            ' add64 x60, x61, x50, x51); \
            Add64 alngWord(12 + lngStep), _
                  alngWord(13 + lngStep), _
                  alngWord(10 + lngStep), _
                  alngWord(11 + lngStep)

            ' sub64 x70, x71, x60^0x89ABCDEF, x61^0x01234567);
            Sub64 alngWord(14 + lngStep), _
                  alngWord(15 + lngStep), _
                  (alngWord(12 + lngStep) Xor K_89), _
                  (alngWord(13 + lngStep) Xor K_12)

            ' Increment index pointer to allow
            ' processing of all word array elements
            lngStep = lngStep + 16

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For   ' exit For..Next loop
            End If

        Next lngLoop
    End With

End Sub

' #define mul5(s0, s1) \
Private Sub Mul5(ByRef s0 As Long, _
                 ByRef s1 As Long)

    Dim lngTemp0 As Long
    Dim lngTemp1 As Long

    ' tempt0 = s0<<2; \
    ' tempt1 = (s1<<2)|(s0>>30); \
    With mobjMath32
        lngTemp0 = .w32Shift(s0, 2)
        lngTemp1 = (.w32Shift(s1, 2)) Or (.w32Shift(s0, -30))
    End With

    ' add64(s0, s1, tempt0, tempt1);
    Add64 s0, s1, lngTemp0, lngTemp1

End Sub

' #define mul7(s0, s1) \
Private Sub Mul7(ByRef s0 As Long, _
                 ByRef s1 As Long)

    Dim lngTemp0 As Long
    Dim lngTemp1 As Long

    ' tempt0 = s0<<3; \
    ' tempt1 = (s1<<3)|(s0>>29); \
    With mobjMath32
        lngTemp0 = .w32Shift(s0, 3)
        lngTemp1 = (.w32Shift(s1, 3)) Or (.w32Shift(s0, -29))
    End With

    ' sub64(tempt0, tempt1, s0, s1); \
    Sub64 lngTemp0, lngTemp1, s0, s1
    s0 = lngTemp0  ' s0 = tempt0; \
    s1 = lngTemp1  ' s1 = tempt1;

End Sub

' #define mul9(s0, s1) \
Private Sub Mul9(ByRef s0 As Long, _
                 ByRef s1 As Long)

    Dim lngTemp0 As Long
    Dim lngTemp1 As Long

    ' tempt0 = s0<<3; \
    ' tempt1 = (s1<<3)|(s0>>29); \
    With mobjMath32
        lngTemp0 = .w32Shift(s0, 3)
        lngTemp1 = (.w32Shift(s1, 3)) Or (.w32Shift(s0, -29))
    End With

    ' add64(s0, s1, tempt0, tempt1);
    Add64 s0, s1, lngTemp0, lngTemp1

End Sub

' #define add64(s0, s1, p0, p1) \
' s0,s1 values are manipulated and returned
Private Sub Add64(ByRef s0 As Long, _
                  ByRef s1 As Long, _
                  ByVal p0 As Long, _
                  ByVal p1 As Long)

    Dim lngTemp  As Long
    Dim lngCarry As Long

    With mobjMath32
        lngTemp = p0                    ' temps0 = (p0); \
        s0 = .UnsignedAdd(s0, lngTemp)  ' s0 += temps0; \

        ' tcarry = s0 < temps0; \
        If s0 < lngTemp Then
            lngCarry = s0
        Else
            lngCarry = lngTemp
        End If

        ' s1+=(p1) + tcarry;
        s1 = .UnsignedAdd(s1, .UnsignedAdd(p1, lngCarry))
    End With

End Sub

' #define sub64(s0, s1, p0, p1) \
' s0,s1 values are manipulated and returned
Private Sub Sub64(ByRef s0 As Long, _
                  ByRef s1 As Long, _
                  ByVal p0 As Long, _
                  ByVal p1 As Long)

    Dim lngTemp  As Long
    Dim lngCarry As Long

    lngTemp = p0  ' temps0 = (p0); \

    ' tcarry = s0 < temps0; \
    If s0 < lngTemp Then
        lngCarry = s0
    Else
        lngCarry = lngTemp
    End If

    ' s0 -= temps0; \
    ' s1 -= (p1) + tcarry;
    With mobjMath32
        s0 = .UnsignedSubtract(s0, lngTemp)
        s1 = .UnsignedSubtract(s1, (.UnsignedAdd(p1, lngCarry)))
    End With

End Sub

' #define xor64(s0, s1, p0, p1) \
' s0,s1 values are manipulated and returned
Private Sub Xor64(ByRef s0 As Long, _
                  ByRef s1 As Long, _
                  ByVal p0 As Long, _
                  ByVal p1 As Long)

    s0 = s0 Xor p0  ' s0 ^= (p0); \
    s1 = s1 Xor p1  ' s1 ^= (p1);

End Sub

' **************************************************************************
' Routine:       GetBlockSize
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
' 24-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine 
' ***************************************************************************
Private Function GetBlockSize(ByVal curAmtLeft As Currency) As Long

    ' Determine amount of data to process
    Select Case curAmtLeft
           Case Is >= KB_32: GetBlockSize = KB_32
           Case Else:        GetBlockSize = CLng(curAmtLeft)
    End Select

End Function

' ***************************************************************************
' Routine:       Initialize
'
' Description:   Preload the Work array and Constants array for 32-bit 
'                processing with predetermined values to be used in the 
'                hashing calculations. 
' 
'                Got this explanation from Ask Dr. Math web site on how to
'                calculate The hex representation of the fractional parts of
'                a whole number.
'                http://mathforum.org/dr.math/ 
' 
'                There is a simple process for converting a base 10 decimal
'                fraction to a "decimal" fraction in another base.  In this
'                case, hex (16) will be our new base.  We repeatedly multiply
'                our decimal fraction by the new base, picking off the whole
'                number part each time as the next digit of the final output. 
'
' ---------------------------------------------------------------------------- 
'            My idea for data to fill the work arrays is to calculate the
'            hex representation of the fractional part of the CUBE ROOTS
'            of the 512th thru 589th prime numbers.  Tiger-128 and Tiger-160
'            will use the last 32-bits of the hex output for 512th through
'            523rd primes.  Tiger-192 will retain its original values.
'            Tiger-224 through Tiger-512 will use last 32-bits of the hex
'            output for 530th through 589th primes.
'
' Example:   512th Prime number = 3671
'            CUBE ROOT of 3671 = 15.426288936629509912318786195
'
'            0.426288936629509912318786195 * 16 =  6.82062298607215859710057912
'            0.82062298607215859710057912  * 16 = 13.12996777715453755360926592
'            0.12996777715453755360926592  * 16 =  2.07948443447260085774825472
'            0.07948443447260085774825472  * 16 =  1.27175095156161372397207552
'            0.27175095156161372397207552  * 16 =  4.34801522498581958355320832
'            0.34801522498581958355320832  * 16 =  5.56824359977311333685133312
'            0.56824359977311333685133312  * 16 =  9.09189759636981338962132992
'            0.09189759636981338962132992  * 16 =  1.47036154191701423394127872
'            0.47036154191701423394127872  * 16 =  7.52578467067222774306045952
'            0.52578467067222774306045952  * 16 =  8.41255473075564388896735232
'            0.41255473075564388896735232  * 16 =  6.60087569209030222347763712
'            0.60087569209030222347763712  * 16 =  9.61401107344483557564219392
'            0.61401107344483557564219392  * 16 =  9.82417717511736921027510272
'            0.82417717511736921027510272  * 16 = 13.18683480187790736440164352
'            0.18683480187790736440164352  * 16 =  2.98935683004651783042629632
'            0.98935683004651783042629632  * 16 = 15.82970928074428528682074112
'                                                  |
'              Whole numbers are converted to hex__|
'
'            The hex representation of the fractional parts
'            of the CUBE ROOTS of 3671:  6d21459178699d2f
'
'            Ex:  malngHash(0) = 78699d2f   <- last 32-bits (8 characters)
' 
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAY-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 10-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Combined all Tiger hashes into a single module.
'              - Created unique initialization values for each of the Tiger
'                hashes (128, 160, 224, 256, 384 512).  Tiger-192 bit
'                initialization data is original data.
'              - Kept original data for the constant arrays.
' ***************************************************************************
Private Sub Initialize()

    Dim lngIdx   As Long
    Dim lngIndex As Long
    Dim strData  As String
    Dim avntData As Variant

    ReDim T1(0 To 1, MAX_BYTE)  ' Size constant arrays
    ReDim T2(0 To 1, MAX_BYTE)
    ReDim T3(0 To 1, MAX_BYTE)
    ReDim T4(0 To 1, MAX_BYTE)

    Select Case mlngHashMethod
           Case eTIGER128
                ' These words were obtained by taking the last thirty-two
                ' bits (8 chars) of the fractional parts of the CUBE ROOTS
                ' of the 512th through 517th prime numbers.
                ' See flowerbox above for more details.
                ReDim malngHash(6)
                malngHash(0) = &H78699D2F
                malngHash(1) = &HAFBA5E23
                malngHash(2) = &H6950D47A
                malngHash(3) = &H935121DF
                malngHash(4) = &H4A5FCDEB
                malngHash(5) = &H2D4A6873

           Case eTIGER160
                ' These words were obtained by taking the last thirty-two
                ' bits (8 chars) of the fractional parts of the CUBE ROOTS
                ' of the 518th through 523rd prime numbers.
                ' See flowerbox above for more details.
                ReDim malngHash(6)
                malngHash(0) = &H45DE1EEA
                malngHash(1) = &HF5358C61
                malngHash(2) = &H9DA35205
                malngHash(3) = &H5A116F23
                malngHash(4) = &H12182662
                malngHash(5) = &HB716845C

           Case eTIGER192    ' Original data
                ReDim malngHash(6)
                malngHash(0) = &H89ABCDEF
                malngHash(1) = &H1234567
                malngHash(2) = &H76543210
                malngHash(3) = &HFEDCBA98
                malngHash(4) = &HC3B2E187
                malngHash(5) = &HF096A5B4

           Case eTIGER224
                ' These words were obtained by taking the last thirty-two
                ' bits (8 chars) of the fractional parts of the CUBE ROOTS
                ' of the 530th through 541st prime numbers.
                ' See flowerbox above for more details.
                ReDim malngHash(12)
                malngHash(0) = &HB54501AE
                malngHash(1) = &HC2B3B044
                malngHash(2) = &H6116716D
                malngHash(3) = &H99B2C5AC
                malngHash(4) = &HCBB5EC26
                malngHash(5) = &H88E13CBC
                malngHash(6) = &H1B77BA24
                malngHash(7) = &H415A1AD3
                malngHash(8) = &HE33A5BA3
                malngHash(9) = &HCA8AABCE
                malngHash(10) = &H3D5A464D
                malngHash(11) = &H2E64F2E

           Case eTIGER256
                ' These words were obtained by taking the last thirty-two
                ' bits (8 chars) of the fractional parts of the CUBE ROOTS
                ' of the 542nd through 553rd prime numbers.
                ' See flowerbox above for more details.
                ReDim malngHash(12)
                malngHash(0) = &H4DE7C5AF
                malngHash(1) = &H27B0150A
                malngHash(2) = &H1BCCD63
                malngHash(3) = &HF24C9F7
                malngHash(4) = &H6CB4E92C
                malngHash(5) = &H22E67181
                malngHash(6) = &H3DAF9FC2
                malngHash(7) = &H3BB50EC3
                malngHash(8) = &H9533E090
                malngHash(9) = &HAB3FA038
                malngHash(10) = &H7F41DB2
                malngHash(11) = &H856E3A8

           Case eTIGER384
                ' These words were obtained by taking the last thirty-two
                ' bits (8 chars) of the fractional parts of the CUBE ROOTS
                ' of the 554th through 571st prime numbers.
                ' See flowerbox above for more details.
                ReDim malngHash(18)
                malngHash(0) = &H629EC872
                malngHash(1) = &H6270F054
                malngHash(2) = &H1C9392CE
                malngHash(3) = &H14ADCCA1
                malngHash(4) = &H9E9F11E1
                malngHash(5) = &H776B865A
                malngHash(6) = &HED449FEC
                malngHash(7) = &HAED905FA
                malngHash(8) = &H1BE27B8C
                malngHash(9) = &HD54DFC34
                malngHash(10) = &H551FF8A9
                malngHash(11) = &HAA755CA9
                malngHash(12) = &HD3794395
                malngHash(13) = &H47600C5A
                malngHash(14) = &H7D08B7E2
                malngHash(15) = &H3EB1B6CE
                malngHash(16) = &HA4675BCD
                malngHash(17) = &HAE901657

           Case eTIGER512
                ' These words were obtained by taking the last thirty-two
                ' bits (8 chars) of the fractional parts of the CUBE ROOTS
                ' of the 572nd through 589th prime numbers.
                ' See flowerbox above for more details.
                ReDim malngHash(18)
                malngHash(0) = &H748B3C4B
                malngHash(1) = &H7EBA1D39
                malngHash(2) = &HB06DDF92
                malngHash(3) = &HA168C4C1
                malngHash(4) = &H40BEF5FD
                malngHash(5) = &H2171CD6E
                malngHash(6) = &H35918AE4
                malngHash(7) = &HB72ED706
                malngHash(8) = &H8BD85572
                malngHash(9) = &HE7A9DF4D
                malngHash(10) = &H6167D94D
                malngHash(11) = &H35ADD8CA
                malngHash(12) = &HCC396EFB
                malngHash(13) = &H2E113BF0
                malngHash(14) = &HFE386A17
                malngHash(15) = &H8B9CBCDF
                malngHash(16) = &HEC9E9646
                malngHash(17) = &H2D545F3A
    End Select

    strData = vbNullString  ' Verify string variable is empty
    avntData = Empty        ' Verify variant variable is empty
    lngIdx = 0              ' Set array index to zero

    ' Original data for constant arrays
    ' Load T1 table array
    strData = strData & "F7E90C5E 02AAB17C E243A8EC AC424B03 0DD5FCD3 72CD5BE3 F6F97F3A 6D019B93 "
    strData = strData & "D21F9193 CD9978FF 708029E2 7573A1C9 922A83C3 B164326B 04915870 46883EEE "
    strData = strData & "7103ECE6 EAACE305 08A3535C C54169B8 8DDEC47C 4CE75491 DC0DF40C 0AA2F4DF "
    strData = strData & "A74DBEFA 10B76F18 5AD1AB6A C6CCB623 572FE2FF 13726121 199D921E 1A488C6F "
    strData = strData & "DA0007CA 4BC9F9F4 E85241C7 26F5E6F6 EA5947B6 859079DB C99E8C92 4F1885C5 "
    strData = strData & "A96F864B D78E761E 52B5C17D 8E36428C 373063C1 69CF6827 9BB4C56E B607C93D "
    strData = strData & "0E76B5EA 7D820E76 F07FDC42 645C9CC6 243342E0 BF38A078 9D2E7D04 5F6B343C "
    strData = strData & "600B0EC6 F2C28AEB 7254BCAC 6C0ED85F A4DB4FE5 71592281 CE0FED9F 1967FA69 "
    strData = strData & "B96545DB FD5293F8 F2A7600B C879E9D7 0193194E 86024892 2D9CC0B3 A4F9533B "
    strData = strData & "15957613 9053836C FC357BF1 DB6DCF8A 7A370F57 18BEEA7A 50B99066 037117CA "
    strData = strData & "74424A35 6AB30A97 E325249B F4E92F02 061CCAE1 7739DB07 ECA42A05 D8F3B49C "
    strData = strData & "51382F73 BD56BE3F 43B0BB28 45FAED58 11BF1F83 1C813D5C D75FA169 8AF0E4B6 "
    strData = strData & "87AD9999 33EE18A4 B1C94410 3C26E8EA C0A822F9 B510102B 0CE6123B 141EEF31 "
    strData = strData & "59DDB154 FC65B900 C5E0E607 E0158640 26C3A3CF 884E0798 23C535FD 930D0D95 "
    strData = strData & "4E9A2B00 35638D75 40469DD5 4085FCCF 8BE23A4C C4B17AD2 6A3E6A2E CAB2F0FC "
    strData = strData & "6B943FCD 2860971A 12E30446 3DDE6EE2 E01765AE 6222F32A 478308FE 5D550BB5 "
    strData = strData & "A0EDA22A A9EFA98D 86C40DA7 C351A716 9C867C84 1105586D FDA22853 DCFFEE85 "
    strData = strData & "2C5EEF76 CCFBD026 8990D201 BAF294CB 2AFAD975 E69464F5 DF133E14 94B013AF "
    strData = strData & "2823C958 06A7D1A3 30F61119 6F95FE51 462C06C0 D92AB34E 887C71D2 ED7BDE33 "
    strData = strData & "6518393E 79746D6E 5D713329 5BA41938 48A97564 7C1BA6B9 7BFDAC67 31987C19 "
    strData = strData & "4B053D02 DE6C23C4 D002D64D 581C49FE 38261571 DD474D63 E473D062 AA4546C3 "
    strData = strData & "9455F860 928FCE34 CAAB94D9 48161BBA 770E6F68 63912430 02C6641C 6EC8A5E6 "
    strData = strData & "337DDD2B 87282515 034B701B 2CDA6B42 81CB096D B03D37C1 66C71C6F E1084382 "
    strData = strData & "EB51B255 2B3180C7 96C08BBC DF92B82F A632F3BA 5C68C8C0 1C3D0556 5504CC86 "
    strData = strData & "5FB26B8F ABBFA4E5 B3BACEB4 41848B0A AA445D32 B334A273 A85AD881 BCA696F0 "
    strData = strData & "B528D56C 24F6EC65 90F4524A 0CE1512E 5506D35A 4E9DD79D C6CE9779 258905FA "
    strData = strData & "3E109B33 2019295B 73A054CC F8A9478B 34417EB0 2924F2F9 536D1BC4 3993357D "
    strData = strData & "1DB6FF8B 38A81AC2 7D6016BF 47C4FBF1 7667E3F5 1E0FAADD 938BEB96 7ABCFF62 "
    strData = strData & "8FC179C9 A78DAD94 2911E50D 8F1F98B7 27121A91 61E48EAE 31859808 4D62F7AD "
    strData = strData & "EF5CEAEB ECEBA345 BC9684CE F5CEB25E B7F76221 F633E20C AB8293E4 A32CDF06 "
    strData = strData & "A5EE2CA4 985A202C CC8A8FB1 CF0B8447 979859A3 9F765244 A1240017 A8D516B1 "
    strData = strData & "BB5DC726 0BD7BA3E B86ADB39 E54BCA55 6C478063 1D7A3AFD E7669EDD 519EC608 "
    strData = strData & "D149AA23 0E5715A2 848FF194 177D4571 41014C22 EEB55F32 3A6E2EC2 0F5E5CA1 "
    strData = strData & "75F5C361 8029927B C3D6E436 AD139FAB 4CCF402F 0D5DF1A9 BEA5DFC8 3E8BD948 "
    strData = strData & "BD3FF77E A5A0D357 1F74F645 A2D12E25 5E81A082 66FD9E52 7F687A49 2E0C90CE "
    strData = strData & "BA973BC5 C2E8BCBE E509745F 000001BC E6DAB3D6 423777BB AEF06EB5 D1661C7E "
    strData = strData & "4DAACFD8 A1781F35 2B16AFFC 2D11284A FA891D1F F1FC4F67 CB920ADA 73ECC25D "
    strData = strData & "C2A12651 AE610C22 D356B78A 96E0A810 2FE7870F 5A9A381F E94E5530 D5AD62ED "
    strData = strData & "368D1427 D225E5E8 C7AF4631 65977B70 DE39D74F 99F889B2 54E1D143 233F30BF "
    strData = strData & "D9A63C97 9A9675D3 F334F9A8 5470554F 4A4F5688 166ACB74 B2E4AEAD 70C74CAA "
    strData = strData & "6F294D12 F0D09164 684031D1 57B82A89 61BE0B6B EFD95A5A 69F2F29A 2FBD12E9 "
    strData = strData & "FEFF9FE8 9BD37013 D6085A06 3F9B0404 166CFE15 4940C1F3 CDF3DEFB 09542C4D "
    strData = strData & "85CD5CE3 B4C52183 4462A641 C935B7DC 8ED3B63F 3417F8A6 5B215B40 B8095929 "
    strData = strData & "3B8C8572 F99CDAEF F8FCB95D 018C0614 1A3ACDF3 1B14ACCD 00BB732D 84D471F2 "
    strData = strData & "95E8DA16 C1A3110E BF1A82B8 430A7220 39DF210E B77E090D 3CD05E9D 5EF4BD9F "
    strData = strData & "7E57A444 9D4FF6DA 83D4A5F8 DA1D60E1 17998E47 B287C384 1BB31886 FE3EDC12 "
    strData = strData & "980CCBEF C7FE3CCC 189BFD03 E46FB590 9A4C57DC 3732FD46 7CF1AD65 7EF700A0 "
    strData = strData & "A31D8859 59C64468 D45B61F6 762FB0B4 99047718 155BAED0 3D50BAA6 68755E4C "
    strData = strData & "22D8B4DF E9214E7F 2EAC95F4 2ADDBF53 B4BD0109 32AE3909 B08E3450 834DF537 "
    strData = strData & "4220728D FA209DA8 9EFE23F7 9E691D9B C4AE8D7F 0446D288 E169785B 7B4CC524 "
    strData = strData & "35CA1385 21D87F01 137B8AA5 CEBB400F 580796BE 272E2B66 25C2B0DE 36122641 "
    strData = strData & "AD1EFBB2 057702BD ACF84BE9 D4BABB8E 641BC67B 91583139 8036E024 8BDC2DE0 "
    strData = strData & "F49F68ED 603C8156 DBEF5111 F7D236F7 8AD21E80 9727C459 670A5FD7 A08A0896 "
    strData = strData & "09EBA9CB CB4A8F43 0F7036A1 81AF564B 78199ABD C0B99AA7 3FC8E952 959F1EC8 "
    strData = strData & "794A81B9 8C505077 056338F0 3ACAAF8F 627A6778 07B43F50 F5ECCC77 4A44AB49 "
    strData = strData & "B679EE98 3BC3D6E4 CF14108C 9CC0D4D1 206BC8A0 4406C00B C8D72D89 82A18854 "
    strData = strData & "5C3C432C 67E366B3 102B37F2 B923DD61 D884271D 56AB2779 FF1525AF BE83E1B0 "
    strData = strData & "217E49A9 FB7C65D4 6D48E7D4 6BDBE0E7 45D9179E 08DF8287 DD53BD34 22EA6A9A "
    strData = strData & "5622200A E36E141C 8CB750EE 7F805D1B 9F58E837 AFE5C7A5 4FB1C23C E27F996A "
    strData = strData & "0775F0D0 D3867DFB 6E88891A D0E673DE AFB86C25 123AEB9E C145B895 30F1D5D5 "
    strData = strData & "EE7269E7 BB434A2D F931FA38 78CB67EC 323BBF9C F33B0372 FB279C74 52D66336 "
    strData = strData & "0AFB4EAA 505F33AC A2CCE187 E8A5CD99 1E2D30BB 53497480 D5876D90 8D2D5711 "
    strData = strData & "91BC038E 1F1A4128 82E56648 D6E2E71D 497732B7 74036C3A 6361F5AB 89B67ED9 "
    strData = strData & "F1EA02A2 FFED95D8 1464D43D E72B3BD6 0BDC4820 A6300F17 ED78A77A EBC18760 "

    avntData = Split(strData, Chr$(32))   ' Load variant array

    ' load T1 table array
    For lngIndex = 0 To 511 Step 2

        ' Convert variant data to Long Integer
        T1(0, lngIdx) = Val("&H" & avntData(lngIndex))
        T1(1, lngIdx) = Val("&H" & avntData(lngIndex + 1))
        lngIdx = lngIdx + 1

    Next lngIndex

    strData = vbNullString       ' Verify string is empty
    avntData = Empty   ' Verify variant is empty
    lngIdx = 0

    ' Load T2 table array
    strData = strData & "05A12138 E6A6BE5A B4F87C98 B5A122A5 140B6990 563C6089 391F5DD5 4C46CB2E "
    strData = strData & "C9B79434 D932ADDB 2015AFF5 08EA70E4 3E478CF1 D765A667 AB278D99 C4FB757E "
    strData = strData & "2D6E0692 DF11C686 0D7F3B16 DDEB84F1 A665EA04 6F2EF604 F0E0DFB3 4A8E0F0F "
    strData = strData & "3DBCBA51 A5EDEEF8 0EA4371E FC4F0A2A 5CB38429 E83E1DA8 BA1B1CE2 DC8FF882 "
    strData = strData & "8353E80D CD45505E D4DB0717 18D19A00 A5F38101 34A0CFED 8887CAF2 0BE77E51 "
    strData = strData & "B3C45136 1E341438 9089CCF9 E05797F4 F2591D14 FFD23F9D 8595C5CD 543DDA22 "
    strData = strData & "99052A33 661F81FD DB0F7B76 8736E641 418E5307 15227725 162EB2FA E25F7F46 "
    strData = strData & "6C13D9FE 48A8B212 92E76EEA AFDC5417 C6D1898F 03D912BF 1B83F51B 31B1AAFA "
    strData = strData & "E42AB7D9 F1AC2796 FCD2EBAC 40A3A7D7 0AFBBCC5 1056136D 9A6D0C85 7889E1DD "
    strData = strData & "2A7974AA D3352578 078AC09B A7E25D09 EAC6EDD0 BD4138B3 71EB9E70 920ABFBE "
    strData = strData & "4FC2625C A2A5D0F5 0B1290A3 C054E36B 62FE932B F6DD59FF 11A8AC7D 35373545 "
    strData = strData & "72FADCD4 CA845E91 329D20DC 84F82B60 CD672F18 79C62CE1 D124642C 8B09A2AD "
    strData = strData & "19D9E726 D0C1E96A 4BA9500C 5A786A9B 634C43F3 0E020336 EB66D822 C17B474A "
    strData = strData & "EC9BAAC2 6A731AE3 E0840258 8226667A 91CAECA5 67D45676 4875ADB5 1D94155C "
    strData = strData & "5B813FDF 6D00FD98 B774CD06 51286EFC 1FA744AF 5E883447 E761AE2E F72CA0AE "
    strData = strData & "AEE8E09A BE40E4CD 5118F665 E9970BBB 33DF1964 726E4BEB 29199762 703B0007 "
    strData = strData & "F5EF30A7 4631D816 1504A6BE B880B5B5 7ED84B6C 641793C3 F6E97D96 7B21ED77 "
    strData = strData & "2EF96B73 77630631 E86FF3F4 AE528948 86A3F8F8 53DBD7F2 4CFC1063 16CADCE7 "
    strData = strData & "FA52C6DD 005C19BD 64D46AD3 68868F5D CF1E186A 3A9D512C 385660AE 367E62C2 "
    strData = strData & "77DCB1D7 E359E7EA 749ABE6E 526C0773 D09F734B 735AE5F9 8A558BA8 493FC7CC "
    strData = strData & "3041AB45 B0B9C153 470A59BD 321958BA 5F46C393 852DB00B D336B0E5 91209B2B "
    strData = strData & "659EF19F 6E604F7D 782CCB24 B99A8AE2 C814C4C7 CCF52AB6 BE11727B 4727D9AF "
    strData = strData & "0121B34D 7E950D0C 70AD471F 756F4356 615A6849 F5ADD442 80B9957A 4E87E099 "
    strData = strData & "50AEE355 2ACFA1DF FD2FD556 D898263A D80C8FD6 C8F4924D 754A173A CF99CA3D "
    strData = strData & "AF91BF3C FE477BAC D690C12D ED5371F6 5E687094 831A5C28 3708A0A4 C5D3C90A "
    strData = strData & "17D06580 0F7F9037 B8FDF27F 19F9BB13 4D502843 B1BD6F1B 8FFF4012 1C761BA3 "
    strData = strData & "E2E21F3B 0D1530C4 A7372C8A 8943CE69 FEB5CE66 E5184E11 BD736621 618BDB80 "
    strData = strData & "8B574D0B 7D29BAD6 25E6FE5B 81BB613E BC07913F 071C9C10 09AC2D97 C7BEEB79 "
    strData = strData & "3BC5D757 C3E58D35 F38F61E8 EB017892 9B1CC21A D4EFFB9C F494F7AB 99727D26 "
    strData = strData & "956B3E03 A3E063A2 4AA09C30 9D4A8B9A 00090FB4 3F6AB7D5 57268AC0 9CC0F2A0 "
    strData = strData & "EDBF42D1 3DEE9D2D 7960A972 330F49C8 87421B41 C6B27202 7C00369C 0AC59EC0 "
    strData = strData & "CB353425 EF4EAC49 EF0129D8 F450244E CAF4DEB6 8ACC46E5 989263F7 2FFEAB63 "
    strData = strData & "5D7A4578 8F7CB9FE 4E634635 5BD8F764 BF2DC900 427A7315 2125261C 17D0C4AA "
    strData = strData & "93518E50 3992486C A2D7D4C3 B4CBFEE0 2C5DDD8D 7C75D620 E35B6C61 DBC295D8 "
    strData = strData & "02032B19 60B369D3 DCE44132 CE42685F DDF65610 06F3DDB9 B5E148F0 8EA4D21D "
    strData = strData & "2FCD496F 20B0FCE6 58B0EE31 2C1B9123 18F5A308 B28317B8 9CA6D2CF A89C1E18 "
    strData = strData & "6AAADBC8 0C6B1857 1299FAE3 B65DEAA9 7F1027E7 FB2B794B 443B5BEB 04E4317F "
    strData = strData & "5939D0A6 4B852D32 FB207FFC D5AE6BEE 81C7D374 309682B2 94C3B475 BAE309A1 "
    strData = strData & "13B49F05 8CC3F97B F8293967 98A9422F 1076FF7C 244B16B0 663D67EE F8BF571C "
    strData = strData & "EEE30DA1 1F0D6758 7ADEB9B7 C9B611D9 7B6C57A2 B7AFD588 6B984FE1 6290AE84 "
    strData = strData & "ACC1A5FD 94DF4CDE C5483AFF 058A5BD1 42BA3C37 63166CC1 B2F76F40 8DB8526E "
    strData = strData & "6F0D6D4E E1088003 971D311D 9E0523C9 CC7CD691 45EC2824 E62382C9 575B8359 "
    strData = strData & "C4889995 FA9E400D 45721568 D1823ECB 8206082F DAFD983B 2386A8CB AA7D2908 "
    strData = strData & "03B87588 269FCD44 28BDD1E0 1B91F5F7 040201F6 E4669F39 8CF04ADE 7A1D7C21 "
    strData = strData & "D79CE5CE 65623C29 96C00BB1 23684490 9DA503BA AB9BF187 A458058E BC23ECB1 "
    strData = strData & "BB401ECC 9A58DF01 A85F143D A070E868 7DF2239E 4FF18830 1A641183 14D565B4 "
    strData = strData & "52701602 EE133374 3F285E09 950E3DCF B9C80953 59930254 8930DA6D 3BF29940 "
    strData = strData & "53691387 A955943F A9CB8784 A15EDECA 352BE9A0 29142127 FF4E7AFB 76F0371F "
    strData = strData & "274F2228 0239F450 1D5E868B BB073AF0 C10E96C1 BFC80571 68222E23 D2670885 "
    strData = strData & "8E80B5B0 9671A3D4 E193BB81 55B5D38A A18B04B8 693AE2D0 ADD5335F 5C48B4EC "
    strData = strData & "4916A1CA FD743B19 34BE98C4 25770181 3C54A4AD E77987E8 DA33E1B9 28E11014 "
    strData = strData & "226AA213 270CC59E 6D1A5F60 71495F75 60AFEF77 9BE853FB F7443DBF ADC786A7 "
    strData = strData & "73B29A82 09044561 C232BD5E 58BC7A66 673AC8B2 F306558C B6C9772A 41F639C6 "
    strData = strData & "9FDA35DA 216DEFE9 1C7BE615 11640CC7 565C5527 93C43694 46777839 EA038E62 "
    strData = strData & "5A3E2469 F9ABF3CE 0FD312D2 741E768D CED652C6 0144B883 A33F8552 C20B5A5B "
    strData = strData & "C3435A9D 1AE69633 088CFDEC 97A28CA4 1E96F420 8824A43C 6EEEA746 37612FA6 "
    strData = strData & "F9CF0E5A 6B4CB165 A0ABFB4A 43AA1C06 F162796B 7F4DC26F 54ED9B0F 6CBACC8E "
    strData = strData & "D2BB253E A6B7FFEF B0A29D4F 2E25BC95 DEF1388C 86D6A58B 76B6F054 DED74AC5 "
    strData = strData & "2B45805D 8030BDBC E94D9289 3C81AF70 9E3100DB 3EFF6DDA DFCC8847 B38DC39F "
    strData = strData & "8D17B87E 12388552 40B1B642 F2DA0ED2 D54BF9A9 44CEFADC 433C7EE6 1312200E "
    strData = strData & "3A78C748 9FFCC84F 248576BB F0CD1F72 3638CFE4 EC697405 0CEC4E4C 2BA7B67C "
    strData = strData & "E5CE32ED AC2F4DF3 26EA4C11 CB33D143 C77E58BC A4E9044C D934FCEF 5F513293 "
    strData = strData & "06E55444 5DC96455 317DE40A 50DE418F 69DDE259 388CB31A 55820A86 2DB4A834 "
    strData = strData & "84711AE9 9010A91E B1498371 4DF7F0B7 C0977179 D62A2EAB AA8D5C0E 22FAC097 "

    avntData = Split(strData, Chr$(32))   ' Load variant array

    ' load T2 table array
    For lngIndex = 0 To 511 Step 2

        ' Convert variant data to Long Integer
        T2(0, lngIdx) = Val("&H" & avntData(lngIndex))
        T2(1, lngIdx) = Val("&H" & avntData(lngIndex + 1))
        lngIdx = lngIdx + 1

    Next lngIndex

    strData = vbNullString       ' Verify string is empty
    avntData = Empty   ' Verify variant is empty
    lngIdx = 0

    ' Load T3 table array
    strData = strData & "F1DAF39B F49FCC2F 6FF29281 487FD5C6 FCDCA83F E8A30667 D2FCCE63 2C9B4BE3 "
    strData = strData & "93FBBBC2 DA3FF74B FE70BA66 2FA165D2 970E93D4 A103E279 B0E45E71 BECDEC77 "
    strData = strData & "3985E497 CFB41E72 5EF75017 B70AAA02 3840B8E0 D42309F0 35898579 8EFC1AD0 "
    strData = strData & "E2B2ABC5 96C6920B 375A9172 66AF4163 CA7127FB 2174ABDC 4A72FF41 B33CCEA6 "
    strData = strData & "083066A5 F04A4933 D7289AF5 8D970ACD 31C8C25E 8F96E8E0 76875D47 F3FEC022 "
    strData = strData & "056190DD EC7BF310 BB0F1491 F5ADB0AE 0FD58892 9B50F885 58B74DE8 49754883 "
    strData = strData & "91531C61 A3354FF6 81D2C6EE 0702BBE4 7DEDED98 89FB2405 8596E902 AC307513 "
    strData = strData & "172772ED 1D2D3580 8E6BC30D EB738FC2 63044326 5854EF8F 5ADD3BBE 9E5C5232 "
    strData = strData & "325C4623 90AA53CF 349DD067 C1D24D51 A69EA624 2051CFEE 862E7E4F 13220F0A "
    strData = strData & "04E04864 CE393994 7086FCB7 D9C42CA4 8A03E7CC 685AD223 AB2FF1DB 066484B2 "
    strData = strData & "EFBF79EC FE9D5D70 9C481854 5B13B9DD ED1509AD 15F0D475 0EC79851 0BEBCD06 "
    strData = strData & "183AB7F8 D58C6791 52F3EEE4 D1187C50 E54E82FF C95D1192 B9AC6CA2 86EEA14C "
    strData = strData & "53677D5D 3485BEB1 1F8C492A DD191D78 A784EBF9 F60866BA A2D08C74 518F643B "
    strData = strData & "E1087C22 8852E956 C410AE8D A768CB8D BFEC8E1A 38047726 CD3B45AA A67738B4 "
    strData = strData & "EC0DDE19 AD16691C 80462E07 C6D43193 0BA61938 C5A5876D A58FD840 16B9FA1F "
    strData = strData & "3CA74F18 188AB117 C99C021F ABDA2F98 134AE816 3E0580AB 73645ABB 5F3B05B7 "
    strData = strData & "5575F2F6 2501A2BE 4E7E8BA9 1B2F7400 71E8D953 1CD75803 62764E30 7F6ED895 "
    strData = strData & "596F003D B15926FF A8C5D6B9 9F65293D D690F84C 6ECEF04D FF33AF88 4782275F "
    strData = strData & "3F820801 E4143308 9A1AF9B5 FD0DFE40 2CDB396B 4325A334 B301B252 8AE77E62 "
    strData = strData & "6655615A C36F9E9F 92D32C09 85455A2D 49477485 F2C7DEA9 33A39EBA 63CFB4C1 "
    strData = strData & "6EBC5462 83B040CC FDB326B0 3B9454C8 87FFD78C 56F56A9E 99F42BC6 2DC2940D "
    strData = strData & "6B096E2D 98F7DF09 3AD852BF 19A6E01E DBD4B40B 42A99CCB 45E9C559 A59998AF "
    strData = strData & "07D93186 366295E8 FAA1F773 6B48181B 157A0A1D 1FEC57E2 F6201AD5 4667446A "
    strData = strData & "CFB0F075 E615EBCA 68290778 B8F31F4F CE22D11E 22713ED6 2EC3C93B 3057C1A7 "
    strData = strData & "7C3F1F2F CB46ACC3 02AAF50E DBB893FD 600B9FCF 331FD92E 48EA3AD6 A498F961 "
    strData = strData & "8B6A83EA A8D8426E B7735CDC A089B274 1E524A11 87F6B373 CBC96749 118808E5 "
    strData = strData & "B19BD394 9906E4C7 9B24A20C AFED7F7E EB3644A7 6509EADE E8EF0EDE 6C1EF1D3 "
    strData = strData & "E9798FB4 B9C97D43 740C28A3 A2F2D784 6197566F 7B849647 B65F069D 7A5BE3E6 "
    strData = strData & "78BE6F10 F96330ED 7A076A15 EEE60DE7 A08B9BD0 2B4BEE4A C7B8894E 6A56A63E "
    strData = strData & "BA34FEF4 02121359 283703FC 4CBF99F8 0CAF30C8 39807135 F017687A D0A77A89 "
    strData = strData & "9E423569 F1C1A9EB 2DEE8199 8C797628 DD1F7ABD 5D1737A5 09A9FA80 4F53433C "
    strData = strData & "DF7CA1D9 FA8B0C53 886CCB77 3FD9DCBC A91B4720 C040917C F9D1DCDF 7DD00142 "
    strData = strData & "4F387B58 8476FC1D F3316503 23F8E7C5 E7E37339 032A2244 50F5A74B 5C87A5D7 "
    strData = strData & "3698992E 082B4CC4 B858F63C DF917BEC 5BF86DDA 3270B8FC 29B5DD76 10AE72BB "
    strData = strData & "7700362B 576AC94E C61EFB8F 1AD112DA C5FAA427 691BC30E CC327143 FF246311 "
    strData = strData & "30E53206 3142368E E02CA396 71380E31 0AAD76F1 958D5C96 C16DA536 F8D6F430 "
    strData = strData & "1BE7E1D2 C8FFD13F 004DDBE1 7578AE66 067BE646 05833F01 3BFE586D BB34B5AD "
    strData = strData & "A12B97F0 095F34C9 25D60CA8 247AB645 017477D1 DCDBC6F3 DECAD24D 4A2E14D4 "
    strData = strData & "BE0A1EEB BDB5E6D9 794301AB 2A7E70F7 270540FD DEF42D8A A34C22C1 01078EC0 "
    strData = strData & "F4C16387 E5DE511A BD9A330A 7EBB3A52 AA7D6435 77697857 03AE4C32 004E8316 "
    strData = strData & "AD78E312 E7A21020 6AB420F2 9D41A70C EA1141E6 28E06C18 984F6B28 D2B28CBD "
    strData = strData & "446E9D83 26B75F6C 4D418D7F BA47568C E6183D8E D80BADBF 5F166044 0E206D7F "
    strData = strData & "11CBCA3E E258A439 B21DC0BC 723A1746 F5D7CDD3 C7CAA854 3D261D9C 7CAC3288 "
    strData = strData & "23BA942C 7690C264 478042B8 17E55524 56A2389F E0BE4776 67AB2DA0 4D289B5E "
    strData = strData & "8FBBFD31 44862B9C 9D141365 B47CC804 2B91C793 822C1B36 FB13DFD8 4EB14655 "
    strData = strData & "14E2A97B 1ECBBA07 5CDE5F14 6143459D D5F0AC89 53A8FBF1 1C5E5B00 97EA04D8 "
    strData = strData & "D4FDB3F3 622181A8 572A1208 E9BCD341 43CCE58A 14112586 A4C6E0A4 9144C5FE "
    strData = strData & "65CF620F 0D33D065 9F219CA1 54A48D48 6D63C821 C43E5EAC 72770DAF A9728B3A "
    strData = strData & "20DF87EF D7934E7B 1A3E86E5 E35503B6 C819D504 CAE321FB AC60BFA6 129A50B3 "
    strData = strData & "7E9FB6C3 CD5E68EA 9483B1C7 B01C9019 C295376C 3DE93CD5 2AB9AD13 AED52EDF "
    strData = strData & "C0A07884 2E60F512 E36210C9 BC3D86A3 163951CE 35269D9B D0CDB5FA 0C7D6E2A "
    strData = strData & "D87F5733 59E86297 898DB0E7 298EF221 D1A5AA7E 55000029 B5061B45 8BC08AE1 "
    strData = strData & "6C92703A C2C31C2B AF25EF42 94CC596B 22540456 0A1D73DB D9C4179A 04B6A0F9 "
    strData = strData & "AE3D3C60 EFFDAFA2 B49496C4 F7C8075B 1D1CD4E3 9CC5C714 218E5534 78BD1638 "
    strData = strData & "F850246A B2F11568 9502BC29 EDFABCFA DA23051B 796CE5F2 DC93537C AAE128B0 "
    strData = strData & "EE4B29AE 3A493DA0 416895D7 B5DF6B2C 122D7F37 FCABBD25 105DC4B1 70810B58 "
    strData = strData & "F7882A90 E10FDD37 518A3F5C 524DCAB5 8451255B 3C9E8587 19BD34E2 40298281 "
    strData = strData & "5D3CECCB 74A05B6F 42E13ECA B6100215 2F59E2AC 0FF979D1 E4F9CC50 6037DA27 "
    strData = strData & "0DF1847D 5E92975A D3E623FE D66DE190 7B568048 5032D6B8 8235216E 9A36B7CE "
    strData = strData & "24F64B4A 80272A7A 8C6916F7 93EFED8B 4CCE1555 37DDBFF4 4B99BD25 4B95DB5D "
    strData = strData & "69812FC0 92D3FDA1 90660BB6 FB1A4A9A 46A4B9B2 730C1969 7F49DA68 81E289AA "
    strData = strData & "83B1A05F 64669A0F 9644F48B 27B3FF7D 8DB675B3 CC6B615C BCEBBE95 674F20B9 "
    strData = strData & "75655982 6F312382 3E45CF05 5AE48871 54C21157 BF619F99 40A8EAE9 EABAC460 "
    strData = strData & "F2C0C1CD 454C6FE9 6412691C 419CF649 265B0F70 D3DC3BEF C3578A9E 6D0E60F5 "

    avntData = Split(strData, Chr$(32))   ' Load variant array

    ' load T3 table array
    For lngIndex = 0 To 511 Step 2

        ' Convert variant data to Long Integer
        T3(0, lngIdx) = Val("&H" & avntData(lngIndex))
        T3(1, lngIdx) = Val("&H" & avntData(lngIndex + 1))
        lngIdx = lngIdx + 1

    Next lngIndex

    strData = vbNullString       ' Verify string is empty
    avntData = Empty   ' Verify variant is empty
    lngIdx = 0

    ' Load T4 table array
    strData = strData & "26323C55 5B0E6085 FA1B59F5 1A46C1A9 7C4C8FFA A9E245A1 DB2955D7 65CA5159 "
    strData = strData & "CE35AFC2 05DB0A76 A9113D45 81EAC77E B6AC0A0D 528EF88A 597BE3FF A09EA253 "
    strData = strData & "AC48CD56 430DDFB3 F45CE46F C4B3A67A FBE2D05E 4ECECFD8 B39935F0 3EF56F10 "
    strData = strData & "9CD619C6 0B22D682 74DF2069 17FD460A 8510ED40 6CF8CC8E 3A6ECAA7 D6C824BF "
    strData = strData & "1A817049 61243D58 BBC163A2 048BACB6 7D44CC32 D9A38AC2 AAF410AB 7FDDFF5B "
    strData = strData & "A804824B AD6D495A 2D8C9F94 E1A6A74F 35DEE8E3 D4F78512 6540D893 FD4B7F88 "
    strData = strData & "2AA4BFDA 247C2004 17D1327C 096EA1C5 361A6685 D56966B4 1221057D 277DA5C3 "
    strData = strData & "A43ACFF7 94D59893 CDC02281 64F0C51C FF6189DB 3D33BCC4 4CE66AF1 E005CB18 "
    strData = strData & "1DB99BEA FF5CCD1D FE42980F B0B854A7 718D4B9F 7BD46A6A 22A5FD8C D10FA8CC "
    strData = strData & "2BE4BD31 D3148495 CB243847 C7FA975F 5846C407 4886ED1E 1EB70B04 28CDDB79 "
    strData = strData & "F573417F C2B00BE2 2180F877 5C959045 F370EB00 7A6BDDFF D6D9D6A4 CE509E38 "
    strData = strData & "647FA702 EBEB0F00 76606F06 1DCC06CF A286FF0A E4D9F28B C918C262 D85A305D "
    strData = strData & "32225F54 475B1D87 68CCB5FE 2D4FB516 D72BBA20 A679B9D9 912D43A5 53841C0D "
    strData = strData & "BF12A4E8 3B7EAA48 F22F1DDF 781E0E47 0AB50973 EFF20CE6 9DFFB742 20D261D1 "
    strData = strData & "062A2E39 16A12B03 39650495 1960EB22 D50EB8B8 251C16FE F826016E 9AC0C330 "
    strData = strData & "953E7671 ED152665 A6369570 02D63194 94B1C987 5074F083 90B25CE1 70BA598C "
    strData = strData & "0B9742F6 794A1581 FCAF8C6C 0D5925E9 D868744E 3067716C E8D7731B 910AB077 "
    strData = strData & "5AC42F61 6A61BBDB F0851567 93513EFB 9E83E9D5 F494724B 5C09648D E887E198 "
    strData = strData & "75370CFD 34B1D3C6 BC0D255D DC35E433 34131BE0 D0AAB842 B48B7EAF 08042A50 "
    strData = strData & "44A3AB35 9997C4EE 201799D0 829A7B49 B7C54441 263B8307 FD6A6CA6 752F95F4 "
    strData = strData & "2C08C6E5 92721740 A795D9EE 2A8AB754 2F72943D A442F755 19781208 2C31334E "
    strData = strData & "EAEE6291 4FA98D7C 665DB309 55C3862F 5D53B1F3 BD061017 40413F27 46FE6CB8 "
    strData = strData & "DF0CFA59 3FE03792 2EB85E8F CFE70037 ADBCE118 A7BE29E7 DE8431DD E544EE5C "
    strData = strData & "41F1873E 8A781B1B A0D2F0E7 A5C94C78 77B60728 39412E28 AFC9A62C A1265EF3 "
    strData = strData & "6A2506C5 BCC2770C DCE1CE12 3AB66DD5 4A675B37 E65499D0 81BFD216 7D8F5234 "
    strData = strData & "EC15F389 0F6F64FC 8B5B13C8 74EFBE61 14273E1D ACDC82B7 03199D17 DD40BFE0 "
    strData = strData & "E7E061F8 37E99257 04775AAA FA526269 463D56F9 8BBBF63A 43A26E64 F0013F15 "
    strData = strData & "879EC898 A8307E9F 150177CC CC4C27A4 CA1D3348 1B432F2C 9F6FA013 DE1D1F8F "
    strData = strData & "47A7DDD6 606602A0 CC1CB2C7 D237AB64 25FCD1D3 9B938E72 8E0FF476 EC4E0370 "
    strData = strData & "3D03C12D FEB2FBDA EE43889A AE0BCED2 EBFB4F43 22CB8923 3CF7396D 69360D01 "
    strData = strData & "D2D4E022 855E3602 D01F784C 073805BA 3852F546 33E17A13 8AC7B638 DF487405 "
    strData = strData & "678AA14A BA92B29C 6CFAADCD 0CE89FC7 08339E34 5F9D4E09 1F5923B9 F1AFE929 "
    strData = strData & "0F4A265F 6E3480F6 B29B841C EEBF3A2A 8F91B4AD E21938A8 45C6D3C3 57DFEFF8 "
    strData = strData & "F62CAAF2 2F006B0B 6F75EE78 62F479EF 1C8916A9 11A55AD4 84FED453 F229D290 "
    strData = strData & "16B000E6 42F1C27B 9823C074 2B1F7674 C2745360 4B76ECA3 B91691BD 8C98F463 "
    strData = strData & "F1ADE66A 14BCC93C 6D458397 8885213E 274D4711 8E177DF0 503F2951 B49B73B5 "
    strData = strData & "C3F96B6B 10168168 63CAB0AE 0E3D963B 55A1DB14 8DFC4B56 6E14DE5C F789F135 "
    strData = strData & "4E51DAC1 683E68AF 8D4B0FD9 C9A84F9D 52A0F9D1 3691E03F E1878E80 5ED86E46 "
    strData = strData & "99D07150 3C711A0E 0C4E9310 5A0865B2 E4F0682E 56FBFC1F 105EDF9B EA8D5DE3 "
    strData = strData & "2379187A 71ABFDB1 BEE77B9C 2EB99DE1 33CF4523 21ECC0EA 1805C7A1 59A4D752 "
    strData = strData & "56AE7C72 3896F5EB B18F75DC AA638F3D ABE9808E 9F39358D C00B72AC B7DEFA91 "
    strData = strData & "62492D92 6B5541FD F92E4D5B 6DC6DEE8 C4BEEA7E 353F57AB DA5690CE 735769D6 "
    strData = strData & "42391484 0A234AA6 28F80D9D F6F95080 7AB3F215 B8E319A2 51341A4D 31AD9C11 "
    strData = strData & "7BEF5805 773C22A5 07968633 45C7561A 249DBE36 F913DA9E 78A64C68 DA652D9B "
    strData = strData & "3BC334EF 4C27A97F E66B17F4 76621220 9ACD7D0B 96774389 E0ED6782 F3EE5BCA "
    strData = strData & "00C879FC 409F7536 B5926DB6 06D09A39 317AC588 6F83AEB0 86381F21 01E6CA4A "
    strData = strData & "D19F3025 66FF3462 DDFD3BFB 72207C24 E2ECE2EB 4AF6B6D3 C7EA08DE 9C994DBE "
    strData = strData & "B09A8BC4 49ACE597 CF0797BA B38C4766 C57C2A75 131B9373 61931E58 B1822CCE "
    strData = strData & "09BA1C0C 9D7555B9 937D11D2 127FAFDD C66D92E4 29DA3BAD 54C2ECBC A2C1D571 "
    strData = strData & "82F6FE24 58C5134D 5B62274F 1C3AE351 01CB8126 E907C82E 13E37FCB F8ED0919 "
    strData = strData & "C80046C9 3249D8F9 E388FB63 80CF9BED 116CF19E 1881539A 6BD52457 5103F3F7 "
    strData = strData & "AE47F7A8 15B7E6F5 D47E9CCF DBD7C6DE 0228BB1A 44E55C41 5EDB4E99 B647D425 "
    strData = strData & "B8AAFC30 5D11882B 29D3212A F5098BBB E90296B3 8FB5EA14 57DD025A 677B9421 "
    strData = strData & "A390ACB5 FB58E7C0 83BD4A01 89D3674C 4BF3B93B 9E2DA4DF 8CAB4829 FCC41E32 "
    strData = strData & "BA582C52 03F38C96 7FD85DB2 CAD1BDBD 6082AE83 BBB442C1 A5DA9AB0 B95FE86B "
    strData = strData & "3771A93F B22E0467 493152D8 845358C9 97B4541E BE2A4886 D38E6966 95A2DC2D "
    strData = strData & "923C852B C02C11AC 0DF2A87B 2388B199 1B4F37BE 7C8008FA 4D54E503 1F70D0C8 "
    strData = strData & "7ECE57D4 5490ADEC D9063A3A 002B3C27 8030A2BF 7EAEA384 ED2003C0 C602326D "
    strData = strData & "69A94086 83A7287D 30F57A8A C57A5FCB 79EBE779 B56844E4 05DCBCE9 A373B40F "
    strData = strData & "88570EE2 D71A786E BDE8F6A0 879CBACD C164A32F 976AD1BC 9666D78B AB21E25E "
    strData = strData & "E5E5C33C 901063AA 48698D90 9818B344 3E1E8ABB E36487AE 893BDCB4 AFBDF931 "
    strData = strData & "5FBBD519 6345A0DC 9B9465CA 8628FE26 3F9C51EC 1E5D0160 A15049B7 4DE44006 "
    strData = strData & "F776CBB1 BF6C70E5 EF552BED 411218F2 705A36A3 CB0C0708 4F986044 E74D1475 "
    strData = strData & "0EA8280E CD56D943 535F5065 C12591D7 720AEF96 C83223F1 7363A51F C3A0396F "

    avntData = Split(strData, Chr$(32))   ' Load variant array

    ' load the T4 table array
    For lngIndex = 0 To 511 Step 2

        ' Convert variant data to Long Integer
        T4(0, lngIdx) = Val("&H" & avntData(lngIndex))
        T4(1, lngIdx) = Val("&H" & avntData(lngIndex + 1))
        lngIdx = lngIdx + 1

    Next lngIndex

    strData = vbNullString  ' Verify string variable is empty
    avntData = Empty        ' Verify variant variable is empty
    lngIdx = 0              ' Set array index to zero

End Sub

Private Sub ResetVariables()

    Erase malngHash()
    Erase T1()
    Erase T2()
    Erase T3()
    Erase T4()

    mcurAccrued = 0@
    mcurBitCount = 0@
    mcurCurrBitCnt = 0@
    mstrBitsInHex = vbNullString

End Sub

Private Sub Class_Initialize()

    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objTiger3 = new cTiger3

    ResetVariables
    Set mobjMath32 = Nothing      ' Free class object from memory
    Set mobjMath32 = New cMath32  ' Instantiate class object

    ' Preset property values
    ReturnLowercase = False       ' Default return of hashed data
    HashRounds = 6                ' Set default number of passes
    HashMethod = eTIGER192        ' Set default hash algorithm

End Sub

Private Sub Class_Terminate()

    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objTiger3 = Nothing

    ResetVariables
    Set mobjMath32 = Nothing

End Sub



