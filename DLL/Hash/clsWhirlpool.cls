VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWhirlpool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:      Whirlpool Hash
'
'              Whirlpool versions for -224, -256, -384 bit output are my
'              experiment.  Any problems with these outputs, please email
'              me at:
'
'                        Kenneth Ives  kenaso@tx.rr.com
'
'              DO NOT CONTACT   Vincent Rijmen
'                               Paulo S.L.M.Barreto
'
'              because they did not write this module nor are they
'              responsible in any manner as to its content.
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
' Description:   The WHIRLPOOL Hashing Function
'                Designed by Vincent Rijmen and Paulo S. L. M. Barreto
'
'                The WHIRLPOOL hashing function is named after the Whirlpool
'                galaxy in Canes Venatici (M51, or NGC 5194), the first one
'                recognized to have spiral structure by William Parsons,
'                third Earl of Rosse, in April 1845 (cf. M. Hoskin, "The
'                Cambridge Illustrated History of Astronomy," Cambridge
'                University Press, 1997).
'
'                Whirlpool is a cryptographic hash function adopted by
'                the International Standards Organization (ISO) and
'                International Electrotechnical Commission (IEC) as part
'                of the joint ISO/IEC 10118-3 international standard.
'                It takes an arbritrary block of data and returns a 512
'                bit digest that can be used as a digital fingerprint for
'                message authentication.  Compliance with the standard
'                may be verified at:
'                http://hash.online-convert.com/whirlpool-generator
'
'                Whirlpool is a hash designed after the Square block cipher.
'                Whirlpool is a Miyaguchi-Preneel construction based on a
'                substantially modified Advanced Encryption Standard (AES).
'                It takes a message of any length less than 2256 bits and
'                returns a 512-bit message digest.
'
'                The authors have declared that "WHIRLPOOL is not (and will
'                never be) patented and may be used free of charge for any
'                purpose. The reference implementations are in the public
'                domain."
'
'                Using the reference C implementation on a 1 GHz Pentium
'                III platform, we observe that Whirlpool operates at
'                about 73 cycles per hashed byte.
'
'                The compression function runs at about 56 cycles per
'                hashed byte. Many factors explain the observed
'                performance. First, a 32-bit processor was used to test
'                a native 64-bit implementation; better results are
'                expected by merely running the speed measurement on an
'                Alpha or Itanium processor. Second, it seems that the
'                pipe parallelism capabilities of the Pentium were not
'                fully used; this may reflect a non-optimising
'                implementation of 64-bit arithmetic support by the C
'                compiler, and might be overcome by an assembler
'                implementation. Third, the tables employed in the
'                reference implementation are quite large, and the
'                built-in processor cache might not be enough to hold
'                them, the data being hashed, and the hashing code at
'                once, thus degrading processing speed.
'
'                Whirlpool is much more scalable than most modern
'                hashing functions. Even though is not specifically
'                oriented toward any platform, it is rather efficient on
'                many of them, its structure favouring extensively
'                parallel execution of the component mappings. At the
'                same time, it does not require excessive storage space
'                (either for code or for tables), and can therefore be
'                efficiently implemented in quite constrained
'                environments like smart cards, although it can benefit
'                from larger cache memory available on modern processors
'                to achieve higher performance. It does not use
'                expensive or unusual instructions that must be built in
'                the processor. The mathematical simplicity of the
'                primitive resulting from the design strategy tends to
'                make analysis easier. And finally, it has a very long
'                hash length; this not only provides increased
'                protection against birthday attacks, but also offers a
'                larger internal state for entropy containment, as is
'                needed for certain classes of pseudo-random number
'                generators.
'
' Reference:     The WHIRLPOOL Hash Function (Home page)
'                http://www.larc.usp.br/~pbarreto/WhirlpoolPage.html
'
'                Whirlpool (cryptography)
'                http://en.wikipedia.org/wiki/Whirlpool_(cryptography)
'
'                Whirlpool Hashing Function in Visual Basic 6.0
'                John Korejwa  <korejwa@tiac.net>
'                31-Dec-2010
'                http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=73638&lngWId=1
'
' ***************************************************************************
'
'  Whirlpool hash output using one (1) pass
'
'  Whirlpool-512 in accordance with ISO test vectors tested to be correct.
'  Whirlpool-224, -256, -384 are dynamic (unofficial) and can be changed
'  by the hash distributor.
'
'  "a"
'        8aca2602792aec6f11a67206531fb7d7f0dff59413145e6973c45001d0087b42d11bc645413aeff63a42391a39145a591a92200d560195e53b478584fdae231a
'
'  "abc"
'        4e2448a4c6f486bb16b6562c73b4020bf3043e3a731bce721ae1b303d97e6d4c7181eebdb6c57e277d0e34957114cbd6c797fc9d95d8b582d225292076d4eef5
'
'  "message digest"
'        378c84a4126e2dc6e56dcc7458377aac838d00032230f53ce1f5700c0ffb4d3b8421557659ef55c106b4b52ac5a4aaa692ed920052838f3362e86dbd37a8903e
'
'  "abcdefghijklmnopqrstuvwxyz"
'        f1d754662636ffe92c82ebb9212a484a8d38631ead4238f5442ee13b8054e41b08bf2a9251c30b6a0b8aae86177ab4a6f68f673e7207865d5d9819a3dba4eb3b
'
'  "abcdbcdecdefdefgefghfghighijhijk"
'        2a987ea40f917061f5d6f0a0e4644f488a7a5a52deee656207c562f988e95c6916bdc8031bc5be1b7b947639fe050b56939baaa0adff9ae6745b7b181c3be3fd
'
'  "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
'        dc37e008cf9ee69bf11f00ed9aba26901dd7c28cdec066cc6af42e40f82f3a1e08eba26629129d8fb7cb57211b9281a65517cc879d7b962142c65f5a7af01467
'
'  8 times "1234567890"
'        466ef18babb0154d25b9d38a6414f5c08784372bccb204d6549c4afadb6014294d5bd8df2a6c44e538cd047b2681a51a2c60481e88c5a20b2c2a80cf3a9a083b
'
'  1 million letter "a"
'        0c99005beb57eff50a7cf005560ddf5d29057fd86b20bfd62deca0f1ccea4af51fc15490eddc47af32bb2b66c34ff9ad8c6008ad677f77126953b226e4ed8b01
'
'  1 million binary zeroes
'        f3008d58b928890f6c902cc7e69f292cfd4870a015e7d4200612a5a3a7eb5b71be0bf894d2a030b59d7a608e63daf69cfd3d69cd77e446ebf4952b9470d820af
'
'  10 million binary zeroes  (Non ISO - exceeds 5mb limit in this module)
'        6d578ff856536058c50b8c737a71217126a0777e9a5c0443d10dfacd28da6c7f8988dfe02e325aa94b3ada8ae8b52a1dd199c3e47ce05efa673cc4182f63af1b
'
' ***************************************************************************
'
'  To create a 10mb test file, use the following code:
'
'     Dim hFile As Long
'     hFile = Freefile
'     Open "C:\Temp\MB_10.tst" For Binary Access Write As #hFile
'     Put #hFile, 10000000, Chr$(0)    ' Fill with binary 0's
'     Close #hFile
'
'  Got same results hashing 10mb file with this module and at web site
'  http://hash.online-convert.com/whirlpool-generator
'
'  Creating a file this way is fast and easy.  The CHR$(0) is the last
'  character in the file.  All previous characters are null values.  You
'  can create files up to one gigabyte in size.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 31-Dec-2010  John Korejwa  korejwa@tiac.net
'              Whirlpool Hashing Function in Visual Basic 6.0
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=73638&lngWId=1
' 05-Feb-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 30-Mar-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modified to produce 256 and 384 bit hashes.  See Initialize()
'              and FormatOutput() routines.
' 13-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modified to produce 224 bit hashes.  See Initialize() and
'              FormatOutput() routines.
' 20-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated number of data mixing iterations. See Initialize()
'              routine.
' 04-May-2012  Kenneth Ives  kenaso@tx.rr.com
'              Added new property ReturnLowercse().  Designates if returned
'              hashed data string should be in upper or lowercase format.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "clsWhirlpool"
  Private Const MAX_ROUNDS  As Long = 10
  Private Const BLOCK_SIZE  As Long = 64
  Private Const WORK_SIZE   As Long = 15
  Private Const BIT_CHUNK   As Long = 512
  Private Const MAX_BYTE    As Long = 256
  Private Const FIVE_MB     As Long = 5000000  ' Must be multiple of 64

' ***************************************************************************
' Enumerations
' ***************************************************************************
  Private Enum enumWHIRLPOOL_ALGORITHM
      eWHIRLPOOL224  ' 0
      eWHIRLPOOL256  ' 1
      eWHIRLPOOL384  ' 2
      eWHIRLPOOL512  ' 3  Default
  End Enum

' ***************************************************************************
' Module Variables
'
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (Byte)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a byt Hashed
' Variable name:     mabytHashed
'
' ***************************************************************************
  Private mlngHashMethod    As enumWHIRLPOOL_ALGORITHM
  Private mblnRetLowercase  As Boolean
  Private mcurAccrued       As Currency   ' Calculating progress
  Private mcurMaxSize       As Currency
  Private mlngHashRounds    As Long       ' Extend hash iterations (see property HashRounds)
  Private mlngTotalBitsLow  As Long       ' Accumulators
  Private mlngTotalBitsHigh As Long
  Private malngKey()        As Long       ' Round key
  Private malngHash()       As Long       ' Hashed data as long integer
  Private malngWord()       As Long       ' Data converted into long integers
  Private malngRound()      As Long       ' Round constants
  Private malngCTab0()      As Long       ' Circulant table constants
  Private malngCTab1()      As Long
  Private malngCTab2()      As Long
  Private malngCTab3()      As Long
  Private malngCTab4()      As Long
  Private malngCTab5()      As Long
  Private malngCTab6()      As Long
  Private malngCTab7()      As Long
  Private mabytHashed()     As Byte       ' Hashed results

' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************

' Update progress bar
Public Event HashProgress(ByVal lngProgress As Long)


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let HashMethod(ByVal lngData As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashMethod = eWHIRLPOOL512
    Select Case lngData
           Case 0 To 3: mlngHashMethod = lngData
           Case Else:   mlngHashMethod = eWHIRLPOOL512  ' Use default value
    End Select

End Property

Public Property Let HashRounds(ByVal lngData As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashRounds = 1
    Select Case lngData
           Case 1 To MAX_ROUNDS: mlngHashRounds = lngData   ' Good selection
           Case Else:            mlngHashRounds = 1         ' Default value
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
' Description:   Function to create a unique hex string representation of 
'                the data passed.  Can process files larger than 2 gb.
'                See Readme.txt file for more information. 
'
' Syntax:
'    With mobjWHIRLPOOL
'        abytData() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)  ' convert file location to byte array 
'        mabytHashed() = .HashFile(abytData())                        ' hash data and return as Byte array
'        txtData.Text = StrConv(mabytHashed(), vbUnicode)             ' convert byte array to string data
'    End With
' 
' Parameters:    abytData() - Path/filename in byte array to be hashed
'
' Returns:       Hashed data string in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-Feb-2011  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function HashFile(ByRef abytFile() As Byte) As Byte()

    Dim hFile        As Long
    Dim lngByteCnt   As Long
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

    mcurAccrued = 0@  ' Init accrued progress
    curFilePos = 0@   ' set to first position in the file

    Set objBigFiles = New cBigFiles  ' Instantiate class object
    ResetVariables                   ' Reset class variables
    Initialize                       ' load work and constant arrays

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo HashFile_CleanUp
    End If

    With objBigFiles

        strSource = ByteArrayToString(abytFile())    ' Convert byte array to string
        .CalcFileSize strSource, curAmtLeft          ' Get the size of the file
        mcurMaxSize = curAmtLeft                     ' Capture file size

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

        ' If file exceeds 5mb then
        ' process data in chunks
        Do While (curAmtLeft > FIVE_MB)

            lngBlockSize = FIVE_MB                         ' Adjust block size
            ReDim abytData(lngBlockSize - 1)               ' Size data array
            curAmtLeft = curAmtLeft - CCur(lngBlockSize)   ' Adjust amount of file left to process

            ' Read source file
            If Not .API_ReadFile(hFile, curFilePos, abytData()) Then
                gblnStopProcessing = True
                Exit Do  ' exit Do..Loop
            End If

            ' Adjust pointers accordingly
            curFilePos = curFilePos + CCur(UBound(abytData) + 1)
            lngByteCnt = 0

            Do While (lngByteCnt < lngBlockSize)

                CopyMemory malngWord(0), abytData(lngByteCnt), BLOCK_SIZE   ' Load word array
                lngByteCnt = lngByteCnt + BLOCK_SIZE                        ' Increment index pointer
                mlngTotalBitsLow = mlngTotalBitsLow + BIT_CHUNK             ' Increment bit count
                Transform                                                   ' Hash data

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

            ' Update accumulators
            mlngTotalBitsHigh = mlngTotalBitsHigh + (mlngTotalBitsLow \ &H1000000)
            mlngTotalBitsLow = mlngTotalBitsLow And &HFFFFFF

        Loop

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            GoTo HashFile_CleanUp
        End If

        ' See if any data is left
        ' or file is less than 5mb
        If (curAmtLeft > 0@) Then

            lngBlockSize = CLng(curAmtLeft)    ' Final block size
            Erase abytData()                   ' empty data array
            ReDim abytData(lngBlockSize - 1)   ' Size receiving array

            ' Read source file
            If Not .API_ReadFile(hFile, curFilePos, abytData()) Then
                gblnStopProcessing = True
                GoTo HashFile_CleanUp
            End If

            ' Process message data
            If Not LoadWordArray(abytData()) Then
                gblnStopProcessing = True
            End If

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                GoTo HashFile_CleanUp
            End If

        End If

        .API_CloseFile hFile   ' Close file opened by this routine

    End With

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo HashFile_CleanUp
    End If

    ' Concatenate the state array into one data
    ' string and return the digital signature.
    HashFile = FormatOutput()

HashFile_CleanUp:
    objBigFiles.API_CloseFile hFile  ' Close file opened by this routine
    Set objBigFiles = Nothing        ' Always free objects from memory
    ResetVariables                   ' Reset class variables

End Function

' ***************************************************************************
' Routine:       HashString
'
' Description:   Function to digest a text string and output the result as
'                a string of hexadecimal characters.  Got basic idea from
'                David Ireland of DI Management Services Pty Ltd
'                <www.di-mgt.com.au>.
'
' Syntax:
'    With mobjWHIRLPOOL
'        abytData() = StrConv("abc", vbFromUnicode)       ' convert string data to byte array 
'        mabytHashed() = .HashString(abytData())          ' hash data and return as Byte array
'        txtData.Text = StrConv(mabytHashed(), vbUnicode) ' convert byte array to string data
'    End With
' 
' Parameters:    abytData() - Path/filename in byte array to be hashed
'
' Returns:       Hashed data string in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-Feb-2011  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
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

    mcurAccrued = 0@                                             ' Init accrued progress
    mcurMaxSize = CCur((UBound(abytData) + 1) * mlngHashRounds)  ' Capture data size

    ' Process message data
    If Not LoadWordArray(abytData()) Then
        gblnStopProcessing = True    ' Set flag to stop processing
    Else
        HashString = FormatOutput()  ' Format final output
    End If

HashString_CleanUp:
    ResetVariables   ' Reset class variables

End Function


' ***************************************************************************
' ****                  Internal functions and Procedures                ****
' ***************************************************************************

' ***************************************************************************
' Routine:       FormatOutput
'
' Description:   Concatenate the hashed values into one string.
'
'                Output length in hex format:
'                    WHIRLPOOL-224 =  56 characters
'                    WHIRLPOOL-256 =  64 characters
'                    WHIRLPOOL-384 =  96 characters
'                    WHIRLPOOL-512 = 128 characters
'
' Returns:       Hex data in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-Feb-2011  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 30-Mar-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modified to produce 256 and 384 bit hashes.
' 13-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modified to produce 224 bit hash.
' ***************************************************************************
Private Function FormatOutput() As Byte()

    Dim strHex     As String
    Dim intIndex   As Integer
    Dim intPointer As Integer

    strHex = Space$(150)  ' preload output buffer
    intPointer = 1        ' Start position in output buffer

    For intIndex = 0 To BLOCK_SIZE - 1
        Mid$(strHex, intPointer, 2) = Right$("00" & Hex$(mabytHashed(intIndex)), 2)
        intPointer = intPointer + 2
    Next intIndex

    ' Return appropriate amount of data
    Select Case mlngHashMethod
           Case eWHIRLPOOL224: strHex = Left$(strHex, 56)
           Case eWHIRLPOOL256: strHex = Left$(strHex, 64)
           Case eWHIRLPOOL384: strHex = Left$(strHex, 96)
           Case eWHIRLPOOL512: strHex = Left$(strHex, 128)
    End Select

    If mblnRetLowercase Then
        strHex = LCase$(Trim$(strHex))   ' Convert string to lowercase
    Else
        strHex = UCase$(Trim$(strHex))   ' Convert string to uppercase
    End If

    FormatOutput = StringToByteArray(strHex)  ' convert string data to byte array

    RaiseEvent HashProgress(100)   ' Update progress bar

End Function

Private Function LoadWordArray(ByRef abytData() As Byte) As Boolean

    Dim lngIdx      As Long
    Dim lngIndex    As Long
    Dim lngBitCnt   As Long
    Dim lngBitValue As Long
    Dim abytTemp()  As Byte

    Const ROUTINE_NAME As String = "LoadWordArray"

    On Error GoTo LoadWordArray_Error

    LoadWordArray = False   ' Preset to FALSE

    ReDim abytTemp(BLOCK_SIZE - 1)  ' Size temp work array

    lngBitCnt = (UBound(abytData) + 1) * 8   ' Get bit count for incoming data

    mlngTotalBitsLow = mlngTotalBitsLow + lngBitCnt                         ' Increment accumulator
    mlngTotalBitsHigh = mlngTotalBitsHigh + (mlngTotalBitsLow \ &H1000000)  ' Increment accumulator
    mlngTotalBitsLow = mlngTotalBitsLow And &HFFFFFF                        ' Adjust low side accumulator
    lngIndex = 0                                                            ' Set array pointer

    Do While lngBitCnt >= BIT_CHUNK

        CopyMemory malngWord(0), abytData(lngIndex), BLOCK_SIZE   ' Load word array
        lngIndex = lngIndex + BLOCK_SIZE                          ' Increment index pointer
        lngBitCnt = lngBitCnt - BIT_CHUNK                         ' Decrement bit count
        Transform                                                 ' Hash data

    Loop

    lngIdx = lngBitCnt \ 8   ' Add remaining 8 bit chunks

    If lngIdx > 0 Then
        CopyMemory abytTemp(0), abytData(lngIndex), lngIdx
    End If

    lngBitCnt = lngBitCnt - 8 * lngIdx  ' Adjust bit counter
    lngBitValue = &H80&                 ' Add remaining bits
    abytTemp(lngIdx) = 0

    Do While lngBitCnt > 0

        If (abytData(lngIndex) And lngBitValue) Then
            abytTemp(lngIdx) = abytTemp(lngIdx) Or lngBitValue
        End If

        lngBitValue = lngBitValue \ 2
        lngBitCnt = lngBitCnt - 1

    Loop

    abytTemp(lngIdx) = abytTemp(lngIdx) Or lngBitValue    ' Append "1" bit
    lngIdx = lngIdx + 1                                   ' Increment pointer

    If lngIdx >= (BLOCK_SIZE \ 2) Then

        ' clear the rest of abytTemp()
        Do While lngIdx < (BLOCK_SIZE - 1)
            abytTemp(lngIdx) = 0
            lngIdx = lngIdx + 1
        Loop

        CopyMemory malngWord(0), abytTemp(0), BLOCK_SIZE   ' Load word array
        Transform                                          ' Hash data

        ' Reset temp array values
        For lngIdx = 0 To BLOCK_SIZE - 1
            abytTemp(lngIdx) = 0
        Next lngIdx

    Else

        ' Clear rest of temp array
        Do While lngIdx < BLOCK_SIZE
            abytTemp(lngIdx) = 0
            lngIdx = lngIdx + 1
        Loop

    End If

    ' append (right-justified) bit length of entire message
    abytTemp(63) = mlngTotalBitsLow And &HFF&
    abytTemp(62) = (mlngTotalBitsLow \ &H100&) And &HFF&
    abytTemp(61) = (mlngTotalBitsLow \ &H10000) And &HFF&
    abytTemp(60) = mlngTotalBitsHigh And &HFF&
    abytTemp(59) = (mlngTotalBitsHigh \ &H100&) And &HFF&
    abytTemp(58) = (mlngTotalBitsHigh \ &H10000) And &HFF&

    CopyMemory malngWord(0), abytTemp(0), BLOCK_SIZE  ' Load word array
    Transform                                         ' Hash data

    ' Store hashed results in byte array
    CopyMemory mabytHashed(0), malngHash(0), BLOCK_SIZE

    ' Reset hash data array
    For lngIndex = 0 To WORK_SIZE
        malngHash(lngIndex) = 0
    Next lngIndex

    LoadWordArray = True   ' Good finish

LoadWordArray_CleanUp:
    On Error GoTo 0      ' Nullify error trap in this routine
    Exit Function

LoadWordArray_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume LoadWordArray_CleanUp

End Function

Private Sub Transform()

    Dim lngIndex    As Long
    Dim lngRounds   As Long
    Dim lngProgress As Long
    Dim alngState() As Long   ' Cipher state
    Dim abytIdx()   As Byte   ' Index pointer

    Const ROUTINE_NAME As String = "Transform"

    On Error GoTo Transform_Error

    ReDim abytIdx(BLOCK_SIZE)    ' Size work arrays
    ReDim alngState(WORK_SIZE)

    CopyMemory malngKey(0), malngHash(0), BLOCK_SIZE  ' Load key array
    mcurAccrued = mcurAccrued + BLOCK_SIZE            ' Increment accumulator

    ' Load state array
    For lngIndex = 0 To WORK_SIZE
        alngState(lngIndex) = malngWord(lngIndex) Xor malngHash(lngIndex)
    Next lngIndex

    ' For added security, use multiple iterations
    ' by updating property HashRounds()
    ' Default = 1 iteration
    For lngRounds = 1 To mlngHashRounds

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            Exit For   ' Exit For..Next loop
        End If

        For lngIndex = 0 To 9

            CopyMemory abytIdx(0), malngKey(0), BLOCK_SIZE  ' Load index pointer array

            malngKey(0) = malngCTab0(0, abytIdx(0)) Xor malngCTab1(0, abytIdx(57)) Xor _
                          malngCTab2(0, abytIdx(50)) Xor malngCTab3(0, abytIdx(43)) Xor _
                          malngCTab4(0, abytIdx(36)) Xor malngCTab5(0, abytIdx(29)) Xor _
                          malngCTab6(0, abytIdx(22)) Xor malngCTab7(0, abytIdx(15)) Xor malngRound(0, lngIndex)

            malngKey(1) = malngCTab0(1, abytIdx(0)) Xor malngCTab1(1, abytIdx(57)) Xor _
                          malngCTab2(1, abytIdx(50)) Xor malngCTab3(1, abytIdx(43)) Xor _
                          malngCTab4(1, abytIdx(36)) Xor malngCTab5(1, abytIdx(29)) Xor _
                          malngCTab6(1, abytIdx(22)) Xor malngCTab7(1, abytIdx(15)) Xor malngRound(1, lngIndex)

            malngKey(2) = malngCTab0(0, abytIdx(8)) Xor malngCTab1(0, abytIdx(1)) Xor _
                          malngCTab2(0, abytIdx(58)) Xor malngCTab3(0, abytIdx(51)) Xor _
                          malngCTab4(0, abytIdx(44)) Xor malngCTab5(0, abytIdx(37)) Xor _
                          malngCTab6(0, abytIdx(30)) Xor malngCTab7(0, abytIdx(23))

            malngKey(3) = malngCTab0(1, abytIdx(8)) Xor malngCTab1(1, abytIdx(1)) Xor _
                          malngCTab2(1, abytIdx(58)) Xor malngCTab3(1, abytIdx(51)) Xor _
                          malngCTab4(1, abytIdx(44)) Xor malngCTab5(1, abytIdx(37)) Xor _
                          malngCTab6(1, abytIdx(30)) Xor malngCTab7(1, abytIdx(23))

            malngKey(4) = malngCTab0(0, abytIdx(16)) Xor malngCTab1(0, abytIdx(9)) Xor _
                          malngCTab2(0, abytIdx(2)) Xor malngCTab3(0, abytIdx(59)) Xor _
                          malngCTab4(0, abytIdx(52)) Xor malngCTab5(0, abytIdx(45)) Xor _
                          malngCTab6(0, abytIdx(38)) Xor malngCTab7(0, abytIdx(31))

            malngKey(5) = malngCTab0(1, abytIdx(16)) Xor malngCTab1(1, abytIdx(9)) Xor _
                          malngCTab2(1, abytIdx(2)) Xor malngCTab3(1, abytIdx(59)) Xor _
                          malngCTab4(1, abytIdx(52)) Xor malngCTab5(1, abytIdx(45)) Xor _
                          malngCTab6(1, abytIdx(38)) Xor malngCTab7(1, abytIdx(31))

            malngKey(6) = malngCTab0(0, abytIdx(24)) Xor malngCTab1(0, abytIdx(17)) Xor _
                          malngCTab2(0, abytIdx(10)) Xor malngCTab3(0, abytIdx(3)) Xor _
                          malngCTab4(0, abytIdx(60)) Xor malngCTab5(0, abytIdx(53)) Xor _
                          malngCTab6(0, abytIdx(46)) Xor malngCTab7(0, abytIdx(39))

            malngKey(7) = malngCTab0(1, abytIdx(24)) Xor malngCTab1(1, abytIdx(17)) Xor _
                          malngCTab2(1, abytIdx(10)) Xor malngCTab3(1, abytIdx(3)) Xor _
                          malngCTab4(1, abytIdx(60)) Xor malngCTab5(1, abytIdx(53)) Xor _
                          malngCTab6(1, abytIdx(46)) Xor malngCTab7(1, abytIdx(39))

            malngKey(8) = malngCTab0(0, abytIdx(32)) Xor malngCTab1(0, abytIdx(25)) Xor _
                          malngCTab2(0, abytIdx(18)) Xor malngCTab3(0, abytIdx(11)) Xor _
                          malngCTab4(0, abytIdx(4)) Xor malngCTab5(0, abytIdx(61)) Xor _
                          malngCTab6(0, abytIdx(54)) Xor malngCTab7(0, abytIdx(47))

            malngKey(9) = malngCTab0(1, abytIdx(32)) Xor malngCTab1(1, abytIdx(25)) Xor _
                          malngCTab2(1, abytIdx(18)) Xor malngCTab3(1, abytIdx(11)) Xor _
                          malngCTab4(1, abytIdx(4)) Xor malngCTab5(1, abytIdx(61)) Xor _
                          malngCTab6(1, abytIdx(54)) Xor malngCTab7(1, abytIdx(47))

            malngKey(10) = malngCTab0(0, abytIdx(40)) Xor malngCTab1(0, abytIdx(33)) Xor _
                           malngCTab2(0, abytIdx(26)) Xor malngCTab3(0, abytIdx(19)) Xor _
                           malngCTab4(0, abytIdx(12)) Xor malngCTab5(0, abytIdx(5)) Xor _
                           malngCTab6(0, abytIdx(62)) Xor malngCTab7(0, abytIdx(55))

            malngKey(11) = malngCTab0(1, abytIdx(40)) Xor malngCTab1(1, abytIdx(33)) Xor _
                           malngCTab2(1, abytIdx(26)) Xor malngCTab3(1, abytIdx(19)) Xor _
                           malngCTab4(1, abytIdx(12)) Xor malngCTab5(1, abytIdx(5)) Xor _
                           malngCTab6(1, abytIdx(62)) Xor malngCTab7(1, abytIdx(55))

            malngKey(12) = malngCTab0(0, abytIdx(48)) Xor malngCTab1(0, abytIdx(41)) Xor _
                           malngCTab2(0, abytIdx(34)) Xor malngCTab3(0, abytIdx(27)) Xor _
                           malngCTab4(0, abytIdx(20)) Xor malngCTab5(0, abytIdx(13)) Xor _
                           malngCTab6(0, abytIdx(6)) Xor malngCTab7(0, abytIdx(63))

            malngKey(13) = malngCTab0(1, abytIdx(48)) Xor malngCTab1(1, abytIdx(41)) Xor _
                           malngCTab2(1, abytIdx(34)) Xor malngCTab3(1, abytIdx(27)) Xor _
                           malngCTab4(1, abytIdx(20)) Xor malngCTab5(1, abytIdx(13)) Xor _
                           malngCTab6(1, abytIdx(6)) Xor malngCTab7(1, abytIdx(63))

            malngKey(14) = malngCTab0(0, abytIdx(56)) Xor malngCTab1(0, abytIdx(49)) Xor _
                           malngCTab2(0, abytIdx(42)) Xor malngCTab3(0, abytIdx(35)) Xor _
                           malngCTab4(0, abytIdx(28)) Xor malngCTab5(0, abytIdx(21)) Xor _
                           malngCTab6(0, abytIdx(14)) Xor malngCTab7(0, abytIdx(7))

            malngKey(15) = malngCTab0(1, abytIdx(56)) Xor malngCTab1(1, abytIdx(49)) Xor _
                           malngCTab2(1, abytIdx(42)) Xor malngCTab3(1, abytIdx(35)) Xor _
                           malngCTab4(1, abytIdx(28)) Xor malngCTab5(1, abytIdx(21)) Xor _
                           malngCTab6(1, abytIdx(14)) Xor malngCTab7(1, abytIdx(7))

            CopyMemory abytIdx(0), alngState(0), BLOCK_SIZE  ' Load index pointer array

            alngState(0) = malngCTab0(0, abytIdx(0)) Xor malngCTab1(0, abytIdx(57)) Xor _
                           malngCTab2(0, abytIdx(50)) Xor malngCTab3(0, abytIdx(43)) Xor _
                           malngCTab4(0, abytIdx(36)) Xor malngCTab5(0, abytIdx(29)) Xor _
                           malngCTab6(0, abytIdx(22)) Xor malngCTab7(0, abytIdx(15)) Xor malngKey(0)

            alngState(1) = malngCTab0(1, abytIdx(0)) Xor malngCTab1(1, abytIdx(57)) Xor _
                           malngCTab2(1, abytIdx(50)) Xor malngCTab3(1, abytIdx(43)) Xor _
                           malngCTab4(1, abytIdx(36)) Xor malngCTab5(1, abytIdx(29)) Xor _
                           malngCTab6(1, abytIdx(22)) Xor malngCTab7(1, abytIdx(15)) Xor malngKey(1)

            alngState(2) = malngCTab0(0, abytIdx(8)) Xor malngCTab1(0, abytIdx(1)) Xor _
                           malngCTab2(0, abytIdx(58)) Xor malngCTab3(0, abytIdx(51)) Xor _
                           malngCTab4(0, abytIdx(44)) Xor malngCTab5(0, abytIdx(37)) Xor _
                           malngCTab6(0, abytIdx(30)) Xor malngCTab7(0, abytIdx(23)) Xor malngKey(2)

            alngState(3) = malngCTab0(1, abytIdx(8)) Xor malngCTab1(1, abytIdx(1)) Xor _
                           malngCTab2(1, abytIdx(58)) Xor malngCTab3(1, abytIdx(51)) Xor _
                           malngCTab4(1, abytIdx(44)) Xor malngCTab5(1, abytIdx(37)) Xor _
                           malngCTab6(1, abytIdx(30)) Xor malngCTab7(1, abytIdx(23)) Xor malngKey(3)

            alngState(4) = malngCTab0(0, abytIdx(16)) Xor malngCTab1(0, abytIdx(9)) Xor _
                           malngCTab2(0, abytIdx(2)) Xor malngCTab3(0, abytIdx(59)) Xor _
                           malngCTab4(0, abytIdx(52)) Xor malngCTab5(0, abytIdx(45)) Xor _
                           malngCTab6(0, abytIdx(38)) Xor malngCTab7(0, abytIdx(31)) Xor malngKey(4)

            alngState(5) = malngCTab0(1, abytIdx(16)) Xor malngCTab1(1, abytIdx(9)) Xor _
                           malngCTab2(1, abytIdx(2)) Xor malngCTab3(1, abytIdx(59)) Xor _
                           malngCTab4(1, abytIdx(52)) Xor malngCTab5(1, abytIdx(45)) Xor _
                           malngCTab6(1, abytIdx(38)) Xor malngCTab7(1, abytIdx(31)) Xor malngKey(5)

            alngState(6) = malngCTab0(0, abytIdx(24)) Xor malngCTab1(0, abytIdx(17)) Xor _
                           malngCTab2(0, abytIdx(10)) Xor malngCTab3(0, abytIdx(3)) Xor _
                           malngCTab4(0, abytIdx(60)) Xor malngCTab5(0, abytIdx(53)) Xor _
                           malngCTab6(0, abytIdx(46)) Xor malngCTab7(0, abytIdx(39)) Xor malngKey(6)

            alngState(7) = malngCTab0(1, abytIdx(24)) Xor malngCTab1(1, abytIdx(17)) Xor _
                           malngCTab2(1, abytIdx(10)) Xor malngCTab3(1, abytIdx(3)) Xor _
                           malngCTab4(1, abytIdx(60)) Xor malngCTab5(1, abytIdx(53)) Xor _
                           malngCTab6(1, abytIdx(46)) Xor malngCTab7(1, abytIdx(39)) Xor malngKey(7)

            alngState(8) = malngCTab0(0, abytIdx(32)) Xor malngCTab1(0, abytIdx(25)) Xor _
                           malngCTab2(0, abytIdx(18)) Xor malngCTab3(0, abytIdx(11)) Xor _
                           malngCTab4(0, abytIdx(4)) Xor malngCTab5(0, abytIdx(61)) Xor _
                           malngCTab6(0, abytIdx(54)) Xor malngCTab7(0, abytIdx(47)) Xor malngKey(8)

            alngState(9) = malngCTab0(1, abytIdx(32)) Xor malngCTab1(1, abytIdx(25)) Xor _
                           malngCTab2(1, abytIdx(18)) Xor malngCTab3(1, abytIdx(11)) Xor _
                           malngCTab4(1, abytIdx(4)) Xor malngCTab5(1, abytIdx(61)) Xor _
                           malngCTab6(1, abytIdx(54)) Xor malngCTab7(1, abytIdx(47)) Xor malngKey(9)

            alngState(10) = malngCTab0(0, abytIdx(40)) Xor malngCTab1(0, abytIdx(33)) Xor _
                            malngCTab2(0, abytIdx(26)) Xor malngCTab3(0, abytIdx(19)) Xor _
                            malngCTab4(0, abytIdx(12)) Xor malngCTab5(0, abytIdx(5)) Xor _
                            malngCTab6(0, abytIdx(62)) Xor malngCTab7(0, abytIdx(55)) Xor malngKey(10)

            alngState(11) = malngCTab0(1, abytIdx(40)) Xor malngCTab1(1, abytIdx(33)) Xor _
                            malngCTab2(1, abytIdx(26)) Xor malngCTab3(1, abytIdx(19)) Xor _
                            malngCTab4(1, abytIdx(12)) Xor malngCTab5(1, abytIdx(5)) Xor _
                            malngCTab6(1, abytIdx(62)) Xor malngCTab7(1, abytIdx(55)) Xor malngKey(11)

            alngState(12) = malngCTab0(0, abytIdx(48)) Xor malngCTab1(0, abytIdx(41)) Xor _
                            malngCTab2(0, abytIdx(34)) Xor malngCTab3(0, abytIdx(27)) Xor _
                            malngCTab4(0, abytIdx(20)) Xor malngCTab5(0, abytIdx(13)) Xor _
                            malngCTab6(0, abytIdx(6)) Xor malngCTab7(0, abytIdx(63)) Xor malngKey(12)

            alngState(13) = malngCTab0(1, abytIdx(48)) Xor malngCTab1(1, abytIdx(41)) Xor _
                            malngCTab2(1, abytIdx(34)) Xor malngCTab3(1, abytIdx(27)) Xor _
                            malngCTab4(1, abytIdx(20)) Xor malngCTab5(1, abytIdx(13)) Xor _
                            malngCTab6(1, abytIdx(6)) Xor malngCTab7(1, abytIdx(63)) Xor malngKey(13)

            alngState(14) = malngCTab0(0, abytIdx(56)) Xor malngCTab1(0, abytIdx(49)) Xor _
                            malngCTab2(0, abytIdx(42)) Xor malngCTab3(0, abytIdx(35)) Xor _
                            malngCTab4(0, abytIdx(28)) Xor malngCTab5(0, abytIdx(21)) Xor _
                            malngCTab6(0, abytIdx(14)) Xor malngCTab7(0, abytIdx(7)) Xor malngKey(14)

            alngState(15) = malngCTab0(1, abytIdx(56)) Xor malngCTab1(1, abytIdx(49)) Xor _
                            malngCTab2(1, abytIdx(42)) Xor malngCTab3(1, abytIdx(35)) Xor _
                            malngCTab4(1, abytIdx(28)) Xor malngCTab5(1, abytIdx(21)) Xor _
                            malngCTab6(1, abytIdx(14)) Xor malngCTab7(1, abytIdx(7)) Xor malngKey(15)

        Next lngIndex

    Next lngRounds

    ' apply the Miyaguchi-Preneel compression function
    For lngIndex = 0 To WORK_SIZE
        malngHash(lngIndex) = malngHash(lngIndex) Xor alngState(lngIndex) Xor malngWord(lngIndex)
        malngWord(lngIndex) = 0
    Next lngIndex

    ' Update progress bar
    lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
    RaiseEvent HashProgress(lngProgress)

Transform_CleanUp:
    On Error GoTo 0
    Exit Sub

Transform_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Transform_CleanUp

End Sub

Private Sub Initialize()

    Dim strData    As String
    Dim avntData   As Variant
    Dim lngIdx     As Long
    Dim lngIndex   As Long
    Dim abytBase() As Byte
    Dim abytTemp() As Byte

    On Error GoTo Initialize_Error

    ' Size arrays
    ReDim mabytHashed(BLOCK_SIZE)  ' Hashed output
    ReDim malngKey(WORK_SIZE)      ' Round key
    ReDim malngHash(WORK_SIZE)     ' Hashed data as long integer
    ReDim malngWord(WORK_SIZE)     ' Data converted into long integers
    ReDim malngRound(0 To 1, 9)    ' Round constants

    ReDim malngCTab0(0 To 1, MAX_BYTE)  ' Circulant tables
    ReDim malngCTab1(0 To 1, MAX_BYTE)
    ReDim malngCTab2(0 To 1, MAX_BYTE)
    ReDim malngCTab3(0 To 1, MAX_BYTE)
    ReDim malngCTab4(0 To 1, MAX_BYTE)
    ReDim malngCTab5(0 To 1, MAX_BYTE)
    ReDim malngCTab6(0 To 1, MAX_BYTE)
    ReDim malngCTab7(0 To 1, MAX_BYTE)

    avntData = Empty  ' Always start with empty variants

    ReDim abytBase(MAX_BYTE)   ' Size work arrays
    ReDim abytTemp(WORK_SIZE)

    strData = vbNullString
    mlngTotalBitsLow = 0    ' Empty accumulators
    mlngTotalBitsHigh = 0

    ' 01-Jan-2012 Bug fix - Duplicated a character in original data while
    '             attempting to make it look cleaner
    '
    ' These are ASCII values 0-255 mixed in hex format with no duplicates.
    ' See Whirlpool_tbl.txt for additional sets of data strings from which
    ' to choose.  See application BuildTables.vbp to create new data.
    '
    ' Whirlpool-224, -256, -384 are dynamic and can be modified/changed
    ' by the hash distributor.
    '
    ' Each case statement:
    '    1.  Consists of all ASCII values (0-255) in two
    '        character hex equivalent with no duplicates
    '    2.  Mixed 100-600 iterations using Knuth Shuffle
    Select Case mlngHashMethod
           Case eWHIRLPOOL224
                strData = strData & "14 A8 B9 FE 9F 6B 41 03 D8 59 6F 95 D4 FC E1 A6 19 53 E5 55 69 1A 2D 51 38 3A 87 17 CC 25 11 0F "
                strData = strData & "B3 78 54 BB 9A D9 9D 81 A1 B4 DB 00 7F AE 27 F4 6A CE 2E E7 96 79 85 76 A7 56 98 0D 12 64 DF DE "
                strData = strData & "83 7A 30 EC 9C 43 8E C2 2B 8D B8 AB 7E 75 C0 FF E0 02 F2 C5 3B E6 F7 39 93 48 B2 16 24 BE 9B 77 "
                strData = strData & "0A BF B1 46 8A 3F 74 D7 4C 5D C9 2A 82 28 C8 4D D5 F6 68 CA 15 88 F1 DA 6E EF 5A 1F 22 35 EE 0B "
                strData = strData & "01 BD E8 40 C7 C1 D0 6D 32 D1 45 5E D3 D6 E3 34 FD 73 86 91 AC 7C 1D ED 57 21 1E 8B 08 20 33 06 "
                strData = strData & "3D EB 18 92 D2 A4 90 49 1B 2C C3 B5 70 FA 37 F5 DC 5C 4B 3C A0 4A 7B 10 71 E9 E2 99 52 89 4E 4F "
                strData = strData & "C4 0C CF 80 AF BC 66 A3 BA 2F 9E F9 07 29 F3 23 09 72 6C 26 8F 62 A5 58 31 50 13 B6 A9 47 04 1C "
                strData = strData & "CB 61 5F 7D 3E 65 B0 E4 B7 05 44 AD 8C 63 F8 C6 67 AA FB 97 60 42 5B EA CD 84 0E A2 DD 36 F0 94 "

           Case eWHIRLPOOL256
                strData = strData & "22 40 F0 82 B8 8B 61 B4 FE 5B 73 C1 4C A1 0D 96 AD 67 5F 24 E7 47 38 8A DA B7 B6 6B ED 51 A6 7B "
                strData = strData & "CC 06 11 1F 3B BB BA 28 C4 9B 53 4A 05 93 49 52 6A F3 37 B3 D4 76 1C 20 57 81 0F CA C8 1E 07 08 "
                strData = strData & "26 34 55 FA E4 30 F7 14 FF 23 86 2D 8F E5 D7 63 F6 DC 0E 79 74 92 CB 0B 39 03 10 2F 83 D3 AF 42 "
                strData = strData & "A0 18 A9 E1 32 CE 48 01 98 72 C2 3D B0 99 C6 54 31 3C 4E 12 87 C7 EB 78 62 43 70 2B C3 17 AE 35 "
                strData = strData & "2E AB 60 3A 6E 6F 1A 50 F9 F8 56 0C 6D 80 44 D8 3E BE 89 95 36 66 15 A5 6C 04 E0 A8 27 E9 25 9E "
                strData = strData & "02 29 7C 1D B5 9A 7D D0 BD 0A CD 45 FC 88 A4 58 FB DD 8C 21 64 A2 97 F4 AC EC B2 41 EE 33 71 E2 "
                strData = strData & "7E BF 75 DB 9F 8E 4B 8D EF 68 5C D2 9D 69 2A D1 19 85 CF 3F F5 D9 00 C5 13 A3 E8 9C 16 59 E3 BC "
                strData = strData & "91 B9 5D 84 DE 4F 5E 94 77 F1 F2 65 EA C9 5A 90 E6 AA 09 D5 B1 DF 7A 2C D6 46 FD 4D C0 A7 1B 7F "

           Case eWHIRLPOOL384
                strData = strData & "D4 EA EF 02 71 B8 7E 2A AD D6 2F F9 55 98 EE 75 C7 E0 8E F0 C1 19 82 6D 95 08 4C 7D 09 3F 07 30 "
                strData = strData & "8F C0 36 3E 28 2B 76 D1 62 AE 6C CE DE 23 E5 F3 A4 31 50 7C 1E 9E 04 B1 99 56 C6 C4 A3 C2 EB E6 "
                strData = strData & "A5 67 1B 2C 89 9B 26 A6 D0 B5 ED 9F 52 A8 CF 6A 47 3D 85 C8 00 66 E7 87 E8 6B 8C CB 69 E2 FC 1C "
                strData = strData & "03 06 17 3B 78 86 81 AB D3 68 0A CA 7B F2 D9 77 0E F4 D2 BD 70 61 5F 57 C5 A9 3A 9D 97 E1 83 49 "
                strData = strData & "D8 48 18 12 4B FF DF 54 AA 41 8A EC DA 5D DB CD B2 43 20 8D DC 42 F7 79 39 4D 91 25 33 B7 F5 5B "
                strData = strData & "A7 2D B4 72 64 32 BF 60 58 FD AF 38 B0 BC 22 10 35 6F 74 B3 E9 9A F6 7F 73 0D FA BA A0 37 6E 88 "
                strData = strData & "14 29 4E A2 F1 CC C9 8B FB 84 40 BE 1F AC 11 45 C3 A1 2E E3 13 63 53 B6 90 21 24 44 0F 92 93 27 "
                strData = strData & "51 34 5C 05 DD B9 5E D7 65 15 46 1A 4F D5 F8 9C 1D 0B 16 01 59 FE 96 7A 3C E4 4A 94 5A 80 BB 0C "

           Case eWHIRLPOOL512   ' Original data - DO NOT MODIFY
                strData = strData & "18 23 C6 E8 87 B8 01 4F 36 A6 D2 F5 79 6F 91 52 60 BC 9B 8E A3 0C 7B 35 1D E0 D7 C2 2E 4B FE 57 "
                strData = strData & "15 77 37 E5 9F F0 4A DA 58 C9 29 0A B1 A0 6B 85 BD 5D 10 F4 CB 3E 05 67 E4 27 41 8B A7 7D 95 D8 "
                strData = strData & "FB EE 7C 66 DD 17 47 9E CA 2D BF 07 AD 5A 83 33 63 02 AA 71 C8 19 49 D9 F2 E3 5B 88 9A 26 32 B0 "
                strData = strData & "E9 0F D5 80 BE CD 34 48 FF 7A 90 5F 20 68 1A AE B4 54 93 22 64 F1 73 12 40 08 C3 EC DB A1 8D 3D "
                strData = strData & "97 00 CF 2B 76 82 D6 1B B5 AF 6A 50 45 F3 30 EF 3F 55 A2 EA 65 BA 2F C0 DE 1C FD 4D 92 75 06 8A "
                strData = strData & "B2 E6 0E 1F 62 D4 A8 96 F9 C5 25 59 84 72 39 4C 5E 78 38 8C D1 A5 E2 61 B3 21 9C 1E 43 C7 FC 04 "
                strData = strData & "51 99 6D 0D FA DF 7E 24 3B AB CE 11 8F 4E B7 EB 3C 81 94 F7 B9 13 2C D3 E7 6E C4 03 56 44 7F A9 "
                strData = strData & "2A BB C1 53 DC 0B 9D 6C 31 74 F6 46 AC 89 14 E1 16 3A 69 09 70 B6 D0 ED CC 42 98 A4 28 5C F8 86 "
    End Select

    avntData = Split(strData, Chr$(32))  ' Load data into an array

    ' Load byte array.  Convert
    ' data to a smaller footprint.
    For lngIndex = 0 To (MAX_BYTE - 1)
        abytBase(lngIndex) = CByte("&H" & avntData(lngIndex))
    Next lngIndex

    avntData = Empty        ' Always empty variants when not needed
    strData = vbNullString  ' Clear string variable

    ' Calculate data to be inserted into constant arrays
    For lngIndex = 0 To (MAX_BYTE - 1)

        abytTemp(0) = abytBase(lngIndex)
        abytTemp(1) = abytBase(lngIndex)
        abytTemp(3) = abytBase(lngIndex)

        If (abytTemp(0) And &H80&) = 0& Then
            abytTemp(6) = abytTemp(0) * 2&
        Else
            abytTemp(6) = (abytTemp(0) * 2&) Xor &H11D&
        End If

        If (abytTemp(6) And &H80&) = 0& Then
            abytTemp(2) = abytTemp(6) * 2&
        Else
            abytTemp(2) = (abytTemp(6) * 2&) Xor &H11D&
        End If

        abytTemp(5) = abytTemp(2) Xor abytTemp(0)

        If (abytTemp(2) And &H80&) = 0& Then
            abytTemp(4) = abytTemp(2) * 2&
        Else
            abytTemp(4) = (abytTemp(2) * 2&) Xor &H11D&
        End If

        abytTemp(7) = abytTemp(4) Xor abytTemp(0)

        ' Build circulant table
        CopyMemory malngCTab0(0, lngIndex), abytTemp(0), 4&
        CopyMemory malngCTab0(1, lngIndex), abytTemp(4), 4&

        ' Load circulant table constants
        CopyMemory abytTemp(8), abytTemp(0), 8&
        CopyMemory malngCTab1(0, lngIndex), abytTemp(7), 4&
        CopyMemory malngCTab1(1, lngIndex), abytTemp(7 + 4), 4&
        CopyMemory malngCTab2(0, lngIndex), abytTemp(6), 4&
        CopyMemory malngCTab2(1, lngIndex), abytTemp(6 + 4), 4&
        CopyMemory malngCTab3(0, lngIndex), abytTemp(5), 4&
        CopyMemory malngCTab3(1, lngIndex), abytTemp(5 + 4), 4&
        CopyMemory malngCTab4(0, lngIndex), abytTemp(4), 4&
        CopyMemory malngCTab4(1, lngIndex), abytTemp(4 + 4), 4&
        CopyMemory malngCTab5(0, lngIndex), abytTemp(3), 4&
        CopyMemory malngCTab5(1, lngIndex), abytTemp(3 + 4), 4&
        CopyMemory malngCTab6(0, lngIndex), abytTemp(2), 4&
        CopyMemory malngCTab6(1, lngIndex), abytTemp(2 + 4), 4&
        CopyMemory malngCTab7(0, lngIndex), abytTemp(1), 4&
        CopyMemory malngCTab7(1, lngIndex), abytTemp(1 + 4), 4&

    Next lngIndex

    ' Build round constants
    For lngIndex = 0 To 9

        lngIdx = 8 * lngIndex

        malngRound(0, lngIndex) = (malngCTab0(0, lngIdx) And &HFF&) Or _
                                  (malngCTab1(0, lngIdx + 1) And &HFF00&) Or _
                                  (malngCTab2(0, lngIdx + 2) And &HFF0000) Or _
                                  (malngCTab3(0, lngIdx + 3) And &HFF000000)

        malngRound(1, lngIndex) = (malngCTab4(1, lngIdx + 4) And &HFF&) Or _
                                  (malngCTab5(1, lngIdx + 5) And &HFF00&) Or _
                                  (malngCTab6(1, lngIdx + 6) And &HFF0000) Or _
                                  (malngCTab7(1, lngIdx + 7) And &HFF000000)
    Next lngIndex

    ' Preload output hashed data array
    For lngIndex = 0 To WORK_SIZE
        malngHash(lngIndex) = 0
    Next lngIndex

Initialize_CleanUp:
    avntData = Empty  ' Always empty variants when not needed
    On Error GoTo 0
    Exit Sub

Initialize_Error:
    gblnStopProcessing = True
    Resume Initialize_CleanUp

End Sub

Private Sub ResetVariables()

    Erase malngKey()     ' Round key
    Erase malngHash()    ' Hashed data as long integer
    Erase malngWord()    ' Data converted into long integers
    Erase malngCTab0()   ' Circulant table constants
    Erase malngCTab1()
    Erase malngCTab2()
    Erase malngCTab3()
    Erase malngCTab4()
    Erase malngCTab5()
    Erase malngCTab6()
    Erase malngCTab7()
    Erase malngRound()   ' Round constants
    Erase mabytHashed()  ' Hashed output

End Sub

Private Sub Class_Initialize()

    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objWhirlpool = new cWhirlpool

    ResetVariables   ' Always start with empty arrays

    ' Preset property values
    HashRounds = 1            ' default number of rounds
    ReturnLowercase = False   ' Default return of hashed data

End Sub

Private Sub Class_Terminate()

    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objWhirlpool = Nothing

    ResetVariables   ' Always empty arrays when not needed

End Sub



