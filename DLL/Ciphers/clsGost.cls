VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Project:       GOST 28147-89 Encryption/Decryption
'
' Module:        clsGOST.cls
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
' Description:   GOST (Gosudarstvennyi Standard Soyuza SSR 28147-89) is a Soviet
'                block cipher. The Government Standard of the U.S.S.R. 28147-89,
'                Cryptographic protection for Data Protection Systems, appears
'                to have played a role in the Soviet Union similar to that played
'                by the U.S. Data Encryption Standard (FIPS 46). When issued, it
'                bore the minimal classification 'For Official Use', but is now
'                said to be widely available in software both in the Former Soviet
'                Union and elsewhere. In apparent contrast to DES's explicit
'                limitation to unclassified information, the introduction to GOST
'                28147-89 contains the intriguing remark that the cryptographic
'                transformation algorithm "does not place any limitations on the
'                secrecy level of the protected information."
'
'                The algorithms are similar in that both operate on 64-bit blocks
'                by successively modifying half of the bits with a function of the
'                other half. Beyond that, the similarity declines and several
'                differences are visible.
'
'                   - The Soviet System has 32 rounds rather than the 16 of DES.
'
'                   - Each round is somewhat similar to a round of DES.
'                     In the MixBits() function, 32 bits of text are added modulo
'                     32 to 32 bits of key, transformed by a block of eight,
'                     4-bit to 4-bit S-boxes and rotated 11 bits to the left.
'
'                   - In contrast to DES's meager 56 bits of key, GOST 28147-89
'                     has 256 bits of primary key and 512 bits of secondary key.
'                     The secondary key is the block of eight S-boxes, which are
'                     specific to individual networks and are not included in
'                     the standard.
'
'                   - In place of complex key schedule of DES, the primary key
'                     is divided into eight 32-bit words. For the first
'                     twenty-four rounds, these are used cyclically in ascending
'                     order. For the last eight, they are used in descending order.
'
'                The standard is also somewhat broader than FIPS 46.  It includes
'                output feedback and cipher feedback modes of operation, both
'                limited to 64-bit blocks, and a mode for producing message
'                authentication codes.
'
'                The translation is internally colloquial, preferring standard
'                English terms to charmingly quaint cognates of the Russian.
'                The objective is to make it easy for English speaking
'                cryptographers to identify what is novel with minimal effort.
'                In some places, particularly the glossary, which has changed
'                the wording considerably and a few remarks on some of the freer
'                choices, seems in order.
'
'                **********************************************************
' SPECIAL NOTE:  The encryption process has been enhanced.
'
'                - Encryption process can be performed multiple times.
'                  See property EncryptRounds().
'
'                - Password key length is expanded to meet and exceed AES
'                  (Advanced Encryption Standard) standards.  Key lengths
'                  are 128 to 448 in 32 bit increments and 448 to 1024 in
'                  64 bit increments.  See property KeyLength().
'
'                - Primary key has been increased in size from eight (8)
'                  words to sixteen (16) words.  These values are run
'                  thru BlockEncrypt() routine so that each value is unique.
'
'                - Primary key can be initialized one to five times in order
'                  to create unique values within the array.  See property
'                  PrimaryKeyRounds().  Used in EvaluateKey() routine.
'
'                - Modified the BlockEncrypt()/BlockDecrypt() routines to
'                  be able to access all components of the new primary key.
'                  For encryption, first forty-eight (48 = 16 * 3) rounds,
'                  primary key is used cyclically in ascending order.  For
'                  the last sixteen (16), in descending order.  Reverse
'                  this order for decryption.
'
'                - Created multiple sets of S-Box values to be determined by
'                  the key length property.  See LoadWorkBoxes() routine.
'                **********************************************************
'
' Warning:       The weakest link in any encryption is the password. This
'                includes the security, length and components that make up
'                the password.  Password length and components are usually
'                dictated by organizational standards.  Security is the
'                responsibility of the user. Do not write down or discuss
'                your password with others.  Lock your screensaver when not
'                at your workstation.
'
' Note:          After encryption, data sizes will not match original sizes.
'                This is due to internal padding and the storing of information
'                required to decrypt the data later.
'
' References:    Jetico software
'                http://www.jetico.com/index.htm#/gost.htm
'
'                Wikipedia GOST (Block cipher)
'                http://en.wikipedia.org/wiki/GOST_28147-89
'
'                The GOST 28147-89 cipher
'                http://www.koders.com/c/fid6A28BBAD70A1964DFD21D17F277632B04835D376.aspx?s=ftp
'
'                An Introduction to Using Keys in Cryptography
'                By DI Management Services Pty Ltd, Sydney Australia
'                http://www.di-mgt.com.au/cryptokeys.html
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-JUN-2004  Kenneth Ives  kenaso@tx.rr.com
'              Documented, enhanced and modified
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 08-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added GetWord() and PutWord() routines.  Fixed a logic bug
'              in BlockDecrypt() routine.
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified GOST_File() and GOST_String() routines to
'              encrypt appended data information for additional security.
'              Moved ExpandData() routine to basCommon.bas module.
' 25-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added property EncryptRounds().  Number of encryption rounds
'              is now user-defined.  See constant MAX_ROUNDS.
'              Added property KeyLength() to meet and exceed AES standards.
'              (128, 192, 256, 384, 512, 1024 bits)  Modified EvaluateKey()
'              routine to use utilize the new key lengths.
' 27-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Corrected a logic bug in BlockEcrypt()/BlockDecrypt() routines.
'              Reverted to previous code.
' 01-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Fixed OffSet range in EncryptData() and DecryptData() routines.
'              Moved GetBlockSize() routine to common module.
'              Set key length to increments of 32 bits for 128 to 448 and
'              increments of 64 bits for 448 to 1024.
'              Modified the mixing of the password key.  See EvaluateKey()
'              routine.
' 08-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Set data block size to 16 bytes.
'              Updated documentation.
' 15-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added additional looping to initializing key array in
'              EvaluateKey() routine.
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Removed obsolete code. Updated documentation.
'              Correct data size capture for decryption in GOST_String().
' 29-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Modified LoadWorkBoxes() routine with two sets of S-Box values
'              to be determined by key length.
' 15-Mar-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated EvaluateKey() and LoadWorkBoxes() routines.
' 01-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced Err.Raise with InfoMsg so programs will continue to
'              flow smoothly.
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Moved number of rounds from EncryptData() and DecryptData()
'              to BlockEncrypt() and BlockDecrypt() routines for security and
'              speed. Modified EvaluateKey() and LoadWorkBoxes() routines for
'              password key mixing.
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added ability to see progress display while encrypting or
'              decrypting data (EncryptData() and DecryptData() routines).
'              Removed obsolete code.  Updated documentation.
' 19-May-2010  Kenneth Ives  kenaso@tx.rr.com
'              Updated EvaluateKey(), MixBits() routines.
' 10-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Removed MixBytes() routine.  I created this routine and now
'              feel that it is no longer needed.
' 20-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated number of data mixing iterations. See LoadWorkBoxes()
'              routine.
'              Fixed a bug in BlockEncrypt() and BlockDecrypt() routines.
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Made many updates as per Joe Sova's suggestions
'                1. Removed most of the erase array as they were not needed
'                2. Moved call to display progress to outside processing loop
'                   in EncryptData() and DecryptData() routines.
'                3. Moved Decryption size testing to earlier position within
'                   within GOST_File() routine.
'                4. Split size testing in GOST_String() routine.
'              - Changed (KEY_ARRAY - 1) to (KEY_ARRAY - 8)
' 21-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Added Public Event CurrentBlockSize() to track current
'                amount of data being processed.
'              - Updated GetBlockSize() routine.
'              - Moved progress bar code from BlockEncrypt() and
'                BlockDecrypt() routines to EncryptData() and DecryptData()
'                routines.
'              - Changed Primary Key array size from 32 to 16 for speed.
'                Still maintains a high level of data security.
'              - Added new Property CreateNewFile().  Designates if creating
'                a new output file or overwriting the source file after
'                encryption or decryption.  Default value is TRUE.
'              - Updated GOST_File() routine to reference new property.
'              - Added new property to designate number of rounds to mix
'                primary key array.
'              - Increased speed by inserting bitwise routines into this
'                class and removing some IF..Then decisions.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME    As String = "clsGOST"
  Private Const MAX_FILE_SIZE  As Currency = 10000000000@   ' Approx 10 Gig
  Private Const MAX_STRING_LEN As Long = &HFFFE&            ' 65534
  Private Const KB_32          As Long = &H8000&            ' 32768
  Private Const BLOCK_SIZE     As Long = 8
  
  Private Const KEY_ARRAY      As Long = 16  ' Strong encryption. This is double normal size.
  'Private Const KEY_ARRAY      As Long = 32  ' Very strong encryption. Four times normal size. (Runs slower)
                                             ' Update documentation in Declaration Section ("SPECIAL NOTE:"),
                                             ' BlockEncrypt() and BlockDecrypt() routines.
                                             ' Change:  forty-eight (48 = 16 * 3) -> ninety-six (96 = 32 * 3)
  Private Const DATA_BLOCK     As Long = 16
  Private Const MAX_ROUNDS     As Long = 10
  Private Const MAX_BYTE       As Long = 256
  Private Const MAX_LONG       As Long = &H7FFFFFFF    '  2147483647
  Private Const MIN_LONG       As Long = &H80000000    ' -2147483648
  Private Const GB_4           As Double = (2# ^ 32)   '  4294967296  (== 4.2 Gig)

' ***************************************************************************
' Key dependant p-boxes
' ***************************************************************************
  Private mabytK87() As Byte
  Private mabytK65() As Byte
  Private mabytK43() As Byte
  Private mabytK21() As Byte

' ***************************************************************************
' Module Variables
'
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (Byte)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a byt Password
' Variable name:     mabytPassword
'
' ***************************************************************************
  Private mblnLoaded        As Boolean
  Private mblnCreateNewFile As Boolean
  Private mcurMaxSize       As Currency
  Private mcurAccrued       As Currency
  Private mcurOrigLength    As Currency
  Private mcurCurrByteCnt   As Currency
  Private mlngRounds        As Long
  Private mlngKeyRounds     As Long
  Private mlngKeyLength     As Long
  Private malngPrimaryKey() As Long
  Private mabytValue(8)     As Byte
  Private mabytPower2(8)    As Byte
  Private mabytPassword()   As Byte

' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************
    
' Update progress bar
Public Event CipherProgress(ByVal lngProgress As Long)

' Track amount of data being processed.
Public Event CurrentBlockSize(ByVal lngBlockSize As Long)


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let CreateNewFile(ByVal blnNewFile As Boolean)
   
    ' 21-Feb-2012 Added Property to desiginate if input file
    '             is to be overwritten after encryption/decryption.
    '
    ' TRUE - Create new file to hold encrypted/decrypted data
    ' FALSE - Overwrite input file after encryption/decryption
    
    mblnCreateNewFile = blnNewFile

End Property

Public Property Let KeyLength(ByVal lngKeyLength As Long)
        
    ' 01-Jan-2009 Set increments of 32 bits for 128 to 448
    '             Set increments of 64 bits for 448 to 1024
    Select Case lngKeyLength
           Case 128, 160, 192, 224, 256, 288, 320, 352, 384, 416, _
                448, 512, 576, 640, 704, 768, 832, 896, 960, 1024

                mlngKeyLength = lngKeyLength
                
           Case Else
                mlngKeyLength = 128  ' Use default value
    End Select
    
End Property

Public Property Let EncryptRounds(ByVal lngData As Long)
   
    ' 26-Dec-2008 Added multiple pass functionality.
    Select Case lngData
           Case 1 To MAX_ROUNDS: mlngRounds = lngData   ' Good selection
           Case Else:            mlngRounds = 1         ' Default value
    End Select
    
End Property

Public Property Let PrimaryKeyRounds(ByVal lngData As Long)
   
    ' 21-Feb-2011 Number of rounds to mix primary key array data
    '             within EvaluateKey() routine
    Select Case lngData
           Case 1 To 5: mlngKeyRounds = lngData   ' Good selection
           Case Else:   mlngKeyRounds = 1         ' Default value
    End Select
    
End Property

Public Property Let Password(ByVal vntData As Variant)

    ' Used when assigning a value to the password property, on the
    ' left side of an assignment.
    '
    ' Syntax:     X.Password = "User Supplied Password"
    '                            or
    '             X.Password = abytPwd()  <-- Highly recommend using a byte array
    
    Dim intIndex As Integer
    
    Const ROUTINE_NAME As String = "Password Property"
    
    On Error GoTo Password_Error

    Erase mabytPassword()
    
    ' Is password data in an array
    If CBool(IsArrayInitialized(vntData)) Then
    
        ReDim mabytPassword(UBound(vntData))
        
        For intIndex = 0 To UBound(vntData) - 1
            mabytPassword(intIndex) = CByte(vntData(intIndex))
        Next intIndex
    
    ' Is password data in raw string format
    ElseIf Len(Trim$(vntData)) > 0 Then
        mabytPassword() = StringToByteArray(CStr(Trim$(vntData)))
            
    ' no data passed
    Else
        InfoMsg "Missing or invalid password." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
    End If

Password_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Property

Password_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume Password_CleanUp
    
End Property


' ***************************************************************************
' ****                         Methods                                   ****
' ***************************************************************************

' ***************************************************************************
' Routine:       GOST_File
'
' Description:   Function to encrypt/decrypt a file.
' 
' Syntax:        abytFile() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)  ' convert file location to byte array 
'                objGOST.GOST_File abytFile(), blnEncrypt                     ' Encrypt/Decrypt data
' 
' Parameters:    abytFile() - Path/filename in byte array to be hashed
'                blnEncrypt - [Optional] - TRUE - Encrypt incoming data
'                             FALSE - Decrypt incoming data
'                             Default=TRUE
'
' Returns:       TRUE - Successful completion
'                FALSE - An error occurred
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2004  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array for security purposes
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 08-MAY-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added call for file verification.
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added call to MixAppendedData() routine to encrypt/decrypt
'              size information that is to be appended to original data.
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Removed obsolete code. Updated documentation.
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Moved Decryption size testing to earlier an position as per
'              Joe Sova's suggestion.
' ***************************************************************************
Public Function GOST_File(ByRef abytFile() As Byte, _
                 Optional ByVal blnEncrypt As Boolean = True) As Boolean

    Dim strPath      As String
    Dim strSource    As String
    Dim strTarget    As String  ' Not used when overlaying source file
    Dim strTempFile  As String
    Dim strFileName  As String
    Dim hFile1       As Long
    Dim hFile2       As Long
    Dim lngBlockSize As Long
    Dim lngLastBlock As Long
    Dim curFilePos1  As Currency
    Dim curFilePos2  As Currency
    Dim curAmtLeft   As Currency
    Dim abytData()   As Byte
    Dim abytSize(11) As Byte
    Dim objBigFiles  As cBigFiles

    Const ROUTINE_NAME As String = "GOST_File"
    
    On Error GoTo GOST_File_Error
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
    
    Set objBigFiles = New cBigFiles   ' Instantiate class object
    Erase abytData()                  ' Always start with empty arrays
    Erase abytSize()
    ResetVariables                    ' Reset class variables
    
    If Not EvaluateKey() Then
        GoTo GOST_File_CleanUp
    End If
    
    strSource = ByteArrayToString(abytFile())           ' convert incoming array to string
    strPath = QualifyPath(GetPath(strSource))           ' Capture source path
    strFileName = GetFilename(strSource)                ' Capture source filename
    SetFileAttributes strSource, FILE_ATTRIBUTE_NORMAL  ' Reset file attributes to normal
    
    ' See if a new file is to be created after
    ' performing encryption or decryption
    If mblnCreateNewFile Then
        If blnEncrypt Then
            ' Append "ENC" extension to target file
            ' designating the file is encrypted.
            strTarget = strPath & strFileName & ENCRYPT_EXT
        Else
            ' Append "DEC" extension to target file
            ' designating the file is decrypted.
            strTarget = strPath & strFileName & DECRYPT_EXT
        End If
    End If
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo GOST_File_CleanUp
    End If

    With objBigFiles
        .CalcFileSize strSource, curAmtLeft  ' Get size of file
        
        ' If zero byte file then leave
        If curAmtLeft < 1 Then
            gblnStopProcessing = True
            GoTo GOST_File_CleanUp
        End If
            
        ' If performing decryption then do not
        ' process the last few bytes of the file
        ' because this is where we stored the
        ' size of the original file.
        If blnEncrypt Then
            mcurOrigLength = curAmtLeft   ' Capture total byte count
        Else
            curAmtLeft = curAmtLeft - (UBound(abytSize) + 1)   ' Adjust amount that has to be decrypted
            curFilePos1 = curAmtLeft                           ' Set pointer to capture the last 12 bytes
            
            If Not .OpenReadOnly(strSource, hFile1) Then      ' Open the source file
                gblnStopProcessing = True
            End If
            
            .API_ReadFile hFile1, curFilePos1, abytSize()     ' Capture last 12 bytes of source file
            .API_CloseFile hFile1                             ' close source file
            
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                GoTo GOST_File_CleanUp
            End If
    
            MixAppendedData abytSize()                         ' Decrypt appended data
            CopyMemory lngLastBlock, abytSize(0), 4&           ' Size of last block of data without padding [0-3]
            CopyMemory mcurOrigLength, abytSize(4), 8&         ' Capture original file size [4-11]
            
            ' Decryption error test
            If (lngLastBlock < 0) Or _
               (lngLastBlock > MAX_FILE_SIZE) Or _
               (curAmtLeft < 1) Or _
               (mcurOrigLength < 0) Or _
               (mcurOrigLength > MAX_FILE_SIZE) Then
            
                InfoMsg "This file cannot be Decrypted, it is" & vbNewLine & _
                        "either not Encrypted, too large or corrupt." & vbNewLine & vbNewLine & _
                        "Current size:  " & Format$(curAmtLeft, "#,##0") & " bytes" & vbNewLine & _
                        "Original size:  " & Format$(mcurOrigLength, "#,##0") & " bytes" & vbNewLine & _
                        "Max allowed size:  " & Format$(MAX_FILE_SIZE, "#,##0") & " bytes" & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                
                gblnStopProcessing = True
                GoTo GOST_File_CleanUp
                
            End If
        End If
        
        mcurMaxSize = mcurOrigLength * mlngRounds   ' Save file size for progress calcs
        
        mcurAccrued = 0@   ' Init accrued progress
        curFilePos1 = 0@   ' Reset pointers to first position in file
        curFilePos2 = 0@
        
        ' Open source file
        If Not .OpenReadOnly(strSource, hFile1) Then
            gblnStopProcessing = True
            GoTo GOST_File_CleanUp
        End If
        
        strTempFile = CreateTempFile()  ' Create a temp target file
        
        ' Open target file
        If Not .OpenReadWrite(strTempFile, hFile2) Then
            gblnStopProcessing = True
            GoTo GOST_File_CleanUp
        End If
        
        Do
            Erase abytData()                              ' empty data array
            lngBlockSize = GetBlockSize(curAmtLeft)       ' Process the source file in chunks
            curAmtLeft = curAmtLeft - CCur(lngBlockSize)  ' Adjust amount of file left to process
            ReDim abytData(lngBlockSize - 1)              ' Size receiving array
            
            If blnEncrypt Then
                lngLastBlock = lngBlockSize               ' Save size of last block of data
            End If
            
            ' Read source file
            If Not .API_ReadFile(hFile1, curFilePos1, abytData()) Then
                gblnStopProcessing = True
                Exit Do  ' exit Do..Loop
            End If
                       
            curFilePos1 = curFilePos1 + CCur(UBound(abytData) + 1)  ' Adjust pointers accordingly
            
            If blnEncrypt Then
        
                ' Encrypt data
                If Not EncryptData(abytData()) Then
                    gblnStopProcessing = True
                    Exit Do  ' exit Do..Loop
                End If
            
            Else
                ' Decrypt data
                If DecryptData(abytData()) Then
                
                    If curAmtLeft < 1 Then
                        ReDim Preserve abytData(lngLastBlock - 1)   ' Resize last data chunk
                    End If
                
                Else
                    gblnStopProcessing = True
                    Exit Do  ' exit Do..Loop
                End If
            
            End If
                
            ' Write to target file
            If Not .API_WriteFile(hFile2, curFilePos2, abytData()) Then
                gblnStopProcessing = True
                Exit Do  ' exit Do..Loop
            End If
                       
            curFilePos2 = curFilePos2 + CCur(UBound(abytData) + 1)    ' Adjust pointers accordingly
            
            ' If nothing else to read then leave
            DoEvents
            If curAmtLeft < 1 Then
                Exit Do  ' exit Do..Loop
            End If
            
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do   ' exit Do..Loop
            End If
    
        Loop
        
        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            GoTo GOST_File_CleanUp
        End If
    
        .API_CloseFile hFile1    ' Close all files opened by this routine
        .API_CloseFile hFile2
        
        ' Add important data to end of encrypted file
        If blnEncrypt Then
            
            CopyMemory abytSize(0), lngLastBlock, 4&         ' Save size of last block of data without padding [0-3]
            CopyMemory abytSize(4), mcurOrigLength, 8&       ' Save original file size [4-11]
            .CalcFileSize strTempFile, curFilePos2           ' Get size of target file
            MixAppendedData abytSize()                       ' Encrypt appended data
            
            ' Excessive DoEvents are to keep the write
            ' process in sync with the CPU speed
            DoEvents
            
            .OpenReadWrite strTempFile, hFile2               ' open target file
            .API_WriteFile hFile2, curFilePos2, abytSize()   ' Append to target file
            .API_CloseFile hFile2                            ' Close target file
            DoEvents
        Else
            ' Final decryption verification
            curAmtLeft = 0@                         ' Verify empty variable
            .CalcFileSize strTempFile, curAmtLeft   ' Get size of decrypted file
            
            ' Compare current file size with
            ' what it is supposed to be
            If (curAmtLeft <> mcurOrigLength) Then
            
                InfoMsg "This file cannot be Decrypted." & vbNewLine & vbNewLine & _
                        "Current size:  " & Format$(curAmtLeft, "#,##0") & " bytes" & vbNewLine & _
                        "Original size:  " & Format$(mcurOrigLength, "#,##0") & " bytes" & vbNewLine & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                
                gblnStopProcessing = True
                GoTo GOST_File_CleanUp
                
            End If

        End If
    End With
    
    ' Move temp file to destination location
    DoEvents
    On Error Resume Next
    DoEvents

    ' See if a new file is to be created after
    ' performing encryption or decryption
    If mblnCreateNewFile Then
        ' Create new file
        MoveFileEx strTempFile, strTarget, _
                   MOVEFILE_COPY_ALLOWED Or _
                   MOVEFILE_REPLACE_EXISTING
    Else
        ' Overwrite source file
        MoveFileEx strTempFile, strSource, _
                   MOVEFILE_COPY_ALLOWED Or _
                   MOVEFILE_REPLACE_EXISTING
    End If

    RaiseEvent CipherProgress(100)   ' Update progress bar

GOST_File_CleanUp:
    With objBigFiles
        .API_CloseFile hFile1   ' Close all files opened by this routine
        .API_CloseFile hFile2
    End With
    Set objBigFiles = Nothing   ' Free class object from memory
    
    ResetVariables    ' Reset class variables
    Erase abytData()  ' Always empty arrays when not needed
    Erase abytSize()
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GOST_File = False
    Else
        GOST_File = True
    End If
        
    On Error GoTo 0   ' Nullify this error trap
    Exit Function
    
GOST_File_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume GOST_File_CleanUp

End Function

' ***************************************************************************
' Routine:       GOST_String
'
' Description:   Function to encrypt/decrypt a string of data. 
' 
' Syntax:        abytData() = StrConv("abc", vbFromUnicode)  ' convert string data to byte array
'                objGOST.GOST_String abytData(), blnEncrypt  ' Encrypt/Decrypt data & return as byte array
' 
' Parameters:    abytData() - data string in byte array
'                blnEncrypt - [Optional] - TRUE - Encrypt incoming data
'                             FALSE - Decrypt incoming data
'                             Default=TRUE
'
' Returns:       TRUE - Successful completion
'                FALSE - An error occurred
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2004  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array for security purposes
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added call to MixAppendedData() routine to encrypt/decrypt
'              size information that is to be appended to original data.
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Split size testing as per Joe Sova's suggestion.
' ***************************************************************************
Public Function GOST_String(ByRef abytData() As Byte, _
                   Optional ByVal blnEncrypt As Boolean = True) As Boolean

    Dim curDataLength As Currency
    Dim abytSize(7)   As Byte
    
    Const ROUTINE_NAME As String = "GOST_String"

    On Error GoTo GOST_String_Error
    
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
    
    Erase abytSize()                                   ' Always start with empty arrays
    ResetVariables                                     ' Reset class variables
    mcurMaxSize = CCur(UBound(abytData)) * mlngRounds  ' Save file size for progress calcs
    mcurAccrued = 0@                                   ' Init accrued progress
    
    If Not EvaluateKey() Then
        GoTo GOST_String_CleanUp
    End If
    
    ' encrypt the data
    If blnEncrypt Then
        
        mcurOrigLength = CCur(UBound(abytData))         ' Get string length for progress display
        CopyMemory abytSize(0), mcurOrigLength, 8&      ' Copy original length to temp array
        MixAppendedData abytSize()                      ' Encrypt appended data
        
        If EncryptData(abytData()) Then
            
            curDataLength = UBound(abytData) + 1        ' data length after encryption
            ReDim Preserve abytData(curDataLength + 7)  ' Resize to hold original length
            
            ' Append original data length to end of encrypted data
            CopyMemory abytData(curDataLength), abytSize(0), 8&
            
        Else
            gblnStopProcessing = True
        End If
    
    Else
        ' Decrypt the data
        CopyMemory abytSize(0), abytData(UBound(abytData) - 7), 8&   ' capture the original data length
        ReDim Preserve abytData(UBound(abytData) - 8)                ' Adjust data length
        MixAppendedData abytSize()                                   ' Decrypt appended data
        CopyMemory mcurOrigLength, abytSize(0), 8&                   ' capture the original data length
        
        ' Must have data and cannot exceed size of textbox
        If (UBound(abytData) < 1) Or _
           (mcurOrigLength > MAX_STRING_LEN) Then
           
            InfoMsg "This data cannot be Decrypted, it is" & vbNewLine & _
                    "either not Encrypted, too large or corrupt." & vbNewLine & vbNewLine & _
                    "Original size:  " & Format$(mcurOrigLength, "#,##0") & " bytes" & vbNewLine & _
                    "Max allowed size:  " & Format$(MAX_STRING_LEN, "#,##0") & " bytes" & _
                    vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                
            gblnStopProcessing = True
            GoTo GOST_String_CleanUp
            
        End If
        
        ' Decrypt the data
        If DecryptData(abytData()) Then
                
            ReDim Preserve abytData(mcurOrigLength)  ' Resize data to exact length
            curDataLength = CCur(UBound(abytData))   ' data length after decryption
            
            ' Compare data length
            If curDataLength <> mcurOrigLength Then
                
                InfoMsg "Encrypted data string is corrupted.  Cannot Decrypt." & _
                        vbNewLine & vbNewLine & _
                        "Current size:  " & Format$(curDataLength, "#,##0") & vbNewLine & _
                        "Original size:  " & Format$(mcurOrigLength, "#,##0") & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
    
                gblnStopProcessing = True
            
            End If
        Else
            gblnStopProcessing = True
        End If
    
    End If
    
GOST_String_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GOST_String = False
    Else
        GOST_String = True
    End If
    
    ResetVariables    ' Reset class variables
    Erase abytSize()  ' Always empty arrays when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function
    
GOST_String_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume GOST_String_CleanUp

End Function


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Function EvaluateKey() As Boolean

    ' Called by Gost_File()
    '           Gost_String()
    
    Dim lngHold      As Long
    Dim lngLoop      As Long
    Dim lngTemp      As Long
    Dim lngIndex     As Long
    Dim lngPwdIdx    As Long
    Dim lngKeyLen    As Long
    Dim lngLeftWord  As Long
    Dim lngRightWord As Long
    Dim abytPwd()    As Byte

    Const ROUTINE_NAME As String = "EvaluateKey"

    On Error GoTo EvaluateKey_Error

    EvaluateKey = False  ' Preset flag
    
    If Not CBool(IsArrayInitialized(mabytPassword)) Then
        InfoMsg "Need a password before data can be encrypted/decrypted." & _
                vbNewLine & vbNewLine & _
                "Min length - " & MIN_PWD_LENGTH & vbNewLine & _
                "Max length - " & MAX_PWD_LENGTH & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If
    
    Erase abytPwd()          ' Always start with empty arrays
    Erase malngPrimaryKey()
    
    ' 21-Feb-2012 Changed size of primary key array
    '             from thirty-two to sixteen elements
    ReDim malngPrimaryKey(KEY_ARRAY)                     ' Size primary key array
    lngKeyLen = mlngKeyLength \ 8                        ' Ex:  16 bytes = 128 bits \ 8 bits
    abytPwd() = ExpandData(mabytPassword(), lngKeyLen)   ' Expand password key to desired length
    
    LoadWorkBoxes                   ' Load various work boxes
    lngHold = mlngRounds            ' 02-Nov-2009 Capture number of rounds for block encryption
    EncryptRounds = mlngKeyRounds   ' 02-Nov-2009 Temp reset number of rounds of block
                                    '             encryption (1-5 rounds is a good mix)
    lngPwdIdx = 0                   ' Set temp password array index
    
    ' Load constant primary key array.  Convert password byte array
    ' to long integers.  Multiple looping due to unknown key length.
    ' Key lengths will vary 16-128 bytes (128-1024 bits).
    For lngLoop = 1 To 2
        For lngIndex = 0 To (KEY_ARRAY - 1)
                
            CopyMemory lngTemp, abytPwd(lngPwdIdx Mod lngKeyLen), 4&                     ' Create long integer
            malngPrimaryKey(lngIndex) = UnsignedAdd(malngPrimaryKey(lngIndex), lngTemp)  ' Add data
            lngPwdIdx = lngPwdIdx + 4                                                    ' Increment password index
        
        Next lngIndex
    Next lngLoop
    
    ' 15-Mar-2009  Perform additional mixing so that
    '              all elements in primary key array
    '              will have a unique value
    For lngIndex = 0 To (KEY_ARRAY - 2) Step 2
    
        lngLeftWord = malngPrimaryKey(lngIndex)        ' Capture 2 long integers
        lngRightWord = malngPrimaryKey(lngIndex + 1)
        BlockEncrypt lngLeftWord, lngRightWord         ' Perform block encryption
        malngPrimaryKey(lngIndex) = lngLeftWord        ' Insert data back into array
        malngPrimaryKey(lngIndex + 1) = lngRightWord
        
    Next lngIndex
        
EvaluateKey_CleanUp:
    Erase abytPwd()           ' Always empty arrays when not needed
    EncryptRounds = lngHold   ' 02-Nov-2009 Reset number of rounds of block
                              '             encryption back to original value
    ' An error occurred or user
    ' opted to STOP processing
    If gblnStopProcessing Then
        EvaluateKey = False
    Else
        EvaluateKey = True
    End If

    On Error GoTo 0           ' Nullify this error trap
    Exit Function

EvaluateKey_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume EvaluateKey_CleanUp

End Function

Private Function EncryptData(ByRef abytData() As Byte) As Boolean

    ' Called by Gost_File()
    '           Gost_String()
    
    Dim lngOffSet     As Long
    Dim lngProgress   As Long
    Dim lngLeftWord   As Long
    Dim lngRightWord  As Long
    Dim lngDataLength As Long
    Dim abytInput(7)  As Byte
    Dim abytOutput(7) As Byte

    Const ROUTINE_NAME As String = "EncryptData"

    On Error GoTo EncryptData_Error

    lngDataLength = UBound(abytData) + 1   ' Get size of incoming array
    
    ' The encoded length may have to be padded out.
    ' This normally does not take place until the
    ' last chunk of data is to be processed.
    If (lngDataLength Mod DATA_BLOCK <> 0) Then
        lngDataLength = lngDataLength + DATA_BLOCK - (lngDataLength Mod DATA_BLOCK)
        ReDim Preserve abytData(lngDataLength - 1)
    End If
    
    ' 1-Jan-2009 Fixed lngOffSet range
    ' Encrypt the data
    For lngOffSet = 0 To (lngDataLength - 1) Step BLOCK_SIZE

        ' Transfer incoming data to a temp work array
        CopyMemory abytInput(0), abytData(lngOffSet), BLOCK_SIZE
            
        ' Convert bytes to two long integers
        GetWord lngLeftWord, abytInput(), 0&
        GetWord lngRightWord, abytInput(), 4&

        ' Encrypt block of data
        BlockEncrypt lngLeftWord, lngRightWord
        
        ' Convert two long integers to bytes
        PutWord lngLeftWord, abytOutput(), 0&
        PutWord lngRightWord, abytOutput(), 4&
            
        ' Transfer decrypted data back to original array
        CopyMemory abytData(lngOffSet), abytOutput(0), BLOCK_SIZE
        
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
    Next lngOffSet
    
    ' 19-Feb-2012 Update progress bar. Separate
    '             steps for easier debugging
    mcurAccrued = mcurAccrued + CCur(lngDataLength * mlngRounds)
    lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
    RaiseEvent CipherProgress(lngProgress)
            
EncryptData_CleanUp:
    DoEvents
    If gblnStopProcessing Then
        EncryptData = False   ' An error occurred or user opted to STOP processing
    Else
        EncryptData = True    ' Good Finish
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

EncryptData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume EncryptData_CleanUp

End Function

Private Function DecryptData(ByRef abytData() As Byte) As Boolean

    ' Called by Gost_File()
    '           Gost_String()
    
    Dim lngOffSet     As Long
    Dim lngProgress   As Long
    Dim lngLeftWord   As Long
    Dim lngRightWord  As Long
    Dim lngDataLength As Long
    Dim abytInput(7)  As Byte
    Dim abytOutput(7) As Byte

    Const ROUTINE_NAME As String = "DecryptData"

    On Error GoTo DecryptData_Error

    lngDataLength = UBound(abytData) + 1   ' Get size of incoming array

    ' 1-Jan-2009 Fixed lngOffSet range
    ' Decrypt the data
    For lngOffSet = 0 To (lngDataLength - 1) Step BLOCK_SIZE

        ' Transfer eight bytes to work array
        CopyMemory abytInput(0), abytData(lngOffSet), BLOCK_SIZE
                
        ' Convert bytes to two long integers
        GetWord lngLeftWord, abytInput(), 0&
        GetWord lngRightWord, abytInput(), 4&

        ' Decrypt block of data
        BlockDecrypt lngLeftWord, lngRightWord
        
        ' Convert two long integers to bytes
        PutWord lngLeftWord, abytOutput(), 0&
        PutWord lngRightWord, abytOutput(), 4&
            
        ' Transfer decrypted data back to original array
        CopyMemory abytData(lngOffSet), abytOutput(0), BLOCK_SIZE
        
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
    Next lngOffSet
        
    ' 19-Feb-2012 Update progress bar. Separate
    '             steps for easier debugging
    mcurAccrued = mcurAccrued + CCur(lngDataLength * mlngRounds)
    lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
    RaiseEvent CipherProgress(lngProgress)
            
DecryptData_CleanUp:
    DoEvents
    If gblnStopProcessing Then
        DecryptData = False   ' An error occurred or user opted to STOP processing
    Else
        DecryptData = True    ' Good Finish
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

DecryptData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume DecryptData_CleanUp

End Function

Private Static Sub BlockEncrypt(ByRef lngLeftWord As Long, _
                                ByRef lngRightWord As Long)

    ' Called by EvaluateKey()
    '           EncryptData()
    
    Dim lngLoop  As Long
    Dim lngIndex As Long
    Dim lngRound As Long

    ' 21-Feb-2012 Decreased size of primary key array
    '             from thirty-two to sixteen elements
    ' 02-Nov-2009 Perform encryption for required
    '             number of iterations
    For lngRound = 1 To mlngRounds
    
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
        ' For the first forty-eight (48 = 16 * 3) rounds,
        ' the primary key is used cyclically in ascending
        ' order.  For the last sixteen (16) rounds, the
        ' primary key is used in descending order.
        '
        ' The 3 outer loops are standard.
        For lngLoop = 1 To 3
            
            ' This inner loop has been added so
            ' that all elements of the new expanded
            ' primary key will be accessed
            For lngIndex = 0 To (KEY_ARRAY - 8) Step 8
                
                lngRightWord = lngRightWord Xor MixBits(lngLeftWord, malngPrimaryKey(lngIndex))
                lngLeftWord = lngLeftWord Xor MixBits(lngRightWord, malngPrimaryKey(lngIndex + 1))
                lngRightWord = lngRightWord Xor MixBits(lngLeftWord, malngPrimaryKey(lngIndex + 2))
                lngLeftWord = lngLeftWord Xor MixBits(lngRightWord, malngPrimaryKey(lngIndex + 3))
                lngRightWord = lngRightWord Xor MixBits(lngLeftWord, malngPrimaryKey(lngIndex + 4))
                lngLeftWord = lngLeftWord Xor MixBits(lngRightWord, malngPrimaryKey(lngIndex + 5))
                lngRightWord = lngRightWord Xor MixBits(lngLeftWord, malngPrimaryKey(lngIndex + 6))
                lngLeftWord = lngLeftWord Xor MixBits(lngRightWord, malngPrimaryKey(lngIndex + 7))
    
            Next lngIndex
    
        Next lngLoop
        
        ' This outer loop has been added so
        ' that all elements of the new expanded
        ' primary key will be accessed
        For lngIndex = (KEY_ARRAY - 8) To 0 Step -8
            
            lngRightWord = lngRightWord Xor MixBits(lngLeftWord, malngPrimaryKey(lngIndex + 7))
            lngLeftWord = lngLeftWord Xor MixBits(lngRightWord, malngPrimaryKey(lngIndex + 6))
            lngRightWord = lngRightWord Xor MixBits(lngLeftWord, malngPrimaryKey(lngIndex + 5))
            lngLeftWord = lngLeftWord Xor MixBits(lngRightWord, malngPrimaryKey(lngIndex + 4))
            lngRightWord = lngRightWord Xor MixBits(lngLeftWord, malngPrimaryKey(lngIndex + 3))
            lngLeftWord = lngLeftWord Xor MixBits(lngRightWord, malngPrimaryKey(lngIndex + 2))
            lngRightWord = lngRightWord Xor MixBits(lngLeftWord, malngPrimaryKey(lngIndex + 1))
            lngLeftWord = lngLeftWord Xor MixBits(lngRightWord, malngPrimaryKey(lngIndex))
    
        Next lngIndex
    Next lngRound
    
End Sub

Private Static Sub BlockDecrypt(ByRef lngLeftWord As Long, _
                                ByRef lngRightWord As Long)

    ' Called by DecryptData()
    
    Dim lngLoop  As Long
    Dim lngIndex As Long
    Dim lngRound As Long

    ' 21-Feb-2012 Decreased size of primary key array
    '             from thirty-two to sixteen elements
    ' 02-Nov-2009 Perform decryption for required
    '             number of iterations
    For lngRound = 1 To mlngRounds
    
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
        ' For the first sixteen (16) rounds, the primary key
        ' is used cyclically in ascending order.  For the last
        ' forty-eight (48 = 16 * 3) rounds, the primary key is
        ' used in descending order.
        '
        ' This outer loop has been added so that all elements
        ' of the new expanded primary key will be accessed
        For lngIndex = 0 To (KEY_ARRAY - 8) Step 8
            
            lngLeftWord = lngLeftWord Xor MixBits(lngRightWord, malngPrimaryKey(lngIndex))
            lngRightWord = lngRightWord Xor MixBits(lngLeftWord, malngPrimaryKey(lngIndex + 1))
            lngLeftWord = lngLeftWord Xor MixBits(lngRightWord, malngPrimaryKey(lngIndex + 2))
            lngRightWord = lngRightWord Xor MixBits(lngLeftWord, malngPrimaryKey(lngIndex + 3))
            lngLeftWord = lngLeftWord Xor MixBits(lngRightWord, malngPrimaryKey(lngIndex + 4))
            lngRightWord = lngRightWord Xor MixBits(lngLeftWord, malngPrimaryKey(lngIndex + 5))
            lngLeftWord = lngLeftWord Xor MixBits(lngRightWord, malngPrimaryKey(lngIndex + 6))
            lngRightWord = lngRightWord Xor MixBits(lngLeftWord, malngPrimaryKey(lngIndex + 7))
    
        Next lngIndex
    
        ' The 3 outer loops are standard.
        For lngLoop = 1 To 3
            
            ' This inner loop has been added so
            ' that all elements of the new expanded
            ' primary key will be accessed
            For lngIndex = (KEY_ARRAY - 8) To 0 Step -8
                
                lngLeftWord = lngLeftWord Xor MixBits(lngRightWord, malngPrimaryKey(lngIndex + 7))
                lngRightWord = lngRightWord Xor MixBits(lngLeftWord, malngPrimaryKey(lngIndex + 6))
                lngLeftWord = lngLeftWord Xor MixBits(lngRightWord, malngPrimaryKey(lngIndex + 5))
                lngRightWord = lngRightWord Xor MixBits(lngLeftWord, malngPrimaryKey(lngIndex + 4))
                lngLeftWord = lngLeftWord Xor MixBits(lngRightWord, malngPrimaryKey(lngIndex + 3))
                lngRightWord = lngRightWord Xor MixBits(lngLeftWord, malngPrimaryKey(lngIndex + 2))
                lngLeftWord = lngLeftWord Xor MixBits(lngRightWord, malngPrimaryKey(lngIndex + 1))
                lngRightWord = lngRightWord Xor MixBits(lngLeftWord, malngPrimaryKey(lngIndex))
    
            Next lngIndex
        Next lngLoop
    Next lngRound
    
End Sub

Private Static Function MixBits(ByVal lngWord As Long, _
                                ByVal lngPrimaryKey As Long) As Long

    ' Called by BlockEncrypt()
    '           BlockDecrypt()
    
    Dim lngTemp     As Long
    Dim abytKey(4)  As Byte
    Dim abytTemp(4) As Byte

    lngTemp = UnsignedAdd(lngWord, lngPrimaryKey)   ' Add word value to primary key constant
    CopyMemory abytTemp(0), lngTemp, 4&             ' Convert temp value to byte array
 
    ' Key-dependant substitutions.  Use
    ' temp array values as indices to
    ' load key array.
    abytKey(0) = mabytK21(abytTemp(3))
    abytKey(1) = mabytK43(abytTemp(2))
    abytKey(2) = mabytK65(abytTemp(1))
    abytKey(3) = mabytK87(abytTemp(0))

    ' Mix each byte
    abytTemp(0) = ((abytKey(3) And 31) * 8) Or ((abytKey(2) And 224) \ 32)
    abytTemp(1) = ((abytKey(0) And 31) * 8) Or ((abytKey(3) And 224) \ 32)
    abytTemp(2) = ((abytKey(1) And 31) * 8) Or ((abytKey(0) And 224) \ 32)
    abytTemp(3) = ((abytKey(2) And 31) * 8) Or ((abytKey(1) And 224) \ 32)

    CopyMemory MixBits, abytTemp(0), 4&   ' Convert byte array to long integer
     
End Function

' ***************************************************************************
' Routine:       UnsignedAdd
'
' Description:   Function to add two unsigned numbers together as in C.
'                Overflows are ignored!
'
' Parameters:    dblValue1 - Value of A
'                dblValue2 - Value of B
'
' Returns:       Calculated value
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Apr-2005  Pablo Mariano Ronchi  pmronchi@yahoo.com.ar
'              Routine created
' 19-Dec-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified variable names
' ***************************************************************************
Private Function UnsignedAdd(ByVal dblValue1 As Double, _
                             ByVal dblValue2 As Double) As Long
                             
    Dim dblTemp As Double

    dblTemp = dblValue1 + dblValue2

    If dblTemp < MIN_LONG Then
        UnsignedAdd = CLng(GB_4 + dblTemp)
    Else
        If dblTemp > MAX_LONG Then
            UnsignedAdd = CLng(dblTemp - GB_4)
        Else
            UnsignedAdd = CLng(dblTemp)
        End If
    End If

End Function

Private Sub GetWord(ByRef lngValue As Long, _
                    ByRef abytData() As Byte, _
                    ByVal lngIndex As Long)
                    
    ' Called by EncryptData()
    '           DecryptData()
    '
    ' Convert byte array to long integer in
    ' Big-Endian format
    
    Dim abytTemp(4) As Byte
    
    abytTemp(3) = abytData(lngIndex)       ' Invert the bytes
    abytTemp(2) = abytData(lngIndex + 1)
    abytTemp(1) = abytData(lngIndex + 2)
    abytTemp(0) = abytData(lngIndex + 3)
    
    CopyMemory lngValue, abytTemp(0), 4&   ' Convert byte array to long integer
  
End Sub

Private Sub PutWord(ByVal lngValue As Long, _
                    ByRef abytData() As Byte, _
                    ByVal lngIndex As Long)

    ' Called by EncryptData()
    '           DecryptData()
    '
    ' Convert long integer to byte array in
    ' Little-Endian format
        
    Dim abytTemp(4) As Byte
    
    CopyMemory abytTemp(0), lngValue, 4&   ' Convert long integer to byte array
    
    abytData(lngIndex) = abytTemp(3)       ' Invert the bytes
    abytData(lngIndex + 1) = abytTemp(2)
    abytData(lngIndex + 2) = abytTemp(1)
    abytData(lngIndex + 3) = abytTemp(0)
    
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
    
    RaiseEvent CurrentBlockSize(GetBlockSize)
    
End Function

Private Sub LoadWorkBoxes()

    ' Called by EvaluateKey()
    
    Dim avntData   As Variant
    Dim lngIdx     As Long
    Dim lngIndex   As Long
    Dim abytSBox() As Byte   ' Key dependant s-boxes

    Const ROUTINE_NAME As String = "LoadWorkBoxes"

    On Error GoTo LoadWorkBoxes_Error

    avntData = Empty       ' Start with empty variants
    ReDim avntData(8)      ' Size arrays
    ReDim abytSBox(8, 16)
    ReDim mabytK21(MAX_BYTE)
    ReDim mabytK43(MAX_BYTE)
    ReDim mabytK65(MAX_BYTE)
    ReDim mabytK87(MAX_BYTE)
    
    ' 02-Nov-2009 For added security, use multiple sets of
    '             values based on user defined key length.
    '
    ' S-Box values are 0-15 in random order with no duplicates.  See
    ' GOST_Tbl.txt file for additional S-Box tables.  See application
    ' BuildTables.vbp to create different CASE statements.
    '
    ' The S-boxes accept a four-bit input and produce a four-bit output.
    ' The S-box substitution in the round function consists of eight
    ' 4×4 S-boxes. The S-boxes are implementation-dependent - parties
    ' that want to secure their communications using GOST must be using
    ' the same S-boxes.  For extra security, the S-boxes can be kept secret.
    ' In the original standard where GOST was specified, no S-boxes were
    ' given, but they were to be supplied somehow. This led to speculation
    ' that organizations the government wished to spy on were given weak
    ' S-boxes.  One GOST chip manufacturer reported that he generated
    ' S-boxes himself using a pseudorandom number generator (Schneier, 1996).
    '
    ' Each data line:
    '    1.  Consists of values (0-15) with no duplicates
    '    2.  Mixed 100-200 iterations using Knuth Shuffle
    Select Case mlngKeyLength
           Case 128, 448
                avntData(0) = Array(8, 11, 12, 2, 15, 14, 3, 13, 0, 7, 6, 1, 10, 9, 4, 5)
                avntData(1) = Array(1, 4, 11, 10, 15, 3, 12, 8, 14, 0, 13, 5, 6, 9, 7, 2)
                avntData(2) = Array(13, 1, 7, 4, 3, 12, 2, 5, 0, 11, 10, 6, 9, 14, 15, 8)
                avntData(3) = Array(15, 14, 10, 12, 2, 5, 4, 6, 7, 0, 11, 9, 13, 3, 1, 8)
                avntData(4) = Array(1, 7, 6, 0, 15, 14, 2, 3, 10, 11, 8, 9, 4, 13, 5, 12)
                avntData(5) = Array(1, 15, 11, 13, 5, 2, 9, 8, 6, 0, 14, 3, 12, 7, 4, 10)
                avntData(6) = Array(11, 8, 7, 2, 1, 10, 5, 14, 15, 12, 3, 6, 9, 4, 13, 0)
                avntData(7) = Array(2, 5, 0, 10, 3, 6, 13, 1, 4, 14, 12, 15, 9, 7, 11, 8)
           Case 160, 512
                avntData(0) = Array(2, 5, 10, 15, 12, 7, 11, 1, 8, 3, 13, 14, 4, 9, 6, 0)
                avntData(1) = Array(11, 5, 2, 13, 8, 12, 9, 10, 7, 3, 14, 0, 6, 1, 15, 4)
                avntData(2) = Array(2, 7, 3, 11, 15, 1, 8, 14, 13, 6, 9, 5, 12, 10, 0, 4)
                avntData(3) = Array(1, 6, 2, 9, 4, 14, 0, 12, 10, 7, 3, 11, 13, 5, 15, 8)
                avntData(4) = Array(10, 8, 7, 2, 15, 0, 12, 1, 5, 6, 3, 13, 4, 14, 9, 11)
                avntData(5) = Array(14, 3, 5, 1, 12, 7, 8, 6, 2, 4, 0, 15, 13, 9, 11, 10)
                avntData(6) = Array(7, 11, 10, 6, 14, 5, 3, 4, 12, 2, 15, 0, 9, 13, 8, 1)
                avntData(7) = Array(14, 1, 6, 2, 5, 7, 0, 4, 13, 8, 10, 12, 11, 9, 15, 3)
           Case 192, 576
                avntData(0) = Array(14, 7, 13, 2, 1, 3, 0, 8, 9, 4, 10, 11, 15, 5, 6, 12)
                avntData(1) = Array(12, 3, 11, 13, 10, 14, 5, 8, 9, 0, 7, 6, 15, 1, 4, 2)
                avntData(2) = Array(9, 6, 15, 1, 8, 12, 0, 5, 4, 3, 2, 14, 7, 11, 13, 10)
                avntData(3) = Array(3, 7, 12, 1, 2, 6, 4, 14, 8, 9, 13, 10, 11, 15, 5, 0)
                avntData(4) = Array(15, 12, 1, 8, 0, 13, 14, 9, 10, 3, 6, 11, 5, 7, 4, 2)
                avntData(5) = Array(4, 5, 6, 0, 3, 15, 2, 8, 14, 10, 1, 11, 12, 7, 9, 13)
                avntData(6) = Array(6, 13, 1, 2, 10, 3, 14, 7, 12, 9, 4, 5, 0, 11, 15, 8)
                avntData(7) = Array(0, 9, 7, 3, 1, 8, 14, 13, 15, 6, 12, 2, 11, 10, 4, 5)
           Case 224, 640
                avntData(0) = Array(12, 13, 2, 14, 11, 0, 9, 3, 1, 4, 5, 6, 7, 10, 15, 8)
                avntData(1) = Array(9, 15, 4, 8, 14, 11, 2, 7, 1, 6, 0, 12, 13, 3, 5, 10)
                avntData(2) = Array(7, 15, 6, 1, 2, 4, 13, 14, 0, 11, 10, 8, 9, 12, 5, 3)
                avntData(3) = Array(15, 11, 2, 3, 1, 0, 5, 13, 14, 10, 7, 4, 8, 12, 9, 6)
                avntData(4) = Array(2, 12, 11, 14, 7, 3, 15, 0, 6, 8, 4, 13, 1, 10, 9, 5)
                avntData(5) = Array(14, 8, 0, 10, 13, 11, 4, 7, 2, 1, 9, 5, 15, 12, 6, 3)
                avntData(6) = Array(2, 6, 13, 11, 0, 3, 8, 9, 12, 10, 4, 14, 5, 7, 15, 1)
                avntData(7) = Array(9, 14, 7, 0, 1, 3, 8, 13, 10, 2, 15, 4, 12, 6, 11, 5)
           Case 256, 704
                avntData(0) = Array(8, 9, 13, 6, 12, 15, 1, 11, 14, 0, 5, 4, 3, 2, 7, 10)
                avntData(1) = Array(12, 0, 13, 4, 14, 8, 7, 1, 10, 2, 6, 9, 11, 3, 15, 5)
                avntData(2) = Array(10, 13, 7, 12, 15, 3, 9, 14, 1, 5, 2, 0, 6, 11, 8, 4)
                avntData(3) = Array(4, 8, 14, 7, 3, 6, 1, 5, 13, 15, 10, 9, 2, 0, 11, 12)
                avntData(4) = Array(5, 15, 13, 14, 6, 0, 12, 2, 1, 7, 10, 9, 3, 4, 11, 8)
                avntData(5) = Array(2, 4, 13, 1, 3, 9, 12, 15, 8, 11, 14, 10, 6, 5, 0, 7)
                avntData(6) = Array(2, 9, 11, 12, 4, 6, 0, 15, 10, 8, 1, 5, 14, 3, 13, 7)
                avntData(7) = Array(11, 12, 9, 13, 1, 7, 15, 10, 4, 0, 5, 3, 8, 2, 14, 6)
           Case 288, 768
                avntData(0) = Array(9, 10, 0, 2, 4, 11, 1, 5, 13, 7, 3, 8, 6, 12, 14, 15)
                avntData(1) = Array(15, 1, 8, 11, 10, 0, 12, 7, 2, 4, 6, 3, 5, 14, 9, 13)
                avntData(2) = Array(9, 5, 0, 11, 6, 2, 14, 15, 10, 12, 8, 7, 3, 4, 13, 1)
                avntData(3) = Array(11, 5, 10, 14, 7, 13, 12, 8, 15, 0, 1, 3, 2, 9, 6, 4)
                avntData(4) = Array(8, 13, 15, 0, 11, 14, 2, 12, 5, 3, 4, 9, 7, 6, 1, 10)
                avntData(5) = Array(11, 8, 12, 9, 13, 14, 10, 2, 5, 3, 0, 15, 4, 7, 6, 1)
                avntData(6) = Array(15, 2, 4, 5, 12, 10, 6, 8, 0, 14, 1, 3, 9, 11, 7, 13)
                avntData(7) = Array(0, 8, 9, 15, 4, 13, 2, 12, 10, 6, 14, 1, 5, 11, 3, 7)
           Case 320, 832
                avntData(0) = Array(3, 14, 11, 8, 12, 9, 13, 4, 10, 5, 15, 1, 6, 0, 2, 7)
                avntData(1) = Array(12, 6, 3, 7, 0, 2, 10, 9, 4, 11, 14, 1, 5, 13, 15, 8)
                avntData(2) = Array(1, 3, 11, 8, 12, 15, 13, 6, 5, 0, 14, 10, 4, 7, 2, 9)
                avntData(3) = Array(10, 11, 5, 8, 6, 1, 0, 14, 4, 9, 3, 12, 13, 7, 15, 2)
                avntData(4) = Array(4, 12, 11, 0, 14, 9, 10, 7, 2, 13, 3, 15, 5, 8, 1, 6)
                avntData(5) = Array(10, 8, 14, 2, 13, 7, 11, 1, 5, 0, 12, 15, 3, 6, 9, 4)
                avntData(6) = Array(4, 10, 8, 1, 2, 7, 15, 13, 14, 5, 11, 6, 3, 0, 12, 9)
                avntData(7) = Array(10, 4, 3, 11, 15, 9, 1, 6, 12, 5, 8, 0, 2, 7, 14, 13)
           Case 352, 896
                avntData(0) = Array(5, 8, 9, 14, 10, 6, 7, 13, 4, 0, 3, 11, 1, 12, 2, 15)
                avntData(1) = Array(4, 12, 8, 10, 14, 5, 9, 1, 13, 2, 11, 15, 0, 3, 7, 6)
                avntData(2) = Array(13, 3, 8, 12, 11, 15, 9, 1, 2, 0, 4, 7, 14, 5, 10, 6)
                avntData(3) = Array(12, 5, 11, 1, 10, 0, 7, 4, 3, 8, 14, 9, 13, 6, 15, 2)
                avntData(4) = Array(14, 12, 15, 7, 2, 13, 8, 10, 1, 9, 4, 6, 3, 5, 0, 11)
                avntData(5) = Array(5, 8, 9, 14, 7, 13, 4, 0, 3, 15, 2, 11, 6, 12, 10, 1)
                avntData(6) = Array(0, 7, 2, 4, 9, 14, 13, 15, 6, 10, 12, 5, 3, 1, 8, 11)
                avntData(7) = Array(6, 14, 10, 4, 15, 3, 0, 13, 8, 2, 1, 11, 12, 7, 5, 9)
           Case 384, 960
                avntData(0) = Array(8, 3, 2, 10, 12, 0, 6, 13, 14, 9, 1, 15, 11, 7, 5, 4)
                avntData(1) = Array(1, 6, 14, 10, 4, 11, 8, 9, 2, 5, 12, 13, 3, 15, 7, 0)
                avntData(2) = Array(8, 14, 4, 2, 12, 6, 0, 10, 13, 15, 7, 11, 5, 1, 9, 3)
                avntData(3) = Array(5, 13, 2, 11, 15, 1, 4, 7, 8, 3, 14, 6, 10, 12, 9, 0)
                avntData(4) = Array(13, 11, 12, 7, 14, 6, 3, 15, 1, 0, 8, 5, 10, 4, 2, 9)
                avntData(5) = Array(11, 12, 1, 15, 14, 2, 3, 8, 10, 9, 7, 5, 4, 6, 0, 13)
                avntData(6) = Array(14, 0, 5, 3, 1, 13, 9, 6, 8, 11, 15, 2, 12, 7, 4, 10)
                avntData(7) = Array(9, 1, 2, 7, 5, 11, 13, 15, 3, 0, 12, 8, 10, 6, 4, 14)
           Case 416, 1024
                avntData(0) = Array(4, 0, 15, 5, 3, 6, 1, 7, 2, 14, 13, 9, 11, 12, 8, 10)
                avntData(1) = Array(11, 0, 7, 6, 8, 5, 14, 9, 10, 12, 4, 3, 2, 15, 1, 13)
                avntData(2) = Array(4, 7, 9, 11, 15, 8, 10, 3, 6, 12, 5, 1, 13, 14, 0, 2)
                avntData(3) = Array(14, 13, 5, 0, 9, 6, 11, 7, 3, 4, 1, 8, 10, 2, 15, 12)
                avntData(4) = Array(6, 5, 10, 2, 4, 0, 8, 1, 9, 11, 7, 12, 3, 13, 14, 15)
                avntData(5) = Array(6, 5, 9, 7, 0, 12, 3, 10, 15, 13, 8, 2, 4, 1, 14, 11)
                avntData(6) = Array(0, 6, 2, 11, 13, 15, 14, 5, 8, 1, 12, 4, 7, 9, 3, 10)
                avntData(7) = Array(2, 11, 13, 0, 6, 7, 15, 1, 14, 4, 3, 5, 12, 8, 10, 9)
    End Select
    
    ' Load S-Box array
    For lngIdx = 0 To 7
        For lngIndex = 0 To 15
            abytSBox(lngIdx, lngIndex) = CByte(avntData(lngIdx)(lngIndex))
        Next lngIndex
    Next lngIdx

    ' Calculate substitution arrays
    For lngIdx = 0 To (MAX_BYTE - 1)

        mabytK87(lngIdx) = w8Shift(abytSBox(7, (w8Shift(lngIdx, -4))), 4) Or _
                                   abytSBox(6, lngIdx And &HF)
        mabytK65(lngIdx) = w8Shift(abytSBox(5, (w8Shift(lngIdx, -4))), 4) Or _
                                   abytSBox(4, lngIdx And &HF)
        mabytK43(lngIdx) = w8Shift(abytSBox(3, (w8Shift(lngIdx, -4))), 4) Or _
                                   abytSBox(2, lngIdx And &HF)
        mabytK21(lngIdx) = w8Shift(abytSBox(1, (w8Shift(lngIdx, -4))), 4) Or _
                                   abytSBox(0, lngIdx And &HF)
    Next lngIdx
    
LoadWorkBoxes_CleanUp:
    avntData = Empty    ' Always empty variants when not needed
    On Error GoTo 0     ' Nullify this error trap
    Exit Sub

LoadWorkBoxes_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume LoadWorkBoxes_CleanUp

End Sub

' ***************************************************************************
' Routine:       w8Shift (Byte shifting)
'
' Description:   Shifts the bits to the right or left the specified number
'                of positions and returns the new value.  Bits "falling off"
'                the edge do not wrap around on the opposite side.  Some
'                common languages like C/C++ or Java have an operator for
'                this job:  ">>" or "<<".
'
' Parameters:    bytValue    - numeric value to be evaluated
'                intBitShift - number of bit positions
'                              Positive value = left shift
'                              Negative value = right shift
'
' Returns:       Reformatted value
'
'                Byte    Binary
' Original:        60   00111100
'   Left 3:       224   11100000
'  Right 3:         7   00000111
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 12-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 01-Feb-2009  Kenneth Ives  kenaso@tx.rr.com
'              Corrected logic in testing constant arrays
' ***************************************************************************
Private Function w8Shift(ByVal bytValue As Byte, _
                         ByVal intBitShift As Integer) As Byte
                               
    ' Called by LoadWorkBoxes()
    
    ' Test amount of bit shifting
    If intBitShift = 0 Then
        w8Shift = bytValue   ' Nothing to do, return input value
    ElseIf intBitShift > 7 Then
        w8Shift = 0          ' return zero too many bit shift positions
    
    ElseIf intBitShift > 0 Then
        ' Positive bit shift means shift left
        If intBitShift = 7 Then
        
            ' Test for non-zero
            If bytValue And 1 Then
                bytValue = &H80
            Else
                bytValue = 0
            End If
        Else
            If Not mblnLoaded Then
                LoadShiftArrays   ' See if arrays are already loaded
            End If
        
            bytValue = ((bytValue And mabytValue(7 - intBitShift)) * mabytPower2(intBitShift))
        End If
                
        w8Shift = bytValue   ' Return new value
    
    ElseIf intBitShift < 0 Then
        ' Negative bit shift means shift right
        '
        ' Make bit shift value positive
        intBitShift = Abs(intBitShift)
        
        If intBitShift = 7 Then
            
            ' Test for non-zero
            If bytValue And &H80 Then
                bytValue = 1
            Else
                bytValue = 0
            End If
        Else
            If Not mblnLoaded Then
                LoadShiftArrays   ' See if arrays are already loaded
            End If
        
            bytValue = bytValue \ mabytPower2(intBitShift)
        End If
    
        w8Shift = bytValue   ' Return new value
    
    End If

End Function

Private Sub LoadShiftArrays()

    ' Called by w8Shift()
    
    On Error Resume Next
    
    mblnLoaded = False   ' Verify flag is FALSE
    
    DoEvents
    If IsArrayInitialized(mabytPower2()) Then
        
        ' See if arrays are already loaded
        If mabytPower2(0) = 1 And mabytValue(0) = 1 Then
            mblnLoaded = True      ' Set flag to TRUE
        Else
            ' Load power 2 array
            mabytPower2(0) = 1     ' 00000001
            mabytPower2(1) = 2     ' 00000010
            mabytPower2(2) = 4     ' 00000100
            mabytPower2(3) = 8     ' 00001000
            mabytPower2(4) = 16    ' 00010000
            mabytPower2(5) = 32    ' 00100000
            mabytPower2(6) = 64    ' 01000000
            mabytPower2(7) = 128   ' 10000000
            
            ' Load byte value array
            mabytValue(0) = 1      ' 00000001
            mabytValue(1) = 3      ' 00000011
            mabytValue(2) = 7      ' 00000111
            mabytValue(3) = 15     ' 00001111
            mabytValue(4) = 31     ' 00011111
            mabytValue(5) = 63     ' 00111111
            mabytValue(6) = 127    ' 01111111
            mabytValue(7) = 255    ' 11111111
            
            mblnLoaded = True      ' Set flag to TRUE
        End If
    
    Else
        ' Array has not been initialized
        ' therefore load both arrays
        Erase mabytPower2()    ' Always start with empty arrays
        Erase mabytValue()
    
        ' Load power 2 array
        mabytPower2(0) = 1     ' 00000001
        mabytPower2(1) = 2     ' 00000010
        mabytPower2(2) = 4     ' 00000100
        mabytPower2(3) = 8     ' 00001000
        mabytPower2(4) = 16    ' 00010000
        mabytPower2(5) = 32    ' 00100000
        mabytPower2(6) = 64    ' 01000000
        mabytPower2(7) = 128   ' 10000000
        
        ' Load byte value array
        mabytValue(0) = 1      ' 00000001
        mabytValue(1) = 3      ' 00000011
        mabytValue(2) = 7      ' 00000111
        mabytValue(3) = 15     ' 00001111
        mabytValue(4) = 31     ' 00011111
        mabytValue(5) = 63     ' 00111111
        mabytValue(6) = 127    ' 01111111
        mabytValue(7) = 255    ' 11111111
        
        mblnLoaded = True      ' Set flag to TRUE
    End If
    
    On Error GoTo 0   ' Nullify this error trap

End Sub

Private Sub ResetVariables()
    
    ' Always empty arrays and variants
    ' when not needed.  Saves resources.
    Erase mabytK21()
    Erase mabytK43()
    Erase mabytK65()
    Erase mabytK87()
    Erase malngPrimaryKey()
        
    mcurOrigLength = 0@
    mcurCurrByteCnt = 0@
    
End Sub

Private Sub Class_Initialize()

    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objGost = new cGost
                
    ResetVariables
    mblnLoaded = False     ' Preset flag to FALSE
    Erase mabytPower2()    ' Always start with empty arrays
    Erase mabytValue()
    Erase mabytPassword()
    
    ' Set property values
    EncryptRounds = 1      ' Number of rounds of encryption
    PrimaryKeyRounds = 1   ' Number of rounds to mix primary key array
    KeyLength = 128        ' Minimum key length
    CreateNewFile = True   ' Default is to create a new output file
    
End Sub

Private Sub Class_Terminate()
    
    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objGost = Nothing
            
    ResetVariables
    Erase mabytPower2()    ' Empty arrays when not needed
    Erase mabytValue()
    Erase mabytPassword()
    
End Sub


