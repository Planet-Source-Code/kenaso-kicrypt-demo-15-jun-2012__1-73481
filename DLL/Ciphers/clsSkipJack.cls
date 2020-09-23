VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSkipjack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Project:       Skipjack Encryption/Decryption
'
' Module:        clsSkipjack.cls
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
' Description:   Skipjack is a recent algorithm that was developed in 1987
'                and put into service in 1993.  It is a formerly secret NSA
'                (National Security Agency) encryption algorithm that was
'                declassified on June 24, 1998.  Skipjack is a
'                representative of a family of encryption algorithms
'                developed in 1980 as  part of the NSA suite of "Type I"
'                algorithms, which are  suitable for protecting all levels
'                of classified data.   Type I algorithms are typically
'                extremely secure and are  usually classified as secret.
'                Skipjack was used to encrypt  sensitive, but not
'                classified, government data. It was  implemented in two
'                government encryption devices:  the Clipper chip and
'                Fortezza PC card.  These devices have many uses and are
'                widely employed by agencies such as the FBI and NSA. They
'                provide a high level of security for sensitive
'                communications while enabling the interception of
'                telecommunications by law  enforcement officials for such
'                things as criminal  investigations.  For example, Clipper
'                chips can be used to  provide secure telephone
'                transmissions and Fortezza cards can  be used to encrypt
'                such things as e-mail and network traffic.  The key
'                characteristic of both devices is that they were designed
'                with "back doors" that allow government agents to monitor
'                encrypted transmissions given the proper authority.  This
'                feature is covered under the Escrowed Encryption Standard,
'                of which Skipjack is a part. It is implemented via a
'                mechanism called a LEAF (Law Enforcement Access Field). It
'                is important to note that Skipjack itself has nothing to
'                do with this "back door" functionality!  Skipjack is
'                entirely separate from the LEAF in the Clipper and
'                Fortezza products and is not affected in any way by its
'                presence.
'
'                Skipjack has been extensively cryptanalyzed, and has no
'                weaknesses. There are no known shortcut attacks that can
'                break Skipjack. However, the small key size makes this
'                algorithm inferior to the newer candidate algorithms for
'                the AES (Advanced Encryption Standard) competition being
'                held by NIST (National Institute of Standards and
'                Technology).  The NIST is a federal technology agency that
'                develops and promotes measurement, standards, and
'                technology.  Despite this shortcoming, Skipjack still
'                provides very strong security and it should be many years
'                before the algorithm is broken by a brute force attack.
'                Note that it was declassified in order to provide a
'                software implementation of Fortezza enabled applications.
'                The NSA does not intend  for the algorithm to be a
'                candidate for the AES.  Like  Triple DES, Skipjack is an
'                interim solution to be used until  the final AES is
'                completed and widely implemented. It offers  a safe
'                alternative to DES without having to rely on the AES.
'
'                In 1993 an Interim Report was released that gave a
'                thorough analysis of Skipjack and discussed issues
'                relating to Skipjack's integration in the Clipper and
'                Fortezza hardware. Although this report is many years old,
'                it provides some interesting information and a good
'                perspective on what it would take to break Skipjack with a
'                brute force attack. While reading the document, keep in
'                mind that the algorithm was still a closely guarded secret
'                at the time the report was written. The Final Report
'                mentioned in this document was never written, so this
'                document is the closest thing there is to an official
'                analysis of the Skipjack algorithm.
'
'                Skipjack encrypts and decrypts data in 64-bit blocks,
'                normally using an 80-bit key. It takes a 64-bit block of
'                plaintext as input and outputs a 64-bit block of
'                ciphertext. Skipjack has 32 rounds, meaning the main
'                algorithm is repeated 32 times to produce the ciphertext.
'                It has been found that the number of rounds is
'                exponentially proportional to the amount of time required
'                to find a key using a brute-force attack. So as the number
'                of rounds increases, the security of the algorithm
'                increases exponentially.
'
'                The detailed description of the actual algorithm is
'                contained in the official Skipjack specification provided
'                by the NSA after the algorithm was declassified.  Another
'                algorithm called KEA was declassified at the same time As
'                Skipjack, so the specification contains information about
'                both Skipjack and KEA. The first half of the document
'                concerns Skipjack and the second half focuses on KEA (Key
'                Exchange Algorithm).
'
'                **********************************************************
' SPECIAL NOTE:  The encryption process has been enhanced.
'
'                - Encryption process can be performed multiple times.
'                  See property EncryptRounds().
'
'                - Password key length is expanded to meet and exceed AES
'                  standards.  Key lengths are 128 to 448 in 32 bit
'                  increments and 448 to 1024 in 64 bit increments.  See
'                  property KeyLength().

'                - F-table loading is more dynamic because it is now based
'                  on key length and not a single table.  See LoadTable()
'                  routine.
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
' References:    Skipjack documentation above
'                http://csrc.nist.gov/encryption/skipjack/skipjack.pdf
'
'                Tropical software (Portions of description above)
'                http://www.tropsoft.com/strongenc/skipjack.htm
'
'                Observations on the SkipJack Encryption Algorithm
'                http://www.cs.technion.ac.il/~biham/Reports/SkipJack/
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
'              Documented, modified, updated variable names
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              - Modified Skipjack_File() and Skipjack_String() routines to
'                encrypt appended data information for additional security.
'              - Moved ExpandData() routine to basCommon.bas module.
'              - Renamed enumeration SKIPJACK_ROUNDS and used standard values
'                of 0-4 to represent number of rounds of encryption.
' 06-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              - Modified EvaluateKey() routine for initial table loading.
'              - Updated property EncryptRounds().  Number of encryption
'                rounds is now user-defined.  See constant MAX_ROUNDS.
'              - Rewrote Encrytion/Decryption routines.  I have found several
'                C versions of Skipjack on the web.  This interpretation is
'                based on the version by Paulo Barreto with some modifications
'                by myself because the C language can manipulate numeric values
'                that are not easily done within the realm of Visual Basic 6.0.
'
'                Originally written by Panu Rissanen <bande@lut.fi> 1998.06.24
'                Optimized by Mark Tillotson <markt@chaos.org.uk> 1998.06.25
'                Optimized by Paulo Barreto <pbarreto@nw.com.br> 1998.06.30
'
'                Home page:  http://www.larc.usp.br/~pbarreto/
'                            http://www.larc.usp.br/~pbarreto/skipjack32.zip
'
' 26-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added property KeyLength() to meet and exceed AES standards.
'              (128, 192, 256, 384, 512, 1024 bits)  Modified EvaluateKey()
'              routine to use utilize the new key lengths.
' 01-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Fixed OffSet range in EncryptData() and DecryptData() routines.
'              - Moved GetBlockSize() routine to common module.
'              - Set key length to increments of 32 bits for 128 to 448 and
'                increments of 64 bits for 448 to 1024.
' 08-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Set data block size to 16 bytes.
'              - Updated documentation.
' 15-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Added additional mixing of data within maintTable() array in
'                EvaluateKey() routine.
'              - Added data swapping to EncryptData() and DecryptData()
'                routines for enhanced security.
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Removed obsolete code. Updated documentation.
'              - Correct data size capture for decryption in Skipjack_String().
' 15-Mar-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Enhanced initial loading of key array in EvaluateKey() routine.
'              - Updated LoadTable() routine with a good alternative for
'                loading the table data.
' 29-Aug-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated initial loading of key array in EvaluateKey() routine.
' 01-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced Err.Raise with InfoMsg so programs will continue to
'              flow smoothly.
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Added ability to see progress display while encrypting or
'                decrypting data (EncryptData() and DecryptData() routines).
'              - Removed obsolete code.  Updated documentation.
'              - Updated initial key loading (EvaluateKey() routine).
' 02-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Replaced ShiftInt() routine with call to cMath32 module.
'              - Updated EvaluateKey() routine.
' 10-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Removed MixBytes() routine.  I created this routine and now
'              feel that it is no longer needed.
' 20-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated LoadTable() routine
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Made many updates as per Joe Sova's suggestions
'                1. Removed most of the erase array as they were not needed
'                2. Moved call to display progress to outside processing loop
'                   in EncryptData() and DecryptData() routines.
'                3. Moved Decryption size testing to earlier position within
'                   within Skipjack_File() routine.
'                4. Split size testing in Skipjack_String() routine.
' 21-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Added Public Event CurrentBlockSize() to track current
'                amount of data being processed.
'              - Updated GetBlockSize() routine.
'              - Increased area for number of rounds that will affect
'                multiple pass functionality.  See EncryptData() and
'                DecryptData() routines.
'              - Added new Property CreateNewFile().  Designates if creating
'                a new output file or overwriting the source file after
'                encryption or decryption.  Default value is TRUE.
'              - Updated Skipjack_File() routine to reference new property.
' 06-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Updated EncryptData() and DecryptData() routines to encompass
'              complete process per round if more than one iteration is
'              selected by user.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME    As String = "clsSkipjack"
  Private Const MAX_FILE_SIZE  As Currency = 10000000000@   ' Approx 10 Gig
  Private Const MAX_STRING_LEN As Long = &HFFFE&            ' 65534
  Private Const KB_32          As Long = &H8000&            ' 32768
  Private Const BLOCK_SIZE     As Long = 8
  Private Const DATA_BLOCK     As Long = 16
  Private Const MAX_ROUNDS     As Long = 10        ' Modify as you see fit
  Private Const MAX_BYTE       As Long = 256       ' Max range of byte (0-255)

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
  Private mblnCreateNewFile As Boolean
  Private mcurMaxSize       As Currency
  Private mcurAccrued       As Currency
  Private mcurOrigLength    As Currency
  Private mlngRounds        As Long
  Private mlngKeyLength     As Long
  Private maintTable()      As Integer
  Private mabytPassword()   As Byte
  Private mobjMath32        As cMath32

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
                mlngKeyLength = 128  ' Default value
    End Select
    
End Property

Public Property Let EncryptRounds(ByVal lngData As Long)
   
    ' 26-Dec-2008 Added multiple pass functionality.
    Select Case lngData
           Case 1 To MAX_ROUNDS: mlngRounds = lngData   ' Good selection
           Case Else:            mlngRounds = 1         ' Default value
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
    If IsArray(vntData) Then
    
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
    On Error GoTo 0
    Exit Property

Password_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume Password_CleanUp
    
End Property


' ***************************************************************************
' ****                         Methods                                   ****
' ***************************************************************************

' ***************************************************************************
' Routine:       Skipjack_File
'
' Description:   Function to encrypt/decrypt a file.
' 
' Syntax:        abytFile() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)  ' convert file location to byte array 
'                objSkipjack.Skipjack_File abytFile(), blnEncrypt             ' Encrypt/Decrypt data
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
' 15-Mar-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated EvaluateKey() and LoadTable() routines.
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Moved Decryption size testing to earlier an position as per
'              Joe Sova's suggestion.
' ***************************************************************************
Public Function Skipjack_File(ByRef abytFile() As Byte, _
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

    Const ROUTINE_NAME As String = "Skipjack_File"
    
    On Error GoTo Skipjack_File_Error
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
    
    Set objBigFiles = New cBigFiles   ' Instantiate class object
    ResetVariables                    ' Reset class variables
    Erase abytData()                  ' Always start with empty arrays
    Erase abytSize()
    
    If Not EvaluateKey() Then
        GoTo Skipjack_File_CleanUp
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
        GoTo Skipjack_File_CleanUp
    End If

    With objBigFiles
        .CalcFileSize strSource, curAmtLeft   ' Get size of file
        
        ' If zero byte file then leave
        If curAmtLeft < 1 Then
            gblnStopProcessing = True
            GoTo Skipjack_File_CleanUp
        End If
            
        ' If performing decryption then do not
        ' process the last few bytes of the file
        ' because this is where we stored the
        ' size of the original file.
        If blnEncrypt Then
            mcurOrigLength = curAmtLeft   ' Capture total byte count
        Else
            curAmtLeft = curAmtLeft - (UBound(abytSize) + 1)  ' Adjust amount that has to be decrypted
            curFilePos1 = curAmtLeft                          ' Set pointer to capture the last 12 bytes
            
            If Not .OpenReadOnly(strSource, hFile1) Then      ' Open the source file
                gblnStopProcessing = True
            End If
            
            .API_ReadFile hFile1, curFilePos1, abytSize()     ' Capture last 12 bytes of source file
            .API_CloseFile hFile1                             ' close source file
            
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                GoTo Skipjack_File_CleanUp
            End If
    
            MixAppendedData abytSize()                        ' Decrypt appended data
            CopyMemory lngLastBlock, abytSize(0), 4&          ' Size of last block of data without padding [0-3]
            CopyMemory mcurOrigLength, abytSize(4), 8&        ' Capture original file size [4-11]
            
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
                GoTo Skipjack_File_CleanUp
                
            End If
        End If
        
        mcurMaxSize = mcurOrigLength * mlngRounds   ' Save file size for progress calcs
        
        mcurAccrued = 0@   ' Init accrued progress
        curFilePos1 = 0@   ' Reset pointers to first position in file
        curFilePos2 = 0@
        
        ' Open source file
        If Not .OpenReadOnly(strSource, hFile1) Then
            gblnStopProcessing = True
            GoTo Skipjack_File_CleanUp
        End If
        
        strTempFile = CreateTempFile()  ' Create a temp target file
        
        ' Open target file
        If Not .OpenReadWrite(strTempFile, hFile2) Then
            gblnStopProcessing = True
            GoTo Skipjack_File_CleanUp
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
                Exit Do
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
            GoTo Skipjack_File_CleanUp
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
                GoTo Skipjack_File_CleanUp
                
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

Skipjack_File_CleanUp:
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
        Skipjack_File = False
    Else
        Skipjack_File = True
    End If
        
    On Error GoTo 0   ' Nullify this error trap
    Exit Function
    
Skipjack_File_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Skipjack_File_CleanUp

End Function

' ***************************************************************************
' Routine:       Skipjack_String
'
' Description:   Function to encrypt/decrypt a string of data.
' 
' Syntax:        abytData() = StrConv("abc", vbFromUnicode)          ' convert string data to byte array
'                objSkipjack.Skipjack_String abytData(), blnEncrypt  ' Encrypt/Decrypt data & return as byte array
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
Public Function Skipjack_String(ByRef abytData() As Byte, _
                       Optional ByVal blnEncrypt As Boolean = True) As Boolean

    Dim curDataLength As Currency
    Dim abytSize(7)   As Byte
    
    Const ROUTINE_NAME As String = "Skipjack_String"

    On Error GoTo Skipjack_String_Error
    
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
    
    ResetVariables     ' Reset class variables
    Erase abytSize()   ' Always start with empty arrays
    
    mcurMaxSize = CCur(UBound(abytData)) * mlngRounds   ' Save file size for progress calcs
    mcurAccrued = 0@                                    ' Init accrued progress
    
    If Not EvaluateKey() Then
        GoTo Skipjack_String_CleanUp
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
            GoTo Skipjack_String_CleanUp
            
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
    
Skipjack_String_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Skipjack_String = False
    Else
        Skipjack_String = True
    End If
    
    ResetVariables    ' Reset class variables
    Erase abytSize()  ' Always empty arrays when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function
    
Skipjack_String_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Skipjack_String_CleanUp

End Function


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Function EvaluateKey() As Boolean

    ' Called by Skipjack_File()
    '           Skipjack_String()
    
    Dim lngIdx      As Long
    Dim lngIndex    As Long
    Dim lngKeyLen   As Long
    Dim lngTemp     As Long
    Dim abytPwd()   As Byte
    Dim abytTable() As Byte

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
    
    Erase abytPwd()    ' Always start with empty arrays
    Erase abytTable()
    
    lngKeyLen = mlngKeyLength \ 8                        ' Ex:  16 bytes = 128 bits \ 8 bits
    abytPwd() = ExpandData(mabytPassword(), lngKeyLen)   ' Expand password key to desired length
    abytTable() = LoadTable()                            ' Load F-table data (ASCII 0-255 mixed)
    
    ReDim maintTable(0 To 9, MAX_BYTE)                   ' Size integer table array
    
    ' for (i = 0; i < 10; i++)
    '     byte *t = tab[i], k = key[i];
    '     for (c = 0; c < 256; c++)
    '         t[c] = fTable[c ^ k];
    For lngIndex = 0 To 9
        
        lngTemp = CLng(abytPwd(lngIndex))   ' Use first 10 values of password array
        
        ' Load integer table data
        For lngIdx = 0 To (MAX_BYTE - 1)
            maintTable(lngIndex, lngIdx) = CInt(abytTable((lngIdx Xor lngTemp) Mod MAX_BYTE))
        Next lngIdx
    Next lngIndex
    
EvaluateKey_CleanUp:
    Erase abytPwd()    ' Always empty arrays when not needed
    Erase abytTable()
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        EvaluateKey = False
    Else
        EvaluateKey = True
    End If
    
    On Error GoTo 0    ' Nullify this error trap
    Exit Function

EvaluateKey_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    ResetVariables
    gblnStopProcessing = True
    Resume EvaluateKey_CleanUp

End Function

Private Function EncryptData(ByRef abytData() As Byte) As Boolean

    ' Called by Skipjack_File()
    '           Skipjack_String()
    
    Dim lngRound      As Long
    Dim lngOffSet     As Long
    Dim lngProgress   As Long
    Dim lngDataLength As Long
    Dim intW1         As Integer
    Dim intW2         As Integer
    Dim intW3         As Integer
    Dim intW4         As Integer
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
          
        For lngRound = 1 To mlngRounds

            ' Transfer eight bytes to work array
            CopyMemory abytInput(0), abytData(lngOffSet), BLOCK_SIZE
                    
            ' Convert eight bytes to four integers
            CopyMemory intW1, abytInput(0), 2&
            CopyMemory intW2, abytInput(2), 2&
            CopyMemory intW3, abytInput(4), 2&
            CopyMemory intW4, abytInput(6), 2&
                
            ' Rule A1
            '    /* stepping rule A: */
            '    g0(tab, w1); w4 ^= w1 ^ 1;
            '    g1(tab, w4); w3 ^= w4 ^ 2;
            '    g2(tab, w3); w2 ^= w3 ^ 3;
            '    g3(tab, w2); w1 ^= w2 ^ 4;
            '    g4(tab, w1); w4 ^= w1 ^ 5;
            '    g0(tab, w4); w3 ^= w4 ^ 6;
            '    g1(tab, w3); w2 ^= w3 ^ 7;
            '    g2(tab, w2); w1 ^= w2 ^ 8;
            GG intW1, 0, 1, 2, 3: intW4 = intW4 Xor (intW1 Xor 1)
            GG intW4, 4, 5, 6, 7: intW3 = intW3 Xor (intW4 Xor 2)
            GG intW3, 8, 9, 0, 1: intW2 = intW2 Xor (intW3 Xor 3)
            GG intW2, 2, 3, 4, 5: intW1 = intW1 Xor (intW2 Xor 4)
            GG intW1, 6, 7, 8, 9: intW4 = intW4 Xor (intW1 Xor 5)
            GG intW4, 0, 1, 2, 3: intW3 = intW3 Xor (intW4 Xor 6)
            GG intW3, 4, 5, 6, 7: intW2 = intW2 Xor (intW3 Xor 7)
            GG intW2, 8, 9, 0, 1: intW1 = intW1 Xor (intW2 Xor 8)

            ' Rule B1
            '    /* stepping rule B: */
            '    w2 ^= w1 ^  9; g3(tab, w1);
            '    w1 ^= w4 ^ 10; g4(tab, w4);
            '    w4 ^= w3 ^ 11; g0(tab, w3);
            '    w3 ^= w2 ^ 12; g1(tab, w2);
            '    w2 ^= w1 ^ 13; g2(tab, w1);
            '    w1 ^= w4 ^ 14; g3(tab, w4);
            '    w4 ^= w3 ^ 15; g4(tab, w3);
            '    w3 ^= w2 ^ 16; g0(tab, w2);
            intW2 = intW2 Xor (intW1 Xor 9):  GG intW1, 2, 3, 4, 5
            intW1 = intW1 Xor (intW4 Xor 10): GG intW4, 6, 7, 8, 9
            intW4 = intW4 Xor (intW3 Xor 11): GG intW3, 0, 1, 2, 3
            intW3 = intW3 Xor (intW2 Xor 12): GG intW2, 4, 5, 6, 7
            intW2 = intW2 Xor (intW1 Xor 13): GG intW1, 8, 9, 0, 1
            intW1 = intW1 Xor (intW4 Xor 14): GG intW4, 2, 3, 4, 5
            intW4 = intW4 Xor (intW3 Xor 15): GG intW3, 6, 7, 8, 9
            intW3 = intW3 Xor (intW2 Xor 16): GG intW2, 0, 1, 2, 3

            ' Rule A2
            '    /* stepping rule A: */
            '    g1(tab, w1); w4 ^= w1 ^ 17;
            '    g2(tab, w4); w3 ^= w4 ^ 18;
            '    g3(tab, w3); w2 ^= w3 ^ 19;
            '    g4(tab, w2); w1 ^= w2 ^ 20;
            '    g0(tab, w1); w4 ^= w1 ^ 21;
            '    g1(tab, w4); w3 ^= w4 ^ 22;
            '    g2(tab, w3); w2 ^= w3 ^ 23;
            '    g3(tab, w2); w1 ^= w2 ^ 24;
            GG intW1, 4, 5, 6, 7: intW4 = intW4 Xor (intW1 Xor 17)
            GG intW4, 8, 9, 0, 1: intW3 = intW3 Xor (intW4 Xor 18)
            GG intW3, 2, 3, 4, 5: intW2 = intW2 Xor (intW3 Xor 19)
            GG intW2, 6, 7, 8, 9: intW1 = intW1 Xor (intW2 Xor 20)
            GG intW1, 0, 1, 2, 3: intW4 = intW4 Xor (intW1 Xor 21)
            GG intW4, 4, 5, 6, 7: intW3 = intW3 Xor (intW4 Xor 22)
            GG intW3, 8, 9, 0, 1: intW2 = intW2 Xor (intW3 Xor 23)
            GG intW2, 2, 3, 4, 5: intW1 = intW1 Xor (intW2 Xor 24)
    
            ' Perform 8 more iterations of mixing.
            ' Rule B2
            '    /* stepping rule B: */
            '    w2 ^= w1 ^ 25; g4(tab, w1);
            '    w1 ^= w4 ^ 26; g0(tab, w4);
            '    w4 ^= w3 ^ 27; g1(tab, w3);
            '    w3 ^= w2 ^ 28; g2(tab, w2);
            '    w2 ^= w1 ^ 29; g3(tab, w1);
            '    w1 ^= w4 ^ 30; g4(tab, w4);
            '    w4 ^= w3 ^ 31; g0(tab, w3);
            '    w3 ^= w2 ^ 32; g1(tab, w2);
            intW2 = intW2 Xor (intW1 Xor 25): GG intW1, 6, 7, 8, 9
            intW1 = intW1 Xor (intW4 Xor 26): GG intW4, 0, 1, 2, 3
            intW4 = intW4 Xor (intW3 Xor 27): GG intW3, 4, 5, 6, 7
            intW3 = intW3 Xor (intW2 Xor 28): GG intW2, 8, 9, 0, 1
            intW2 = intW2 Xor (intW1 Xor 29): GG intW1, 2, 3, 4, 5
            intW1 = intW1 Xor (intW4 Xor 30): GG intW4, 6, 7, 8, 9
            intW4 = intW4 Xor (intW3 Xor 31): GG intW3, 0, 1, 2, 3
            intW3 = intW3 Xor (intW2 Xor 32): GG intW2, 4, 5, 6, 7
                
            ' Convert four integers to bytes
            CopyMemory abytOutput(0), intW1, 2&
            CopyMemory abytOutput(2), intW2, 2&
            CopyMemory abytOutput(4), intW3, 2&
            CopyMemory abytOutput(6), intW4, 2&
                    
            ' Transfer encrytped data back to the temporary work array
            CopyMemory abytData(lngOffSet), abytOutput(0), BLOCK_SIZE

            ' An error occurred or user opted to STOP processing
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
    
        Next lngRound
    
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
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        EncryptData = False
    Else
        EncryptData = True
    End If

    On Error GoTo 0
    Exit Function

EncryptData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume EncryptData_CleanUp

End Function

Private Function DecryptData(ByRef abytData() As Byte) As Boolean

    ' Called by Skipjack_File()
    '           Skipjack_String()
    
    Dim lngRound      As Long
    Dim lngOffSet     As Long
    Dim lngProgress   As Long
    Dim lngDataLength As Long
    Dim intW1         As Integer
    Dim intW2         As Integer
    Dim intW3         As Integer
    Dim intW4         As Integer
    Dim abytInput(7)  As Byte
    Dim abytOutput(7) As Byte

    Const ROUTINE_NAME  As String = "DecryptData"

    On Error GoTo DecryptData_Error

    lngDataLength = UBound(abytData) + 1   ' Get size of incoming array
    
    ' 1-Jan-2009 Fixed lngOffSet range
    ' Decrypt the data
    For lngOffSet = 0 To (lngDataLength - 1) Step BLOCK_SIZE

        ' Added feature.
        ' Perform requested number of rounds
        For lngRound = 1 To mlngRounds
            
            ' Transfer eight bytes to work array
            CopyMemory abytInput(0), abytData(lngOffSet), BLOCK_SIZE
                    
            ' Convert eight bytes to four integers
            CopyMemory intW1, abytInput(0), 2&
            CopyMemory intW2, abytInput(2), 2&
            CopyMemory intW3, abytInput(4), 2&
            CopyMemory intW4, abytInput(6), 2&
        
            ' Perform 8 iterations of mixing.
            ' Inverse of Rule B2
            '    /* stepping rule A: */
            '    h1(tab, w2); w3 ^= w2 ^ 32;
            '    h0(tab, w3); w4 ^= w3 ^ 31;
            '    h4(tab, w4); w1 ^= w4 ^ 30;
            '    h3(tab, w1); w2 ^= w1 ^ 29;
            '    h2(tab, w2); w3 ^= w2 ^ 28;
            '    h1(tab, w3); w4 ^= w3 ^ 27;
            '    h0(tab, w4); w1 ^= w4 ^ 26;
            '    h4(tab, w1); w2 ^= w1 ^ 25;
            HH intW2, 4, 5, 6, 7: intW3 = intW3 Xor (intW2 Xor 32)
            HH intW3, 0, 1, 2, 3: intW4 = intW4 Xor (intW3 Xor 31)
            HH intW4, 6, 7, 8, 9: intW1 = intW1 Xor (intW4 Xor 30)
            HH intW1, 2, 3, 4, 5: intW2 = intW2 Xor (intW1 Xor 29)
            HH intW2, 8, 9, 0, 1: intW3 = intW3 Xor (intW2 Xor 28)
            HH intW3, 4, 5, 6, 7: intW4 = intW4 Xor (intW3 Xor 27)
            HH intW4, 0, 1, 2, 3: intW1 = intW1 Xor (intW4 Xor 26)
            HH intW1, 6, 7, 8, 9: intW2 = intW2 Xor (intW1 Xor 25)
    
            ' Use 24 iterations to mix the data for each round.
            '
            ' Inverse of Rule A2
            '    /* stepping rule B: */
            '    w1 ^= w2 ^ 24; h3(tab, w2);
            '    w2 ^= w3 ^ 23; h2(tab, w3);
            '    w3 ^= w4 ^ 22; h1(tab, w4);
            '    w4 ^= w1 ^ 21; h0(tab, w1);
            '    w1 ^= w2 ^ 20; h4(tab, w2);
            '    w2 ^= w3 ^ 19; h3(tab, w3);
            '    w3 ^= w4 ^ 18; h2(tab, w4);
            '    w4 ^= w1 ^ 17; h1(tab, w1);
            intW1 = intW1 Xor (intW2 Xor 24): HH intW2, 2, 3, 4, 5
            intW2 = intW2 Xor (intW3 Xor 23): HH intW3, 8, 9, 0, 1
            intW3 = intW3 Xor (intW4 Xor 22): HH intW4, 4, 5, 6, 7
            intW4 = intW4 Xor (intW1 Xor 21): HH intW1, 0, 1, 2, 3
            intW1 = intW1 Xor (intW2 Xor 20): HH intW2, 6, 7, 8, 9
            intW2 = intW2 Xor (intW3 Xor 19): HH intW3, 2, 3, 4, 5
            intW3 = intW3 Xor (intW4 Xor 18): HH intW4, 8, 9, 0, 1
            intW4 = intW4 Xor (intW1 Xor 17): HH intW1, 4, 5, 6, 7
    
            ' Inverse of Rule B1
            '    /* stepping rule A: */
            '    h0(tab, w2); w3 ^= w2 ^ 16;
            '    h4(tab, w3); w4 ^= w3 ^ 15;
            '    h3(tab, w4); w1 ^= w4 ^ 14;
            '    h2(tab, w1); w2 ^= w1 ^ 13;
            '    h1(tab, w2); w3 ^= w2 ^ 12;
            '    h0(tab, w3); w4 ^= w3 ^ 11;
            '    h4(tab, w4); w1 ^= w4 ^ 10;
            '    h3(tab, w1); w2 ^= w1 ^  9;
            HH intW2, 0, 1, 2, 3: intW3 = intW3 Xor (intW2 Xor 16)
            HH intW3, 6, 7, 8, 9: intW4 = intW4 Xor (intW3 Xor 15)
            HH intW4, 2, 3, 4, 5: intW1 = intW1 Xor (intW4 Xor 14)
            HH intW1, 8, 9, 0, 1: intW2 = intW2 Xor (intW1 Xor 13)
            HH intW2, 4, 5, 6, 7: intW3 = intW3 Xor (intW2 Xor 12)
            HH intW3, 0, 1, 2, 3: intW4 = intW4 Xor (intW3 Xor 11)
            HH intW4, 6, 7, 8, 9: intW1 = intW1 Xor (intW4 Xor 10)
            HH intW1, 2, 3, 4, 5: intW2 = intW2 Xor (intW1 Xor 9)

            ' Inverse of Rule A1
            '    /* stepping rule B: */
            '    w1 ^= w2 ^ 8; h2(tab, w2);
            '    w2 ^= w3 ^ 7; h1(tab, w3);
            '    w3 ^= w4 ^ 6; h0(tab, w4);
            '    w4 ^= w1 ^ 5; h4(tab, w1);
            '    w1 ^= w2 ^ 4; h3(tab, w2);
            '    w2 ^= w3 ^ 3; h2(tab, w3);
            '    w3 ^= w4 ^ 2; h1(tab, w4);
            '    w4 ^= w1 ^ 1; h0(tab, w1);
            intW1 = intW1 Xor (intW2 Xor 8): HH intW2, 8, 9, 0, 1
            intW2 = intW2 Xor (intW3 Xor 7): HH intW3, 4, 5, 6, 7
            intW3 = intW3 Xor (intW4 Xor 6): HH intW4, 0, 1, 2, 3
            intW4 = intW4 Xor (intW1 Xor 5): HH intW1, 6, 7, 8, 9
            intW1 = intW1 Xor (intW2 Xor 4): HH intW2, 2, 3, 4, 5
            intW2 = intW2 Xor (intW3 Xor 3): HH intW3, 8, 9, 0, 1
            intW3 = intW3 Xor (intW4 Xor 2): HH intW4, 4, 5, 6, 7
            intW4 = intW4 Xor (intW1 Xor 1): HH intW1, 0, 1, 2, 3

            ' Convert four integers to bytes
            CopyMemory abytOutput(0), intW1, 2&
            CopyMemory abytOutput(2), intW2, 2&
            CopyMemory abytOutput(4), intW3, 2&
            CopyMemory abytOutput(6), intW4, 2&
                    
            ' Transfer decrypted data back to the temporary work array
            CopyMemory abytData(lngOffSet), abytOutput(0), BLOCK_SIZE

            ' An error occurred or user opted to STOP processing
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If
    
        Next lngRound
    
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
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        DecryptData = False
    Else
        DecryptData = True
    End If

    On Error GoTo 0
    Exit Function

DecryptData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume DecryptData_CleanUp

End Function

'    /**
'    * The key-dependent permutation GG on V^16 is a four-round Feistel network.
'    * The round function is a fixed byte-substitution table (permutation on V^8),
'    * the F-table.  Each round of GG incorporates a single byte from the key.
'    */
'    #define g(tab, w, i, j, k, l) \
'    { \
'       w ^= (word32)tab[i][w & 0xff] << 8; \
'       w ^= (word32)tab[j][w >>   8]; \
'       w ^= (word32)tab[k][w & 0xff] << 8; \
'       w ^= (word32)tab[l][w >>   8]; \
Private Sub GG(ByRef intWn As Integer, _
               ByVal a As Byte, _
               ByVal b As Byte, _
               ByVal c As Byte, _
               ByVal d As Byte)
                   
    With mobjMath32
        intWn = intWn Xor (maintTable(a, (.w16Shift((intWn And &HFF&), 8) Mod MAX_BYTE)))  ' Shift left
        intWn = intWn Xor (maintTable(b, (.w16Shift(intWn, -8) Mod MAX_BYTE)))             ' Shift right
        intWn = intWn Xor (maintTable(c, (.w16Shift((intWn And &HFF&), 8) Mod MAX_BYTE)))
        intWn = intWn Xor (maintTable(d, (.w16Shift(intWn, -8) Mod MAX_BYTE)))
    End With
    
End Sub

'    * The inverse of the GG permutation.
'    #define h(tab, w, i, j, k, l) \
'    { \
'        w ^= (word32)tab[l][w >>   8]; \
'        w ^= (word32)tab[k][w & 0xff] << 8; \
'        w ^= (word32)tab[j][w >>   8]; \
'        w ^= (word32)tab[i][w & 0xff] << 8; \
Private Sub HH(ByRef intWn As Integer, _
               ByVal a As Byte, _
               ByVal b As Byte, _
               ByVal c As Byte, _
               ByVal d As Byte)
               
    With mobjMath32
        intWn = intWn Xor (maintTable(d, (.w16Shift(intWn, -8) Mod MAX_BYTE)))              ' Shift right
        intWn = intWn Xor (maintTable(c, (.w16Shift((intWn And &HFF&), 8) Mod MAX_BYTE)))   ' Shift left
        intWn = intWn Xor (maintTable(b, (.w16Shift(intWn, -8) Mod MAX_BYTE)))
        intWn = intWn Xor (maintTable(a, (.w16Shift((intWn And &HFF&), 8) Mod MAX_BYTE)))
    End With

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

' ***************************************************************************
' Routine:       LoadTable
'
' Description:   F-table data consist of ASCII 0-255 values mixed with
'                no duplicates.
'
'                This is an alternate way of loading the F-table array.
'
'                For added security, I have made this table a little more
'                dynamic by creating several data strings and select an
'                appropriate one based on key length provided by user.
'
' Returns:       Mixed byte array of ASCII values 0-255
'
' ---------------------------------------------------------------------------
'    ' Original Skipjack table data
'    strData = strData & "A3 D7 09 83 F8 48 F6 F4 B3 21 15 78 99 B1 AF F9 "
'    strData = strData & "E7 2D 4D 8A CE 4C CA 2E 52 95 D9 1E 4E 38 44 28 "
'    strData = strData & "0A DF 02 A0 17 F1 60 68 12 B7 7A C3 E9 FA 3D 53 "
'    strData = strData & "96 84 6B BA F2 63 9A 19 7C AE E5 F5 F7 16 6A A2 "
'    strData = strData & "39 B6 7B 0F C1 93 81 1B EE B4 1A EA D0 91 2F B8 "
'    strData = strData & "55 B9 DA 85 3F 41 BF E0 5A 58 80 5F 66 0B D8 90 "
'    strData = strData & "35 D5 C0 A7 33 06 65 69 45 00 94 56 6D 98 9B 76 "
'    strData = strData & "97 FC B2 C2 B0 FE DB 20 E1 EB D6 E4 DD 47 4A 1D "
'    strData = strData & "42 ED 9E 6E 49 3C CD 43 27 D2 07 D4 DE C7 67 18 "
'    strData = strData & "89 CB 30 1F 8D C6 8F AA C8 74 DC C9 5D 5C 31 A4 "
'    strData = strData & "70 88 61 2C 9F 0D 2B 87 50 82 54 64 26 7D 03 40 "
'    strData = strData & "34 4B 1C 73 D1 C4 FD 3B CC FB 7F AB E6 3E 5B A5 "
'    strData = strData & "AD 04 23 9C 14 51 22 F0 29 79 71 7E FF 8C 0E E2 "
'    strData = strData & "0C EF BC 72 75 6F 37 A1 EC D3 8E 62 8B 86 10 E8 "
'    strData = strData & "08 77 11 BE 92 4F 24 C5 32 36 9D CF F3 A6 BB AC "
'    strData = strData & "5E 6C A9 13 57 25 B5 E3 BD A8 3A 01 05 59 2A 46 "
' ---------------------------------------------------------------------------
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 12-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              Documented routine
' 20-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated number of data mixing iterations
' ***************************************************************************
Private Function LoadTable() As Byte()

    ' Called by EvaluateKey()

    Dim lngIndex  As Long
    Dim strData   As String
    Dim avntData  As Variant
    Dim abytTmp() As Byte
    
    avntData = Empty        ' Always start with empty variants
    strData = vbNullString  ' Verify data string is empty
    
    ' See Skipjack_tbl.txt for additional sets of data
    ' strings from which to choose.  See application
    ' BuildTables.vbp to create different style CASE
    ' statements.
    '
    ' Determine which string of data to use based on key
    ' length selected by the user.
    '
    ' Each case statement:
    '    1.  Consists of all ASCII values (0-255) in two
    '        character hex equivalent with no duplicates
    '    2.  Mixed 100-600 iterations using Knuth Shuffle
    '
    Select Case mlngKeyLength
           Case 128, 448
                strData = strData & "C9 06 92 B8 AC 23 F6 25 86 16 E7 41 51 5F 83 61 39 B0 87 38 B3 E1 1F FB 8D B5 33 99 4A 81 0E 67 "
                strData = strData & "A5 E5 09 77 AA 40 97 96 B9 B7 BA C3 A3 78 E6 36 85 08 0A AB BE 9D 55 21 0B 2B F8 74 37 17 48 8E "
                strData = strData & "FA 90 72 3A DD 3F 01 D9 0D 0F A4 B2 C1 64 12 CD 24 73 19 03 B1 7E BF 9B 91 E8 F4 DC 71 F0 D2 47 "
                strData = strData & "94 34 7F 2A 5C A8 A2 1D AF 46 3B CA 14 C7 11 7D CC 84 F9 6B F5 9E 58 EA DF D6 65 29 50 E9 89 76 "
                strData = strData & "52 EC D8 4E 1B DE 45 D4 15 8A C2 30 E3 E0 8C A1 07 66 3C D5 4F C6 9C 6F FE 62 49 F2 AD CF B6 69 "
                strData = strData & "05 57 2C F7 EE 75 54 18 F3 5D 5B 7C BC 2F 59 F1 EB 7A 0C 6D 6A 3E 10 98 B4 31 D3 C0 43 2E FF 9F "
                strData = strData & "DA 4B 42 6C 1E 9A 2D 5A 56 3D 26 C4 FD 8B 00 6E 02 ED 88 27 E4 5E E2 35 D7 60 8F 93 AE A0 1C C8 "
                strData = strData & "22 FC 28 C5 04 A6 BB CB 4C EF 70 44 A7 4D 13 BD 20 95 7B DB D0 A9 80 D1 68 CE 82 63 53 1A 32 79 "
           Case 160, 512
                strData = strData & "AD 96 6F A0 1A C4 46 1B 7F 3E 19 16 DC 7B DE DA EC F2 EA 23 49 62 55 06 C2 FB 4C 2E 59 A8 04 8A "
                strData = strData & "48 99 76 C1 5A 50 B1 34 82 E4 2A F0 BC B5 3C 64 44 86 8B B4 72 4E 15 C0 0C C9 42 89 71 27 47 43 "
                strData = strData & "B6 77 30 F6 FC B9 BF 52 60 5D 45 4D 61 A2 2F 54 5B 1E 75 FF 3A 69 05 DF 93 4A D5 1F 9E 28 20 63 "
                strData = strData & "E7 0A 24 3F 8E 07 A4 09 18 41 A3 29 9C C7 5C 36 D6 D2 26 0D BB 66 83 F9 1C E8 0F 97 EE 94 9D 02 "
                strData = strData & "38 3B 21 F5 FA E3 32 22 E1 E6 73 D8 C6 7C 78 12 9B C5 4F 85 9A 7D 01 D0 51 B7 D9 6E A9 5E E9 CD "
                strData = strData & "B0 CA DD 65 9F D1 2C F1 AE A6 A1 AF 11 4B 92 8C F7 67 AC D3 00 70 B8 0B 0E 3D 8F 03 AB 90 7E ED "
                strData = strData & "33 40 C3 1D 74 D7 BE C8 F3 98 88 35 F4 E2 31 39 6C B2 FE EF 56 79 A5 BA 58 AA EB 80 CE CF DB A7 "
                strData = strData & "E5 2B 91 17 37 81 FD 57 8D 10 68 B3 F8 CB 6A 2D 95 7A 08 5F CC 13 6B 25 6D E0 84 D4 BD 53 87 14 "
           Case 192, 576
                strData = strData & "9E 4A 44 97 71 5F A3 1E 76 7F 9C AA 43 48 2D EF 00 D4 EB F9 18 08 86 3A 2E E8 85 BE C7 F4 9A 88 "
                strData = strData & "53 38 AF 75 30 6B AB E4 27 42 66 69 36 B5 62 BB DA B2 98 77 68 D1 96 CB E0 15 8D 7E 7C 2A BF 73 "
                strData = strData & "B3 6E 29 DD A0 46 C2 6D 7A FD 5C 70 0A 33 B9 E7 16 F1 8B C1 95 E9 5E E6 9F 41 DC 14 4F 54 8C 4C "
                strData = strData & "25 1D 02 A2 04 AE 61 CA 0E DF EC E1 1C 03 6A 90 D7 3F A1 F5 58 CD 8A 59 32 74 24 63 23 ED 52 28 "
                strData = strData & "19 D3 B0 91 C3 65 CF 21 D2 9B D8 FB A9 01 8E 3E 4E F3 67 8F 79 C8 13 2C 84 56 B6 4B 6F C6 31 0D "
                strData = strData & "D5 A6 34 AD E2 AC FA 11 DB 82 1B BC 64 07 20 37 87 1F 0F 40 7D 55 99 83 7B D0 57 60 49 1A A7 FE "
                strData = strData & "DE 26 80 B1 05 E5 6C EA B7 A8 22 C9 A5 B8 3C 89 F7 45 93 5A BD A4 0B C4 FF 17 B4 0C 12 EE 78 47 "
                strData = strData & "50 39 C5 FC 10 81 92 BA 9D F8 D9 3B F2 72 2B 35 CE 09 06 F6 94 C0 3D D6 F0 CC 4D 5D 51 2F 5B E3 "
           Case 224, 640
                strData = strData & "DE 06 7F F5 75 66 FA F8 6E E0 EF F1 C8 AB FC 6D BB E5 7E 8F 78 0C C6 5F 80 8B 64 A6 9A F6 84 04 "
                strData = strData & "FF A8 D0 1D 1B F0 65 FD 26 51 73 0F CD 7C 94 27 A1 1C AE 4C 8E 24 C7 E6 13 EC 77 83 CC 2F 49 1E "
                strData = strData & "B0 3B A5 8A 1A 70 95 B5 D2 D1 41 68 20 9F 4D D4 62 FB EB 82 5E 4F 5C 03 85 FE BD 55 0E 48 ED 47 "
                strData = strData & "00 CB 97 36 6F 2E 18 50 D9 02 B7 15 1F 29 10 12 5A 2C C9 E9 08 A2 23 A9 DB 42 DF 79 B1 D6 C3 B6 "
                strData = strData & "88 31 F2 BE 2D CA 33 61 93 46 92 9B 56 57 C1 74 32 A3 28 40 86 76 D8 19 22 05 53 87 A4 2B 89 BF "
                strData = strData & "07 3D B2 34 45 7D D3 6A E3 90 63 69 CF AC 25 B9 4A EA E4 E8 6B E7 B8 0D F9 F7 96 14 AD 43 E1 AA "
                strData = strData & "11 C4 3F 09 9E 3E 54 8C 4E DA 5D BC C5 A7 30 F4 58 B3 CE 7B 4B 3C 16 44 D7 2A 98 17 71 B4 59 37 "
                strData = strData & "91 5B C0 0B 3A 67 01 A0 BA 0A 7A EE 9D D5 9C 52 72 8D 39 38 DC AF F3 81 60 E2 6C 21 C2 35 DD 99 "
           Case 256, 704
                strData = strData & "A5 89 EF D5 A1 C8 B0 B6 4B 65 82 29 13 45 51 46 D9 A0 AF 37 4E F1 71 E1 C3 02 48 FA D2 57 F6 80 "
                strData = strData & "99 49 3A 17 76 67 21 6A 87 CD 36 50 9D F9 C1 14 54 5C BF 30 81 66 C4 04 FD 09 60 32 F3 B5 A3 C9 "
                strData = strData & "E4 6F C7 7B 06 86 6D 23 40 2F AC 22 FB 16 5B 43 A7 8C 9C DD 98 E7 62 3D B7 1D 03 31 59 F2 EE 39 "
                strData = strData & "FE 7C FC B1 88 94 0F D8 05 70 4A E9 18 C5 9B D6 34 93 12 7D 75 6C 74 0B EB 27 5A D0 1E 5D AE 33 "
                strData = strData & "77 D4 BA ED BC 72 7E D7 2B A6 B9 C0 7F 69 3E 53 8E 84 91 DC 52 F5 35 DB E8 B4 68 1F 8F 00 61 96 "
                strData = strData & "B2 0A A2 3C EC E5 1B 2E 4D 63 CC 01 B8 8D 64 CA A8 9A 97 78 08 BD 5F 8A 44 25 0C 8B 5E F0 90 F7 "
                strData = strData & "E0 10 C2 3F CE CF 95 DE 0D 42 26 58 79 47 F4 D1 FF B3 E2 55 A9 2C 38 C6 7A 20 85 15 1A 6E 3B E6 "
                strData = strData & "19 56 9E 4C 2D 6B AA 41 2A 28 CB F8 DA AD 11 D3 0E 1C 24 AB 4F 83 9F BB EA DF A4 BE E3 92 07 73 "
           Case 288, 768
                strData = strData & "18 1E C7 64 10 C3 A8 8C CB 1D 61 B7 2D 82 CD 35 D7 BD BF 6F E6 60 3A AE 80 4C 11 6B 02 DA 77 0A "
                strData = strData & "5F 39 51 3F 49 69 E7 6D E0 E2 C1 88 78 34 A6 00 EF E3 43 A2 9F 5E 04 B3 2A 1B 58 D5 0E 30 B8 6C "
                strData = strData & "CA 31 63 2F F6 4B 07 70 8E 06 E9 46 DC D1 84 EB 8B 91 B5 23 6E F8 75 28 1C BE A9 CE E8 FA 9B 25 "
                strData = strData & "B4 DB 8D C0 7C F4 7F 9A 44 99 87 86 03 F0 73 B0 4E 68 52 50 F7 48 24 D6 3C 7D 22 65 6A F3 21 FB "
                strData = strData & "8F B1 2E 5B D3 54 15 3E 01 3B DF B9 57 83 CF 41 BC 85 93 C8 9E 94 2C ED 09 FD 3D 29 32 1F AA 4D "
                strData = strData & "F5 56 A4 F9 AC 7A 47 C2 D4 0F 0B FC 1A EA FF A3 13 53 C4 AD 08 97 71 F2 40 67 72 76 5A D0 7B F1 "
                strData = strData & "EE C5 05 20 27 26 79 98 EC 17 62 C9 4A A0 8A 9C 7E DD B2 DE D9 E1 D8 FE A1 0C 96 CC 45 C6 14 89 "
                strData = strData & "42 37 B6 BA 66 36 4F 9D 55 12 AF E5 BB 90 E4 D2 59 2B 92 A5 81 5D A7 33 5C AB 16 95 19 38 74 0D "
           Case 320, 832
                strData = strData & "FF 6E 8E 41 F4 D4 BE 95 A7 A3 07 BB 1A 33 9B CF D1 8D 43 F8 4E EC F5 8A 74 5D 69 9A 23 C7 17 C3 "
                strData = strData & "73 90 55 1D 2A AF 35 A8 71 44 3E 54 6F 05 D0 93 87 98 30 C2 F1 C4 E4 FB 7F 80 AC CB C8 61 59 DC "
                strData = strData & "E7 E8 D3 36 21 56 E6 F6 CE E2 DD 25 22 83 0B B3 53 97 4C 94 91 AE 1E 32 1F C0 B5 3C 28 D6 D8 2C "
                strData = strData & "45 B4 E0 15 EB 31 AA E5 2B D9 82 29 C9 13 A9 3F 16 92 C1 8F 4A 08 F0 DF 2D B1 2F 88 9C 5F 70 63 "
                strData = strData & "99 50 FA FD 46 84 1B 11 51 B6 8B CC 64 E9 BC DB 03 39 86 67 60 6B ED FC 47 9F EA 5B 9D F9 77 57 "
                strData = strData & "D2 7E FE 66 A4 A5 96 24 F2 68 89 6D 09 BF 02 20 EE 0C 58 85 DE AD C5 0E 19 75 3B 3D 79 A6 12 D5 "
                strData = strData & "4D 37 00 4F 48 49 01 5E CD AB A0 BD 78 F7 42 EF 7A 65 40 C6 B7 81 27 B9 E3 06 38 62 F3 0F 7D A2 "
                strData = strData & "B8 BA 7B 1C 76 14 E1 3A B0 52 0D 6C 34 9E 7C CA 8C A1 72 4B B2 5C 6A 18 D7 DA 26 2E 0A 04 10 5A "
           Case 352, 896
                strData = strData & "5B 85 F6 83 6D EA D4 D1 14 BC 70 E1 F4 C4 5D F0 66 1F 9A D9 67 E4 92 A9 3C DE 57 05 69 32 B2 26 "
                strData = strData & "F7 42 25 45 B3 9D B4 04 72 50 CC 10 C3 D3 4A 01 82 68 BE 90 D2 EF 96 E7 13 8C CE D5 31 3D 23 B7 "
                strData = strData & "2E 94 1D 40 6A 16 46 C7 55 18 7F 08 29 48 CD F8 F3 E0 38 5A 6B 81 11 17 4C 4E DB 77 41 47 0B 86 "
                strData = strData & "0D FC 73 75 1E DD C5 22 74 87 80 30 D0 02 5C 27 EB A7 20 A2 95 54 FA 0A CB 3B FF 49 1B 19 4D 00 "
                strData = strData & "C8 7E E9 59 97 28 5E 0F EE AC 58 A6 79 3A DC A3 2B 7C B1 FB F9 39 07 8E 1C CF DA BD CA 03 8F C6 "
                strData = strData & "E3 FD 53 99 56 65 C9 BA DF A1 A5 71 0E A4 9C AF 60 E6 6C C1 D8 B5 36 09 B0 7B 61 AE 15 E8 3F 64 "
                strData = strData & "44 2D 06 B8 21 98 8D 9F 8B 43 BB AD 51 C2 93 7A 5F 88 84 78 D6 0C 89 F5 A0 A8 8A C0 4F 2A 35 F1 "
                strData = strData & "12 BF E5 63 37 4B B6 52 91 EC 33 E2 D7 24 B9 6F 9E 76 7D 2C 9B 34 ED 62 F2 FE 2F AB AA 6E 1A 3E "
           Case 384, 960
                strData = strData & "13 B1 18 CF 71 6E 80 58 20 40 8C 9C 77 FB 0F 8F BB DB 16 F8 7D B7 D3 B6 9E D2 D9 B2 2C 69 F2 78 "
                strData = strData & "3C F6 BD C7 64 2B 1B 9B 73 FE 5C 8A 4F B9 14 A5 48 F7 89 31 7A AE 6A D7 36 EA 10 3B FD 65 8B B8 "
                strData = strData & "3A A4 3E 23 22 BA 4A 92 12 EF A6 D1 4B 66 1C FA C8 95 5B 96 B3 33 3F 1F B5 A0 84 24 DA 63 A3 35 "
                strData = strData & "03 C3 F1 52 85 49 E9 E1 72 29 74 76 43 CC 45 AB 42 B4 BE 97 A8 11 55 AD 83 F4 E4 94 C1 6D C9 BF "
                strData = strData & "87 0D CB 44 2E E3 D8 91 AF 9D E6 D0 07 27 A7 7B B0 C6 59 90 60 02 6B E8 1A 7F 05 C0 15 82 A2 88 "
                strData = strData & "BC 99 5E F3 A9 68 75 ED 01 38 C5 5F EC FF F0 E2 F9 17 C4 EE AA 50 D5 AC 25 9A 21 5D 3D 47 DC 7C "
                strData = strData & "E0 9F 4E DD D6 1D 56 79 E7 4D 53 CD CE 70 8D 61 FC DF 67 F5 08 93 7E 4C 86 39 6F 41 0B 51 DE 1E "
                strData = strData & "A1 30 6C 37 26 62 C2 04 06 54 57 09 2A D4 CA 98 0E 00 28 8E E5 0A 2D 34 5A 0C 46 2F 32 EB 81 19 "
           Case 416, 1024
                strData = strData & "56 95 F3 F4 AD 6C 6F 76 86 4D 98 7B 7E E9 17 3E 7D 74 52 6D 01 FA 53 B5 85 BE F6 FD 62 E4 55 DC "
                strData = strData & "2B 3D 6B C1 A5 07 88 E8 32 12 A0 77 0E A2 24 5E 54 EC 43 AF 93 B6 3F 41 65 F7 8A 5C 18 64 ED 94 "
                strData = strData & "33 AB D1 72 BC E1 6E FF 7C DD DB 63 6A 92 02 1C C4 14 A8 BD E0 BF 80 51 8E FC 29 78 37 D3 22 2E "
                strData = strData & "57 84 CA 58 C9 DE 75 A9 9E 49 C5 82 B9 E2 28 27 40 04 87 99 35 0B 73 F8 E6 34 1D 09 9C 5A 9F 15 "
                strData = strData & "7F 50 30 21 4B 83 91 9B 8F 3B 48 71 D2 A6 CF F2 1B 5B D7 BA AA 47 2A 3C 4E 08 C3 FE C7 23 61 00 "
                strData = strData & "B7 0A D9 5F 68 EE A3 C2 A4 31 60 89 F0 C0 26 2F CC CD 8D 3A EF 03 A1 C8 DF D4 D6 0F 46 AC 69 9D "
                strData = strData & "DA 38 16 CB 90 B4 BB B3 25 10 05 B8 CE 11 4C 4A 36 1A F9 7A 9A 19 13 8C AE 42 59 0C 20 44 06 C6 "
                strData = strData & "70 EB B2 D5 2D 66 1E 39 F1 A7 97 67 5D FB 4F D0 B1 1F 96 F5 D8 79 45 E7 E3 8B 2C EA 81 E5 B0 0D "
    End Select

    avntData = Split(strData, Chr$(32))   ' convert data string into an array
    ReDim abytTmp(MAX_BYTE)               ' Resize temp byte array
    
    ' Convert data to a smaller footprint
    For lngIndex = 0 To (MAX_BYTE - 1)
        abytTmp(lngIndex) = CByte("&H" & avntData(lngIndex))
    Next lngIndex
    
    LoadTable = abytTmp()    ' Return data in byte array format
    
    Erase abytTmp()          ' Always empty arrays when not needed
    avntData = Empty         ' Always empty variants when not needed
    strData = vbNullString   ' Empty data string
    
End Function

Private Sub ResetVariables()

    Erase maintTable()  ' Always empty arrays when not needed
    mcurOrigLength = 0@
    
End Sub

Private Sub Class_Initialize()

    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objSkipjack = new cSkipjack
                
    ResetVariables
    Erase mabytPassword()   ' Always start with empty arrays
         
    Set mobjMath32 = Nothing       ' Always free object from memory
    Set mobjMath32 = New cMath32   ' Instantiate new class object
    
    ' Set property values
    EncryptRounds = 1      ' Number of rounds of encryption
    KeyLength = 128        ' Minimum key length
    CreateNewFile = True   ' Default is to create a new output file
    
End Sub

Private Sub Class_Terminate()
    
    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objSkipjack = Nothing
            
    ResetVariables
    Erase mabytPassword()      ' Always empty arrays when not needed
    Set mobjMath32 = Nothing   ' Always free object from memory

End Sub
