Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | mSha1Hash.vba                                               |
'| EntryPoint   | varies (GetSha1Hash, GetSha1HashLeftHalf, GetSha1HashBytes) |
'| Purpose      | Compute Sha1 Hash for various text strings                  |
'| Inputs       | string value                                                |
'| Outputs      | SHA-1 Hash lowercase hexidecimal represnetation of input str|
'| Dependencies | bcrypt.dll                                                  |
'| By Name,Date | T.Sciple, 9/8/2025                                          |

' 64-bit API declarations for bcrypt.dll
#If VBA7 Then
    Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" (ByRef phAlgorithm As LongPtr, ByVal pszAlgId As LongPtr, ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptCreateHash Lib "bcrypt.dll" (ByVal hAlgorithm As LongPtr, ByRef phHash As LongPtr, ByVal pbHashObject As LongPtr, ByVal cbHashObject As Long, ByVal pbSecret As LongPtr, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptHashData Lib "bcrypt.dll" (ByVal hHash As LongPtr, ByVal pbInput As LongPtr, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptFinishHash Lib "bcrypt.dll" (ByVal hHash As LongPtr, ByVal pbOutput As LongPtr, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptDestroyHash Lib "bcrypt.dll" (ByVal hHash As LongPtr) As Long
    Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" (ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptGetProperty Lib "bcrypt.dll" (ByVal hObject As LongPtr, ByVal pszProperty As LongPtr, ByVal pbOutput As LongPtr, ByVal cbOutput As Long, ByRef pcbResult As Long, ByVal dwFlags As Long) As Long
#Else
' 32-bit API declarations (for older Office versions)
    Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" (ByRef phAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptCreateHash Lib "bcrypt.dll" (ByVal hAlgorithm As Long, ByRef phHash As Long, ByVal pbHashObject As Long, ByVal cbHashObject As Long, ByVal pbSecret As Long, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptHashData Lib "bcrypt.dll" (ByVal hHash As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptFinishHash Lib "bcrypt.dll" (ByVal hHash As Long, ByVal pbOutput As Long, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDestroyHash Lib "bcrypt.dll" (ByVal hHash As Long) As Long
    Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptGetProperty Lib "bcrypt.dll" (ByVal hObject As Long, ByVal pszProperty As Long, ByVal pbOutput As Long, ByVal cbOutput As Long, ByRef pcbResult As Long, ByVal dwFlags As Long) As Long
#End If

' Constants for bcrypt.dll functions
Private Const BCRYPT_ALG_HANDLE_HMAC_FLAG = &H8
Private Const BCRYPT_OBJECT_LENGTH = "ObjectLength"
Private Const BCRYPT_HASH_LENGTH = "HashLength"
Private Const STATUS_SUCCESS As Long = 0


' SHA-1 algorithm string (wide-char/unicode pointer)
#If VBA7 Then
    Private Const BCRYPT_SHA1_ALGORITHM_WIDE As LongPtr = 1000 ' Pointer to "SHA1" wide string
#Else
    Private Const BCRYPT_SHA1_ALGORITHM_WIDE As Long = 1000 ' Pointer to "SHA1" wide string
#End If


Public Function GetSha1Hash(ByVal str As String) As String
    If str = "" Then
        GetSha1Hash = "Error - Empty String"
        Exit Function
    Else
        Dim hashResult As String
        hashResult = ComputeSHA1_BCrypt(str, False)
        
        GetSha1Hash = hashResult
    End If
End Function


Public Function GetSha1HashLeftHalf(ByVal str As String) As String
    If str = "" Then
        GetSha1HashLeftHalf = "Error - Empty String"
        Exit Function
    
    Else
        Dim hashResult As String
        hashResult = ComputeSHA1_BCrypt(str, False)
    
        Dim hash_truncated As String
        hash_truncated = Left(hashResult, 20)
    End If
    GetSha1HashLeftHalf = hash_truncated
End Function


Public Function GetSha1HashBytesHexString(ByVal str As String) As String
    Dim hashBytes As Variant
    hashBytes = GetSha1HashBytes(str)
    Dim i As Integer
    
    Dim hashBytesStr As String
    For i = LBound(hashBytes) To UBound(hashBytes)
        hashBytesStr = hashBytesStr & Right("0" & Hex(hashBytes(i)), 2) & " "
    Next i
    hashBytesStr = LCase(hashBytesStr)
    GetSha1HashBytesHexString = hashBytesStr
End Function


Public Function GetSha1HashBytes(ByVal str As String) As Variant
    Dim hashBytes As Variant
    hashBytes = ComputeSHA1_BCrypt(str, True)
    GetSha1HashBytes = hashBytes
End Function


Function ComputeSHA1_BCrypt(ByVal sInput As String, Optional ByVal returnByteArray As Boolean = False) As Variant
    #If VBA7 Then
        Dim hAlgorithm As LongPtr
        Dim hHash As LongPtr
        Dim lPtrSHA1Wide As LongPtr
    #Else
        Dim hAlgorithm As Long
        Dim hHash As Long
        Dim lPtrSHA1Wide As Long
    #End If
    
    Dim lStatus As Long
    Dim lObjSize As Long
    Dim lHashSize As Long
    Dim lcbResult As Long
    Dim bHashObject() As Byte
    Dim bHashValue() As Byte
    Dim bInput() As Byte
    Dim i As Long

    ' Pointer to the wide string "SHA1"
    lPtrSHA1Wide = StrPtr("SHA1")

    ' Open the SHA1 algorithm provider
    lStatus = BCryptOpenAlgorithmProvider(hAlgorithm, lPtrSHA1Wide, 0, 0)
    If lStatus <> STATUS_SUCCESS Then
        ComputeSHA1_BCrypt = "Error: OpenAlgorithmProvider, Status: " & Hex(lStatus)
        BCryptCloseAlgorithmProvider hAlgorithm, 0
        Exit Function
    End If
    'Debug.Print "Algorithm Handle: " & hAlgorithm

    ' Get the size of the hash object
    lStatus = BCryptGetProperty(hAlgorithm, StrPtr("ObjectLength"), VarPtr(lObjSize), 4, lcbResult, 0)
    If lStatus <> STATUS_SUCCESS Then
        ComputeSHA1_BCrypt = "Error: GetProperty ObjectLength, Status: " & Hex(lStatus)
        BCryptCloseAlgorithmProvider hAlgorithm, 0
        Exit Function
    End If
    ReDim bHashObject(0 To lObjSize - 1)

    ' Hardcode SHA-1 hash length (20 bytes)
    lHashSize = 20
    ReDim bHashValue(0 To lHashSize - 1)

    ' Create the hash object
    lStatus = BCryptCreateHash(hAlgorithm, hHash, VarPtr(bHashObject(0)), lObjSize, 0, 0, 0)
    If lStatus <> STATUS_SUCCESS Then
        ComputeSHA1_BCrypt = "Error: CreateHash, Status: " & Hex(lStatus)
        BCryptCloseAlgorithmProvider hAlgorithm, 0
        Exit Function
    End If
    
    ' Get the input string as a byte array (UTF-8 encoding)
    bInput = StrConv(sInput, vbFromUnicode)

    ' Hash the input data
    lStatus = BCryptHashData(hHash, VarPtr(bInput(0)), UBound(bInput) + 1, 0)
    If lStatus <> STATUS_SUCCESS Then
        ComputeSHA1_BCrypt = "Error: HashData, Status: " & Hex(lStatus)
        BCryptDestroyHash hHash
        BCryptCloseAlgorithmProvider hAlgorithm, 0
        Exit Function
    End If
    
    ' Finalize the hash computation and get the value
    lStatus = BCryptFinishHash(hHash, VarPtr(bHashValue(0)), lHashSize, 0)
    If lStatus <> STATUS_SUCCESS Then
        ComputeSHA1_BCrypt = "Error: FinishHash, Status: " & Hex(lStatus)
        BCryptDestroyHash hHash
        BCryptCloseAlgorithmProvider hAlgorithm, 0
        Exit Function
    End If

    ' Return either byte array or lowercase hexadecimal string
    If returnByteArray Then
        ComputeSHA1_BCrypt = bHashValue
    Else
        Dim sHexHash As String
        For i = 0 To lHashSize - 1
            sHexHash = sHexHash & LCase(Right("0" & Hex(bHashValue(i)), 2))
        Next i
        ComputeSHA1_BCrypt = sHexHash
    End If

    ' Clean up the handles
    BCryptDestroyHash hHash
    BCryptCloseAlgorithmProvider hAlgorithm, 0
End Function