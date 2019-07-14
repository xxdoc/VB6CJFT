Attribute VB_Name = "basFileTransmit"
Option Explicit


'''���ϳ���MD5���ܽ���
'''���ӵ�ַ��https://www.cnblogs.com/youyouran/p/5381050.html

Public Declare Function CryptAcquireContext Lib "advapi32.dll" _
    Alias "CryptAcquireContextA" ( _
    ByRef phProv As Long, _
    ByVal pszContainer As String, _
    ByVal pszProvider As String, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long

Public Declare Function CryptReleaseContext Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal dwFlags As Long) As Long

Public Declare Function CryptCreateHash Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal Algid As Long, _
    ByVal hKey As Long, _
    ByVal dwFlags As Long, _
    ByRef phHash As Long) As Long

Public Declare Function CryptDestroyHash Lib "advapi32.dll" ( _
    ByVal hHash As Long) As Long

Public Declare Function CryptHashData Lib "advapi32.dll" ( _
    ByVal hHash As Long, _
    pbData As Any, _
    ByVal dwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Public Declare Function CryptDeriveKey Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal Algid As Long, _
    ByVal hBaseData As Long, _
    ByVal dwFlags As Long, _
    ByRef phKey As Long) As Long

Public Declare Function CryptDestroyKey Lib "advapi32.dll" ( _
    ByVal hKey As Long) As Long

Public Declare Function CryptEncrypt Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal hHash As Long, _
    ByVal Final As Long, _
    ByVal dwFlags As Long, _
    pbData As Any, _
    ByRef pdwDataLen As Long, _
    ByVal dwBufLen As Long) As Long

Public Declare Function CryptDecrypt Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal hHash As Long, _
    ByVal Final As Long, _
    ByVal dwFlags As Long, _
    pbData As Any, _
    ByRef pdwDataLen As Long) As Long

Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Dest As Any, _
    Src As Any, _
    ByVal Ln As Long)

Private Const PROV_RSA_FULL = 1

Private Const CRYPT_NEWKEYSET = &H8

Private Const ALG_CLASS_HASH = 32768
Private Const ALG_CLASS_DATA_ENCRYPT = 24576&

Private Const ALG_TYPE_ANY = 0
Private Const ALG_TYPE_BLOCK = 1536&
Private Const ALG_TYPE_STREAM = 2048&

Private Const ALG_SID_MD2 = 1
Private Const ALG_SID_MD4 = 2
Private Const ALG_SID_MD5 = 3
Private Const ALG_SID_SHA1 = 4

Private Const ALG_SID_DES = 1
Private Const ALG_SID_3DES = 3
Private Const ALG_SID_RC2 = 2
Private Const ALG_SID_RC4 = 1

Public Enum HASHALGORITHM
   MD2 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2
   MD4 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4
   MD5 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5
   SHA1 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1
End Enum

Public Enum ENCALGORITHM
   DES = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_DES
   [3DES] = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_3DES
   RC2 = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_RC2
   RC4 = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4
End Enum

Dim HexMatrix(15, 15) As Byte

'================================================
'����
'================================================
Public Function EncryptString(ByVal str As String, password As String) As String
    Dim byt() As Byte
    Dim HASHALGORITHM As HASHALGORITHM
    Dim ENCALGORITHM As ENCALGORITHM
On Error GoTo LineERR
    byt = str
    HASHALGORITHM = MD5
    ENCALGORITHM = RC4
    EncryptString = BytesToHex(Encrypt(byt, password, HASHALGORITHM, ENCALGORITHM))
LineERR:
    If Err.Number <> 0 Then
        Call gsAlarmAndLog("�����쳣", False)
    End If
End Function

Public Function EncryptByte(byt() As Byte, password As String) As Byte()
    Dim HASHALGORITHM As HASHALGORITHM
    Dim ENCALGORITHM As ENCALGORITHM
    HASHALGORITHM = MD5
    ENCALGORITHM = RC4
    EncryptByte = Encrypt(byt, password, HASHALGORITHM, ENCALGORITHM)
End Function

Public Function Encrypt(data() As Byte, ByVal password As String, Optional ByVal HASHALGORITHM As HASHALGORITHM = MD5, Optional ByVal ENCALGORITHM As ENCALGORITHM = RC4) As Byte()
    Dim lRes As Long
    Dim hProv As Long
    Dim hHash As Long
    Dim hKey As Long
    Dim lBufLen As Long
    Dim lDataLen As Long
    Dim abData() As Byte
    lRes = CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_FULL, 0)
    If lRes = 0 And Err.LastDllError = &H80090016 Then lRes = CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_NEWKEYSET)
    If lRes <> 0 Then
        lRes = CryptCreateHash(hProv, HASHALGORITHM, 0, 0, hHash)
        If lRes <> 0 Then
            lRes = CryptHashData(hHash, ByVal password, Len(password), 0)
            If lRes <> 0 Then
                lRes = CryptDeriveKey(hProv, ENCALGORITHM, hHash, 0, hKey)
                If lRes <> 0 Then
                    lBufLen = UBound(data) - LBound(data) + 1
                    lDataLen = lBufLen
                    lRes = CryptEncrypt(hKey, 0&, 1, 0, ByVal 0&, lBufLen, 0)
                    If lRes <> 0 Then
                        If lBufLen < lDataLen Then lBufLen = lDataLen
                        ReDim abData(0 To lBufLen - 1)
                        MoveMemory abData(0), data(LBound(data)), lDataLen
                        lRes = CryptEncrypt(hKey, 0&, 1, 0, abData(0), lBufLen, lDataLen)
                        If lRes <> 0 Then
                            If lDataLen <> lBufLen Then ReDim Preserve abData(0 To lBufLen - 1)
                            Encrypt = abData
                        End If
                    End If
                End If
                CryptDestroyKey hKey
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hProv, 0
    End If
    If lRes = 0 Then Err.Raise Err.LastDllError
 End Function
 
 
'================================================
'����
'================================================
Public Function DecryptString(ByVal str As String, password As String) As String
    Dim byt() As Byte
    Dim HASHALGORITHM As HASHALGORITHM
    Dim ENCALGORITHM As ENCALGORITHM
On Error GoTo LineERR
    byt = HexToBytes(str)
    HASHALGORITHM = MD5
    ENCALGORITHM = RC4
    DecryptString = Decrypt(byt, password, HASHALGORITHM, ENCALGORITHM)
LineERR:
    If Err.Number <> 0 Then
        Call gsAlarmAndLog("�����쳣", False)
    End If
End Function

Public Function DecryptByte(byt() As Byte, password As String) As Byte()
    Dim HASHALGORITHM As HASHALGORITHM
    Dim ENCALGORITHM As ENCALGORITHM
    HASHALGORITHM = MD5
    ENCALGORITHM = RC4
    DecryptByte = Decrypt(byt, password, HASHALGORITHM, ENCALGORITHM)
End Function

Public Function Decrypt(data() As Byte, ByVal password As String, Optional ByVal HASHALGORITHM As HASHALGORITHM = MD5, Optional ByVal ENCALGORITHM As ENCALGORITHM = RC4) As Byte()
    Dim lRes As Long
    Dim hProv As Long
    Dim hHash As Long
    Dim hKey As Long
    Dim lBufLen As Long
    Dim abData() As Byte
    lRes = CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_FULL, 0)
    If lRes = 0 And Err.LastDllError = &H80090016 Then lRes = CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_NEWKEYSET)
    If lRes <> 0 Then
        lRes = CryptCreateHash(hProv, HASHALGORITHM, 0, 0, hHash)
        If lRes <> 0 Then
            lRes = CryptHashData(hHash, ByVal password, Len(password), 0)
            If lRes <> 0 Then
                lRes = CryptDeriveKey(hProv, ENCALGORITHM, hHash, 0, hKey)
                If lRes <> 0 Then
                    lBufLen = UBound(data) - LBound(data) + 1
                    ReDim abData(0 To lBufLen - 1)
                    MoveMemory abData(0), data(LBound(data)), lBufLen
                    lRes = CryptDecrypt(hKey, 0&, 1, 0, abData(0), lBufLen)
                    If lRes <> 0 Then
                        ReDim Preserve abData(0 To lBufLen - 1)
                        Decrypt = abData
                    End If
                End If
                CryptDestroyKey hKey
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hProv, 0
    End If
    If lRes = 0 Then Err.Raise Err.LastDllError
End Function

'================================================
'�ֽ���ʮ�������ַ�����ת��
'================================================
Public Function BytesToHex(bits() As Byte) As String
    Dim I As Long
    Dim b
    Dim s As String
    For Each b In bits
        If b < 16 Then
            s = s & "0" & Hex(b)
        Else
            s = s & Hex(b)
        End If
    Next
    BytesToHex = s
End Function

Public Function HexToBytes(sHex As String) As Byte()
    Dim b() As Byte
    Dim rst() As Byte
    Dim I As Long
    Dim n As Long
    Dim m1 As Byte
    Dim m2 As Byte
    If HexMatrix(15, 15) = 0 Then Call MatrixInitialize
    b = StrConv(sHex, vbFromUnicode)
    I = (UBound(b) + 1) / 2 - 1
    ReDim rst(I)
    For I = 0 To UBound(b) Step 2
        If b(I) > 96 Then
            m1 = b(I) - 87
        ElseIf b(I) > 64 Then
            m1 = b(I) - 55
        ElseIf b(I) > 47 Then
            m1 = b(I) - 48
        End If
        If b(I + 1) > 96 Then
            m2 = b(I + 1) - 87
        ElseIf b(I + 1) > 64 Then
            m2 = b(I + 1) - 55
        ElseIf b(I + 1) > 47 Then
            m2 = b(I + 1) - 48
        End If
        rst(n) = HexMatrix(m1, m2)
        n = n + 1
    Next I
    HexToBytes = rst
End Function

Public Sub MatrixInitialize()
    HexMatrix(0, 0) = &H0:    HexMatrix(0, 1) = &H1:    HexMatrix(0, 2) = &H2:    HexMatrix(0, 3) = &H3:    HexMatrix(0, 4) = &H4:    HexMatrix(0, 5) = &H5:    HexMatrix(0, 6) = &H6:    HexMatrix(0, 7) = &H7
    HexMatrix(0, 8) = &H8:    HexMatrix(0, 9) = &H9:    HexMatrix(0, 10) = &HA:   HexMatrix(0, 11) = &HB:   HexMatrix(0, 12) = &HC:   HexMatrix(0, 13) = &HD:   HexMatrix(0, 14) = &HE:   HexMatrix(0, 15) = &HF
    HexMatrix(1, 0) = &H10:   HexMatrix(1, 1) = &H11:   HexMatrix(1, 2) = &H12:   HexMatrix(1, 3) = &H13:   HexMatrix(1, 4) = &H14:   HexMatrix(1, 5) = &H15:   HexMatrix(1, 6) = &H16:   HexMatrix(1, 7) = &H17
    HexMatrix(1, 8) = &H18:   HexMatrix(1, 9) = &H19:   HexMatrix(1, 10) = &H1A:  HexMatrix(1, 11) = &H1B:  HexMatrix(1, 12) = &H1C:  HexMatrix(1, 13) = &H1D:  HexMatrix(1, 14) = &H1E:  HexMatrix(1, 15) = &H1F
    HexMatrix(2, 0) = &H20:   HexMatrix(2, 1) = &H21:   HexMatrix(2, 2) = &H22:   HexMatrix(2, 3) = &H23:   HexMatrix(2, 4) = &H24:   HexMatrix(2, 5) = &H25:   HexMatrix(2, 6) = &H26:   HexMatrix(2, 7) = &H27
    HexMatrix(2, 8) = &H28:   HexMatrix(2, 9) = &H29:   HexMatrix(2, 10) = &H2A:  HexMatrix(2, 11) = &H2B:  HexMatrix(2, 12) = &H2C:  HexMatrix(2, 13) = &H2D:  HexMatrix(2, 14) = &H2E:  HexMatrix(2, 15) = &H2F
    HexMatrix(3, 0) = &H30:   HexMatrix(3, 1) = &H31:   HexMatrix(3, 2) = &H32:   HexMatrix(3, 3) = &H33:   HexMatrix(3, 4) = &H34:   HexMatrix(3, 5) = &H35:   HexMatrix(3, 6) = &H36:   HexMatrix(3, 7) = &H37
    HexMatrix(3, 8) = &H38:   HexMatrix(3, 9) = &H39:   HexMatrix(3, 10) = &H3A:  HexMatrix(3, 11) = &H3B:  HexMatrix(3, 12) = &H3C:  HexMatrix(3, 13) = &H3D:  HexMatrix(3, 14) = &H3E:  HexMatrix(3, 15) = &H3F
    HexMatrix(4, 0) = &H40:   HexMatrix(4, 1) = &H41:   HexMatrix(4, 2) = &H42:   HexMatrix(4, 3) = &H43:   HexMatrix(4, 4) = &H44:   HexMatrix(4, 5) = &H45:   HexMatrix(4, 6) = &H46:   HexMatrix(4, 7) = &H47
    HexMatrix(4, 8) = &H48:   HexMatrix(4, 9) = &H49:   HexMatrix(4, 10) = &H4A:  HexMatrix(4, 11) = &H4B:  HexMatrix(4, 12) = &H4C:  HexMatrix(4, 13) = &H4D:  HexMatrix(4, 14) = &H4E:  HexMatrix(4, 15) = &H4F
    HexMatrix(5, 0) = &H50:   HexMatrix(5, 1) = &H51:   HexMatrix(5, 2) = &H52:   HexMatrix(5, 3) = &H53:   HexMatrix(5, 4) = &H54:   HexMatrix(5, 5) = &H55:   HexMatrix(5, 6) = &H56:   HexMatrix(5, 7) = &H57
    HexMatrix(5, 8) = &H58:   HexMatrix(5, 9) = &H59:   HexMatrix(5, 10) = &H5A:  HexMatrix(5, 11) = &H5B:  HexMatrix(5, 12) = &H5C:  HexMatrix(5, 13) = &H5D:  HexMatrix(5, 14) = &H5E:  HexMatrix(5, 15) = &H5F
    HexMatrix(6, 0) = &H60:   HexMatrix(6, 1) = &H61:   HexMatrix(6, 2) = &H62:   HexMatrix(6, 3) = &H63:   HexMatrix(6, 4) = &H64:   HexMatrix(6, 5) = &H65:   HexMatrix(6, 6) = &H66:   HexMatrix(6, 7) = &H67
    HexMatrix(6, 8) = &H68:   HexMatrix(6, 9) = &H69:   HexMatrix(6, 10) = &H6A:  HexMatrix(6, 11) = &H6B:  HexMatrix(6, 12) = &H6C:  HexMatrix(6, 13) = &H6D:  HexMatrix(6, 14) = &H6E:  HexMatrix(6, 15) = &H6F
    HexMatrix(7, 0) = &H70:   HexMatrix(7, 1) = &H71:   HexMatrix(7, 2) = &H72:   HexMatrix(7, 3) = &H73:   HexMatrix(7, 4) = &H74:   HexMatrix(7, 5) = &H75:   HexMatrix(7, 6) = &H76:   HexMatrix(7, 7) = &H77
    HexMatrix(7, 8) = &H78:   HexMatrix(7, 9) = &H79:   HexMatrix(7, 10) = &H7A:  HexMatrix(7, 11) = &H7B:  HexMatrix(7, 12) = &H7C:  HexMatrix(7, 13) = &H7D:  HexMatrix(7, 14) = &H7E:  HexMatrix(7, 15) = &H7F
    HexMatrix(8, 0) = &H80:   HexMatrix(8, 1) = &H81:   HexMatrix(8, 2) = &H82:   HexMatrix(8, 3) = &H83:   HexMatrix(8, 4) = &H84:   HexMatrix(8, 5) = &H85:   HexMatrix(8, 6) = &H86:   HexMatrix(8, 7) = &H87
    HexMatrix(8, 8) = &H88:   HexMatrix(8, 9) = &H89:   HexMatrix(8, 10) = &H8A:  HexMatrix(8, 11) = &H8B:  HexMatrix(8, 12) = &H8C:  HexMatrix(8, 13) = &H8D:  HexMatrix(8, 14) = &H8E:  HexMatrix(8, 15) = &H8F
    HexMatrix(9, 0) = &H90:   HexMatrix(9, 1) = &H91:   HexMatrix(9, 2) = &H92:   HexMatrix(9, 3) = &H93:   HexMatrix(9, 4) = &H94:   HexMatrix(9, 5) = &H95:   HexMatrix(9, 6) = &H96:   HexMatrix(9, 7) = &H97
    HexMatrix(9, 8) = &H98:   HexMatrix(9, 9) = &H99:   HexMatrix(9, 10) = &H9A:  HexMatrix(9, 11) = &H9B:  HexMatrix(9, 12) = &H9C:  HexMatrix(9, 13) = &H9D:  HexMatrix(9, 14) = &H9E:  HexMatrix(9, 15) = &H9F
    HexMatrix(10, 0) = &HA0:  HexMatrix(10, 1) = &HA1:  HexMatrix(10, 2) = &HA2:  HexMatrix(10, 3) = &HA3:  HexMatrix(10, 4) = &HA4:  HexMatrix(10, 5) = &HA5:  HexMatrix(10, 6) = &HA6:  HexMatrix(10, 7) = &HA7
    HexMatrix(10, 8) = &HA8:  HexMatrix(10, 9) = &HA9:  HexMatrix(10, 10) = &HAA: HexMatrix(10, 11) = &HAB: HexMatrix(10, 12) = &HAC: HexMatrix(10, 13) = &HAD: HexMatrix(10, 14) = &HAE: HexMatrix(10, 15) = &HAF
    HexMatrix(11, 0) = &HB0:  HexMatrix(11, 1) = &HB1:  HexMatrix(11, 2) = &HB2:  HexMatrix(11, 3) = &HB3:  HexMatrix(11, 4) = &HB4:  HexMatrix(11, 5) = &HB5:  HexMatrix(11, 6) = &HB6:  HexMatrix(11, 7) = &HB7
    HexMatrix(11, 8) = &HB8:  HexMatrix(11, 9) = &HB9:  HexMatrix(11, 10) = &HBA: HexMatrix(11, 11) = &HBB: HexMatrix(11, 12) = &HBC: HexMatrix(11, 13) = &HBD: HexMatrix(11, 14) = &HBE: HexMatrix(11, 15) = &HBF
    HexMatrix(12, 0) = &HC0:  HexMatrix(12, 1) = &HC1:  HexMatrix(12, 2) = &HC2:  HexMatrix(12, 3) = &HC3:  HexMatrix(12, 4) = &HC4:  HexMatrix(12, 5) = &HC5:  HexMatrix(12, 6) = &HC6:  HexMatrix(12, 7) = &HC7
    HexMatrix(12, 8) = &HC8:  HexMatrix(12, 9) = &HC9:  HexMatrix(12, 10) = &HCA: HexMatrix(12, 11) = &HCB: HexMatrix(12, 12) = &HCC: HexMatrix(12, 13) = &HCD: HexMatrix(12, 14) = &HCE: HexMatrix(12, 15) = &HCF
    HexMatrix(13, 0) = &HD0:  HexMatrix(13, 1) = &HD1:  HexMatrix(13, 2) = &HD2:  HexMatrix(13, 3) = &HD3:  HexMatrix(13, 4) = &HD4:  HexMatrix(13, 5) = &HD5:  HexMatrix(13, 6) = &HD6:  HexMatrix(13, 7) = &HD7
    HexMatrix(13, 8) = &HD8:  HexMatrix(13, 9) = &HD9:  HexMatrix(13, 10) = &HDA: HexMatrix(13, 11) = &HDB: HexMatrix(13, 12) = &HDC: HexMatrix(13, 13) = &HDD: HexMatrix(13, 14) = &HDE: HexMatrix(13, 15) = &HDF
    HexMatrix(14, 0) = &HE0:  HexMatrix(14, 1) = &HE1:  HexMatrix(14, 2) = &HE2:  HexMatrix(14, 3) = &HE3:  HexMatrix(14, 4) = &HE4:  HexMatrix(14, 5) = &HE5:  HexMatrix(14, 6) = &HE6:  HexMatrix(14, 7) = &HE7
    HexMatrix(14, 8) = &HE8:  HexMatrix(14, 9) = &HE9:  HexMatrix(14, 10) = &HEA: HexMatrix(14, 11) = &HEB: HexMatrix(14, 12) = &HEC: HexMatrix(14, 13) = &HED: HexMatrix(14, 14) = &HEE: HexMatrix(14, 15) = &HEF
    HexMatrix(15, 0) = &HF0:  HexMatrix(15, 1) = &HF1:  HexMatrix(15, 2) = &HF2:  HexMatrix(15, 3) = &HF3:  HexMatrix(15, 4) = &HF4:  HexMatrix(15, 5) = &HF5:  HexMatrix(15, 6) = &HF6:  HexMatrix(15, 7) = &HF7
    HexMatrix(15, 8) = &HF8:  HexMatrix(15, 9) = &HF9:  HexMatrix(15, 10) = &HFA: HexMatrix(15, 11) = &HFB: HexMatrix(15, 12) = &HFC: HexMatrix(15, 13) = &HFD: HexMatrix(15, 14) = &HFE: HexMatrix(15, 15) = &HFF
End Sub

'''===============================================================================
'Ҫ��Winsock�ؼ��ڿͻ��������˶����뽨�����飬����Indexֵ���Ӧ������������±�Ҫ��ͬ
'''===============================================================================

Public Function gfAppExist(ByVal strName As String) As Boolean
    'ָ��Ӧ�ó�������Ƿ����
    
    Dim RetVal As Long
    Dim objWMIService As Object
    Dim colProcessList As Object
    Dim objProcess As Object
    
    On Error GoTo LineERR
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("select * from Win32_Process where Name='" & strName & "' ")
    For Each objProcess In colProcessList
        gfAppExist = True   '���ڸý�����ʱ
    Next
    
LineERR:
    Set objProcess = Nothing
    Set colProcessList = Nothing
    Set objWMIService = Nothing
End Function

Public Function gfBackFileName(Optional ByVal enumType As genumCharType = udUpperCase, Optional ByVal lngLen As Long = 30) As String
    '����lngLen������ַ���ɵ��ļ���
    Dim K As Long
    Dim strName As String
    
    For K = 1 To lngLen
        strName = strName & gfBackOneChar(enumType)
    Next
    gfBackFileName = strName
End Function

Public Function gfBackVersion(ByVal strFile As String) As String
    '�����ļ��İ汾��
    Dim objFile As Scripting.FileSystemObject
    
    If Not gfDirFile(strFile) Then Exit Function
    Set objFile = New FileSystemObject
    gfBackVersion = objFile.GetFileVersion(strFile)

    Set objFile = Nothing
End Function


Public Function gfCheckIP(ByVal strIP As String) As String
    Dim K As Long
    Dim arrIP() As String
    
    arrIP = Split(strIP, ".")
    If UBound(arrIP) <> 3 Then GoTo LineOver
    For K = 0 To 3
        If Not IsNumeric(arrIP(K)) Then GoTo LineOver
        If Val(arrIP(K)) < 0 Or Val(arrIP(K)) > 255 Then GoTo LineOver
        arrIP(K) = CStr(Val(arrIP(K)))
    Next
    gfCheckIP = arrIP(0) & "." & arrIP(1) & "." & arrIP(2) & "." & arrIP(3)
    Exit Function
    
LineOver:
    gfCheckIP = "127.0.0.1"
End Function

Public Function gfCloseApp(ByVal strName As String) As Boolean
    '�ر�ָ��Ӧ�ó������
    
    Dim winHwnd As Long
    Dim RetVal As Long
    Dim objWMIService As Object
    Dim colProcessList As Object
    Dim objProcess As Object
    
    On Error GoTo LineERR
    
''    winHwnd = FindWindow(vbNullString, strName) '���Ҵ��ڣ�strName���ݼ��������Ͽ����Ĵ��ڱ���
''    If winHwnd <> 0 Then    '��Ϊ0��ʾ�ҵ�����
''        RetVal = PostMessage(winHwnd, WM_CLOSE, 0&, 0&) '���͹رմ�����Ϣ,����ֵΪ0��ʾ�ر�ʧ��
''    End If
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("select * from Win32_Process where Name='" & strName & "' ")
    For Each objProcess In colProcessList
        RetVal = objProcess.Terminate
        If RetVal <> 0 Then Exit Function   '���۲�=0ʱ�رս��̳ɹ������ɹ�ʱ����ֵ��Ϊ��
    Next
    
    gfCloseApp = True   'ȫ���رճɹ��򲻴��ڸý�����ʱ
    
LineERR:
    Set objWMIService = Nothing
    Set colProcessList = Nothing
    Set objProcess = Nothing
End Function


Public Function gfDirFile(ByVal strFile As String) As Boolean
    Dim strDir As String
    
    strFile = Trim(strFile)
    If Len(strFile) = 0 Then Exit Function
    
    On Error GoTo LineERR
    
    strDir = Dir(strFile, vbHidden + vbReadOnly + vbSystem)
    If Len(strDir) > 0 Then
        SetAttr strFile, vbNormal
        gfDirFile = True
    End If
    
    Exit Function
LineERR:
    Debug.Print "Error:gfDirFile--" & Err.Number & "  " & Err.Description & ";" & strFile
    Call gsAlarmAndLog("�ļ�·��ʶ���쳣", False)
End Function

Public Function gfDirFolder(ByVal strFolder As String) As Boolean
    Dim strDir As String
    
    strFolder = Trim(strFolder)
    If Len(strFolder) = 0 Then Exit Function
    
    On Error GoTo LineERR
    
    strDir = Dir(strFolder, vbHidden + vbReadOnly + vbSystem + vbDirectory)
    If Len(strDir) = 0 Then
        MkDir strFolder
    Else
        SetAttr strFolder, vbNormal
    End If
    gfDirFolder = True
    
    Exit Function
LineERR:
    Debug.Print "Error:gfDirFolder--" & Err.Number & "  " & Err.Description & ";" & strFolder
    Call gsAlarmAndLog("�ļ���·��ʶ���쳣", False)
End Function

Public Function gfDatabaseInfoJoin(Optional ByVal blnJoin As Boolean = True) As String
    '���ݿ�������Ϣ���ܺ�ƴ��
    
    If Not blnJoin Then Exit Function
    With gVar
        gfDatabaseInfoJoin = .PTDBDataSource & EncryptString(.ConSource, .EncryptKey) & _
                             .PTDBDatabase & EncryptString(.ConDatabase, .EncryptKey) & _
                             .PTDBUserID & EncryptString(.ConUserID, .EncryptKey) & _
                             .PTDBPassword & EncryptString(.ConPassword, .EncryptKey)
    End With
    
End Function

Public Function gfFileInfoJoin(ByVal intIndex As Integer, Optional ByVal enmType As genumFileTransimitType = ftSend) As String
    '�ļ���Ϣƴ��
    Dim strType As String
    
    strType = IIf(enmType = ftReceive, gVar.PTFileReceive, gVar.PTFileSend) 'ȷ���ļ��������͡�վ�ڿͻ��˽Ƕ�ȷ����
    With gArr(intIndex)
        gfFileInfoJoin = gVar.PTFileFolder & .FileFolder & gVar.PTFileName & .FileName & gVar.PTFileSize & .FileSizeTotal & strType
    End With
    
End Function

Public Function gfLoadPicture(ByRef ImageLoad As VB.Image, ByVal strPath As String) As Boolean
    '����ͼƬ
    
    On Error GoTo LineERR
    
    ImageLoad.Picture = LoadPicture(strPath)
    gfLoadPicture = True
    
    Exit Function
LineERR:
    Call gsAlarmAndLog("����ͼƬ�쳣")
End Function

Public Function gfNotifyIconAdd(ByRef frmCur As Form) As Boolean
    '��������ͼ��
    With gNotifyIconData
        .hwnd = frmCur.hwnd
        .uID = frmCur.Icon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP Or NIF_INFO
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = frmCur.Icon.Handle
        .szTip = App.Title & " " & App.Major & "." & App.Minor & _
            "." & App.Revision & vbNullChar   '����ƶ�����ͼ��ʱ��ʾ��Tip��Ϣ
        .cbSize = Len(gNotifyIconData)
    End With
    Call Shell_NotifyIcon(NIM_ADD, gNotifyIconData)
End Function

Public Function gfNotifyIconBalloon(ByRef frmCur As Form, ByVal BalloonInfo As String, _
    ByVal BalloonTitle As String, Optional IconFlag As genumNotifyIconFlag = NIIF_INFO) As Boolean
    '����ͼ�굯��������Ϣ
    With gNotifyIconData
        .dwInfoFlags = IconFlag
        .szInfoTitle = BalloonTitle & vbNullChar
        .szInfo = BalloonInfo & vbNullChar
        .cbSize = Len(gNotifyIconData)
    End With
    Call gfNotifyIconModify(gNotifyIconData)
End Function

Public Function gfNotifyIconDelete(ByRef frmCur As Form) As Boolean
    'ɾ������ͼ��
    Call Shell_NotifyIcon(NIM_DELETE, gNotifyIconData)
End Function

Public Function gfNotifyIconModify(nfIconData As gtypeNOTIFYICONDATA) As Boolean
    '�޸�����ͼ����Ϣ
    gNotifyIconData = nfIconData
    Call Shell_NotifyIcon(NIM_MODIFY, gNotifyIconData)
End Function

Public Function gfRegOperate(ByVal RegHKEY As genumRegRootDirectory, ByVal lpSubKey As String, _
    ByVal lpValueName As String, Optional ByVal lpType As genumRegDataType = REG_SZ, _
    Optional ByRef lpValue As String, Optional ByVal lpOp As genumRegOperateType = RegRead) As Boolean
    '
    Dim Ret As Long, hKey As Long, lngLength As Long
    Dim Buff() As Byte
    
    
    Ret = RegOpenKey(RegHKEY, lpSubKey, hKey)
    If Ret = 0 Then
        Select Case lpOp
            Case RegDelete
                Ret = RegDeleteValue(hKey, lpValueName)
                If Ret = 0 Then
                    gfRegOperate = True
                End If
                
            Case RegWrite
                lngLength = LenB(StrConv(lpValue, vbFromUnicode))   '����LenB��StrConv�Ļ�lpValue�ַ������ȶԲ���
                Ret = RegSetValueEx(hKey, lpValueName, 0, lpType, ByVal lpValue, lngLength)
                If Ret = 0 Then
                    gfRegOperate = True
'Debug.Print "W", lpValue, lngLength
                End If
                
            Case Else
                Ret = RegQueryValueEx(hKey, lpValueName, 0, lpType, ByVal 0, lngLength) '��ȡֵ�ĳ���
                If Ret = 0 And lngLength > 0 Then
                    ReDim Buff(lngLength - 1)   '�ض��建���С
                    Ret = RegQueryValueEx(hKey, lpValueName, 0, lpType, Buff(0), lngLength) 'ȡֵ
                    If Ret = 0 And lngLength > 1 Then
                        ReDim Preserve Buff(lngLength - 2)
                        lpValue = StrConv(Buff, vbUnicode)
                        gfRegOperate = True
'Debug.Print "R", lpValue, lngLength - 1
                    End If
                End If
                
        End Select
    End If
    
    Call RegCloseKey(hKey)
    
End Function

Public Function gfRestoreDBInfo(ByVal strInfo As String) As Boolean
    '��ԭ���ݿ�������Ϣ
    Dim lngSrc As Long, lngDB As Long, lngID As Long, lngPWD As Long
    Dim strSrc As String, strDB As String, strID As String, strPWD As String
        
    On Error Resume Next    '���ܷ���������ʱ���ܳ���
    
    With gVar
        lngSrc = InStr(strInfo, .PTDBDataSource) '���ݿ��������ַ
        lngDB = InStr(strInfo, .PTDBDatabase)   '���ݿ���
        lngID = InStr(strInfo, .PTDBUserID)     '���ݿ�����˺�
        lngPWD = InStr(strInfo, .PTDBPassword)  '���ݿ��������
        
        If Not (lngSrc > 0 And lngDB > lngSrc And lngID > lngDB And lngPWD > lngID) Then Exit Function '��Ϣ�е�˳�򲻶�
        strSrc = Mid(strInfo, Len(.PTDBDataSource) + lngSrc, lngDB - lngSrc - Len(.PTDBDataSource))
        .ConSource = DecryptString(strSrc, .EncryptKey)
        
        strDB = Mid(strInfo, lngDB + Len(.PTDBDatabase), lngID - lngDB - Len(.PTDBDatabase))
        .ConDatabase = DecryptString(strDB, .EncryptKey)
        
        strID = Mid(strInfo, lngID + Len(.PTDBUserID), lngPWD - lngID - Len(.PTDBUserID))
        .ConUserID = DecryptString(strID, .EncryptKey)
        
        strPWD = Mid(strInfo, lngPWD + Len(.PTDBPassword))
        .ConPassword = DecryptString(strPWD, .EncryptKey)
                     
        .ConString = "Provider=SQLOLEDB.1;" & _
                     "Persist Security Info=False;" & _
                     "User ID=" & .ConUserID & ";Password=" & .ConPassword & ";" & _
                     "Initial Catalog=" & .ConDatabase & ";" & _
                     "Data Source=" & .ConSource & ";"
    End With
    If Err.Number = 0 Then gfRestoreDBInfo = True
End Function

Public Function gfRestoreInfo(ByVal strInfo As String, sckGet As MSWinsockLib.Winsock) As Boolean
    '��ԭ���յ����ļ���Ϣ
    Dim lngFod As Long, lngFile As Long, lngSize As Long
    Dim lngSend As Long, lngReceive As Long, lngType As Long
    Dim strFod As String, strSize As String, strType As String
            
    With gArr(sckGet.Index)
        If InStr(strInfo, gVar.PTFileFolder) > 0 Then   'һ���ļ���
            lngFod = InStr(strInfo, gVar.PTFileFolder)
            lngFile = InStr(strInfo, gVar.PTFileName)
            lngSize = InStr(strInfo, gVar.PTFileSize)
            lngSend = InStr(strInfo, gVar.PTFileSend)
            lngReceive = InStr(strInfo, gVar.PTFileReceive)
            
            If lngFile > 0 Then '�����ļ���
                gArr(sckGet.Index) = gArr(0)    '�ȳ�ʼ���ļ��������Ϊ����Ϣ
                
                If (lngSend > 0 And lngReceive > 0) Or (lngSend = 0 And lngReceive = 0) Then
                    Exit Function   '�����뷢��Э�����ҽ�������һ������
                End If
                strType = IIf(lngSend > 0, gVar.PTFileSend, gVar.PTFileReceive)
                lngType = IIf(lngSend > 0, lngSend, lngReceive)
                
                .FileFolder = Mid(strInfo, lngFod + Len(gVar.PTFileFolder), lngFile - (lngFod + Len(gVar.PTFileFolder)))
                strFod = gVar.AppPath & IIf(Len(.FileFolder) = 0, "", .FileFolder & "\")
                If Not gfDirFolder(strFod) Then Exit Function
                
                .FileName = Mid(strInfo, lngFile + Len(gVar.PTFileName), lngSize - (lngFile + Len(gVar.PTFileName)))
                
                strSize = Mid(strInfo, lngSize + Len(gVar.PTFileSize), lngType - (lngSize + Len(gVar.PTFileSize)))
                If Not IsNumeric(strSize) Then Exit Function
                
                If strType <> Mid(strInfo, lngType) Then Exit Function
                
                If strType = gVar.PTFileSend Then   '�ļ������նˡ��򡾷��Ͷˡ����ݿ�ʼ����ָ�
                    .FileSizeTotal = CLng(strSize)
                    .FilePath = strFod & .FileName
                    Call gfSendInfo(gVar.PTFileStart, sckGet)
                    .FileTransmitState = True
                    
                ElseIf strType = gVar.PTFileReceive Then    '�ļ������Ͷˡ����յ������նˡ�����������Ҫָ���ļ���ָ�
                    .FilePath = strFod & .FileName
                    If gfDirFile(.FilePath) Then
                        .FileSizeTotal = FileLen(.FilePath)
                        Call gfSendInfo(gVar.PTFileExist & gVar.PTFileSize & .FileSizeTotal, sckGet)    '֪ͨ���ն��ļ�����
                        Call gfSendInfo(gfFileInfoJoin(sckGet.Index, ftSend), sckGet)   '�ٴ��ط������ն���Ҫ���ļ���Ϣ��Ȩ��֮���÷���
                    Else
                        gArr(sckGet.Index) = gArr(0)
                        Call gfSendInfo(gVar.PTFileNoExist, sckGet)
                    End If
                End If
                gfRestoreInfo = True
            End If
        ElseIf InStr(strInfo, gVar.PTVersionNotUpdate) > 0 Then '������
            
        End If
    End With

End Function

Public Function gfSaveFile(ByRef frmSave As Form) As String
    '�����ļ���Ϣ�����ݿ���
    Dim rsFile As ADODB.Recordset
    Dim strSQL As String, strFileID As String, strMsg As String
    
    strSQL = "SELECT FileID ,FileClassify ,FileExtension ,FileOldName ,FileSaveName ,FileSize ," & _
             "FileSaveLocation ,FileUploadMen ,FileUploadTime FROM tb_FT_Lib_File   " & _
             "WHERE FileSaveName ='" & gVar.FTUploadFileNameNew & "' AND FileSaveLocation ='" & gVar.FTUploadFileFolder & "' "
    Set rsFile = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
    If rsFile.State = adStateOpen Then
        If rsFile.RecordCount > 0 Then
            rsFile.Close
            MsgBox "�ļ���Ϣ�ڿ����Ѵ��ڣ��������ϴ���", vbCritical, "����"
        Else
            On Error GoTo LineERR
            rsFile.AddNew
            rsFile.Fields("FileClassify") = gVar.FTUploadFileClassify
            rsFile.Fields("FileExtension") = gVar.FTUploadFileExtension
            rsFile.Fields("FileOldName") = gVar.FTUploadFileNameOld
            rsFile.Fields("FileSaveName") = gVar.FTUploadFileNameNew
            rsFile.Fields("FileSize") = gVar.FTUploadFileSize
            rsFile.Fields("FileSaveLocation") = gVar.FTUploadFileFolder
            rsFile.Fields("FileUploadMen") = gVar.UserFullName
            rsFile.Fields("FileUploadTime") = Now
            rsFile.Update
            strFileID = CStr(rsFile.Fields("FileID"))    '��ȡID
            rsFile.Close
            strMsg = "�ϴ��ļ���" & strFileID & "����" & gVar.FTUploadFileNameNew & "��"
            Call gsLogAdd(frmSave, udInsert, "tb_FT_Lib_File", strMsg)
            gfSaveFile = strFileID  '�����ɵ��ļ�IDֵ���ø������ķ���ֵ
        End If
    End If
LineERR:
    If Not rsFile Is Nothing Then If rsFile.State = adStateOpen Then rsFile.Close
    Set rsFile = Nothing
    If Err.Number > 0 Then
        Call gsAlarmAndLog("�ϴ��ļ��쳣")
    End If
End Function

Public Function gfSendClientInfo(ByVal strPC As String, ByVal strLogin As String, _
        ByVal FullName As String, ByRef sckInfo As MSWinsockLib.Winsock) As Boolean
    '���Ϳͻ�����Ϣ������ˡ�����˽��պ���ʾ�������б���
    
    Dim strInfo As String
    
    strInfo = gVar.PTClientUserComputerName & strPC & gVar.PTClientUserLoginName & strLogin & gVar.PTClientUserFullName & FullName
    If gfSendInfo(strInfo, sckInfo) Then gfSendClientInfo = True
    
End Function


Public Function gfSendFile(ByVal strFile As String, sckSend As MSWinsockLib.Winsock) As Boolean
    Dim lngSendSize As Long, lngRemain As Long
    Dim byteSend() As Byte
    
    With gArr(sckSend.Index)
        If .FileNumber = 0 Then
            .FileNumber = FreeFile
            Open strFile For Binary As #.FileNumber
            .FileTransmitState = True
        End If
        
        lngSendSize = gVar.FTChunkSize
        lngRemain = .FileSizeTotal - Loc(.FileNumber)
        If lngSendSize > lngRemain Then lngSendSize = lngRemain
        
        ReDim byteSend(lngSendSize - 1)
        Get #.FileNumber, , byteSend
        sckSend.SendData byteSend
        
        .FileSizeCompleted = .FileSizeCompleted + lngSendSize
        If .FileSizeCompleted = .FileSizeTotal Then Close #.FileNumber
        
    End With
    
End Function

Public Function gfSendInfo(ByVal strInfo As String, sckSend As MSWinsockLib.Winsock) As Boolean
    If sckSend.State = 7 Then
        sckSend.SendData strInfo
        DoEvents    '�ƺ����������
        gfSendInfo = True
    End If
End Function

Public Function gfShell(ByVal strFile As String, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus) As Boolean
    '����Shell�����쳣
    
    Dim Ret
    
    On Error Resume Next
    
    Ret = Shell(strFile, WindowStyle)

    If Ret > 0 Then gfShell = True
    
End Function

Public Function gfShellExecute(ByVal strFile As String) As Boolean
    'ִ�г������ļ����ļ���
    '''Call ShellExecute(Me.hwnd, "open", strFile, vbNullString, vbNullString, 1)

    Dim lngRet As Long
    Dim strDir As String
    
    lngRet = ShellExecute(GetDesktopWindow, "open", strFile, vbNullString, vbNullString, vbNormalFocus)

    ' û�й����ĳ���
    If lngRet = SE_ERR_NOASSOC Then
         strDir = Space$(260)
         lngRet = GetSystemDirectory(strDir, Len(strDir))
         strDir = Left$(strDir, lngRet)
       ' ��ʾ�򿪷�ʽ����
         lngRet = ShellExecute(GetDesktopWindow, vbNullString, "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & strFile, strDir, vbNormalFocus)
    End If
    
    If lngRet > 32 Then gfShellExecute = True
    
End Function


Public Function gfStartUpSet(Optional ByVal blnSet As Boolean = True, Optional ByVal OpType As genumRegOperateType = RegRead) As Boolean
    '��������������������������ȡ����ӡ�ɾ��
    
    Dim strReg As String, strCur As String
    Dim blnReg As Boolean
    
    If Not blnSet Then Exit Function    '�������κβ���
    
    strCur = Chr(34) & gVar.AppPath & App.EXEName & ".exe" & Chr(34) & "-s"
    blnReg = gfRegOperate(HKEY_LOCAL_MACHINE, HKEY_USER_RUN, App.EXEName, REG_SZ, strReg, RegRead)
    If blnReg Then
        If LCase(strCur) = LCase(strReg) Then   '�Ѵ���
            gfStartUpSet = True
        Else    '������
            blnReg = False
'''Debug.Print LCase(strCur),LCase(strReg)
        End If
    End If
    If OpType = RegWrite Then   '���������
        If blnReg Then
            gfStartUpSet = True
        Else
            blnReg = gfRegOperate(HKEY_LOCAL_MACHINE, HKEY_USER_RUN, App.EXEName, REG_SZ, strCur, RegWrite)
            If blnReg Then
                gfStartUpSet = True
            Else
                Call gsAlarmAndLog("���ÿ����Զ�������ʧ�ܣ�")
                gfStartUpSet = False
            End If
        End If
    ElseIf OpType = RegDelete Then  'ɾ��������
        If blnReg Then
            blnReg = gfRegOperate(HKEY_LOCAL_MACHINE, HKEY_USER_RUN, App.EXEName, REG_SZ, strCur, RegDelete)
            If blnReg Then
                gfStartUpSet = True
            Else
                Call gsAlarmAndLog("ɾ�������Զ�������ʧ�ܣ�")
                gfStartUpSet = False
            End If
        Else
            gfStartUpSet = True
        End If
    End If
End Function

Public Function gfVersionCompare(ByVal strVerCL As String, ByVal strVerSV As String) As String
    '�¾ɰ汾�űȽ�
    Dim ArrCL() As String, ArrSV() As String
    Dim K As Long, C As Long
    
    ArrCL = Split(strVerCL, ".")
    ArrSV = Split(strVerSV, ".")
    K = UBound(ArrCL)
    C = UBound(ArrSV)
    If K = C And K = 3 Then
        For K = 0 To C
            If Not IsNumeric(ArrCL(K)) Then
                gfVersionCompare = "�ͻ��˰汾�쳣" '����ֵ
                Exit For
            End If
            If Not IsNumeric(ArrSV(K)) Then
                gfVersionCompare = "����˰汾�쳣" '����ֵ
                Exit For
            End If
            
            If Val(ArrSV(K)) > Val(ArrCL(K)) Then
                gfVersionCompare = "1" ''����ֵ˵�����°汾
                Exit For
            End If
        Next
        If K = C + 1 Then gfVersionCompare = "0" ''����ֵ˵��û���°棬���ø���
    Else
        If K = 3 And C <> 3 Then
            gfVersionCompare = "����˰汾��ȡ�쳣" '����ֵ
        ElseIf C = 3 And K <> 3 Then
            gfVersionCompare = "�ͻ��˰汾��ȡ�쳣" '����ֵ
        Else
            gfVersionCompare = "�汾��ȡ�쳣"   '����ֵ
        End If
    End If
    
End Function


Public Sub gsConnectToServer(ByRef sckCon As MSWinsockLib.Winsock, Optional ByVal blnConnect As Boolean = False)
    '�����������������
    
    If Not blnConnect Then Exit Sub
    With sckCon
        If .State <> 0 Then .Close
        .RemoteHost = gVar.TCPSetIP
        .RemotePort = gVar.TCPSetPort
        .Connect
    End With
End Sub

Public Sub gsFormEnable(frmCur As Form, Optional ByVal blnState As Boolean = False)
    With frmCur
        If blnState Then
            .Enabled = True
            .MousePointer = 0
        Else
            .Enabled = False
            .MousePointer = 13
        End If
    End With
End Sub

Public Sub gsLoadFileInfo(Optional ByVal arrIndex As Long = 1, Optional ByVal blnUpload As Boolean = True)
    '�����ϴ��������ļ���Ϣ
    
    With gArr(arrIndex)
        If blnUpload Then   '�ϴ��ļ���Ϣ
            .FilePath = gVar.FTUploadFilePath       '�����Ͷˡ������͵��ļ�·��
            .FileName = gVar.FTUploadFileNameNew    '�����Ͷˡ����͹�ȥ���ļ��ڡ����նˡ�������ļ���
            .FileFolder = gVar.FTUploadFileFolder   '�����Ͷˡ����͹�ȥ���ļ��ڡ����նˡ�������ļ�������
            .FileSizeTotal = gVar.FTUploadFileSize  '�����Ͷˡ������͵��ļ���С�򡾽��նˡ����յ��ļ���С
            .FileTransmitNotOver = True     '�����Ͷˡ�����δ������ʶ
            gVar.FTUploadOrDownload = True  '�����Ͷˡ��ϴ�״̬
        Else    '�����ļ���Ϣ
            .FilePath = gVar.FTDownloadFilePath
            .FileName = gVar.FTDownloadFileNameNew
            .FileFolder = gVar.FTDownloadFileFolder
            .FileSizeTotal = gVar.FTDownloadFileSize
            .FileTransmitNotOver = True
            gVar.FTUploadOrDownload = False '����״̬
        End If
    End With
End Sub

