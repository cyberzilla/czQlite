VERSION 5.00
Begin VB.UserControl czQlite 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   420
   ScaleWidth      =   420
End
Attribute VB_Name = "czQlite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' czQlite.ctl - Single-Instance SQLite UserControl for VB6
'
' Drop onto Form -> Set DatabasePath -> Query/Execute. That's it.
'
' Features:
'   SQLite Engine  - via sqlite3.dll (ship alongside EXE, or use winsqlite3.dll on Win10+)
'   JSON Engine    - Adapted from mdJson.bas by wqweto (Phase 2)
'   File Encrypt   - AES-256 via bcrypt.dll (Phase 3)
'   Zero dependency. Single file. Invisible at runtime.
'
' Author: Cyberzilla
' License: MIT
'==============================================================================
Option Explicit

'==============================================================================
' SQLITE3 API DECLARES (sqlite3.dll - download from sqlite.org)
' All handles/pointers are Long (VB6 x86 = 32-bit)
'==============================================================================
'--- Lifecycle
Private Declare Function sqlite3_open_v2 Lib "sqlite3" (ByVal filename As Long, ppDb As Long, ByVal flags As Long, ByVal zVfs As Long) As Long
Private Declare Function sqlite3_close_v2 Lib "sqlite3" (ByVal db As Long) As Long
'--- Prepare / Execute
Private Declare Function sqlite3_prepare_v2 Lib "sqlite3" (ByVal db As Long, ByVal zSql As Long, ByVal nByte As Long, ppStmt As Long, pzTail As Long) As Long
Private Declare Function sqlite3_step Lib "sqlite3" (ByVal stmt As Long) As Long
Private Declare Function sqlite3_finalize Lib "sqlite3" (ByVal stmt As Long) As Long
Private Declare Function sqlite3_reset Lib "sqlite3" (ByVal stmt As Long) As Long
'--- Bind Parameters
Private Declare Function sqlite3_bind_text Lib "sqlite3" (ByVal stmt As Long, ByVal idx As Long, ByVal pText As Long, ByVal nBytes As Long, ByVal xDel As Long) As Long
Private Declare Function sqlite3_bind_int Lib "sqlite3" (ByVal stmt As Long, ByVal idx As Long, ByVal value As Long) As Long
Private Declare Function sqlite3_bind_int64 Lib "sqlite3" (ByVal stmt As Long, ByVal idx As Long, ByVal value As Currency) As Long
Private Declare Function sqlite3_bind_double Lib "sqlite3" (ByVal stmt As Long, ByVal idx As Long, ByVal value As Double) As Long
Private Declare Function sqlite3_bind_blob Lib "sqlite3" (ByVal stmt As Long, ByVal idx As Long, ByVal pBlob As Long, ByVal nBytes As Long, ByVal xDel As Long) As Long
Private Declare Function sqlite3_bind_null Lib "sqlite3" (ByVal stmt As Long, ByVal idx As Long) As Long
Private Declare Function sqlite3_bind_parameter_count Lib "sqlite3" (ByVal stmt As Long) As Long
Private Declare Function sqlite3_bind_parameter_index Lib "sqlite3" (ByVal stmt As Long, ByVal zName As Long) As Long
Private Declare Function sqlite3_clear_bindings Lib "sqlite3" (ByVal stmt As Long) As Long
'--- Read Columns
Private Declare Function sqlite3_column_count Lib "sqlite3" (ByVal stmt As Long) As Long
Private Declare Function sqlite3_column_type Lib "sqlite3" (ByVal stmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_name Lib "sqlite3" (ByVal stmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_text Lib "sqlite3" (ByVal stmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_int Lib "sqlite3" (ByVal stmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_int64 Lib "sqlite3" (ByVal stmt As Long, ByVal iCol As Long) As Currency
Private Declare Function sqlite3_column_double Lib "sqlite3" (ByVal stmt As Long, ByVal iCol As Long) As Double
Private Declare Function sqlite3_column_blob Lib "sqlite3" (ByVal stmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_bytes Lib "sqlite3" (ByVal stmt As Long, ByVal iCol As Long) As Long
'--- Error / Info
Private Declare Function sqlite3_errmsg Lib "sqlite3" (ByVal db As Long) As Long
Private Declare Function sqlite3_errcode Lib "sqlite3" (ByVal db As Long) As Long
Private Declare Function sqlite3_changes Lib "sqlite3" (ByVal db As Long) As Long
Private Declare Function sqlite3_last_insert_rowid Lib "sqlite3" (ByVal db As Long) As Currency
Private Declare Function sqlite3_libversion Lib "sqlite3" () As Long
Private Declare Function sqlite3_busy_timeout Lib "sqlite3" (ByVal db As Long, ByVal ms As Long) As Long
'--- Backup
Private Declare Function sqlite3_backup_init Lib "sqlite3" (ByVal pDest As Long, ByVal zDestName As Long, ByVal pSource As Long, ByVal zSourceName As Long) As Long
Private Declare Function sqlite3_backup_step Lib "sqlite3" (ByVal pBackup As Long, ByVal nPage As Long) As Long
Private Declare Function sqlite3_backup_finish Lib "sqlite3" (ByVal pBackup As Long) As Long
'--- Serialize / Deserialize (SQLite 3.36+)
Private Declare Function sqlite3_serialize Lib "sqlite3" (ByVal db As Long, ByVal zSchema As Long, piSize As Currency, ByVal mFlags As Long) As Long
Private Declare Function sqlite3_deserialize Lib "sqlite3" (ByVal db As Long, ByVal zSchema As Long, ByVal pData As Long, ByVal szDb As Currency, ByVal szBuf As Currency, ByVal mFlags As Long) As Long
Private Declare Function sqlite3_malloc64 Lib "sqlite3" (ByVal n As Currency) As Long
Private Declare Sub sqlite3_free Lib "sqlite3" (ByVal p As Long)

'==============================================================================
' KERNEL32 HELPERS (UTF-8 marshaling)
'==============================================================================
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub memcpy Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hModule As Long) As Long
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long

'==============================================================================
' BCRYPT API DECLARES (bcrypt.dll - built-in Windows Vista+)
'==============================================================================
Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt" (phAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptGetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, pbOutput As Any, ByVal cbOutput As Long, pcbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptSetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptGenerateSymmetricKey Lib "bcrypt" (ByVal hAlgorithm As Long, phKey As Long, ByVal pbKeyObject As Long, ByVal cbKeyObject As Long, ByVal pbSecret As Long, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroyKey Lib "bcrypt" (ByVal hKey As Long) As Long
Private Declare Function BCryptEncrypt Lib "bcrypt" (ByVal hKey As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal pPaddingInfo As Long, ByVal pbIV As Long, ByVal cbIV As Long, ByVal pbOutput As Long, ByVal cbOutput As Long, pcbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDecrypt Lib "bcrypt" (ByVal hKey As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal pPaddingInfo As Long, ByVal pbIV As Long, ByVal cbIV As Long, ByVal pbOutput As Long, ByVal cbOutput As Long, pcbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDeriveKeyPBKDF2 Lib "bcrypt" (ByVal hPrf As Long, ByVal pbPassword As Long, ByVal cbPassword As Long, ByVal pbSalt As Long, ByVal cbSalt As Long, ByVal cIterations As Currency, ByVal pbDerivedKey As Long, ByVal cbDerivedKey As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptGenRandom Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal pbBuffer As Long, ByVal cbBuffer As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCreateHash Lib "bcrypt" (ByVal hAlgorithm As Long, phHash As Long, ByVal pbHashObject As Long, ByVal cbHashObject As Long, ByVal pbSecret As Long, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptHashData Lib "bcrypt" (ByVal hHash As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptFinishHash Lib "bcrypt" (ByVal hHash As Long, ByVal pbOutput As Long, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroyHash Lib "bcrypt" (ByVal hHash As Long) As Long

'==============================================================================
' CONSTANTS
'==============================================================================
'--- Result codes
Private Const SQLITE_OK             As Long = 0
Private Const SQLITE_ERROR          As Long = 1
Private Const SQLITE_BUSY           As Long = 5
Private Const SQLITE_ROW            As Long = 100
Private Const SQLITE_DONE           As Long = 101
'--- Column types
Private Const SQLITE_TYPE_INTEGER   As Long = 1
Private Const SQLITE_TYPE_FLOAT     As Long = 2
Private Const SQLITE_TYPE_TEXT      As Long = 3
Private Const SQLITE_TYPE_BLOB      As Long = 4
Private Const SQLITE_TYPE_NULL      As Long = 5
'--- Open flags
Private Const SQLITE_OPEN_READWRITE As Long = &H2
Private Const SQLITE_OPEN_CREATE    As Long = &H4
Private Const SQLITE_OPEN_READONLY  As Long = &H1
'--- Destructor sentinel
Private Const SQLITE_TRANSIENT      As Long = -1
'--- Deserialize flags
Private Const SQLITE_DESERIALIZE_FREEONCLOSE As Long = 1
Private Const SQLITE_DESERIALIZE_RESIZEABLE  As Long = 2
'--- UTF-8 codepage
Private Const CP_UTF8               As Long = 65001
'--- Encryption constants
Private Const AES_KEYLEN            As Long = 32    ' AES-256
Private Const AES_IVLEN             As Long = 16    ' CBC block size
Private Const AES_SALT_LEN          As Long = 16
Private Const AES_HMAC_LEN          As Long = 32    ' SHA-256
Private Const AES_HMAC_KEYLEN       As Long = 32
Private Const AES_DERIVED_KEY_LEN   As Long = 80    ' 32+16+32
Private Const AES_KDF_ITER          As Long = 100000
Private Const BCRYPT_USE_SYSTEM_PREFERRED_RNG As Long = 2
Private Const BCRYPT_BLOCK_PADDING  As Long = 1
Private Const BCRYPT_ALG_HANDLE_HMAC_FLAG As Long = 8
Private Const CZQ_MAGIC_0           As Byte = &H43  ' C
Private Const CZQ_MAGIC_1           As Byte = &H5A  ' Z
Private Const CZQ_MAGIC_2           As Byte = &H51  ' Q
Private Const CZQ_MAGIC_3           As Byte = &H1   ' version 1
Private Const CZQ_HEADER_LEN        As Long = 52    ' 4 + 16 + 32

'--- JSON Engine Types ---
Private Type SAFEARRAY1D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As Long
    cElements           As Long
    lLbound             As Long
End Type

Private Type JsonContext
    StrictMode          As Boolean
    Text()              As Integer
    Pos                 As Long
    Error               As String
    LastChar            As Integer
    TextArray           As SAFEARRAY1D
End Type

Private Enum VbCollectionOffsets
    o_pFirstIndexedItem = &H18
    o_pRootTreeItem = &H24
    o_pEndTreePtr = &H28
    o_pvUnk5 = &H2C
    o_KeyPtr = &H10
    o_pNextIndexedItem = o_pFirstIndexedItem
    o_pRightBranch = &H24
    o_pLeftBranch = &H28
End Enum

Private Const JSON_IDX_OFFSET       As Long = 1
Private Const JSON_VARIANT_ALPHABOOL As Long = 2
Private Const JSON_DEF_IGNORE_CASE  As Boolean = True

'==============================================================================
' EVENTS
'==============================================================================
Public Event BeforeOpen(Cancel As Boolean)
Public Event AfterOpen()
Public Event BeforeClose()
Public Event AfterClose()
Public Event OnError(ByVal ErrCode As Long, ByVal ErrMsg As String)

'==============================================================================
' MEMBER VARIABLES
'==============================================================================
'--- SQLite state
Private m_hDb                   As Long
Private m_bIsOpen               As Boolean
Private m_sLastError            As String
Private m_lAffectedRows         As Long
Private m_lLastInsertID         As Long
'--- Design-time properties
Private m_sDatabasePath         As String
Private m_sPassword             As String
Private m_lBusyTimeout          As Long
Private m_bAutoOpen             As Boolean
Private m_bMapNullEmpty         As Boolean
'--- Internal recordset state
Private m_vData()               As Variant
Private m_aFieldNames()         As String
Private m_aFieldTypes()         As Long
Private m_lRowCount             As Long
Private m_lColCount             As Long
Private m_lPosition             As Long
Private m_bHasData              As Boolean
'--- JSON Engine Cache
Private m_oJsonRoot             As Object
'--- Encryption state
Private m_bEncrypted            As Boolean

'==============================================================================
' DESIGN-TIME PROPERTIES
'==============================================================================
Public Property Get DatabasePath() As String
    DatabasePath = m_sDatabasePath
End Property

Public Property Let DatabasePath(ByVal sPath As String)
    m_sDatabasePath = sPath
    PropertyChanged "DatabasePath"
End Property

Public Property Get Password() As String
    Password = m_sPassword
End Property

Public Property Let Password(ByVal sPass As String)
    m_sPassword = sPass
    PropertyChanged "Password"
End Property

Public Property Get BusyTimeout() As Long
    BusyTimeout = m_lBusyTimeout
End Property

Public Property Let BusyTimeout(ByVal lMs As Long)
    m_lBusyTimeout = lMs
    PropertyChanged "BusyTimeout"
    If m_bIsOpen Then
        sqlite3_busy_timeout m_hDb, m_lBusyTimeout
    End If
End Property

Public Property Get AutoOpen() As Boolean
    AutoOpen = m_bAutoOpen
End Property

Public Property Let AutoOpen(ByVal bVal As Boolean)
    m_bAutoOpen = bVal
    PropertyChanged "AutoOpen"
End Property

Public Property Get MapNullToEmpty() As Boolean
    MapNullToEmpty = m_bMapNullEmpty
End Property

Public Property Let MapNullToEmpty(ByVal bVal As Boolean)
    m_bMapNullEmpty = bVal
    PropertyChanged "MapNullToEmpty"
End Property

'==============================================================================
' READ-ONLY PROPERTIES
'==============================================================================
Public Property Get IsOpen() As Boolean
    IsOpen = m_bIsOpen
End Property

Public Property Get AffectedRows() As Long
    AffectedRows = m_lAffectedRows
End Property

Public Property Get LastInsertID() As Long
    LastInsertID = m_lLastInsertID
End Property

Public Property Get SQLiteVersion() As String
    SQLiteVersion = pvFromUtf8Ptr(sqlite3_libversion())
End Property

Public Property Get LastError() As String
    LastError = m_sLastError
End Property

Public Property Get IsEncrypted() As Boolean
    IsEncrypted = m_bEncrypted
End Property

Public Function SaveEncrypted() As Boolean
    '--- Manually save encrypted :memory: DB to file without closing
    If Not m_bEncrypted Or Not m_bIsOpen Then Exit Function
    If Len(m_sPassword) = 0 Or m_hDb = 0 Then Exit Function
    Dim cSize       As Currency
    Dim lSerSize    As Long
    Dim pSerialized As Long
    Dim baDbBytes() As Byte
    Dim baEncOut()  As Byte
    Dim baSchemaS() As Byte
    baSchemaS = pvToUtf8("main" & vbNullChar)
    pSerialized = sqlite3_serialize(m_hDb, VarPtr(baSchemaS(0)), cSize, 0)
    lSerSize = CLng(CDec(cSize) * CDec(10000))
    If pSerialized <> 0 And lSerSize > 0 Then
        ReDim baDbBytes(0 To lSerSize - 1) As Byte
        CopyMemory baDbBytes(0), ByVal pSerialized, lSerSize
        sqlite3_free pSerialized
        pvEncryptBytes baDbBytes, m_sPassword, baEncOut
        pvWriteFile m_sDatabasePath, baEncOut
        Erase baDbBytes
        Erase baEncOut
        SaveEncrypted = True
    End If
End Function

'==============================================================================
' CORE SQLITE METHODS
'==============================================================================
Public Function OpenDB(Optional ByVal sPath As String = "") As Boolean
    Dim bCancel         As Boolean
    Dim baFile()        As Byte
    Dim lRc             As Long

    If m_bIsOpen Then
        CloseDB
    End If
    RaiseEvent BeforeOpen(bCancel)
    If bCancel Then
        Exit Function
    End If
    '--- use explicit path or property
    If Len(sPath) > 0 Then
        m_sDatabasePath = sPath
    End If
    If Len(m_sDatabasePath) = 0 Then
        m_sLastError = "DatabasePath is empty"
        RaiseEvent OnError(0, m_sLastError)
        Exit Function
    End If
    '--- encrypted database flow: decrypt file → open :memory: → deserialize
    If Len(m_sPassword) > 0 Then
        m_bEncrypted = True
        If Dir$(m_sDatabasePath) <> "" Then
            '--- file exists: read → decrypt → deserialize into :memory:
            Dim baEncFile()  As Byte
            Dim baPlain()    As Byte
            baEncFile = pvReadFile(m_sDatabasePath)
            If Not pvDecryptBytes(baEncFile, m_sPassword, baPlain) Then
                m_sLastError = "Decryption failed: " & m_sLastError
                RaiseEvent OnError(0, m_sLastError)
                Exit Function
            End If
            '--- open :memory: database
            Dim baMemTag()   As Byte
            baMemTag = pvToUtf8(":memory:" & vbNullChar)
            lRc = sqlite3_open_v2(VarPtr(baMemTag(0)), m_hDb, SQLITE_OPEN_READWRITE Or SQLITE_OPEN_CREATE, 0)
            If lRc <> SQLITE_OK Then
                m_sLastError = "Cannot open :memory: database"
                If m_hDb <> 0 Then sqlite3_close_v2 m_hDb: m_hDb = 0
                RaiseEvent OnError(lRc, m_sLastError)
                Exit Function
            End If
            '--- deserialize decrypted bytes into :memory:
            Dim lPlainLen    As Long
            Dim pBuf         As Long
            Dim baMain()     As Byte
            lPlainLen = UBound(baPlain) + 1
            pBuf = sqlite3_malloc64(lPlainLen / 10000@)
            If pBuf = 0 Then
                m_sLastError = "Cannot allocate memory for deserialization"
                sqlite3_close_v2 m_hDb: m_hDb = 0
                RaiseEvent OnError(0, m_sLastError)
                Exit Function
            End If
            CopyMemory ByVal pBuf, baPlain(0), lPlainLen
            baMain = pvToUtf8("main" & vbNullChar)
            lRc = sqlite3_deserialize(m_hDb, VarPtr(baMain(0)), pBuf, _
                    lPlainLen / 10000@, lPlainLen / 10000@, _
                    SQLITE_DESERIALIZE_FREEONCLOSE Or SQLITE_DESERIALIZE_RESIZEABLE)
            If lRc <> SQLITE_OK Then
                m_sLastError = "Deserialization failed: " & pvGetDbError()
                sqlite3_close_v2 m_hDb: m_hDb = 0
                RaiseEvent OnError(lRc, m_sLastError)
                Exit Function
            End If
            Erase baPlain
            Erase baEncFile
        Else
            '--- new encrypted DB: just open :memory:
            Dim baMemNew()   As Byte
            baMemNew = pvToUtf8(":memory:" & vbNullChar)
            lRc = sqlite3_open_v2(VarPtr(baMemNew(0)), m_hDb, SQLITE_OPEN_READWRITE Or SQLITE_OPEN_CREATE, 0)
            If lRc <> SQLITE_OK Then
                m_sLastError = "Cannot open :memory: database"
                If m_hDb <> 0 Then sqlite3_close_v2 m_hDb: m_hDb = 0
                RaiseEvent OnError(lRc, m_sLastError)
                Exit Function
            End If
        End If
    Else
        m_bEncrypted = False
        '--- normal (unencrypted) database open
        baFile = pvToUtf8(m_sDatabasePath & vbNullChar)
        lRc = sqlite3_open_v2(VarPtr(baFile(0)), m_hDb, SQLITE_OPEN_READWRITE Or SQLITE_OPEN_CREATE, 0)
        If lRc <> SQLITE_OK Then
            m_sLastError = "Cannot open database: " & pvGetDbError()
            If m_hDb <> 0 Then sqlite3_close_v2 m_hDb: m_hDb = 0
            RaiseEvent OnError(lRc, m_sLastError)
            Exit Function
        End If
    End If
    '--- set busy timeout
    If m_lBusyTimeout > 0 Then
        sqlite3_busy_timeout m_hDb, m_lBusyTimeout
    End If
    m_bIsOpen = True
    m_sLastError = ""
    RaiseEvent AfterOpen
    OpenDB = True
End Function

Public Sub CloseDB()
    If Not m_bIsOpen Then
        Exit Sub
    End If
    RaiseEvent BeforeClose
    '--- encrypted: serialize → encrypt → write to file
    If m_bEncrypted And Len(m_sPassword) > 0 And m_hDb <> 0 Then
        Dim cSize       As Currency
        Dim lSerSize    As Long
        Dim pSerialized As Long
        Dim baDbBytes() As Byte
        Dim baEncOut()  As Byte
        Dim baSchemaS() As Byte
        baSchemaS = pvToUtf8("main" & vbNullChar)
        pSerialized = sqlite3_serialize(m_hDb, VarPtr(baSchemaS(0)), cSize, 0)
        lSerSize = CLng(CDec(cSize) * CDec(10000))
        If pSerialized <> 0 And lSerSize > 0 Then
            ReDim baDbBytes(0 To lSerSize - 1) As Byte
            CopyMemory baDbBytes(0), ByVal pSerialized, lSerSize
            sqlite3_free pSerialized
            pvEncryptBytes baDbBytes, m_sPassword, baEncOut
            pvWriteFile m_sDatabasePath, baEncOut
            Erase baDbBytes
            Erase baEncOut
        End If
    End If
    '--- clear recordset state
    pvClearRecordset
    '--- close database
    If m_hDb <> 0 Then
        sqlite3_close_v2 m_hDb
        m_hDb = 0
    End If
    m_bIsOpen = False
    m_bEncrypted = False
    RaiseEvent AfterClose
End Sub

Public Function OpenFromBytes(baData() As Byte, Optional ByVal sPassword As String = "") As Boolean
    '--- Open a database from a byte array (raw SQLite or encrypted czQlite format)
    '    Ideal for czStorage integration: read bytes from EXE → open in :memory:
    Dim lRc         As Long
    Dim lDataLen    As Long
    Dim baPlain()   As Byte
    Dim baMemTag()  As Byte
    Dim baMain()    As Byte
    Dim pBuf        As Long
    Dim lPlainLen   As Long

    '--- close any open database first
    If m_bIsOpen Then CloseDB
    On Error GoTo EH
    lDataLen = UBound(baData) - LBound(baData) + 1
    If lDataLen = 0 Then
        m_sLastError = "Empty byte array"
        RaiseEvent OnError(0, m_sLastError)
        Exit Function
    End If
    '--- determine if encrypted or raw
    If Len(sPassword) > 0 Then
        m_bEncrypted = True
        m_sPassword = sPassword
        '--- check if data has czQlite header
        If lDataLen > CZQ_HEADER_LEN And baData(0) = CZQ_MAGIC_0 And _
           baData(1) = CZQ_MAGIC_1 And baData(2) = CZQ_MAGIC_2 And _
           baData(3) = CZQ_MAGIC_3 Then
            '--- encrypted format: decrypt first
            If Not pvDecryptBytes(baData, sPassword, baPlain) Then
                m_sLastError = "Decryption failed: " & m_sLastError
                RaiseEvent OnError(0, m_sLastError)
                Exit Function
            End If
        Else
            '--- raw bytes + password: treat as plain, will encrypt on serialize
            baPlain = baData
        End If
    Else
        m_bEncrypted = False
        m_sPassword = ""
        baPlain = baData
    End If
    '--- open :memory: database
    baMemTag = pvToUtf8(":memory:" & vbNullChar)
    lRc = sqlite3_open_v2(VarPtr(baMemTag(0)), m_hDb, SQLITE_OPEN_READWRITE Or SQLITE_OPEN_CREATE, 0)
    If lRc <> SQLITE_OK Then
        m_sLastError = "Cannot open :memory: database"
        If m_hDb <> 0 Then sqlite3_close_v2 m_hDb: m_hDb = 0
        RaiseEvent OnError(lRc, m_sLastError)
        Exit Function
    End If
    '--- deserialize into :memory:
    lPlainLen = UBound(baPlain) - LBound(baPlain) + 1
    pBuf = sqlite3_malloc64(lPlainLen / 10000@)
    If pBuf = 0 Then
        m_sLastError = "Cannot allocate memory for deserialization"
        sqlite3_close_v2 m_hDb: m_hDb = 0
        RaiseEvent OnError(0, m_sLastError)
        Exit Function
    End If
    CopyMemory ByVal pBuf, baPlain(LBound(baPlain)), lPlainLen
    baMain = pvToUtf8("main" & vbNullChar)
    lRc = sqlite3_deserialize(m_hDb, VarPtr(baMain(0)), pBuf, _
            lPlainLen / 10000@, lPlainLen / 10000@, _
            SQLITE_DESERIALIZE_FREEONCLOSE Or SQLITE_DESERIALIZE_RESIZEABLE)
    If lRc <> SQLITE_OK Then
        m_sLastError = "Deserialization failed: " & pvGetDbError()
        sqlite3_close_v2 m_hDb: m_hDb = 0
        RaiseEvent OnError(lRc, m_sLastError)
        Exit Function
    End If
    Erase baPlain
    '--- set busy timeout
    If m_lBusyTimeout > 0 Then
        sqlite3_busy_timeout m_hDb, m_lBusyTimeout
    End If
    m_bIsOpen = True
    m_sLastError = ""
    RaiseEvent AfterOpen
    OpenFromBytes = True
    Exit Function
EH:
    m_sLastError = "OpenFromBytes error: " & Err.Description
    RaiseEvent OnError(Err.Number, m_sLastError)
End Function

Public Function SerializeToBytes(Optional ByVal sPassword As String = "") As Byte()
    '--- Serialize the current :memory: database to a byte array
    '    If sPassword is provided, the output is encrypted (czQlite format)
    '    If empty, returns raw SQLite bytes
    '    Ideal for czStorage integration: serialize → write bytes to EXE
    Dim cSize       As Currency
    Dim lSerSize    As Long
    Dim pSerialized As Long
    Dim baDbBytes() As Byte
    Dim baSchemaS() As Byte
    Dim baEncOut()  As Byte

    If Not m_bIsOpen Or m_hDb = 0 Then
        m_sLastError = "Database is not open"
        RaiseEvent OnError(0, m_sLastError)
        Exit Function
    End If
    baSchemaS = pvToUtf8("main" & vbNullChar)
    pSerialized = sqlite3_serialize(m_hDb, VarPtr(baSchemaS(0)), cSize, 0)
    lSerSize = CLng(CDec(cSize) * CDec(10000))
    If pSerialized = 0 Or lSerSize = 0 Then
        m_sLastError = "Serialization failed (empty database)"
        RaiseEvent OnError(0, m_sLastError)
        Exit Function
    End If
    ReDim baDbBytes(0 To lSerSize - 1) As Byte
    CopyMemory baDbBytes(0), ByVal pSerialized, lSerSize
    sqlite3_free pSerialized
    '--- encrypt if password provided
    If Len(sPassword) > 0 Then
        pvEncryptBytes baDbBytes, sPassword, baEncOut
        SerializeToBytes = baEncOut
        Erase baDbBytes
        Erase baEncOut
    Else
        '--- use password from property if set
        If Len(m_sPassword) > 0 And m_bEncrypted Then
            pvEncryptBytes baDbBytes, m_sPassword, baEncOut
            SerializeToBytes = baEncOut
            Erase baDbBytes
            Erase baEncOut
        Else
            SerializeToBytes = baDbBytes
            Erase baDbBytes
        End If
    End If
End Function

Public Sub Query(ByVal sSQL As String, ParamArray Params() As Variant)
    Dim aParams()       As Variant

    aParams = Params
    pvDoQuery sSQL, aParams
End Sub

Public Function Execute(ByVal sSQL As String, ParamArray Params() As Variant) As Boolean
    Dim aParams()       As Variant

    aParams = Params
    Execute = pvDoExecute(sSQL, aParams)
End Function

Public Function GetVal(ByVal sSQL As String, ParamArray Params() As Variant) As Variant
    Dim aParams()       As Variant

    aParams = Params
    GetVal = pvDoGetVal(sSQL, aParams)
End Function

Public Function GetRows(ByVal sSQL As String, ParamArray Params() As Variant) As Variant
    Dim aParams()       As Variant

    aParams = Params
    pvDoQuery sSQL, aParams
    If m_bHasData And m_lRowCount > 0 Then
        GetRows = m_vData
    Else
        GetRows = Empty
    End If
End Function

Public Sub BeginTrans()
    pvDoExecSimple "BEGIN"
End Sub

Public Sub CommitTrans()
    pvDoExecSimple "COMMIT"
End Sub

Public Sub RollbackTrans()
    pvDoExecSimple "ROLLBACK"
End Sub

Public Function TableExists(ByVal sTable As String) As Boolean
    Dim aP(0 To 0)      As Variant

    aP(0) = sTable
    TableExists = (pvDoGetVal("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=?", aP) > 0)
End Function

Public Sub Vacuum()
    pvDoExecSimple "VACUUM"
End Sub

Public Function BackupTo(ByVal sDstPath As String) As Boolean
    Dim oDst            As Long
    Dim baFile()        As Byte
    Dim baMain()        As Byte
    Dim hBackup         As Long
    Dim lRc             As Long

    If Not m_bIsOpen Then
        m_sLastError = "Database is not open"
        Exit Function
    End If
    '--- open destination
    baFile = pvToUtf8(sDstPath & vbNullChar)
    lRc = sqlite3_open_v2(VarPtr(baFile(0)), oDst, SQLITE_OPEN_READWRITE Or SQLITE_OPEN_CREATE, 0)
    If lRc <> SQLITE_OK Then
        m_sLastError = "Cannot open backup destination"
        If oDst <> 0 Then
            sqlite3_close_v2 oDst
        End If
        Exit Function
    End If
    '--- perform backup
    baMain = pvToUtf8("main" & vbNullChar)
    hBackup = sqlite3_backup_init(oDst, VarPtr(baMain(0)), m_hDb, VarPtr(baMain(0)))
    If hBackup = 0 Then
        m_sLastError = "Cannot initialize backup"
        sqlite3_close_v2 oDst
        Exit Function
    End If
    lRc = sqlite3_backup_step(hBackup, -1)
    sqlite3_backup_finish hBackup
    sqlite3_close_v2 oDst
    BackupTo = (lRc = SQLITE_DONE)
    If Not BackupTo Then
        m_sLastError = "Backup failed"
    End If
End Function

'==============================================================================
' RECORDSET NAVIGATION (after Query)
'==============================================================================
Public Property Get Field(ByVal vNameOrIndex As Variant) As Variant
    Dim lCol            As Long

    If Not m_bHasData Then
        Exit Property
    End If
    If m_lPosition < 0 Or m_lPosition >= m_lRowCount Then
        Exit Property
    End If
    lCol = pvResolveFieldIndex(vNameOrIndex)
    If lCol < 0 Then
        Exit Property
    End If
    Field = m_vData(m_lPosition, lCol)
End Property

Public Property Get EOF() As Boolean
    If Not m_bHasData Then
        EOF = True
        Exit Property
    End If
    EOF = (m_lPosition >= m_lRowCount)
End Property

Public Property Get BOF() As Boolean
    If Not m_bHasData Then
        BOF = True
        Exit Property
    End If
    BOF = (m_lPosition < 0)
End Property

Public Sub MoveNext()
    If m_bHasData Then
        m_lPosition = m_lPosition + 1
    End If
End Sub

Public Sub MovePrevious()
    If m_bHasData Then
        m_lPosition = m_lPosition - 1
    End If
End Sub

Public Sub MoveFirst()
    If m_bHasData Then
        m_lPosition = 0
    End If
End Sub

Public Sub MoveLast()
    If m_bHasData And m_lRowCount > 0 Then
        m_lPosition = m_lRowCount - 1
    End If
End Sub

Public Property Get RecordCount() As Long
    RecordCount = m_lRowCount
End Property

Public Property Get FieldCount() As Long
    FieldCount = m_lColCount
End Property

Public Property Get FieldName(ByVal lIndex As Long) As String
    If m_bHasData And lIndex >= 0 And lIndex < m_lColCount Then
        FieldName = m_aFieldNames(lIndex)
    End If
End Property

Public Property Get AbsolutePosition() As Long
    AbsolutePosition = m_lPosition
End Property

'==============================================================================
' JSON PUBLIC API (adapted from mdJson.bas by wqweto)
'==============================================================================
Public Function QueryJSON(ByVal sSQL As String, ParamArray Params() As Variant) As String
    Dim aParams()       As Variant

    aParams = Params
    pvDoQuery sSQL, aParams
    QueryJSON = ToJSON()
End Function

Public Function ToJSON(Optional ByVal Pretty As Boolean = False) As String
    Dim lRow            As Long
    Dim lCol            As Long
    Dim oArr            As VBA.Collection
    Dim oRow            As VBA.Collection

    If Not m_bHasData Or m_lRowCount = 0 Then
        ToJSON = "[]"
        Exit Function
    End If
    '--- build JSON array of objects using pvJsonDump
    Set oArr = pvJsonCreateObject(vbTextCompare)
    For lRow = 0 To m_lRowCount - 1
        Set oRow = pvJsonCreateObject(vbBinaryCompare)
        For lCol = 0 To m_lColCount - 1
            oRow.Add m_vData(lRow, lCol), m_aFieldNames(lCol)
        Next
        oArr.Add oRow
    Next
    If Pretty Then
        ToJSON = pvJsonDump(oArr, 0, False, , 2)
    Else
        ToJSON = pvJsonDump(oArr, 0, True)
    End If
End Function

Public Function JsonParse(ByVal sJsonText As String) As Boolean
    '--- Parse JSON string and cache in m_oJsonRoot
    Set m_oJsonRoot = Nothing
    If Len(sJsonText) > 0 Then
        Set m_oJsonRoot = pvJsonDoParseObject(sJsonText)
    End If
    JsonParse = Not (m_oJsonRoot Is Nothing)
End Function

Public Function JsonParseObject(ByVal sJsonText As String) As Object
    '--- Parse JSON string and return Collection directly (stateless)
    Set JsonParseObject = pvJsonDoParseObject(sJsonText)
End Function

Public Property Get JsonGet(ByVal sPath As String) As Variant
    If m_oJsonRoot Is Nothing Then Exit Property
    pvJsonAssign JsonGet, pvJsonValueGet(m_oJsonRoot, sPath)
End Property

Public Function JsonStr(ByVal sPath As String) As String
    If m_oJsonRoot Is Nothing Then Exit Function
    JsonStr = pvJsonC_Str(pvJsonValueGet(m_oJsonRoot, sPath))
End Function

Public Function JsonLng(ByVal sPath As String) As Long
    If m_oJsonRoot Is Nothing Then Exit Function
    JsonLng = pvJsonC_Lng(pvJsonValueGet(m_oJsonRoot, sPath))
End Function

Public Function JsonDbl(ByVal sPath As String) As Double
    If m_oJsonRoot Is Nothing Then Exit Function
    JsonDbl = pvJsonC_Dbl(pvJsonValueGet(m_oJsonRoot, sPath))
End Function

Public Function JsonBool(ByVal sPath As String) As Boolean
    If m_oJsonRoot Is Nothing Then Exit Function
    JsonBool = pvJsonC_Bool(pvJsonValueGet(m_oJsonRoot, sPath))
End Function

Public Function JsonObj(ByVal sPath As String) As Object
    If m_oJsonRoot Is Nothing Then Exit Function
    Dim v As Variant
    pvJsonAssign v, pvJsonValueGet(m_oJsonRoot, sPath)
    If IsObject(v) Then Set JsonObj = v
End Function

Public Property Get JsonHas(ByVal sPath As String) As Boolean
    If m_oJsonRoot Is Nothing Then Exit Property
    Dim v As Variant
    pvJsonAssign v, pvJsonValueGet(m_oJsonRoot, sPath)
    JsonHas = Not IsEmpty(v)
End Property

Public Function JsonGetKeys(Optional ByVal sPath As String = "") As Variant
    If m_oJsonRoot Is Nothing Then
        JsonGetKeys = Array()
        Exit Function
    End If
    JsonGetKeys = pvJsonKeys(m_oJsonRoot, sPath)
End Function

Public Sub JsonSet(ByVal sPath As String, ByVal vValue As Variant)
    If m_oJsonRoot Is Nothing Then Set m_oJsonRoot = pvJsonCreateObject(vbBinaryCompare)
    pvJsonValueLet m_oJsonRoot, sPath, vValue
End Sub

Public Sub JsonRemove(ByVal sPath As String)
    If m_oJsonRoot Is Nothing Then Exit Sub
    pvJsonValueLet m_oJsonRoot, sPath, Empty
End Sub

Public Function JsonDump(Optional ByVal Minimize As Boolean = False) As String
    If m_oJsonRoot Is Nothing Then JsonDump = "{}": Exit Function
    JsonDump = pvJsonDump(m_oJsonRoot, 0, Minimize)
End Function

Public Function JsonDumpPretty(Optional ByVal IndentSize As Long = 4) As String
    If m_oJsonRoot Is Nothing Then JsonDumpPretty = "{}": Exit Function
    JsonDumpPretty = pvJsonDump(m_oJsonRoot, 0, False, , IndentSize)
End Function

Public Sub JsonClose()
    Set m_oJsonRoot = Nothing
End Sub

'==============================================================================
' PRIVATE HELPERS — SQLite Core
'==============================================================================
Private Sub pvDoQuery(sSQL As String, aParams() As Variant)
    Dim hStmt           As Long
    Dim lRc             As Long
    Dim lCol            As Long
    Dim cRows           As Collection
    Dim aRow()          As Variant
    Dim lRow            As Long

    pvClearRecordset
    If Not m_bIsOpen Then
        m_sLastError = "Database is not open"
        RaiseEvent OnError(0, m_sLastError)
        Exit Sub
    End If
    '--- prepare
    hStmt = pvPrepare(sSQL)
    If hStmt = 0 Then
        Exit Sub
    End If
    '--- bind parameters
    On Error GoTo EH
    pvBindParams hStmt, aParams
    '--- read column info
    m_lColCount = sqlite3_column_count(hStmt)
    If m_lColCount > 0 Then
        ReDim m_aFieldNames(0 To m_lColCount - 1) As String
        ReDim m_aFieldTypes(0 To m_lColCount - 1) As Long
        For lCol = 0 To m_lColCount - 1
            m_aFieldNames(lCol) = pvFromUtf8Ptr(sqlite3_column_name(hStmt, lCol))
        Next
    End If
    '--- step through all rows
    Set cRows = New Collection
    Do
        lRc = sqlite3_step(hStmt)
        If lRc = SQLITE_ROW Then
            If m_lColCount > 0 Then
                ReDim aRow(0 To m_lColCount - 1) As Variant
                For lCol = 0 To m_lColCount - 1
                    aRow(lCol) = pvReadColumnValue(hStmt, lCol)
                Next
                cRows.Add aRow
            End If
        ElseIf lRc = SQLITE_DONE Then
            Exit Do
        Else
            m_sLastError = "Query error: " & pvGetDbError()
            RaiseEvent OnError(lRc, m_sLastError)
            Exit Do
        End If
    Loop
    sqlite3_finalize hStmt
    hStmt = 0
    '--- build 2D result array
    m_lRowCount = cRows.Count
    If m_lRowCount > 0 And m_lColCount > 0 Then
        ReDim m_vData(0 To m_lRowCount - 1, 0 To m_lColCount - 1) As Variant
        For lRow = 0 To m_lRowCount - 1
            aRow = cRows(lRow + 1)
            For lCol = 0 To m_lColCount - 1
                If IsNull(aRow(lCol)) And m_bMapNullEmpty Then
                    m_vData(lRow, lCol) = ""
                Else
                    m_vData(lRow, lCol) = aRow(lCol)
                End If
            Next
        Next
    End If
    m_lPosition = 0
    m_bHasData = True
    m_sLastError = ""
    Exit Sub
EH:
    If hStmt <> 0 Then
        sqlite3_finalize hStmt
    End If
    m_sLastError = "Query error: " & Err.Description
    RaiseEvent OnError(Err.Number, m_sLastError)
End Sub

Private Function pvDoExecute(sSQL As String, aParams() As Variant) As Boolean
    Dim hStmt           As Long
    Dim lRc             As Long

    If Not m_bIsOpen Then
        m_sLastError = "Database is not open"
        RaiseEvent OnError(0, m_sLastError)
        Exit Function
    End If
    '--- prepare
    hStmt = pvPrepare(sSQL)
    If hStmt = 0 Then
        Exit Function
    End If
    '--- bind parameters
    On Error GoTo EH
    pvBindParams hStmt, aParams
    '--- execute (step until done, tolerate rows)
    Do
        lRc = sqlite3_step(hStmt)
    Loop While lRc = SQLITE_ROW
    If lRc <> SQLITE_DONE Then
        m_sLastError = "Execute error: " & pvGetDbError()
        sqlite3_finalize hStmt
        RaiseEvent OnError(lRc, m_sLastError)
        Exit Function
    End If
    sqlite3_finalize hStmt
    '--- capture affected rows and last insert id
    m_lAffectedRows = sqlite3_changes(m_hDb)
    m_lLastInsertID = CLng(CDec(sqlite3_last_insert_rowid(m_hDb)) * CDec(10000))
    m_sLastError = ""
    pvDoExecute = True
    Exit Function
EH:
    If hStmt <> 0 Then
        sqlite3_finalize hStmt
    End If
    m_sLastError = "Execute error: " & Err.Description
    RaiseEvent OnError(Err.Number, m_sLastError)
End Function

Private Function pvDoGetVal(sSQL As String, aParams() As Variant) As Variant
    Dim hStmt           As Long
    Dim lRc             As Long

    If Not m_bIsOpen Then
        m_sLastError = "Database is not open"
        pvDoGetVal = Null
        Exit Function
    End If
    '--- prepare
    hStmt = pvPrepare(sSQL)
    If hStmt = 0 Then
        pvDoGetVal = Null
        Exit Function
    End If
    '--- bind parameters
    On Error GoTo EH
    pvBindParams hStmt, aParams
    '--- step one row
    lRc = sqlite3_step(hStmt)
    If lRc = SQLITE_ROW Then
        pvDoGetVal = pvReadColumnValue(hStmt, 0)
        If IsNull(pvDoGetVal) And m_bMapNullEmpty Then
            pvDoGetVal = ""
        End If
    Else
        pvDoGetVal = Null
    End If
    sqlite3_finalize hStmt
    m_sLastError = ""
    Exit Function
EH:
    If hStmt <> 0 Then
        sqlite3_finalize hStmt
    End If
    m_sLastError = "GetVal error: " & Err.Description
    pvDoGetVal = Null
End Function

Private Sub pvDoExecSimple(sSQL As String)
    Dim aEmpty()        As Variant

    ReDim aEmpty(0 To 0) As Variant
    aEmpty(0) = Empty
    '--- use pvDoExecute with no real params
    Dim baSql()         As Byte
    Dim hStmt           As Long
    Dim lRc             As Long

    If Not m_bIsOpen Then
        Exit Sub
    End If
    baSql = pvToUtf8(sSQL & vbNullChar)
    If sqlite3_prepare_v2(m_hDb, VarPtr(baSql(0)), -1, hStmt, 0) <> SQLITE_OK Then
        m_sLastError = pvGetDbError()
        Exit Sub
    End If
    Do
        lRc = sqlite3_step(hStmt)
    Loop While lRc = SQLITE_ROW
    sqlite3_finalize hStmt
    If lRc <> SQLITE_DONE Then
        m_sLastError = pvGetDbError()
    End If
End Sub

'==============================================================================
' PRIVATE HELPERS — Prepare / Bind / Read
'==============================================================================
Private Function pvPrepare(sSQL As String) As Long
    Dim baSql()         As Byte
    Dim hStmt           As Long

    baSql = pvToUtf8(sSQL & vbNullChar)
    If sqlite3_prepare_v2(m_hDb, VarPtr(baSql(0)), -1, hStmt, 0) <> SQLITE_OK Then
        m_sLastError = "Cannot compile SQL: " & pvGetDbError()
        RaiseEvent OnError(sqlite3_errcode(m_hDb), m_sLastError)
        pvPrepare = 0
        Exit Function
    End If
    pvPrepare = hStmt
End Function

Private Sub pvBindParams(ByVal hStmt As Long, aParams() As Variant)
    Dim lIdx            As Long
    Dim lParamIdx       As Long
    Dim lLBound         As Long
    Dim lUBound         As Long
    Dim bNamed          As Boolean
    Dim sName           As String
    Dim baName()        As Byte

    On Error Resume Next
    lLBound = LBound(aParams)
    lUBound = UBound(aParams)
    On Error GoTo 0
    If lLBound > lUBound Then
        Exit Sub
    End If
    '--- skip if only Empty passed (from internal calls)
    If lLBound = lUBound Then
        If IsEmpty(aParams(lLBound)) Then
            Exit Sub
        End If
    End If
    '--- detect named vs positional: if first param is string starting with : @ $
    bNamed = False
    If VarType(aParams(lLBound)) = vbString Then
        sName = CStr(aParams(lLBound))
        If Len(sName) > 0 Then
            Select Case Left$(sName, 1)
            Case ":", "@", "$"
                bNamed = True
            End Select
        End If
    End If
    If bNamed Then
        '--- named parameters: pairs of (name, value)
        lIdx = lLBound
        Do While lIdx + 1 <= lUBound
            sName = CStr(aParams(lIdx))
            baName = pvToUtf8(sName & vbNullChar)
            lParamIdx = sqlite3_bind_parameter_index(hStmt, VarPtr(baName(0)))
            If lParamIdx > 0 Then
                pvBindOne hStmt, lParamIdx, aParams(lIdx + 1)
            End If
            lIdx = lIdx + 2
        Loop
    Else
        '--- positional parameters: bind in order
        lParamIdx = 1
        For lIdx = lLBound To lUBound
            pvBindOne hStmt, lParamIdx, aParams(lIdx)
            lParamIdx = lParamIdx + 1
        Next
    End If
End Sub

Private Sub pvBindOne(ByVal hStmt As Long, ByVal lIdx As Long, vValue As Variant)
    Dim baBuf()         As Byte

    Select Case VarType(vValue)
    Case vbNull, vbEmpty
        sqlite3_bind_null hStmt, lIdx
    Case vbBoolean
        sqlite3_bind_int hStmt, lIdx, IIf(vValue, 1, 0)
    Case vbByte, vbInteger, vbLong
        sqlite3_bind_int hStmt, lIdx, CLng(vValue)
    Case vbSingle, vbDouble
        sqlite3_bind_double hStmt, lIdx, CDbl(vValue)
    Case vbCurrency, vbDecimal
        '--- integral values go through int64; fractional through double
        If vValue = Int(vValue) Then
            sqlite3_bind_int64 hStmt, lIdx, CCur(CDec(vValue) / 10000)
        Else
            sqlite3_bind_double hStmt, lIdx, CDbl(vValue)
        End If
    Case vbDate
        baBuf = pvToUtf8(Format$(vValue, "yyyy-mm-dd hh:nn:ss") & vbNullChar)
        sqlite3_bind_text hStmt, lIdx, VarPtr(baBuf(0)), UBound(baBuf), SQLITE_TRANSIENT
    Case vbByte + vbArray
        baBuf = vValue
        If UBound(baBuf) >= 0 Then
            sqlite3_bind_blob hStmt, lIdx, VarPtr(baBuf(0)), UBound(baBuf) + 1, SQLITE_TRANSIENT
        Else
            sqlite3_bind_null hStmt, lIdx
        End If
    Case Else
        '--- treat everything else as text
        baBuf = pvToUtf8(CStr(vValue))
        sqlite3_bind_text hStmt, lIdx, VarPtr(baBuf(0)), UBound(baBuf) - LBound(baBuf) + 1, SQLITE_TRANSIENT
    End Select
End Sub

Private Function pvReadColumnValue(ByVal hStmt As Long, ByVal lCol As Long) As Variant
    Dim lLen            As Long
    Dim lPtr            As Long
    Dim baBuf()         As Byte

    Select Case sqlite3_column_type(hStmt, lCol)
    Case SQLITE_TYPE_INTEGER
        '--- recover int64: on x86, returns Currency (raw bits = value * 10000)
        pvReadColumnValue = CLng(CDec(sqlite3_column_int64(hStmt, lCol)) * CDec(10000))
    Case SQLITE_TYPE_FLOAT
        pvReadColumnValue = sqlite3_column_double(hStmt, lCol)
    Case SQLITE_TYPE_TEXT
        pvReadColumnValue = pvFromUtf8Ptr(sqlite3_column_text(hStmt, lCol))
    Case SQLITE_TYPE_BLOB
        lLen = sqlite3_column_bytes(hStmt, lCol)
        If lLen > 0 Then
            lPtr = sqlite3_column_blob(hStmt, lCol)
            ReDim baBuf(0 To lLen - 1) As Byte
            CopyMemory baBuf(0), ByVal lPtr, lLen
            pvReadColumnValue = baBuf
        Else
            pvReadColumnValue = vbNullString
        End If
    Case Else
        pvReadColumnValue = Null
    End Select
End Function

'==============================================================================
' PRIVATE HELPERS — UTF-8 Marshaling
'==============================================================================
Private Function pvToUtf8(sText As String) As Byte()
    Dim baRetVal()      As Byte
    Dim lSize           As Long

    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        WideCharToMultiByte CP_UTF8, 0, StrPtr(sText), Len(sText), baRetVal(0), lSize, 0, 0
    Else
        baRetVal = vbNullString
    End If
    pvToUtf8 = baRetVal
End Function

Private Function pvFromUtf8Ptr(ByVal lpUtf8 As Long) As String
    Dim lLen            As Long
    Dim lSize           As Long

    If lpUtf8 = 0 Then
        Exit Function
    End If
    lLen = lstrlenA(lpUtf8)
    If lLen > 0 Then
        pvFromUtf8Ptr = String$(lLen, 0)
        lSize = MultiByteToWideChar(CP_UTF8, 0, ByVal lpUtf8, lLen, StrPtr(pvFromUtf8Ptr), lLen)
        pvFromUtf8Ptr = Left$(pvFromUtf8Ptr, lSize)
    End If
End Function

'==============================================================================
' PRIVATE HELPERS — Recordset / Error / JSON Format
'==============================================================================
Private Sub pvClearRecordset()
    m_lRowCount = 0
    m_lColCount = 0
    m_lPosition = 0
    m_bHasData = False
    Erase m_vData
    Erase m_aFieldNames
    Erase m_aFieldTypes
End Sub

Private Function pvResolveFieldIndex(vNameOrIndex As Variant) As Long
    Dim lIdx            As Long

    pvResolveFieldIndex = -1
    If Not m_bHasData Then
        Exit Function
    End If
    If VarType(vNameOrIndex) = vbString Then
        '--- lookup by name (case-insensitive)
        For lIdx = 0 To m_lColCount - 1
            If StrComp(m_aFieldNames(lIdx), CStr(vNameOrIndex), vbTextCompare) = 0 Then
                pvResolveFieldIndex = lIdx
                Exit Function
            End If
        Next
    Else
        '--- lookup by index
        lIdx = CLng(vNameOrIndex)
        If lIdx >= 0 And lIdx < m_lColCount Then
            pvResolveFieldIndex = lIdx
        End If
    End If
End Function

Private Function pvGetDbError() As String
    If m_hDb <> 0 Then
        pvGetDbError = pvFromUtf8Ptr(sqlite3_errmsg(m_hDb))
    Else
        pvGetDbError = "Database handle is null"
    End If
End Function

'==============================================================================
' JSON PRIVATE ENGINE (adapted from mdJson.bas by wqweto)
'==============================================================================
Private Function pvJsonDoParseObject(sText As String) As Object
    Dim vJson As Variant
    Dim sError As String
    If pvJsonDoParse(sText, vJson, sError) Then
        If IsObject(vJson) Then
            Set pvJsonDoParseObject = vJson
        End If
    End If
    If Not sError = "" Then Debug.Print "JsonParse Error: " & sError
End Function

Private Function pvJsonDoParse(sText As String, Optional RetVal As Variant, Optional Error As String, _
            Optional ByVal StrictMode As Boolean) As Boolean
    Const FADF_AUTO     As Long = 1
    Dim uCtx            As JsonContext

    On Error GoTo EH
    With uCtx
        .StrictMode = StrictMode
        With .TextArray
            .cDims = 1
            .fFeatures = FADF_AUTO
            .cbElements = 2
            .cLocks = 1
            .pvData = StrPtr(sText)
            If .pvData = 0 Then .pvData = StrPtr("")
            .cElements = Len(sText) + 1
        End With
        memcpy ByVal ArrPtr(.Text), VarPtr(.TextArray), 4
        pvJsonAssign RetVal, pvJsonParse(uCtx)
        Error = .Error
        If LenB(Error) <> 0 Then GoTo QH
        If pvJsonGetChar(uCtx) <> 0 Then
            Error = "Extra '" & ChrW$(.LastChar) & "' at position " & .Pos
            GoTo QH
        End If
        pvJsonDoParse = True
QH:
        .TextArray.pvData = 0
        .TextArray.cElements = 0
        memcpy ByVal ArrPtr(.Text), 0&, 4
    End With
    Exit Function
EH:
    Error = "JsonParse: " & Err.Description
    Debug.Print Error
    Resume QH
End Function

Private Function pvJsonParse(uCtx As JsonContext) As Variant
    Dim lIdx            As Long
    Dim sKey            As String
    Dim sText           As String
    Dim vValue          As Variant
    Dim oRetVal         As VBA.Collection

    On Error GoTo EH
    With uCtx
        Select Case pvJsonGetChar(uCtx)
        Case 34 '--- "
            pvJsonParse = pvJsonGetString(uCtx)
            If .LastChar = 0 Then GoTo QH
        Case 91 '--- [
            Set oRetVal = pvJsonCreateObject(vbTextCompare)
            Do
                pvJsonAssign vValue, pvJsonParse(uCtx)
                If LenB(.Error) <> 0 Then
                    If .LastChar = 93 Then
                        If Not .StrictMode Then Exit Do
                        lIdx = oRetVal.Count
                        If lIdx = 0 Then Exit Do
                    End If
                    GoTo QH
                End If
                oRetVal.Add vValue
                Select Case pvJsonGetChar(uCtx)
                Case 44: lIdx = lIdx + 1
                Case 93: Exit Do
                Case Else
                    .Error = "Expected ',' or ']' at position " & .Pos
                    Exit Function
                End Select
            Loop
            .Error = vbNullString
            Set pvJsonParse = oRetVal
        Case 123 '--- {
            Set oRetVal = pvJsonCreateObject(vbBinaryCompare)
            Do
                If pvJsonGetChar(uCtx) <> 34 Then
                    If .LastChar = 125 Then
                        If Not .StrictMode Then Exit Do
                        lIdx = oRetVal.Count
                        If lIdx = 0 Then Exit Do
                    End If
                    .Error = "Missing key at position " & .Pos
                    GoTo QH
                End If
                sKey = pvJsonGetString(uCtx)
                If .LastChar = 0 Then GoTo QH
                If pvJsonGetChar(uCtx) <> 58 Then
                    .Error = "Expected ':' at position " & .Pos
                    GoTo QH
                End If
                pvJsonAssign vValue, pvJsonParse(uCtx)
                If LenB(.Error) <> 0 Then GoTo QH
                Select Case pvJsonGetChar(uCtx)
                Case 44, 125
                    If pvJsonCollIndexByKey(oRetVal, sKey, JSON_DEF_IGNORE_CASE) > 0 Then
                        If .StrictMode Then
                            .Error = "Duplicate key '" & sKey & "' at position " & .Pos
                            GoTo QH
                        End If
                        oRetVal.Remove sKey
                    End If
                    oRetVal.Add vValue, sKey
                    If .LastChar = 125 Then Exit Do
                Case Else
                    .Error = "Expected ',' or '}' at position " & .Pos
                    GoTo QH
                End Select
            Loop
            .Error = vbNullString
            Set pvJsonParse = oRetVal
        Case 116, 84: '--- t T (true)
            If (.Text(.Pos) Or &H20) <> 114 Then GoTo UnexpectedSymbol
            If (.Text(.Pos + 1) Or &H20) <> 117 Then GoTo UnexpectedSymbol
            If (.Text(.Pos + 2) Or &H20) <> 101 Then GoTo UnexpectedSymbol
            .Pos = .Pos + 3
            pvJsonParse = True
        Case 102, 70: '--- f F (false)
            If (.Text(.Pos) Or &H20) <> 97 Then GoTo UnexpectedSymbol
            If (.Text(.Pos + 1) Or &H20) <> 108 Then GoTo UnexpectedSymbol
            If (.Text(.Pos + 2) Or &H20) <> 115 Then GoTo UnexpectedSymbol
            If (.Text(.Pos + 3) Or &H20) <> 101 Then GoTo UnexpectedSymbol
            .Pos = .Pos + 4
            pvJsonParse = False
        Case 110, 78: '--- n N (null)
            If (.Text(.Pos) Or &H20) <> 117 Then GoTo UnexpectedSymbol
            If (.Text(.Pos + 1) Or &H20) <> 108 Then GoTo UnexpectedSymbol
            If (.Text(.Pos + 2) Or &H20) <> 108 Then GoTo UnexpectedSymbol
            .Pos = .Pos + 3
            pvJsonParse = Null
        Case 48 To 57, 43, 45, 46: '--- 0-9 + - .
            For lIdx = 0 To 1000
                Select Case .Text(.Pos + lIdx)
                Case 48 To 57, 43, 45, 46, 101, 69, 120, 88, 97 To 102, 65 To 70
                Case Else: Exit For
                End Select
            Next
            sText = Space$(lIdx + 1)
            memcpy ByVal StrPtr(sText), .Text(.Pos - 1), LenB(sText)
            If LCase$(Left$(sText, 2)) = "0x" Then Mid$(sText, 1, 2) = "&H"
            On Error GoTo ErrorConvert
            pvJsonParse = Val(sText)
            On Error GoTo 0
            .Pos = .Pos + lIdx
        Case 0
            If LenB(.Error) <> 0 Then GoTo QH
        Case Else
            GoTo UnexpectedSymbol
        End Select
QH:
        Exit Function
UnexpectedSymbol:
        .Error = "Unexpected '" & ChrW$(.LastChar) & "' at position " & .Pos
        Exit Function
ErrorConvert:
        .Error = Err.Description & " at position " & .Pos
    End With
    Exit Function
EH:
    Debug.Print "pvJsonParse: " & Err.Description
End Function

Private Function pvJsonGetChar(uCtx As JsonContext) As Integer
    With uCtx
        Do While .Pos <= UBound(.Text)
            .LastChar = .Text(.Pos)
            .Pos = .Pos + 1
            Select Case .LastChar
            Case 0: Exit Function
            Case 9, 10, 13, 32 '--- whitespace
            Case 47 '--- / (comments in non-strict)
                If Not .StrictMode Then
                    Select Case .Text(.Pos)
                    Case 47 '--- //
                        .Pos = .Pos + 1
                        Do
                            .LastChar = .Text(.Pos): .Pos = .Pos + 1
                            If .LastChar = 0 Then Exit Function
                        Loop While Not (.LastChar = 10 Or .LastChar = 13)
                    Case 42 '--- /*
                        Dim lIdx As Long
                        lIdx = .Pos + 1
                        Do
                            .LastChar = .Text(lIdx): lIdx = lIdx + 1
                            If .LastChar = 0 Then
                                .Error = "Unterminated comment at position " & .Pos
                                Exit Function
                            End If
                        Loop While Not (.LastChar = 42 And .Text(lIdx) = 47)
                        .LastChar = .Text(lIdx)
                        .Pos = lIdx + 1
                    Case Else
                        pvJsonGetChar = .LastChar: Exit Do
                    End Select
                Else
                    pvJsonGetChar = .LastChar: Exit Do
                End If
            Case Else
                pvJsonGetChar = .LastChar: Exit Do
            End Select
        Loop
    End With
End Function

Private Function pvJsonGetString(uCtx As JsonContext) As String
    Dim lIdx As Long, nChar As Integer, sText As String

    With uCtx
        For lIdx = 0 To &H7FFFFFFF
            nChar = .Text(.Pos + lIdx)
            Select Case nChar
            Case 0, 34, 92
                sText = Space$(lIdx)
                memcpy ByVal StrPtr(sText), .Text(.Pos), LenB(sText)
                pvJsonGetString = pvJsonGetString & sText
                If nChar = 34 Then
                    .Pos = .Pos + lIdx + 1: Exit For
                ElseIf nChar <> 92 Then
                    nChar = 0: .Pos = .Pos + lIdx + 1
                    .Error = "Missing end of string at position " & .Pos: Exit For
                End If
                lIdx = lIdx + 1
                nChar = .Text(.Pos + lIdx)
                Select Case nChar
                Case 98:  pvJsonGetString = pvJsonGetString & ChrW$(8)
                Case 102: pvJsonGetString = pvJsonGetString & ChrW$(12)
                Case 110: pvJsonGetString = pvJsonGetString & vbLf
                Case 114: pvJsonGetString = pvJsonGetString & vbCr
                Case 116: pvJsonGetString = pvJsonGetString & vbTab
                Case 34:  pvJsonGetString = pvJsonGetString & """"
                Case 92:  pvJsonGetString = pvJsonGetString & "\"
                Case 47:  pvJsonGetString = pvJsonGetString & "/"
                Case 117: '--- \uXXXX
                    pvJsonGetString = pvJsonGetString & ChrW$(CLng("&H" & ChrW$(.Text(.Pos + lIdx + 1)) & ChrW$(.Text(.Pos + lIdx + 2)) & ChrW$(.Text(.Pos + lIdx + 3)) & ChrW$(.Text(.Pos + lIdx + 4))))
                    lIdx = lIdx + 4
                Case Else
                    nChar = 0: .Pos = .Pos + lIdx + 1
                    .Error = "Invalid escape at position " & .Pos: Exit For
                End Select
                .Pos = .Pos + lIdx + 1: lIdx = -1
            End Select
        Next
        .LastChar = nChar
    End With
End Function

Private Function pvJsonValueGet(oJson As Object, ByVal sKey As String) As Variant
    Dim vSplit As Variant, lIdx As Long, vKey As Variant, vItem As Variant
    Dim oParam As VBA.Collection

    On Error GoTo EH
    If oJson Is Nothing Then Exit Function
    If LenB(sKey) = 0 Then
        Set pvJsonValueGet = oJson: Exit Function
    End If
    vSplit = Split(sKey, "/")
    Set oParam = oJson
    For lIdx = 0 To UBound(vSplit)
        vKey = vSplit(lIdx)
        If pvJsonC_Str(vKey) = "-1" Then
            pvJsonValueGet = oParam.Count: Exit Function
        ElseIf pvJsonIsOnlyDigits(vKey) Then
            If pvJsonCompareMode(oParam) <> vbBinaryCompare Then vKey = pvJsonC_Lng(vKey)
        End If
        pvJsonAssign vItem, pvJsonItemGet(oParam, vKey)
        If Not IsEmpty(vItem) Then
            If lIdx < UBound(vSplit) Then
                If Not IsObject(vItem) Then Exit Function
                Set oParam = vItem
            Else
                pvJsonAssign pvJsonValueGet, vItem
            End If
        Else
            If LenB(vKey) = 0 Then Set pvJsonValueGet = oParam
            Exit Function
        End If
    Next
    Exit Function
EH:
End Function

Private Sub pvJsonValueLet(oJson As Object, ByVal sKey As String, vValue As Variant)
    Dim vSplit As Variant, lIdx As Long, vKey As Variant, lKey As Long
    Dim oParam As VBA.Collection

    On Error GoTo EH
    If LenB(sKey) = 0 Then Exit Sub
    vSplit = Split(sKey, "/")
    If oJson Is Nothing Then
        Set oJson = pvJsonCreateObject(-(pvJsonIsOnlyDigits(vSplit(0)) Or vSplit(0) = "-1"))
    End If
    Set oParam = oJson
    For lIdx = 0 To UBound(vSplit)
        vKey = vSplit(lIdx)
        If pvJsonC_Str(vKey) = "-1" Then vKey = oParam.Count
        If pvJsonIsOnlyDigits(vKey) Then
            If pvJsonCompareMode(oParam) <> vbBinaryCompare Then vKey = pvJsonC_Lng(vKey)
        End If
        If lIdx < UBound(vSplit) Then
            If Not IsObject(pvJsonItemGet(oParam, vKey)) Then
                pvJsonItemLet oParam, vKey, pvJsonCreateObject(-(pvJsonIsOnlyDigits(vSplit(lIdx + 1)) Or vSplit(lIdx + 1) = "-1"))
            End If
            Set oParam = pvJsonItemGet(oParam, vKey)
        ElseIf IsEmpty(vValue) Then
            If VarType(vKey) = vbLong Then
                lKey = vKey + JSON_IDX_OFFSET
                If lKey > 0 And lKey <= oParam.Count Then oParam.Remove lKey
            Else
                If pvJsonCollIndexByKey(oParam, vKey, JSON_DEF_IGNORE_CASE) > 0 Then oParam.Remove vKey
            End If
        Else
            pvJsonItemLet oParam, vKey, vValue
        End If
    Next
    Exit Sub
EH:
    Debug.Print "pvJsonValueLet: " & Err.Description
End Sub

Private Function pvJsonKeys(oJson As Object, Optional ByVal Key As String) As Variant
    Dim vSplit As Variant, lIdx As Long, vKey As Variant, vItem As Variant, lCount As Long
    Dim oParam As VBA.Collection

    On Error GoTo EH
    If oJson Is Nothing Then pvJsonKeys = Array(): Exit Function
    If LenB(Key) = 0 Then
        Set oParam = oJson
    Else
        vSplit = Split(Key, "/")
        Set oParam = oJson
        For lIdx = 0 To UBound(vSplit)
            vKey = vSplit(lIdx)
            If pvJsonIsOnlyDigits(vKey) Then
                If pvJsonCompareMode(oParam) <> vbBinaryCompare Then vKey = pvJsonC_Lng(vKey)
            End If
            pvJsonAssign vItem, pvJsonItemGet(oParam, vKey)
            If IsObject(vItem) Then
                Set oParam = vItem
            Else
                pvJsonKeys = Array(): Exit Function
            End If
        Next
    End If
    lCount = oParam.Count
    If lCount = 0 Then pvJsonKeys = Array(): Exit Function
    ReDim vItem(0 To lCount - 1) As Variant
    If pvJsonCompareMode(oParam) = vbBinaryCompare Then
        vItem = pvJsonCollAllKeys(oParam)
    Else
        For lIdx = 0 To UBound(vItem): vItem(lIdx) = lIdx: Next
    End If
    pvJsonKeys = vItem
    Exit Function
EH:
    pvJsonKeys = Array()
End Function

Private Function pvJsonDump(vJson As Variant, Optional ByVal Level As Long, Optional ByVal Minimize As Boolean, _
            Optional CompoundChars As String, Optional IndentSize As Long = 4, Optional ByVal MaxWidth As Long = 100) As String
    Static vTranscode As Variant
    Dim vKeys As Variant, vItems As Variant, lIdx As Long, lSize As Long
    Dim sSpace As String, lAsc As Long, lCompareMode As VbCompareMethod, lCount As Long
    Dim oJson As VBA.Collection

    Select Case VarType(vJson)
    Case vbObject
        Set oJson = vJson
        If oJson Is Nothing Then Exit Function
        lCompareMode = pvJsonCompareMode(oJson)
        If LenB(CompoundChars) = 0 Then CompoundChars = IIf(lCompareMode = vbBinaryCompare, "{}", "[]")
        lCount = oJson.Count
        If lCount <= 0 Then
            pvJsonDump = CompoundChars
        Else
            sSpace = IIf(Minimize, vbNullString, " ")
            ReDim vItems(0 To lCount - 1) As String
            If lCompareMode = vbBinaryCompare Then
                vKeys = pvJsonCollAllKeys(oJson)
                If UBound(vKeys) >= 0 Then
                    If LenB(vKeys(0)) = 0 Then lCompareMode = vbTextCompare
                End If
            End If
            For lIdx = 0 To lCount - 1
                If lCompareMode = vbBinaryCompare Then
                    vItems(lIdx) = pvJsonDump(vKeys(lIdx)) & ":" & sSpace & pvJsonDump(oJson.Item(vKeys(lIdx)), Level + 1, Minimize, , IndentSize, MaxWidth)
                Else
                    vItems(lIdx) = pvJsonDump(oJson.Item(lIdx + JSON_IDX_OFFSET), Level + 1, Minimize, , IndentSize, MaxWidth)
                End If
                lSize = lSize + Len(vItems(lIdx))
            Next
            If lSize > MaxWidth And Not Minimize Then
                pvJsonDump = Left$(CompoundChars, 1) & vbCrLf & _
                    Space$(IIf(Level > -1, Level + 1, 0) * IndentSize) & Join(vItems, "," & vbCrLf & Space$(IIf(Level > -1, Level + 1, 0) * IndentSize)) & vbCrLf & _
                    Space$(IIf(Level > 0, Level, 0) * IndentSize) & Right$(CompoundChars, 1)
            Else
                pvJsonDump = Left$(CompoundChars, 1) & sSpace & Join(vItems, "," & sSpace) & sSpace & Right$(CompoundChars, 1)
            End If
        End If
    Case vbNull, vbEmpty
        pvJsonDump = "null"
    Case vbDate
        pvJsonDump = """" & Format$(vJson, "yyyy\-mm\-dd hh:nn:ss") & """"
        If Left$(pvJsonDump, 12) = """1899-12-30 " Then pvJsonDump = """" & Mid$(pvJsonDump, 13)
    Case vbBoolean
        pvJsonDump = IIf(vJson, "true", "false")
    Case vbString
        If vJson Like "*[?""\" & Chr$(0) & "-" & Chr$(31) & "]*" Then
            If IsEmpty(vTranscode) Then
                vTranscode = Split("\u0000|\u0001|\u0002|\u0003|\u0004|\u0005|\u0006|\u0007|\b|\t|\n|\u000B|\f|\r|\u000E|\u000F|\u0010|\u0011|" & _
                                   "\u0012|\u0013|\u0014|\u0015|\u0016|\u0017|\u0018|\u0019|\u001A|\u001B|\u001C|\u001D|\u001E|\u001F", "|")
            End If
            For lIdx = 1 To Len(vJson)
                lAsc = AscW(Mid$(vJson, lIdx, 1))
                If lAsc = 92 Or lAsc = 34 Then
                    pvJsonDump = pvJsonDump & "\" & ChrW$(lAsc)
                ElseIf lAsc >= 32 And lAsc < 256 Then
                    pvJsonDump = pvJsonDump & ChrW$(lAsc)
                ElseIf lAsc >= 0 And lAsc < 32 Then
                    pvJsonDump = pvJsonDump & vTranscode(lAsc)
                ElseIf Asc(Mid$(vJson, lIdx, 1)) <> 63 Or Mid$(vJson, lIdx, 1) = "?" Then
                    pvJsonDump = pvJsonDump & ChrW$(AscW(Mid$(vJson, lIdx, 1)))
                Else
                    pvJsonDump = pvJsonDump & "\u" & Right$("0000" & Hex$(lAsc), 4)
                End If
            Next
            pvJsonDump = """" & pvJsonDump & """"
        Else
            pvJsonDump = """" & vJson & """"
        End If
    Case Else
        If IsNumeric(vJson) Then
            pvJsonDump = Trim$(Replace(Replace(Str$(vJson), " .", "0."), "-.", "-0."))
        Else
            pvJsonDump = vJson & vbNullString
        End If
    End Select
End Function

'--- JSON Collection Helpers ---
Private Function pvJsonCreateObject(ByVal lCompareMode As VbCompareMethod) As VBA.Collection
    Set pvJsonCreateObject = New VBA.Collection
    memcpy ByVal ObjPtr(pvJsonCreateObject) + o_pvUnk5, lCompareMode, 4
End Function

Private Function pvJsonCompareMode(oJson As VBA.Collection) As VbCompareMethod
    memcpy pvJsonCompareMode, ByVal ObjPtr(oJson) + o_pvUnk5, 4
    pvJsonCompareMode = -(pvJsonCompareMode = vbTextCompare)
End Function

Private Property Get pvJsonItemGet(oParam As VBA.Collection, vKey As Variant) As Variant
    Dim lKey As Long
    On Error GoTo EH
    If VarType(vKey) = vbLong Then
        lKey = vKey + JSON_IDX_OFFSET
        If lKey > 0 And lKey <= oParam.Count Then pvJsonAssign pvJsonItemGet, oParam.Item(lKey)
    Else
        If pvJsonCollIndexByKey(oParam, vKey, JSON_DEF_IGNORE_CASE) > 0 Then
            pvJsonAssign pvJsonItemGet, oParam.Item(vKey)
        End If
    End If
    Exit Property
EH:
End Property

Private Sub pvJsonItemLet(oParam As VBA.Collection, vKey As Variant, vValue As Variant)
    Dim lKey As Long
    On Error GoTo EH
    If VarType(vKey) = vbLong Then
        lKey = vKey + 1
        If lKey > 0 And lKey <= oParam.Count Then oParam.Remove lKey
        If lKey > 0 And lKey <= oParam.Count Then
            oParam.Add vValue, Before:=lKey
        Else
            Do While lKey - 1 > oParam.Count: oParam.Add Empty: Loop
            oParam.Add vValue
        End If
    Else
        lKey = pvJsonCollIndexByKey(oParam, vKey, JSON_DEF_IGNORE_CASE)
        If lKey > 0 Then oParam.Remove lKey
        If lKey > 0 And lKey <= oParam.Count Then
            oParam.Add vValue, vKey, Before:=lKey
        Else
            oParam.Add vValue, vKey
        End If
    End If
    Exit Sub
EH:
    Debug.Print "pvJsonItemLet: " & Err.Description
End Sub

Private Function pvJsonCollAllKeys(oCol As VBA.Collection, Optional ByVal StartIndex As Long) As String()
    Dim lPtr As Long, aRetVal() As String, lIdx As Long, sTemp As String
    If oCol.Count = 0 Then
        aRetVal = Split(vbNullString)
    Else
        ReDim aRetVal(StartIndex To StartIndex + oCol.Count - 1) As String
        lPtr = ObjPtr(oCol)
        For lIdx = LBound(aRetVal) To UBound(aRetVal)
            memcpy lPtr, ByVal lPtr + o_pNextIndexedItem, 4
            memcpy ByVal VarPtr(sTemp), ByVal lPtr + o_KeyPtr, 4
            aRetVal(lIdx) = sTemp
        Next
        memcpy ByVal VarPtr(sTemp), 0&, 4
    End If
    pvJsonCollAllKeys = aRetVal
End Function

Private Function pvJsonCollIndexByKey(oCol As VBA.Collection, ByVal sKey As String, Optional ByVal IgnoreCase As Boolean = True) As Long
    Dim lItemPtr As Long, lEofPtr As Long, lPtr As Long
    Dim sTemp As String, eMethod As VbCompareMethod

    If Not oCol Is Nothing Then
        memcpy lItemPtr, ByVal ObjPtr(oCol) + o_pRootTreeItem, 4
        memcpy lEofPtr, ByVal ObjPtr(oCol) + o_pEndTreePtr, 4
    End If
    eMethod = IIf(IgnoreCase, vbTextCompare, vbBinaryCompare)
    Do While lItemPtr <> lEofPtr
        memcpy ByVal VarPtr(sTemp), ByVal lItemPtr + o_KeyPtr, 4
        Select Case StrComp(sKey, sTemp, eMethod)
        Case Is < 0
            memcpy lItemPtr, ByVal lItemPtr + o_pLeftBranch, 4
        Case Is > 0
            memcpy lItemPtr, ByVal lItemPtr + o_pRightBranch, 4
        Case Else
            lPtr = ObjPtr(oCol)
            Do While lPtr <> lItemPtr
                memcpy lPtr, ByVal lPtr + o_pNextIndexedItem, 4
                pvJsonCollIndexByKey = pvJsonCollIndexByKey + 1
            Loop
            GoTo QH
        End Select
    Loop
QH:
    memcpy ByVal VarPtr(sTemp), 0&, 4
End Function

'--- JSON Type Helpers ---
Private Sub pvJsonAssign(vDest As Variant, vSrc As Variant)
    On Error GoTo QH
    If IsObject(vSrc) Then Set vDest = vSrc Else vDest = vSrc
QH:
End Sub

Private Function pvJsonC_Str(Value As Variant) As String
    Dim vDest As Variant
    If VarType(Value) = vbString Then
        pvJsonC_Str = Value
    ElseIf VariantChangeType(vDest, Value, JSON_VARIANT_ALPHABOOL, vbString) = 0 Then
        pvJsonC_Str = vDest
    End If
End Function

Private Function pvJsonC_Bool(Value As Variant) As Boolean
    Dim vDest As Variant
    If VarType(Value) = vbBoolean Then
        pvJsonC_Bool = Value
    ElseIf VariantChangeType(vDest, Value, JSON_VARIANT_ALPHABOOL, vbBoolean) = 0 Then
        pvJsonC_Bool = vDest
    End If
End Function

Private Function pvJsonC_Lng(Value As Variant) As Long
    Dim vDest As Variant
    If VarType(Value) = vbLong Then
        pvJsonC_Lng = Value
    ElseIf VariantChangeType(vDest, Value, 0, vbLong) = 0 Then
        pvJsonC_Lng = vDest
    End If
End Function

Private Function pvJsonC_Dbl(Value As Variant) As Double
    Dim vDest As Variant
    If VarType(Value) = vbDouble Then
        pvJsonC_Dbl = Value
    ElseIf VariantChangeType(vDest, Value, 0, vbDouble) = 0 Then
        pvJsonC_Dbl = vDest
    End If
End Function

Private Function pvJsonIsOnlyDigits(ByVal sText As String) As Boolean
    If LenB(sText) <> 0 Then pvJsonIsOnlyDigits = Not (sText Like "*[!0-9]*")
End Function

'==============================================================================
' PRIVATE HELPERS — Encryption (AES-256-CBC + PBKDF2-SHA256)
'==============================================================================
Private Sub pvEncryptBytes(baPlain() As Byte, sPassword As String, baOutput() As Byte)
    Dim baSalt(0 To AES_SALT_LEN - 1) As Byte
    Dim baDK(0 To AES_DERIVED_KEY_LEN - 1) As Byte
    Dim baKey(0 To AES_KEYLEN - 1) As Byte
    Dim baIV(0 To AES_IVLEN - 1) As Byte
    Dim baHK(0 To AES_HMAC_KEYLEN - 1) As Byte
    Dim baCipher() As Byte
    Dim baHmac(0 To AES_HMAC_LEN - 1) As Byte
    Dim lPlainLen As Long
    Dim lCipherLen As Long

    lPlainLen = UBound(baPlain) - LBound(baPlain) + 1
    '--- 1. Generate random salt
    BCryptGenRandom 0, VarPtr(baSalt(0)), AES_SALT_LEN, BCRYPT_USE_SYSTEM_PREFERRED_RNG
    '--- 2. Derive key
    pvPBKDF2 sPassword, baSalt, baDK
    '--- 3. Split: AES key (32) + IV (16) + HMAC key (32)
    CopyMemory baKey(0), baDK(0), AES_KEYLEN
    CopyMemory baIV(0), baDK(AES_KEYLEN), AES_IVLEN
    CopyMemory baHK(0), baDK(AES_KEYLEN + AES_IVLEN), AES_HMAC_KEYLEN
    '--- 4. AES-256-CBC encrypt
    pvAesCbc True, baKey, baIV, baPlain, lPlainLen, baCipher, lCipherLen
    '--- 5. HMAC-SHA256(hmac_key, salt + ciphertext)
    pvHmacSha256 baHK, baSalt, AES_SALT_LEN, baCipher, lCipherLen, baHmac
    '--- 6. Build output: magic(4) + salt(16) + hmac(32) + ciphertext
    ReDim baOutput(0 To CZQ_HEADER_LEN + lCipherLen - 1) As Byte
    baOutput(0) = CZQ_MAGIC_0: baOutput(1) = CZQ_MAGIC_1
    baOutput(2) = CZQ_MAGIC_2: baOutput(3) = CZQ_MAGIC_3
    CopyMemory baOutput(4), baSalt(0), AES_SALT_LEN
    CopyMemory baOutput(20), baHmac(0), AES_HMAC_LEN
    CopyMemory baOutput(CZQ_HEADER_LEN), baCipher(0), lCipherLen
End Sub

Private Function pvDecryptBytes(baInput() As Byte, sPassword As String, baPlain() As Byte) As Boolean
    Dim lInputLen As Long
    Dim baSalt(0 To AES_SALT_LEN - 1) As Byte
    Dim baStoredHmac(0 To AES_HMAC_LEN - 1) As Byte
    Dim baDK(0 To AES_DERIVED_KEY_LEN - 1) As Byte
    Dim baKey(0 To AES_KEYLEN - 1) As Byte
    Dim baIV(0 To AES_IVLEN - 1) As Byte
    Dim baHK(0 To AES_HMAC_KEYLEN - 1) As Byte
    Dim baCipher() As Byte
    Dim baComputedHmac(0 To AES_HMAC_LEN - 1) As Byte
    Dim lCipherLen As Long
    Dim lPlainLen As Long
    Dim lIdx As Long

    lInputLen = UBound(baInput) + 1
    '--- 1. Verify minimum size and magic
    If lInputLen < CZQ_HEADER_LEN + 16 Then
        m_sLastError = "File too small to be encrypted"
        Exit Function
    End If
    If baInput(0) <> CZQ_MAGIC_0 Or baInput(1) <> CZQ_MAGIC_1 Or _
       baInput(2) <> CZQ_MAGIC_2 Or baInput(3) <> CZQ_MAGIC_3 Then
        m_sLastError = "Invalid file header (not a czQlite encrypted file)"
        Exit Function
    End If
    '--- 2. Extract salt, stored HMAC, ciphertext
    CopyMemory baSalt(0), baInput(4), AES_SALT_LEN
    CopyMemory baStoredHmac(0), baInput(20), AES_HMAC_LEN
    lCipherLen = lInputLen - CZQ_HEADER_LEN
    ReDim baCipher(0 To lCipherLen - 1) As Byte
    CopyMemory baCipher(0), baInput(CZQ_HEADER_LEN), lCipherLen
    '--- 3. Derive key
    pvPBKDF2 sPassword, baSalt, baDK
    CopyMemory baKey(0), baDK(0), AES_KEYLEN
    CopyMemory baIV(0), baDK(AES_KEYLEN), AES_IVLEN
    CopyMemory baHK(0), baDK(AES_KEYLEN + AES_IVLEN), AES_HMAC_KEYLEN
    '--- 4. Verify HMAC
    pvHmacSha256 baHK, baSalt, AES_SALT_LEN, baCipher, lCipherLen, baComputedHmac
    For lIdx = 0 To AES_HMAC_LEN - 1
        If baStoredHmac(lIdx) <> baComputedHmac(lIdx) Then
            m_sLastError = "HMAC verification failed (wrong password or corrupted file)"
            Exit Function
        End If
    Next
    '--- 5. AES-256-CBC decrypt
    pvAesCbc False, baKey, baIV, baCipher, lCipherLen, baPlain, lPlainLen
    If lPlainLen <= 0 Then
        m_sLastError = "Decryption produced no output"
        Exit Function
    End If
    ReDim Preserve baPlain(0 To lPlainLen - 1) As Byte
    pvDecryptBytes = True
End Function

Private Sub pvPBKDF2(sPassword As String, baSalt() As Byte, baDerivedKey() As Byte)
    Dim hAlg As Long
    Dim baPass() As Byte

    baPass = pvToUtf8(sPassword)
    BCryptOpenAlgorithmProvider hAlg, StrPtr("SHA256"), 0, BCRYPT_ALG_HANDLE_HMAC_FLAG
    BCryptDeriveKeyPBKDF2 hAlg, VarPtr(baPass(0)), UBound(baPass) + 1, _
            VarPtr(baSalt(0)), AES_SALT_LEN, AES_KDF_ITER / 10000@, _
            VarPtr(baDerivedKey(0)), AES_DERIVED_KEY_LEN, 0
    BCryptCloseAlgorithmProvider hAlg, 0
End Sub

Private Sub pvAesCbc(ByVal bEncrypt As Boolean, baKey() As Byte, baIV() As Byte, _
        baInput() As Byte, ByVal lInputLen As Long, baOutput() As Byte, lOutputLen As Long)
    Dim hAlg As Long, hKey As Long
    Dim lKeyObjLen As Long, lDummy As Long
    Dim baKeyObj() As Byte
    Dim baIVCopy(0 To AES_IVLEN - 1) As Byte
    Dim sMode As String

    '--- open AES provider
    BCryptOpenAlgorithmProvider hAlg, StrPtr("AES"), 0, 0
    '--- set CBC mode
    sMode = "ChainingModeCBC"
    BCryptSetProperty hAlg, StrPtr("ChainingMode"), StrPtr(sMode), (Len(sMode) + 1) * 2, 0
    '--- get key object size
    BCryptGetProperty hAlg, StrPtr("ObjectLength"), lKeyObjLen, 4, lDummy, 0
    ReDim baKeyObj(0 To lKeyObjLen - 1) As Byte
    '--- generate key
    BCryptGenerateSymmetricKey hAlg, hKey, VarPtr(baKeyObj(0)), lKeyObjLen, _
            VarPtr(baKey(0)), AES_KEYLEN, 0
    '--- copy IV (BCrypt modifies it in-place)
    CopyMemory baIVCopy(0), baIV(0), AES_IVLEN
    If bEncrypt Then
        '--- get output size first
        BCryptEncrypt hKey, VarPtr(baInput(0)), lInputLen, 0, _
                VarPtr(baIVCopy(0)), AES_IVLEN, 0, 0, lOutputLen, BCRYPT_BLOCK_PADDING
        ReDim baOutput(0 To lOutputLen - 1) As Byte
        '--- reset IV copy (it was modified)
        CopyMemory baIVCopy(0), baIV(0), AES_IVLEN
        '--- encrypt
        BCryptEncrypt hKey, VarPtr(baInput(0)), lInputLen, 0, _
                VarPtr(baIVCopy(0)), AES_IVLEN, VarPtr(baOutput(0)), lOutputLen, lOutputLen, BCRYPT_BLOCK_PADDING
    Else
        '--- get output size first
        BCryptDecrypt hKey, VarPtr(baInput(0)), lInputLen, 0, _
                VarPtr(baIVCopy(0)), AES_IVLEN, 0, 0, lOutputLen, BCRYPT_BLOCK_PADDING
        ReDim baOutput(0 To lOutputLen - 1) As Byte
        '--- reset IV copy
        CopyMemory baIVCopy(0), baIV(0), AES_IVLEN
        '--- decrypt
        BCryptDecrypt hKey, VarPtr(baInput(0)), lInputLen, 0, _
                VarPtr(baIVCopy(0)), AES_IVLEN, VarPtr(baOutput(0)), lOutputLen, lOutputLen, BCRYPT_BLOCK_PADDING
    End If
    '--- cleanup
    BCryptDestroyKey hKey
    BCryptCloseAlgorithmProvider hAlg, 0
End Sub

Private Sub pvHmacSha256(baHmacKey() As Byte, baData1() As Byte, ByVal lData1Len As Long, _
        baData2() As Byte, ByVal lData2Len As Long, baOutput() As Byte)
    Dim hAlg As Long, hHash As Long
    Dim lHashObjLen As Long, lDummy As Long
    Dim baHashObj() As Byte

    BCryptOpenAlgorithmProvider hAlg, StrPtr("SHA256"), 0, BCRYPT_ALG_HANDLE_HMAC_FLAG
    BCryptGetProperty hAlg, StrPtr("ObjectLength"), lHashObjLen, 4, lDummy, 0
    ReDim baHashObj(0 To lHashObjLen - 1) As Byte
    BCryptCreateHash hAlg, hHash, VarPtr(baHashObj(0)), lHashObjLen, _
            VarPtr(baHmacKey(0)), AES_HMAC_KEYLEN, 0
    If lData1Len > 0 Then BCryptHashData hHash, VarPtr(baData1(0)), lData1Len, 0
    If lData2Len > 0 Then BCryptHashData hHash, VarPtr(baData2(0)), lData2Len, 0
    BCryptFinishHash hHash, VarPtr(baOutput(0)), AES_HMAC_LEN, 0
    BCryptDestroyHash hHash
    BCryptCloseAlgorithmProvider hAlg, 0
End Sub

Private Function pvReadFile(sPath As String) As Byte()
    Dim ff As Integer
    Dim baData() As Byte

    ff = FreeFile
    Open sPath For Binary Access Read As #ff
    If LOF(ff) > 0 Then
        ReDim baData(0 To LOF(ff) - 1) As Byte
        Get #ff, , baData
    End If
    Close #ff
    pvReadFile = baData
End Function

Private Sub pvWriteFile(sPath As String, baData() As Byte)
    Dim ff As Integer

    ff = FreeFile
    Open sPath For Binary Access Write As #ff
    Put #ff, , baData
    Close #ff
End Sub

Private Sub pvEnsureSqliteDll()
    '--- Auto-detect sqlite3.dll availability
    '    If not found, try to copy winsqlite3.dll from System32 or SysWOW64 (Win10+)
    '    This allows zero-config deployment on Windows 10+ without shipping sqlite3.dll
    Dim hLib As Long
    hLib = LoadLibraryW(StrPtr("sqlite3.dll"))
    If hLib <> 0 Then
        '--- sqlite3.dll is available (either shipped or already copied)
        FreeLibrary hLib
        Exit Sub
    End If
    '--- sqlite3.dll not found, try winsqlite3.dll fallback (Win10+)
    Dim sWinRoot    As String
    Dim sSrcPath    As String
    Dim sAppDir     As String
    Dim sDstPath    As String
    sWinRoot = Environ$("SystemRoot")
    '--- check SysWOW64 first (32-bit DLL on 64-bit OS), then System32
    sSrcPath = sWinRoot & "\SysWOW64\winsqlite3.dll"
    If Dir$(sSrcPath) = "" Then
        sSrcPath = sWinRoot & "\System32\winsqlite3.dll"
    End If
    If Dir$(sSrcPath) <> "" Then
        '--- determine target directory
        On Error Resume Next
        sAppDir = App.Path
        If Len(sAppDir) = 0 Then sAppDir = CurDir$
        On Error GoTo 0
        sDstPath = sAppDir & "\sqlite3.dll"
        If Dir$(sDstPath) = "" Then
            On Error Resume Next
            FileCopy sSrcPath, sDstPath
            On Error GoTo 0
        End If
    End If
End Sub

'==============================================================================
' USERCONTROL LIFECYCLE
'==============================================================================
Private Sub UserControl_Initialize()
    m_lBusyTimeout = 5000
    m_bMapNullEmpty = True
    pvEnsureSqliteDll
End Sub

Private Sub UserControl_InitProperties()
    m_sDatabasePath = ""
    m_sPassword = ""
    m_lBusyTimeout = 5000
    m_bAutoOpen = False
    m_bMapNullEmpty = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_sDatabasePath = PropBag.ReadProperty("DatabasePath", "")
    m_sPassword = PropBag.ReadProperty("Password", "")
    m_lBusyTimeout = PropBag.ReadProperty("BusyTimeout", 5000)
    m_bAutoOpen = PropBag.ReadProperty("AutoOpen", False)
    m_bMapNullEmpty = PropBag.ReadProperty("MapNullToEmpty", True)
    '--- auto-open if requested and at runtime
    If Ambient.UserMode And m_bAutoOpen And Len(m_sDatabasePath) > 0 Then
        OpenDB
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "DatabasePath", m_sDatabasePath, ""
    PropBag.WriteProperty "Password", m_sPassword, ""
    PropBag.WriteProperty "BusyTimeout", m_lBusyTimeout, 5000
    PropBag.WriteProperty "AutoOpen", m_bAutoOpen, False
    PropBag.WriteProperty "MapNullToEmpty", m_bMapNullEmpty, True
End Sub

Private Sub UserControl_Resize()
    '--- keep control small at design-time, invisible at runtime
    Static bInResize As Boolean

    If bInResize Then
        Exit Sub
    End If
    bInResize = True
    UserControl.Width = 420
    UserControl.Height = 420
    bInResize = False
End Sub

Private Sub UserControl_Terminate()
    JsonClose
    If m_bIsOpen Then
        CloseDB
    End If
End Sub
