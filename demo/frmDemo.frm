VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "czQlite Demo"
   ClientHeight    =   7470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdModifyEnc 
      Caption         =   "Modify Enc"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   12
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpenEnc 
      Caption         =   "Open Enc"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdTestJsonParse 
      Caption         =   "JSON Parse"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdTestEncrypt 
      Caption         =   "Test Encrypt"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdTestJSON 
      Caption         =   "Test JSON"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdTransaction 
      Caption         =   "Transaction"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete ID=1"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Query All"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtSQL 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "SELECT * FROM Users"
      Top             =   1080
      Width           =   8055
   End
   Begin VB.TextBox txtLog 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5835
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1560
      Width           =   8535
   End
   Begin czQliteDemo.czQlite czQlite1 
      Left            =   120
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Label lblSQL 
      AutoSize        =   -1  'True
      Caption         =   "SQL:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8280
      TabIndex        =   2
      Top             =   1125
      Width           =   330
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Log "=== czQlite Demo ==="
    Log ""
    '--- setup database
    czQlite1.DatabasePath = App.Path & "\demo.db"
    czQlite1.BusyTimeout = 5000
    '--- open
    If czQlite1.OpenDB() Then
        Log "Database opened: " & czQlite1.DatabasePath
        Log "SQLite version: " & czQlite1.SQLiteVersion
    Else
        Log "ERROR: " & czQlite1.LastError
        Exit Sub
    End If
    '--- create table if not exists
    czQlite1.Execute "CREATE TABLE IF NOT EXISTS Users (ID INTEGER PRIMARY KEY AUTOINCREMENT, Name TEXT, Age INTEGER, City TEXT)"
    Log "Table 'Users' ready."
    Log "TableExists: " & czQlite1.TableExists("Users")
    Log ""
    '--- show current data
    ShowAllUsers
End Sub

Private Sub cmdQuery_Click()
    Dim sSQL As String
    
    sSQL = Trim$(txtSQL.Text)
    If Len(sSQL) = 0 Then
        sSQL = "SELECT * FROM Users"
    End If
    Log ""
    Log "--- Query: " & sSQL & " ---"
    czQlite1.Query sSQL
    Log "RecordCount: " & czQlite1.RecordCount
    Log "FieldCount: " & czQlite1.FieldCount
    '--- show field names
    Dim lCol As Long
    Dim sHeader As String
    For lCol = 0 To czQlite1.FieldCount - 1
        If Len(sHeader) > 0 Then sHeader = sHeader & " | "
        sHeader = sHeader & czQlite1.FieldName(lCol)
    Next
    Log sHeader
    Log String$(Len(sHeader), "-")
    '--- iterate rows
    Do While Not czQlite1.EOF
        Dim sRow As String
        sRow = ""
        For lCol = 0 To czQlite1.FieldCount - 1
            If Len(sRow) > 0 Then sRow = sRow & " | "
            sRow = sRow & czQlite1.Field(czQlite1.FieldName(lCol))
        Next
        Log sRow
        czQlite1.MoveNext
    Loop
    Log "---"
End Sub

Private Sub cmdInsert_Click()
    Dim sNames As Variant
    Dim sCities As Variant
    
    sNames = Array("Andi", "Budi", "Citra", "Dewi", "Eka", "Fajar", "Gita", "Hadi")
    sCities = Array("Jakarta", "Bandung", "Surabaya", "Yogyakarta", "Semarang", "Malang", "Medan", "Bali")
    '--- insert random user
    Randomize
    Dim sName As String
    Dim lAge As Long
    Dim sCity As String
    sName = sNames(Int(Rnd * 8))
    lAge = 20 + Int(Rnd * 40)
    sCity = sCities(Int(Rnd * 8))
    czQlite1.Execute "INSERT INTO Users (Name, Age, City) VALUES (?, ?, ?)", sName, lAge, sCity
    Log ""
    Log "Inserted: " & sName & ", age " & lAge & ", city " & sCity
    Log "LastInsertID: " & czQlite1.LastInsertID
    Log "AffectedRows: " & czQlite1.AffectedRows
    '--- show updated data
    ShowAllUsers
End Sub

Private Sub cmdDelete_Click()
    czQlite1.Execute "DELETE FROM Users WHERE ID = ?", 1
    Log ""
    Log "Deleted ID=1, AffectedRows: " & czQlite1.AffectedRows
    ShowAllUsers
End Sub

Private Sub cmdTransaction_Click()
    Log ""
    Log "--- Transaction Test ---"
    czQlite1.BeginTrans
    Log "BEGIN"
    czQlite1.Execute "INSERT INTO Users (Name, Age, City) VALUES (?, ?, ?)", "TxnUser1", 99, "TxnCity"
    Log "INSERT TxnUser1"
    czQlite1.Execute "INSERT INTO Users (Name, Age, City) VALUES (?, ?, ?)", "TxnUser2", 88, "TxnCity"
    Log "INSERT TxnUser2"
    czQlite1.CommitTrans
    Log "COMMIT"
    ShowAllUsers
End Sub

Private Sub cmdTestJSON_Click()
    Log ""
    Log "--- JSON Output ---"
    Dim sJson As String
    sJson = czQlite1.QueryJSON("SELECT Name, Age, City FROM Users")
    Log sJson
    Log ""
    Log "--- JSON Pretty ---"
    czQlite1.Query "SELECT Name, Age, City FROM Users LIMIT 2"
    sJson = czQlite1.ToJSON(Pretty:=True)
    Log sJson
    Log ""
    Log "--- GetVal ---"
    Log "User count: " & czQlite1.GetVal("SELECT COUNT(*) FROM Users")
    Log "Max age: " & czQlite1.GetVal("SELECT MAX(Age) FROM Users")
    Log ""
    Log "--- Named Parameter ---"
    czQlite1.Query "SELECT * FROM Users WHERE City = :kota", ":kota", "Jakarta"
    Log "Users in Jakarta: " & czQlite1.RecordCount
    Do While Not czQlite1.EOF
        Log "  " & czQlite1.Field("Name") & " (age " & czQlite1.Field("Age") & ")"
        czQlite1.MoveNext
    Loop
End Sub

Private Sub cmdTestEncrypt_Click()
    Dim sEncPath As String
    Dim sOrigPath As String
    
    Log ""
    Log "=== Encryption Test ==="
    sEncPath = App.Path & "\secure_test.db"
    sOrigPath = czQlite1.DatabasePath
    '--- delete old file if exists
    If Dir$(sEncPath) <> "" Then Kill sEncPath
    '--- close current DB first
    czQlite1.CloseDB
    
    '--- Step 1: Create new encrypted database
    Log "1. Creating encrypted database..."
    czQlite1.DatabasePath = sEncPath
    czQlite1.Password = "test123"
    If czQlite1.OpenDB() Then
        Log "   Opened :memory: (new encrypted DB)"
    Else
        Log "   ERROR: " & czQlite1.LastError
        GoTo RestoreDB
    End If
    czQlite1.Execute "CREATE TABLE Secrets (ID INTEGER PRIMARY KEY, Code TEXT, Value TEXT)"
    czQlite1.Execute "INSERT INTO Secrets (Code, Value) VALUES (?, ?)", "API_KEY", "sk-abc123xyz"
    czQlite1.Execute "INSERT INTO Secrets (Code, Value) VALUES (?, ?)", "DB_PASS", "p@ssw0rd!"
    czQlite1.Execute "INSERT INTO Secrets (Code, Value) VALUES (?, ?)", "TOKEN", "eyJhbGciOi..."
    Log "   Inserted 3 secret records"
    czQlite1.CloseDB
    Log "   Closed + encrypted to file"
    Log "   File size: " & FileLen(sEncPath) & " bytes"
    
    '--- Step 2: Reopen with correct password
    Log "2. Reopening with correct password..."
    czQlite1.DatabasePath = sEncPath
    czQlite1.Password = "test123"
    If czQlite1.OpenDB() Then
        Log "   Decrypted and opened successfully!"
        czQlite1.Query "SELECT * FROM Secrets"
        Log "   Records found: " & czQlite1.RecordCount
        Do While Not czQlite1.EOF
            Log "   [" & czQlite1.Field("Code") & "] = " & czQlite1.Field("Value")
            czQlite1.MoveNext
        Loop
        Dim sJson As String
        sJson = czQlite1.QueryJSON("SELECT Code, Value FROM Secrets")
        Log "   JSON: " & sJson
    Else
        Log "   ERROR: " & czQlite1.LastError
    End If
    czQlite1.CloseDB
    
    '--- Step 3: Try with wrong password
    Log "3. Trying with WRONG password..."
    czQlite1.DatabasePath = sEncPath
    czQlite1.Password = "wrong_password"
    If czQlite1.OpenDB() Then
        Log "   ERROR: Should have failed!"
        czQlite1.CloseDB
    Else
        Log "   Correctly rejected: " & czQlite1.LastError
    End If
    
    Log "=== Encryption Test Complete ==="
    
RestoreDB:
    '--- Restore original database
    czQlite1.Password = ""
    czQlite1.DatabasePath = sOrigPath
    If Len(sOrigPath) > 0 Then czQlite1.OpenDB
End Sub

Private Sub cmdTestJsonParse_Click()
    Log ""
    Log "=== JSON Parse Test ==="
    
    '--- Parse JSON string
    Dim sJson As String
    sJson = "{""name"":""czQlite"",""version"":1.0,""features"":[""sqlite"",""json"",""encryption""],""config"":{""timeout"":5000,""auto_open"":true}}"
    Log "Input: " & sJson
    Log ""
    
    If czQlite1.JsonParse(sJson) Then
        Log "Parsed OK!"
        Log "  name     = " & czQlite1.JsonStr("name")
        Log "  version  = " & czQlite1.JsonDbl("version")
        Log "  feature0 = " & czQlite1.JsonStr("features/0")
        Log "  feature1 = " & czQlite1.JsonStr("features/1")
        Log "  feature2 = " & czQlite1.JsonStr("features/2")
        Log "  count    = " & czQlite1.JsonLng("features/-1")
        Log "  timeout  = " & czQlite1.JsonLng("config/timeout")
        Log "  auto     = " & czQlite1.JsonBool("config/auto_open")
        Log "  has name = " & czQlite1.JsonHas("name")
        Log "  has foo  = " & czQlite1.JsonHas("foo")
    Else
        Log "Parse FAILED!"
    End If
    Log ""
    
    '--- Build JSON
    Log "--- Build JSON ---"
    czQlite1.JsonClose
    czQlite1.JsonSet "app", "Demo"
    czQlite1.JsonSet "db/host", "localhost"
    czQlite1.JsonSet "db/port", 3306
    Log "Minimized: " & czQlite1.JsonDump(Minimize:=True)
    Log "Pretty:"
    Log czQlite1.JsonDumpPretty(2)
    
    czQlite1.JsonClose
    Log "=== JSON Parse Test Complete ==="
End Sub

Private Sub cmdClear_Click()
    txtLog.Text = ""
End Sub

Private Sub cmdOpenEnc_Click()
    Dim sEncPath As String
    Dim sOrigPath As String
    
    Log ""
    Log "=== Open Encrypted DB ==="
    sEncPath = App.Path & "\secure_test.db"
    If Dir$(sEncPath) = "" Then
        Log "File not found: secure_test.db"
        Log "Run 'Test Encrypt' first to create it!"
        Exit Sub
    End If
    sOrigPath = czQlite1.DatabasePath
    czQlite1.CloseDB
    
    czQlite1.DatabasePath = sEncPath
    czQlite1.Password = "test123"
    If czQlite1.OpenDB() Then
        Log "Opened encrypted DB: " & sEncPath
        Log "IsEncrypted: " & czQlite1.IsEncrypted
        czQlite1.Query "SELECT * FROM Secrets ORDER BY ID"
        Log "Records: " & czQlite1.RecordCount
        Log ""
        Do While Not czQlite1.EOF
            Log "  [" & czQlite1.Field("ID") & "] " & czQlite1.Field("Code") & " = " & czQlite1.Field("Value")
            czQlite1.MoveNext
        Loop
        Log ""
        Log "JSON:"
        Log czQlite1.QueryJSON("SELECT * FROM Secrets")
    Else
        Log "ERROR: " & czQlite1.LastError
    End If
    czQlite1.CloseDB
    
    '--- restore original
    czQlite1.Password = ""
    czQlite1.DatabasePath = sOrigPath
    If Len(sOrigPath) > 0 Then czQlite1.OpenDB
    Log "=== Done ==="
End Sub

Private Sub cmdModifyEnc_Click()
    Dim sEncPath As String
    Dim sOrigPath As String
    
    Log ""
    Log "=== Modify Encrypted DB ==="
    sEncPath = App.Path & "\secure_test.db"
    If Dir$(sEncPath) = "" Then
        Log "File not found: secure_test.db"
        Log "Run 'Test Encrypt' first to create it!"
        Exit Sub
    End If
    sOrigPath = czQlite1.DatabasePath
    czQlite1.CloseDB
    
    czQlite1.DatabasePath = sEncPath
    czQlite1.Password = "test123"
    If czQlite1.OpenDB() Then
        Log "Opened encrypted DB"
        '--- show before
        Log "--- Before ---"
        czQlite1.Query "SELECT * FROM Secrets ORDER BY ID"
        Log "Records: " & czQlite1.RecordCount
        
        '--- add new record
        Dim sNewCode As String
        sNewCode = "KEY_" & Format$(Timer * 100, "0")
        czQlite1.Execute "INSERT INTO Secrets (Code, Value) VALUES (?, ?)", sNewCode, "added-" & Format$(Now, "hh:nn:ss")
        Log "Inserted: " & sNewCode
        
        '--- update existing record
        czQlite1.Execute "UPDATE Secrets SET Value = ? WHERE Code = ?", "UPDATED-" & Format$(Now, "hh:nn:ss"), "API_KEY"
        Log "Updated: API_KEY"
        
        '--- show after
        Log "--- After ---"
        czQlite1.Query "SELECT * FROM Secrets ORDER BY ID"
        Log "Records: " & czQlite1.RecordCount
        Do While Not czQlite1.EOF
            Log "  [" & czQlite1.Field("ID") & "] " & czQlite1.Field("Code") & " = " & czQlite1.Field("Value")
            czQlite1.MoveNext
        Loop
        Log ""
        Log "Closing (will re-encrypt to file)..."
    Else
        Log "ERROR: " & czQlite1.LastError
    End If
    czQlite1.CloseDB
    Log "File size: " & FileLen(sEncPath) & " bytes"
    
    '--- restore original
    czQlite1.Password = ""
    czQlite1.DatabasePath = sOrigPath
    If Len(sOrigPath) > 0 Then czQlite1.OpenDB
    Log "=== Modify Complete ==="
End Sub

Private Sub ShowAllUsers()
    Log ""
    Log "--- All Users ---"
    czQlite1.Query "SELECT * FROM Users ORDER BY ID"
    Do While Not czQlite1.EOF
        Log "  [" & czQlite1.Field("ID") & "] " & czQlite1.Field("Name") & _
            ", age " & czQlite1.Field("Age") & ", " & czQlite1.Field("City")
        czQlite1.MoveNext
    Loop
    Log "Total: " & czQlite1.RecordCount & " records"
End Sub

Private Sub czQlite1_OnError(ByVal ErrCode As Long, ByVal ErrMsg As String)
    Log "ERROR [" & ErrCode & "]: " & ErrMsg
End Sub

Private Sub czQlite1_AfterOpen()
    '--- event test
End Sub

Private Sub Form_Unload(Cancel As Integer)
    czQlite1.CloseDB
End Sub

Private Sub Log(sText As String)
    txtLog.Text = txtLog.Text & sText & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub
