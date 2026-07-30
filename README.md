# czQlite

**Truly Single-Instance SQLite UserControl for VB6**

Drop onto Form → Set property → Query. Done.

## Features

| Feature | Description |
|---------|-------------|
| 🗃️ **SQLite Engine** | Full SQLite via `sqlite3.dll` (ship with EXE, ~1.3 MB) |
| 📋 **JSON Engine** | Parse, Get/Set, Dump — adapted from mdJson.bas by wqweto |
| 🔒 **File Encryption** | AES-256-CBC + PBKDF2-SHA256 via `bcrypt.dll` (built-in Vista+) |
| 📦 **Single File** | Just 1 file `czQlite.ctl` + `sqlite3.dll` |
| 👻 **Invisible** | Invisible at runtime, like a Timer control |

## Requirements

- Windows 7 SP1+ with VB6 SP6
- `sqlite3.dll` — download from [sqlite.org](https://sqlite.org/download.html) → "Precompiled Binaries for Windows" → `sqlite-dll-win-x86-*.zip`

> **Auto-fallback on Windows 10+:** If `sqlite3.dll` is not found in the app directory,
> czQlite will automatically copy the built-in `winsqlite3.dll` from System32 as `sqlite3.dll`.
> This means on Windows 10+ you can deploy **without** shipping `sqlite3.dll` at all.
> On Windows 7/8, you must ship `sqlite3.dll` alongside your EXE.
>
> The encryption feature (serialize/deserialize) requires SQLite 3.36+.

## Quick Start

```vb
' 1. Add czQlite.ctl to your VB6 project
' 2. Drop czQlite1 onto a Form (invisible, like Timer)
' 3. Start coding:

Private Sub Form_Load()
    czQlite1.DatabasePath = App.Path & "\mydata.db"
    czQlite1.OpenDB
    
    czQlite1.Execute "CREATE TABLE IF NOT EXISTS Users " & _
        "(ID INTEGER PRIMARY KEY AUTOINCREMENT, Name TEXT, Age INTEGER)"
End Sub

' INSERT
czQlite1.Execute "INSERT INTO Users (Name, Age) VALUES (?, ?)", "Andi", 28

' SELECT + Navigate
czQlite1.Query "SELECT * FROM Users WHERE Age > ?", 25
Do While Not czQlite1.EOF
    Debug.Print czQlite1.Field("Name"), czQlite1.Field("Age")
    czQlite1.MoveNext
Loop

' JSON output directly from query
Dim sJson As String
sJson = czQlite1.QueryJSON("SELECT * FROM Users")
' → [{"Name":"Andi","Age":28},{"Name":"Siti","Age":30}]

' Scalar value
Dim cnt As Long
cnt = czQlite1.GetVal("SELECT COUNT(*) FROM Users")

' Named parameters
czQlite1.Query "SELECT * FROM Users WHERE Name = :name", ":name", "Andi"

' Transaction
czQlite1.BeginTrans
czQlite1.Execute "INSERT INTO ...", ...
czQlite1.CommitTrans   ' or .RollbackTrans

' Cleanup
Private Sub Form_Unload(Cancel As Integer)
    czQlite1.CloseDB
End Sub
```

## Encrypted Database

```vb
czQlite1.DatabasePath = App.Path & "\secure.db"
czQlite1.Password = "secret123"
czQlite1.OpenDB    ' file is decrypted into :memory:

' ... normal operations ...

czQlite1.SaveEncrypted  ' manual flush without closing
czQlite1.CloseDB        ' :memory: is serialized → encrypted → written to file
```

Encrypted files use a proprietary format:
- **Header:** `CZQ\x01` (4-byte magic)
- **Salt:** 16 bytes random (for PBKDF2)
- **HMAC:** 32 bytes SHA-256 (integrity check)
- **Data:** AES-256-CBC ciphertext (PKCS7 padded)

Key derivation: PBKDF2-SHA256 with 100,000 iterations.

## Byte Array / czStorage Integration

czStorage can embed data inside the EXE at design-time and read it at runtime (read-only).
Use `OpenFromBytes` to load an embedded database, and `SerializeToBytes` to export it.

```vb
' ═══════════════════════════════════════
' OPEN FROM BYTES (read embedded DB from czStorage)
' ═══════════════════════════════════════
Dim baDb() As Byte
baDb = czStorageReader1.GetFile("mydb.czq")     ' read encrypted DB from EXE
czQlite1.OpenFromBytes baDb, "secret123"        ' decrypt + load into :memory:

' ... normal query/execute ...

' ═══════════════════════════════════════
' SERIALIZE TO BYTES (export DB — czStorage is read-only at runtime)
' ═══════════════════════════════════════
Dim baOut() As Byte
baOut = czQlite1.SerializeToBytes()             ' encrypted byte array
baOut = czQlite1.SerializeToBytes("")            ' raw SQLite bytes (no encryption)

' Save to external file
Open App.Path & "\data.czq" For Binary Access Write As #1
Put #1, , baOut
Close #1

' Or use czQlite's built-in file save:
czQlite1.DatabasePath = App.Path & "\data.czq"
czQlite1.CloseDB    ' serialize → encrypt → write to file
```

> **Note:** czStorage is read-only at runtime — use `GetFile()` to read embedded data.
> Save modifications to an external encrypted file.

## JSON Engine

```vb
' Parse JSON from any source
czQlite1.JsonParse "{""db"":{""host"":""localhost"",""port"":3306}}"
Debug.Print czQlite1.JsonStr("db/host")    ' → "localhost"
Debug.Print czQlite1.JsonLng("db/port")    ' → 3306
Debug.Print czQlite1.JsonHas("db/name")    ' → False

' Build JSON programmatically
czQlite1.JsonSet "config/app", "MyApp"
czQlite1.JsonSet "config/version", 2
Debug.Print czQlite1.JsonDump()
' → {"config":{"app":"MyApp","version":2}}

' Pretty print
Debug.Print czQlite1.JsonDumpPretty(2)

' Query directly to JSON
sJson = czQlite1.QueryJSON("SELECT Name, Age FROM Users LIMIT 5")
sJson = czQlite1.ToJSON(Pretty:=True)
```

## API Reference

### Design-time Properties

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `DatabasePath` | String | "" | Path to the database file |
| `Password` | String | "" | Encryption password (empty = no encryption) |
| `BusyTimeout` | Long | 5000 | Busy lock timeout in milliseconds |
| `AutoOpen` | Boolean | False | Auto-open database at runtime |
| `MapNullToEmpty` | Boolean | True | Map NULL → "" in query results |

### Read-only Properties

| Property | Type | Description |
|----------|------|-------------|
| `IsOpen` | Boolean | Whether the database is currently open |
| `IsEncrypted` | Boolean | Whether the database is in encrypted mode |
| `AffectedRows` | Long | Rows affected by the last Execute |
| `LastInsertID` | Long | Last auto-increment ID |
| `SQLiteVersion` | String | SQLite engine version |
| `LastError` | String | Last error message |

### SQLite Methods

| Method | Description |
|--------|-------------|
| `OpenDB([path])` | Open a database |
| `CloseDB` | Close the database |
| `Query sql, [params]` | SELECT → internal recordset |
| `Execute sql, [params]` | INSERT/UPDATE/DELETE |
| `GetVal(sql, [params])` | Scalar query (single value) |
| `GetRows(sql, [params])` | Bulk query → 2D Variant array |
| `BeginTrans` | Begin a transaction |
| `CommitTrans` | Commit the transaction |
| `RollbackTrans` | Rollback the transaction |
| `TableExists(name)` | Check if a table exists |
| `Vacuum` | Compact the database |
| `BackupTo(path)` | Backup database to a file |
| `SaveEncrypted` | Flush encrypted DB to file without closing |
| `OpenFromBytes(bytes, [password])` | Open DB from byte array (czStorage integration) |
| `SerializeToBytes([password])` | Serialize DB to byte array (czStorage integration) |

### Recordset Navigation

| Property/Method | Description |
|----------------|-------------|
| `Field(nameOrIndex)` | Get column value by name or index |
| `EOF` | End of recordset |
| `BOF` | Beginning of recordset |
| `MoveNext` | Move to the next row |
| `MovePrevious` | Move to the previous row |
| `MoveFirst` | Move to the first row |
| `MoveLast` | Move to the last row |
| `RecordCount` | Number of rows |
| `FieldCount` | Number of columns |
| `FieldName(index)` | Column name by index |
| `AbsolutePosition` | Current row position |

### JSON Methods

| Method | Description |
|--------|-------------|
| `QueryJSON(sql, [params])` | Query → JSON string directly |
| `ToJSON([pretty])` | Convert last Query result to JSON |
| `JsonParse(json)` | Parse a JSON string |
| `JsonParseObject(json)` | Parse → return Collection object |
| `JsonGet(path)` | Get value by slash-separated path |
| `JsonStr/Lng/Dbl/Bool(path)` | Get typed value by path |
| `JsonObj(path)` | Get sub-object by path |
| `JsonHas(path)` | Check if a key exists |
| `JsonGetKeys([path])` | Get all keys at path |
| `JsonSet(path, value)` | Set a value by path |
| `JsonRemove(path)` | Remove a key by path |
| `JsonDump([minimize])` | Serialize to JSON string |
| `JsonDumpPretty([indent])` | Serialize with pretty formatting |
| `JsonClose` | Clear the JSON cache |

### Events

| Event | Description |
|-------|-------------|
| `BeforeOpen(Cancel)` | Fired before the database opens |
| `AfterOpen` | Fired after the database opens successfully |
| `BeforeClose` | Fired before the database closes |
| `AfterClose` | Fired after the database closes |
| `OnError(ErrCode, ErrMsg)` | Fired when an error occurs |

## Project Structure

```
czQlite/
├── src/
│   └── czQlite.ctl          ← THE ONLY FILE THAT MATTERS
├── demo/
│   ├── czQliteDemo.vbp      ← Demo project
│   └── frmDemo.frm          ← Demo form with 8 test buttons
└── README.md
```

## Credits

- **SQLite engine:** via Windows built-in `winsqlite3.dll`
- **JSON engine:** Adapted from `mdJson.bas` by [wqweto](https://github.com/wqweto)
- **Encryption:** AES-256-CBC + PBKDF2 via Windows built-in `bcrypt.dll`

## License

MIT
