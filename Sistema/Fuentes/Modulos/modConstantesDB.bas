Attribute VB_Name = "modConstantesDB"
'declaro constantes para manejo de tipos de datos ADO
Public Const dbSqlVarchar = 200
Public Const dbSqlSmallint = 2
Public Const dbSqlNumeric = 131
Public Const dbSqlDate = 135
Public Const dbSqlInt = 3
Public Const dbSqlChar = 125

'declaro constantes para manejo de errores de transacciones ADO
Public Const dbSqlDuplicateKey = 2601
Public Const dbSqlLoginFailed = 4002
Public Const dbSqlDataSource = 0
Public Const dbSqlDefaultDrivers = 0
Public Const dbSqlPermission = 229

'declaro constantes con tipos de Motores de Base de Datos
Public Const dbEngineSQLServer = "SQLServer"
Public Const dbEngineInformix5 = "Informix 5"

'declaro sintaxis de SQL de la función de fecha del motor
Public Const dbDateSQLServer = "GetDate()"
Public Const dbDateInformix5 = "CURRENT"


'para cambiar configuracion regional
Public Const LOCALE_USER_DEFAULT = &H400
Public Const LOCALE_SYSTEM_DEFAULT = &H800
Public Const LOCALE_ICURRDIGITS = &H19
Public Const LOCALE_SSHORTDATE = &H1F
Public Const LOCALE_SCURRENCY = &H14
Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long

Public Const LOCALE_SDATE = &H1D
Public Const LOCALE_SDECIMAL = &HE
Public Const LOCALE_STHOUSAND = &HF
Public Const LOCALE_SMONDECIMALSEP = &H16
Public Const LOCALE_SMONTHOUSANDSEP = &H17

