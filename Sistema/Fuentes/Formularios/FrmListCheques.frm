VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmListCheques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Cheques"
   ClientHeight    =   6255
   ClientLeft      =   1365
   ClientTop       =   975
   ClientWidth     =   8595
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmListCheques.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8595
   Begin VB.Frame fraImp 
      Caption         =   "Impresión de Reporte"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6180
      Left            =   30
      TabIndex        =   20
      Top             =   15
      Width           =   8550
      Begin VB.Frame fraImpresion 
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   165
         TabIndex        =   21
         Top             =   4365
         Width           =   8220
         Begin VB.CommandButton CmdCambiarImp 
            Caption         =   "&Configurar Impresora"
            Height          =   420
            Left            =   195
            TabIndex        =   37
            Top             =   675
            Width           =   1890
         End
         Begin VB.OptionButton oImpresora 
            Caption         =   "Impresora"
            Height          =   255
            Left            =   2475
            TabIndex        =   16
            Top             =   270
            Width           =   1215
         End
         Begin VB.OptionButton oPantalla 
            Caption         =   "Pantalla"
            Height          =   255
            Left            =   1245
            TabIndex        =   15
            Top             =   270
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Label LBImpActual 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Impresora Actual"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2235
            TabIndex        =   38
            Top             =   840
            Width           =   1485
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Destino:"
            Height          =   195
            Left            =   480
            TabIndex        =   36
            Top             =   270
            Width           =   600
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4140
         Left            =   3825
         TabIndex        =   24
         Top             =   240
         Width           =   4560
         Begin VB.ComboBox CboBanco 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   615
            Width           =   2775
         End
         Begin VB.TextBox TxtNroCheque 
            Enabled         =   0   'False
            Height          =   330
            Left            =   1650
            TabIndex        =   7
            Top             =   990
            Width           =   1080
         End
         Begin VB.ComboBox CboEstado 
            Height          =   315
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2145
            Width           =   2790
         End
         Begin MSComCtl2.DTPicker TxtFecVtoD 
            Height          =   315
            Left            =   1650
            TabIndex        =   4
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   42401793
            CurrentDate     =   42925
         End
         Begin MSComCtl2.DTPicker TxtFecVtoH 
            Height          =   315
            Left            =   3240
            TabIndex        =   5
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   42401793
            CurrentDate     =   42925
         End
         Begin MSComCtl2.DTPicker TxtFecIngresoH 
            Height          =   315
            Left            =   3270
            TabIndex        =   9
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   42401793
            CurrentDate     =   42925
         End
         Begin MSComCtl2.DTPicker TxtFecIngresoD 
            Height          =   315
            Left            =   1680
            TabIndex        =   8
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   42401793
            CurrentDate     =   42925
         End
         Begin MSComCtl2.DTPicker Fecha1 
            Height          =   315
            Left            =   1680
            TabIndex        =   43
            Top             =   2535
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   42401793
            CurrentDate     =   42925
         End
         Begin MSComCtl2.DTPicker Fecha2 
            Height          =   315
            Left            =   3270
            TabIndex        =   44
            Top             =   2535
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   42401793
            CurrentDate     =   42925
         End
         Begin MSComCtl2.DTPicker Fecha3 
            Height          =   315
            Left            =   1680
            TabIndex        =   45
            Top             =   3000
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   42401793
            CurrentDate     =   42925
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   1680
            TabIndex        =   46
            Top             =   3480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   42401793
            CurrentDate     =   42925
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   3270
            TabIndex        =   47
            Top             =   3480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   42401793
            CurrentDate     =   42925
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Index           =   3
            Left            =   3045
            TabIndex        =   41
            Top             =   3480
            Width           =   120
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            Height          =   195
            Left            =   915
            TabIndex        =   40
            Top             =   3465
            Width           =   510
         End
         Begin VB.Label Label10 
            Caption         =   "(Fecha Cambio Estado)"
            Height          =   585
            Left            =   180
            TabIndex        =   39
            Top             =   2205
            Width           =   600
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "(cheques en cartera)"
            Height          =   195
            Left            =   75
            TabIndex        =   35
            Top             =   405
            Width           =   1515
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   930
            TabIndex        =   34
            Top             =   3000
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            Height          =   195
            Left            =   915
            TabIndex        =   33
            Top             =   2595
            Width           =   510
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Index           =   2
            Left            =   3015
            TabIndex        =   32
            Top             =   2595
            Width           =   120
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Index           =   1
            Left            =   3030
            TabIndex        =   31
            Top             =   270
            Width           =   120
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Index           =   0
            Left            =   3030
            TabIndex        =   30
            Top             =   1425
            Width           =   120
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   885
            TabIndex        =   29
            Top             =   2220
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   915
            TabIndex        =   28
            Top             =   675
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nro de Cheque:"
            Height          =   195
            Left            =   300
            TabIndex        =   27
            Top             =   1020
            Width           =   1140
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Ingreso:"
            Height          =   195
            Left            =   135
            TabIndex        =   26
            Top             =   1455
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Vto.:"
            Height          =   195
            Left            =   375
            TabIndex        =   25
            Top             =   240
            Width           =   1065
         End
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         DisabledPicture =   "FrmListCheques.frx":27A2
         Height          =   420
         Left            =   7305
         TabIndex        =   19
         Top             =   5640
         Width           =   1080
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Aceptar"
         DisabledPicture =   "FrmListCheques.frx":2BEC
         Height          =   420
         Left            =   5115
         TabIndex        =   17
         Top             =   5640
         Width           =   1080
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         DisabledPicture =   "FrmListCheques.frx":34B6
         Height          =   420
         Left            =   6210
         TabIndex        =   18
         Top             =   5640
         Width           =   1080
      End
      Begin VB.Frame fraSentido 
         Caption         =   "Sentido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   165
         TabIndex        =   23
         Top             =   3420
         Width           =   3660
         Begin VB.OptionButton oDescendente 
            Caption         =   "Descendente"
            Height          =   255
            Left            =   1965
            TabIndex        =   14
            Top             =   435
            Width           =   1335
         End
         Begin VB.OptionButton oAscendente 
            Caption         =   "Ascendente"
            Height          =   255
            Left            =   210
            TabIndex        =   13
            Top             =   435
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame fraOrden 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3195
         Left            =   165
         TabIndex        =   22
         Top             =   240
         Width           =   3660
         Begin VB.OptionButton Option2 
            Caption         =   " en Cartera (Historico) al ..."
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   2315
            Width           =   2910
         End
         Begin VB.OptionButton Option6 
            Caption         =   "de quien Recibí a quien Entrege"
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   2730
            Width           =   2910
         End
         Begin VB.OptionButton Option5 
            Caption         =   " en Cartera al ..."
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1903
            Width           =   2910
         End
         Begin VB.OptionButton Option0 
            Caption         =   "... por Fecha de Vencimiento"
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   255
            Value           =   -1  'True
            Width           =   2910
         End
         Begin VB.OptionButton Option4 
            Caption         =   "... por Estado"
            Height          =   330
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1491
            Width           =   2910
         End
         Begin VB.OptionButton Option1 
            Caption         =   "... por Banco y Nro de Cheque"
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   667
            Width           =   2910
         End
         Begin VB.OptionButton Option3 
            Caption         =   "... por Fecha de Ingreso"
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1079
            Width           =   2910
         End
      End
      Begin Crystal.CrystalReport Rep 
         Left            =   3915
         Top             =   5670
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin MSComDlg.CommonDialog CDImpresora 
         Left            =   4350
         Top             =   5610
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Flags           =   64
      End
   End
End
Attribute VB_Name = "FrmListCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Limpio_Campos()
   Me.TxtFecVtoD.Value = Date
   Me.TxtFecVtoH.Value = Date
   Me.CboBanco.ListIndex = -1
   Me.TxtNroCheque.Text = ""
   Me.TxtFecIngresoD.Value = Date
   Me.TxtFecIngresoH.Value = Date
   Me.cboEstado.ListIndex = -1
   Me.fecha1.Value = Date
   Me.Fecha2.Value = Date
   Me.Fecha3.Value = Date
   FechaDesde.Value = Date
   FechaHasta.Value = Date
End Sub

Private Sub ListoChequesPorEntradaSalida()
    Dim mSumaTotChe As Double
    Dim mSumaTotCheS As Double
    Dim mCantChe As Integer
    Dim mCantCheS As Integer
    
    mCantChe = 0
    mSumaTotChe = 0
    mCantCheS = 0
    mSumaTotCheS = 0
    
    'BORRO TEMPORAL TMP_LISTADO_CHEQUES
    sql = "DELETE FROM TMP_LISTADO_CHEQUES"
    DBConn.Execute sql
    '---------CHEQUES QUE ENTRARON--------------
    'BUSCO LOS CHEQUES QUE ENTRARON POR RECIBO
    sql = "SELECT DR.BAN_CODINT, DR.CHE_NUMERO, C.CLI_RAZSOC,"
    sql = sql & " TC.TCO_ABREVIA, R.REC_SUCURSAL, R.REC_NUMERO,"
    sql = sql & " R.REC_FECHA, V.BAN_DESCRI, V.CHE_IMPORT"
    sql = sql & " FROM RECIBO_CLIENTE R, DETALLE_RECIBO_CLIENTE DR,"
    sql = sql & " TIPO_COMPROBANTE TC, CLIENTE C, ChequeEstadoVigente V"
    sql = sql & " WHERE R.TCO_CODIGO=DR.TCO_CODIGO"
    sql = sql & " AND R.REC_NUMERO=DR.REC_NUMERO"
    sql = sql & " AND R.REC_SUCURSAL=DR.REC_SUCURSAL"
    sql = sql & " AND R.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND R.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND DR.BAN_CODINT=V.BAN_CODINT"
    sql = sql & " AND DR.CHE_NUMERO=V.CHE_NUMERO"
    sql = sql & " AND R.EST_CODIGO=3"
    If FechaDesde.Value <> "" Then
        sql = sql & " AND R.REC_FECHA >=" & XDQ(FechaDesde.Value)
    End If
    If FechaHasta.Value <> "" Then
        sql = sql & " AND R.REC_FECHA <=" & XDQ(FechaHasta.Value)
    End If
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Do While Rec.EOF = False
            sql = "INSERT INTO TMP_LISTADO_CHEQUES ("
            sql = sql & " COD_BANCO,CHE_NUMERO,QUIEN,TCO_DESCRI,"
            sql = sql & " COM_NUMERO,COM_FECHA,BAN_DESCRI,CHE_IMPORT,ENTSAL)"
            sql = sql & " VALUES ("
            sql = sql & XN(Rec!BAN_CODINT) & ","
            sql = sql & XS(Rec!CHE_NUMERO) & ","
            sql = sql & XS(Rec!CLI_RAZSOC) & ","
            sql = sql & XS(Rec!TCO_ABREVIA) & ","
            sql = sql & XS(Format(Rec!REC_SUCURSAL, "0000") & "-" & Format(Rec!REC_NUMERO, "00000000")) & ","
            sql = sql & XDQ(Rec!REC_FECHA) & ","
            sql = sql & XS(Rec!BAN_DESCRI) & ","
            sql = sql & XN(Rec!CHE_IMPORT) & ","
            sql = sql & XS("ENTRA") & ")"
            DBConn.Execute sql
            mSumaTotChe = mSumaTotChe + CDbl(Rec!CHE_IMPORT)
            mCantChe = mCantChe + 1
            Rec.MoveNext
        Loop
    End If
    Rec.Close
    
    'BUSCO LOS CHEQUES QUE ENTRARON POR INGRESO DE CAJA
    sql = "SELECT DC.BAN_CODINT, DC.CHE_NUMERO, C.CIGR_DESCRI,"
    sql = sql & " C.CIGR_NUMERO, C.CIGR_FECHA, V.BAN_DESCRI, V.CHE_IMPORT"
    sql = sql & " FROM CAJA_INGRESO C, DETALLE_CAJA_INGRESO DC,"
    sql = sql & " ChequeEstadoVigente V"
    sql = sql & " WHERE C.CIGR_NUMERO=DC.CIGR_NUMERO"
    sql = sql & " AND DC.BAN_CODINT=V.BAN_CODINT"
    sql = sql & " AND DC.CHE_NUMERO=V.CHE_NUMERO"
    If FechaDesde.Value <> "" Then
        sql = sql & " AND C.CIGR_FECHA >=" & XDQ(FechaDesde.Value)
    End If
    If FechaHasta.Value <> "" Then
        sql = sql & " AND C.CIGR_FECHA <=" & XDQ(FechaHasta.Value)
    End If
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Do While Rec.EOF = False
            sql = "INSERT INTO TMP_LISTADO_CHEQUES ("
            sql = sql & " COD_BANCO,CHE_NUMERO,QUIEN,TCO_DESCRI,"
            sql = sql & " COM_NUMERO,COM_FECHA,BAN_DESCRI,CHE_IMPORT,ENTSAL)"
            sql = sql & " VALUES ("
            sql = sql & XN(Rec!BAN_CODINT) & ","
            sql = sql & XS(Rec!CHE_NUMERO) & ","
            sql = sql & XS(Rec!CIGR_DESCRI) & ","
            sql = sql & XS("INGRESO CAJA") & ","
            sql = sql & XS(Format(Rec!CIGR_NUMERO, "00000000")) & ","
            sql = sql & XDQ(Rec!CIGR_FECHA) & ","
            sql = sql & XS(Rec!BAN_DESCRI) & ","
            sql = sql & XN(Rec!CHE_IMPORT) & ","
            sql = sql & XS("ENTRA") & ")"
            DBConn.Execute sql
            mSumaTotChe = mSumaTotChe + CDbl(Rec!CHE_IMPORT)
            mCantChe = mCantChe + 1
            Rec.MoveNext
        Loop
    End If
    Rec.Close
    '-------------CHEQUES QUE SALIERON-------------
    'BUSCO LOS CHEQUES QUE SALIERON POR ORDEN DE PAGO
    sql = "SELECT DO.BAN_CODINT, DO.CHE_NUMERO, P.PROV_RAZSOC,"
    sql = sql & " TC.TCO_ABREVIA, O.OPG_NUMERO,"
    sql = sql & " O.OPG_FECHA, V.BAN_DESCRI, V.CHE_IMPORT"
    sql = sql & " FROM ORDEN_PAGO O, DETALLE_ORDEN_PAGO DO,"
    sql = sql & " TIPO_COMPROBANTE TC, PROVEEDOR P, ChequeEstadoVigente V"
    sql = sql & " WHERE O.TCO_CODIGO=DO.TCO_CODIGO"
    sql = sql & " AND O.OPG_NUMERO=DO.OPG_NUMERO"
    sql = sql & " AND O.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND O.TPR_CODIGO=P.TPR_CODIGO"
    sql = sql & " AND O.PROV_CODIGO=P.PROV_CODIGO"
    sql = sql & " AND DO.BAN_CODINT=V.BAN_CODINT"
    sql = sql & " AND DO.CHE_NUMERO=V.CHE_NUMERO"
    sql = sql & " AND O.EST_CODIGO=3"
    If FechaDesde.Value <> "" Then
        sql = sql & " AND O.OPG_FECHA >=" & XDQ(FechaDesde.Value)
    End If
    If FechaHasta.Value <> "" Then
        sql = sql & " AND O.OPG_FECHA <=" & XDQ(FechaHasta.Value)
    End If
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Do While Rec.EOF = False
            sql = "INSERT INTO TMP_LISTADO_CHEQUES ("
            sql = sql & " COD_BANCO,CHE_NUMERO,QUIEN,TCO_DESCRI,"
            sql = sql & " COM_NUMERO,COM_FECHA,BAN_DESCRI,CHE_IMPORT,ENTSAL)"
            sql = sql & " VALUES ("
            sql = sql & XN(Rec!BAN_CODINT) & ","
            sql = sql & XS(Rec!CHE_NUMERO) & ","
            sql = sql & XS(Rec!PROV_RAZSOC) & ","
            sql = sql & XS(Rec!TCO_ABREVIA) & ","
            sql = sql & XS(Format(Rec!OPG_NUMERO, "00000000")) & ","
            sql = sql & XDQ(Rec!OPG_FECHA) & ","
            sql = sql & XS(Rec!BAN_DESCRI) & ","
            sql = sql & XN(Rec!CHE_IMPORT) & ","
            sql = sql & XS("SALE") & ")"
            DBConn.Execute sql
            mSumaTotCheS = mSumaTotCheS + CDbl(Rec!CHE_IMPORT)
            mCantCheS = mCantCheS + 1
            
            Rec.MoveNext
        Loop
    End If
    Rec.Close
    'BUSCO LOS CHEQUES QUE SALIERON POR EGRESO DE CAJA
    sql = "SELECT DC.BAN_CODINT, DC.CHE_NUMERO, C.CEGR_DESCRI,"
    sql = sql & " C.CEGR_NUMERO, C.CEGR_FECHA, V.BAN_DESCRI, V.CHE_IMPORT"
    sql = sql & " FROM CAJA_EGRESO C, DETALLE_CAJA_EGRESO DC,"
    sql = sql & " ChequeEstadoVigente V"
    sql = sql & " WHERE C.CEGR_NUMERO=DC.CEGR_NUMERO"
    sql = sql & " AND DC.BAN_CODINT=V.BAN_CODINT"
    sql = sql & " AND DC.CHE_NUMERO=V.CHE_NUMERO"
    If FechaDesde.Value <> "" Then
        sql = sql & " AND C.CEGR_FECHA >=" & XDQ(FechaDesde.Value)
    End If
    If FechaHasta.Value <> "" Then
        sql = sql & " AND C.CEGR_FECHA <=" & XDQ(FechaHasta.Value)
    End If
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Do While Rec.EOF = False
            sql = "INSERT INTO TMP_LISTADO_CHEQUES ("
            sql = sql & " COD_BANCO,CHE_NUMERO,QUIEN,TCO_DESCRI,"
            sql = sql & " COM_NUMERO,COM_FECHA,BAN_DESCRI,CHE_IMPORT,ENTSAL)"
            sql = sql & " VALUES ("
            sql = sql & XN(Rec!BAN_CODINT) & ","
            sql = sql & XS(Rec!CHE_NUMERO) & ","
            sql = sql & XS(Rec!CEGR_DESCRI) & ","
            sql = sql & XS("EGRESO CAJA") & ","
            sql = sql & XS(Format(Rec!CEGR_NUMERO, "00000000")) & ","
            sql = sql & XDQ(Rec!CEGR_FECHA) & ","
            sql = sql & XS(Rec!BAN_DESCRI) & ","
            sql = sql & XN(Rec!CHE_IMPORT) & ","
            sql = sql & XS("SALE") & ")"
            DBConn.Execute sql
            mSumaTotCheS = mSumaTotCheS + CDbl(Rec!CHE_IMPORT)
            mCantCheS = mCantCheS + 1
            Rec.MoveNext
        Loop
    End If
    Rec.Close
    
    'BUSCO LOS CHEQUES QUE SE DEPOSITARON
    sql = "SELECT V.BAN_CODINT, V.CHE_NUMERO, B.BOL_NUMERO,"
    sql = sql & " B.BOL_NUMERO, B.BOL_FECHA, V.BAN_DESCRI, V.CHE_IMPORT"
    sql = sql & " FROM BOL_DEPOSITO B, ChequeEstadoVigente V"
    sql = sql & " WHERE B.BAN_CODINT=V.BOL_BAN_CODINT"
    sql = sql & " AND B.CTA_NROCTA=V.CTA_NROCTA"
    sql = sql & " AND B.BOL_NUMERO=V.BOL_NUMERO"
    sql = sql & " AND B.EBO_CODIGO=1"
    If FechaDesde.Value <> "" Then
        sql = sql & " AND B.BOL_FECHA >=" & XDQ(FechaDesde.Value)
    End If
    If FechaHasta.Value <> "" Then
        sql = sql & " AND B.BOL_FECHA <=" & XDQ(FechaHasta.Value)
    End If
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Do While Rec.EOF = False
            sql = "INSERT INTO TMP_LISTADO_CHEQUES ("
            sql = sql & " COD_BANCO,CHE_NUMERO,QUIEN,TCO_DESCRI,"
            sql = sql & " COM_NUMERO,COM_FECHA,BAN_DESCRI,CHE_IMPORT,ENTSAL)"
            sql = sql & " VALUES ("
            sql = sql & XN(Rec!BAN_CODINT) & ","
            sql = sql & XS(Rec!CHE_NUMERO) & ","
            sql = sql & XS(Format(Rec!BOL_NUMERO, "000000000000")) & ","
            sql = sql & XS("BOLETA DE DEPOSITO") & ","
            sql = sql & XS(Format(Rec!BOL_NUMERO, "000000000000")) & ","
            sql = sql & XDQ(Rec!BOL_FECHA) & ","
            sql = sql & XS(Rec!BAN_DESCRI) & ","
            sql = sql & XN(Rec!CHE_IMPORT) & ","
            sql = sql & XS("SALE") & ")"
            DBConn.Execute sql
            mSumaTotCheS = mSumaTotCheS + CDbl(Rec!CHE_IMPORT)
            mCantCheS = mCantCheS + 1
            Rec.MoveNext
        Loop
    End If
    Rec.Close
    
    DBConn.Execute "DELETE FROM TMP_RESUMEN_CUENTA_BANCO"
    'MUESTRO EL LISTADO
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
    Rep.Formulas(3) = ""
    Rep.Formulas(4) = ""
    Rep.Formulas(5) = ""
    Rep.Formulas(0) = "TOTAL='" & Valido_Importe(CStr(mSumaTotChe)) & "'"
    Rep.Formulas(1) = "TOTALS='" & Valido_Importe(CStr(mSumaTotCheS)) & "'"
    Rep.Formulas(2) = "CANTIDAD='" & CStr(mCantChe) & "'"
    Rep.Formulas(3) = "CANTIDADS='" & CStr(mCantCheS) & "'"
    Call MuestroFechaCrystal(Me.FechaDesde, Me.FechaHasta, " ")
    Rep.WindowTitle = "Listado de Cheques Detallado"
    Rep.ReportFileName = DRIVE & DirReport & "cheque_detalle.rpt"
    
    If oImpresora = True Then
       Rep.Destination = 1
   Else
       Rep.Destination = 0
       Rep.WindowMinButton = 0
       Rep.WindowTitle = "Consulta de Cheques Detallada"
       Rep.WindowBorderStyle = 2
   End If
   
     Rep.Action = 1
     
     Rep.SelectionFormula = ""
     Rep.Formulas(0) = ""
     Rep.Formulas(1) = ""
End Sub

Private Sub CboEstado_LostFocus()
    If Me.Option4.Value = True Then fecha1.SetFocus
    If Me.Option0.Value = True Then Me.CmdAgregar.SetFocus
End Sub

Private Sub CmdAgregar_Click()
    sql = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
    Rep.Formulas(3) = ""
    Rep.Formulas(4) = ""
    Rep.Formulas(5) = ""
    'VALIDO LAS FECHAS
    If Option0.Value = True Then
        If TxtFecVtoD.Value = "" Then
            MsgBox "Falta ingresar la fecha desde la cual quiere consultar", vbExclamation, TIT_MSGBOX
            TxtFecVtoD.SetFocus
            Exit Sub
        End If
    ElseIf Option2.Value = True Then
        If Fecha3.Value = "" Then
            MsgBox "Falta ingresar la fecha desde la cual quiere consultar", vbExclamation, TIT_MSGBOX
            Fecha3.SetFocus
            Exit Sub
        End If
    ElseIf Option3.Value = True Then
        If TxtFecIngresoD.Value = "" Then
            MsgBox "Falta ingresar la fecha desde la cual quiere consultar", vbExclamation, TIT_MSGBOX
            TxtFecIngresoD.SetFocus
            Exit Sub
        End If
    ElseIf Option4.Value = True Then
        If fecha1.Value = "" Then
            MsgBox "Falta ingresar la fecha desde la cual quiere consultar", vbExclamation, TIT_MSGBOX
            fecha1.SetFocus
            Exit Sub
        End If
    ElseIf Option5.Value = True Then
        If Fecha3.Value = "" Then
            MsgBox "Falta ingresar la fecha desde la cual quiere consultar", vbExclamation, TIT_MSGBOX
            Fecha3.SetFocus
            Exit Sub
        End If
    ElseIf Option6.Value = True Then
        If FechaDesde.Value = "" Then
            MsgBox "Falta ingresar la fecha desde la cual quiere consultar", vbExclamation, TIT_MSGBOX
            FechaDesde.SetFocus
            Exit Sub
        End If
    End If
    
   On Error GoTo ErrorTrans
   
   Screen.MousePointer = vbHourglass
   
   If Option6.Value = True Then
        ListoChequesPorEntradaSalida
        Screen.MousePointer = vbNormal
        Exit Sub
   End If
   
   'Sentido del Orden
   If oAscendente.Value = True Then
      wSentido = "+"
      Rep.Formulas(1) = "sentido ='Sentido: ASCENDENTE'"
   Else
      wSentido = "-"
      Rep.Formulas(1) = "sentido ='Sentido: DESCENDENTE '"
   End If
   
   If Me.Option0.Value = True Then 'Por Fecha de Vencimiento
       
       If Me.TxtFecVtoD.Value = "" Or Me.TxtFecVtoH.Value = "" Then
          If Me.TxtFecVtoD.Value = "" Then
            Me.TxtFecVtoD.Value = Format(Date, "dd/mm/yyyy")
          ElseIf Me.TxtFecVtoH.Value = "" Then
            Me.TxtFecVtoH.Value = Format(Date, "dd/mm/yyyy")
          End If
       End If
       Call MuestroFechaCrystal(Me.TxtFecVtoD, Me.TxtFecVtoH, "Fecha de Vencimineto  ")
       
       '{ChequeEstadoVigente.ECH_CODIGO} = 1 Unicamente Cheques en Cartera
       sql = sql & "({ChequeEstadoVigente.ECH_CODIGO} = 1 and {ChequeEstadoVigente.CHE_FECVTO} >= DATE(" & Mid(TxtFecVtoD.Value, 7, 4) & "," & _
                                                            Mid(TxtFecVtoD.Value, 4, 2) & "," & _
                                                            Mid(TxtFecVtoD.Value, 1, 2) & ") and " & _
                      "{ChequeEstadoVigente.CHE_FECVTO} <= DATE(" & Mid(TxtFecVtoH.Value, 7, 4) & "," & _
                                                                    Mid(TxtFecVtoH.Value, 4, 2) & "," & _
                                                                    Mid(TxtFecVtoH.Value, 1, 2) & "))"
       wCondicion = wSentido & " {ChequeEstadoVigente.CHE_FECVTO}"
       wCondicion1 = wSentido & " {ChequeEstadoVigente.CHE_NUMERO}"
       Rep.Formulas(0) = "orden ='Ordenado por: FECHA DE VTO. Y NRO DE CHEQUE'"
       
   ElseIf Me.Option1.Value = True Then 'por Banco y Nº de Cheque
      
       If TxtNroCheque.Text = "" Then
          MsgBox "Ingrese el Nro de Cheque", vbExclamation, TIT_MSGBOX
          TxtNroCheque.SetFocus
          Exit Sub
       End If
       
       If CboBanco.List(CboBanco.ListIndex) <> "(Todos)" Then
            sql = sql & "{ChequeEstadoVigente.BAN_CODINT} =  " & XN(Me.CboBanco.ItemData(Me.CboBanco.ListIndex))
       End If
       If sql = "" Then
            sql = " {ChequeEstadoVigente.CHE_NUMERO} =  " & XS(TxtNroCheque.Text)
       Else
            sql = sql & " AND {ChequeEstadoVigente.CHE_NUMERO} =  " & XS(TxtNroCheque.Text)
       End If
       wCondicion = wSentido & " {ChequeEstadoVigente.CHE_NUMERO}"
       wCondicion1 = ""
       Rep.Formulas(0) = "orden ='Ordenado por: NÚMERO DE CHEQUE'"
          
   ElseIf Me.Option3.Value = True Then 'por Fecha de Ingreso
   
       If Me.TxtFecIngresoD.Value = "" Or Me.TxtFecIngresoH.Value = "" Then
          If Me.TxtFecIngresoD.Value = "" Then
            Me.TxtFecIngresoD.Value = Format(Date, "dd/mm/yyyy")
          ElseIf Me.TxtFecIngresoH.Value = "" Then
            Me.TxtFecIngresoH.Value = Format(Date, "dd/mm/yyyy")
          End If
       End If
       Call MuestroFechaCrystal(Me.TxtFecIngresoD, Me.TxtFecIngresoH, "Fecha de Ingreso  ")
       
       sql = sql & "{ChequeEstadoVigente.CHE_FECENT} >= DATE(" & Mid(TxtFecIngresoD.Value, 7, 4) & _
                                                      "," & Mid(TxtFecIngresoD.Value, 4, 2) & _
                                                      "," & Mid(TxtFecIngresoD.Value, 1, 2) & ")and " & _
                   "{ChequeEstadoVigente.CHE_FECENT} <= DATE(" & Mid(TxtFecIngresoH.Value, 7, 4) & "," & _
                                                            Mid(TxtFecIngresoH.Value, 4, 2) & "," & _
                                                            Mid(TxtFecIngresoH.Value, 1, 2) & ")"
       
       wCondicion = wSentido & " {ChequeEstadoVigente.CHE_FECENT}"
       wCondicion1 = wSentido & " {ChequeEstadoVigente.CHE_FECVTO}"
       
       Rep.Formulas(0) = "orden ='Ordenado por: FECHA DE INGRESO y FECHA DE VTO.'"
   
   ElseIf Me.Option4.Value = True Then 'por Estado y Fecha de Cambio de estado
   
       If fecha1.Value = "" Or Fecha2.Value = "" Then
          If fecha1.Value = "" Then
            fecha1.Value = Format(Date, "dd/mm/yyyy")
          ElseIf Fecha2.Value = "" Then
            Fecha2.Value = Format(Date, "dd/mm/yyyy")
          End If
       End If
       Call MuestroFechaCrystal(Me.fecha1, Me.Fecha2, "Fecha de Cambio Estado  ")
       
       sql = sql & " {ChequeEstadoVigente.CES_FECHA} >= DATE(" & Mid(fecha1.Value, 7, 4) & "," & _
                                                                    Mid(fecha1.Value, 4, 2) & "," & _
                                                                    Mid(fecha1.Value, 1, 2) & ") and " & _
                   "{ChequeEstadoVigente.CES_FECHA} <= DATE(" & Mid(Fecha2.Value, 7, 4) & "," & _
                                                                    Mid(Fecha2.Value, 4, 2) & "," & _
                                                                    Mid(Fecha2.Value, 1, 2) & ")"
       'por Estado
       If Me.cboEstado.List(Me.cboEstado.ListIndex) <> "(Todos)" Then
           If Me.cboEstado.List(Me.cboEstado.ListIndex) = "RECHAZADOS TODOS" Then
              sql = sql & " AND {ChequeEstadoVigente.ECH_CODIGO} >= 8 " & _
                            " AND {ChequeEstadoVigente.ECH_CODIGO} <= 24 "
           Else
              sql = sql & " AND {ChequeEstadoVigente.ECH_CODIGO} =  " & XN(Me.cboEstado.ItemData(Me.cboEstado.ListIndex))
           End If
       End If
       wCondicion = wSentido & " {ChequeEstadoVigente.CHE_FECVTO}"
       wCondicion1 = wSentido & " {ChequeEstadoVigente.CHE_NUMERO}"
       Rep.Formulas(0) = "orden ='Ordenado por: FECHA DE VTO. Y NRO. DE CHEQUE'"
   
   ElseIf Me.Option5.Value = True Then 'en Cartera a Fecha
        
       If Fecha3.Value = "" Then Fecha3.Value = Format(Date, "dd/mm/yyyy")
       Rep.Formulas(4) = "FECHA='En Cartera al  " & Fecha3.Value & "'"
       
       sql = sql & " {ChequeEstadoVigente.ECH_CODIGO} = 1"
       sql = sql & " AND {ChequeEstadoVigente.CHE_FECENT} <= DATE(" & Mid(Fecha3.Value, 7, 4) & "," & _
                                                            Mid(Fecha3.Value, 4, 2) & "," & _
                                                            Mid(Fecha3.Value, 1, 2) & ")" 'And " &" _
'                      "{ChequeEstadoVigente.CES_FECHA} <= DATE(" & Mid(Fecha3.value, 7, 4) & "," & _
'                                                                   Mid(Fecha3.value, 4, 2) & "," & _
'                                                                   Mid(Fecha3.value, 1, 2) & ")"
       
       wCondicion = wSentido & " {ChequeEstadoVigente.CHE_FECVTO}"
       wCondicion1 = wSentido & " {ChequeEstadoVigente.CHE_NUMERO}"
       Rep.Formulas(0) = "orden ='Ordenado por: FECHA DE VTO. Y NRO. DE CHEQUE'"
   
   ElseIf Me.Option2.Value = True Then 'CHEQUES EN CARTERA HISTORICO
        Rep.Formulas(4) = "FECHA='En Cartera Historico al  " & Fecha3.Value & "'"
        
        sql = "DELETE FROM TMP_LISTADO_CHEQUES_HISTO"
        DBConn.Execute sql
        
        sql = " INSERT INTO TMP_LISTADO_CHEQUES_HISTO"
        sql = sql & " SELECT CE.CES_FECHA, EC.ECH_DESCRI, B.BAN_DESCRI, B.BAN_BANCO,"
        sql = sql & " B.BAN_LOCALIDAD, B.BAN_SUCURSAL, B.BAN_CODIGO,"
        sql = sql & " C.CHE_NUMERO,C.CHE_IMPORT,C.CHE_FECVTO,C.CHE_MOTIVO"
        sql = sql & " FROM CHEQUE_ESTADOS CE,ESTADO_CHEQUE EC, BANCO B, CHEQUE C"
        sql = sql & " WHERE C.BAN_CODINT=B.BAN_CODINT"
        sql = sql & " AND C.BAN_CODINT=CE.BAN_CODINT"
        sql = sql & " AND C.CHE_NUMERO=CE.CHE_NUMERO"
        sql = sql & " AND CE.ECH_CODIGO=EC.ECH_CODIGO"
        sql = sql & " AND CE.ECH_CODIGO=1"
        sql = sql & " AND CE.CES_FECHA <=" & XDQ(Fecha3.Value)
        sql = sql & " AND C.CHE_NUMERO NOT IN (SELECT CHEQUE.CHE_NUMERO"
        sql = sql & " FROM CHEQUE_ESTADOS,ESTADO_CHEQUE, BANCO, CHEQUE"
        sql = sql & " WHERE CHEQUE.BAN_CODINT=BANCO.BAN_CODINT"
        sql = sql & " AND CHEQUE.BAN_CODINT=CHEQUE_ESTADOS.BAN_CODINT"
        sql = sql & " AND CHEQUE.CHE_NUMERO=CHEQUE_ESTADOS.CHE_NUMERO"
        sql = sql & " AND ESTADO_CHEQUE.ECH_CODIGO=CHEQUE_ESTADOS.ECH_CODIGO"
        sql = sql & " AND CHEQUE_ESTADOS.ECH_CODIGO <> 1"
        sql = sql & " AND CHEQUE_ESTADOS.CES_FECHA <=" & XDQ(Fecha3.Value) & ")"
        DBConn.Execute sql
        sql = ""
'       If Fecha3.value = "" Then Fecha3.value = Format(Date, "dd/mm/yyyy")
'       sql = sql & " {CHEQUE_ESTADOS.ECH_CODIGO} = 1"
'       sql = sql & " AND {CHEQUE_ESTADOS.CES_FECHA} <= DATE(" & Mid(Fecha3.value, 7, 4) & "," & _
'                                                            Mid(Fecha3.value, 4, 2) & "," & _
'                                                            Mid(Fecha3.value, 1, 2) & ")" 'And " &" _
''                      "{CHEQUE_ESTADOS.CES_FECHA} <= DATE(" & Mid(Fecha3.value, 7, 4) & "," & _
''                                                                   Mid(Fecha3.value, 4, 2) & "," & _
''                                                                   Mid(Fecha3.value, 1, 2) & ")"
'
'       wCondicion = wSentido & " {CHEQUE_ESTADOS.CES_FECHA}"
'       wCondicion1 = wSentido & " {CHEQUE.CHE_NUMERO}"
       Rep.Formulas(0) = "orden ='Ordenado por: FECHA DE ESTADO. Y NRO. DE CHEQUE'"
   End If
   
   If oImpresora = True Then
       Rep.Destination = 1
   Else
       Rep.Destination = 0
       Rep.WindowMinButton = 0
       Rep.WindowTitle = "Consulta de Cheques"
       Rep.WindowBorderStyle = 2
   End If
   
   Rep.SortFields(0) = wCondicion
   Rep.SortFields(1) = wCondicion1
   
   Rep.SelectionFormula = sql
   Rep.WindowState = crptMaximized
   Rep.WindowBorderStyle = crptNoBorder
   Rep.Connect = "Provider=MSDASQL.1;Persst Security Info=False;Data Source=" & SERVIDOR
   
   If Option2.Value = False Then
        Rep.ReportFileName = DRIVE & DirReport & "CHEQUE.rpt"
   Else
        DBConn.Execute "DELETE FROM TMP_RESUMEN_CUENTA_BANCO"
        Rep.ReportFileName = DRIVE & DirReport & "CHEQUE_HISTORICO.rpt"
   End If
   Rep.Action = 1
   
   Rep.Formulas(0) = ""
   Rep.Formulas(1) = ""
   Rep.Formulas(2) = ""
   Rep.Formulas(3) = ""
   Rep.Formulas(4) = ""
   Rep.Formulas(5) = ""
   Screen.MousePointer = vbNormal
   Exit Sub

ErrorTrans:
  Screen.MousePointer = 1
  MsgBox "Error intentando armar el reporte. " & Chr(13) & Err.Description, 16, TIT_MSGBOX
End Sub

Private Sub MuestroFechaCrystal(FechaD As Control, FechaH As Control, Tit As String)
    If FechaD.Text <> "" And FechaH.Text <> "" Then
        Rep.Formulas(4) = "FECHA='" & Tit & "Desde: " & FechaD.Text & "   Hasta: " & FechaH.Text & "'"
    ElseIf FechaD.Text <> "" And FechaH.Text = "" Then
        Rep.Formulas(4) = "FECHA='" & Tit & "Desde: " & FechaD.Text & "   Hasta: " & Date & "'"
    ElseIf FechaD.Text = "" And FechaH.Text <> "" Then
        Rep.Formulas(4) = "FECHA='" & Tit & "Desde: Inicio" & "   Hasta: " & FechaH.Text & "'"
    ElseIf FechaD.Text = "" And FechaH.Text = "" Then
        Rep.Formulas(4) = "FECHA='" & Tit & "Desde: Inicio" & "   Hasta: " & Date & "'"
    End If
End Sub
Private Sub CmdCambiarImp_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    LBImpActual.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub cmdCancelar_Click()
    Limpio_Campos
    Option0.Value = True
    Option1.Value = False
    Option3.Value = False
    Option4.Value = False
    oAscendente.Value = True
    oPantalla.Value = True
End Sub

Private Sub cmdSalir_Click()
    Set FrmListCheques = Nothing
    Unload Me
End Sub

Private Sub fecha1_LostFocus()
    'If Me.Option4.Value = True And Fecha1.value = "" Then Fecha1.value = Format(Date, "dd/mm/yyyy")
    'If Me.Option4.Value = True Then Fecha2.SetFocus
End Sub

Private Sub fecha2_LostFocus()
    'If Me.Option4.Value = True And Fecha2.value = "" Then Fecha2.value = Format(Date, "dd/mm/yyyy")
    'If Me.Option4.Value = True Then CmdAgregar.SetFocus
End Sub

Private Sub fecha3_LostFocus()
   'If Me.Option5.Value = True And Fecha3.value = "" Then Fecha3.value = Format(Date, "dd/mm/yyyy")
   'If Me.Option5.Value = True Then CmdAgregar.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then 'avanza de campo
        MySendKeys Chr(9)
        KeyAscii = 0
    End If
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    KeyPreview = True
    
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.Enabled = False
    Me.cboEstado.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    
    Me.Left = 0
    Me.Top = 0
    
    Set Rec = New ADODB.Recordset
    
    cboEstado.Clear
    cboEstado.AddItem "(Todos)"
    sql = "SELECT ECH_CODIGO, ECH_DESCRI FROM ESTADO_CHEQUE"
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Do While Not Rec.EOF
            cboEstado.AddItem Trim(Rec!ECH_DESCRI)
            cboEstado.ItemData(cboEstado.NewIndex) = Trim(Rec!ECH_CODIGO)
            Rec.MoveNext
        Loop
        Me.cboEstado.ListIndex = -1
    End If
    Rec.Close
    cboEstado.AddItem "RECHAZADOS TODOS"
    Me.MousePointer = 1
    
    LBImpActual.Caption = "Impresora Actual: " & Printer.DeviceName
    
    Option0_Click
End Sub

Private Sub oImpresora_Click()
  Me.LBImpActual.Caption = "Impresora Actual: " & Printer.DeviceName
  Me.LBImpActual.Visible = True
End Sub

Private Sub oPantalla_Click()
 ' Me.CDImpresora.Visible = False
  Me.LBImpActual.Visible = False
End Sub

Private Sub Option0_Click()
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = True
    Me.TxtFecVtoH.Enabled = True
    Me.CboBanco.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    Me.cboEstado.Enabled = False
    Me.fecha1.Enabled = False
    Me.Fecha2.Enabled = False
    Me.Fecha3.Enabled = False
    Me.FechaDesde.Enabled = False
    Me.FechaHasta.Enabled = False
    If Me.TxtFecVtoD.Visible = True Then Me.TxtFecVtoD.SetFocus
End Sub

Private Sub Option1_Click()
    Me.CboBanco.Clear
    Set Rec = New ADODB.Recordset
    sql = "SELECT DISTINCT B.BAN_CODINT,B.BAN_DESCRI"
    sql = sql & " FROM BANCO B, CHEQUE C"
    sql = sql & " WHERE C.BAN_CODINT = B.BAN_CODINT"
    sql = sql & " ORDER BY B.BAN_DESCRI"
    
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.RecordCount > 0 Then
        Rec.MoveFirst
            Me.CboBanco.AddItem "(Todos)"
        Do While Not Rec.EOF
            Me.CboBanco.AddItem Trim(Rec!BAN_DESCRI)
            Me.CboBanco.ItemData(Me.CboBanco.NewIndex) = Rec!BAN_CODINT
            Rec.MoveNext
        Loop
        Me.CboBanco.ListIndex = -1
    End If
    Rec.Close
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.ListIndex = -1
    Me.CboBanco.Enabled = True
    Me.TxtNroCheque.Enabled = True
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    Me.cboEstado.Enabled = False
    Me.fecha1.Enabled = False
    Me.Fecha2.Enabled = False
    Me.Fecha3.Enabled = False
    Me.FechaDesde.Enabled = False
    Me.FechaHasta.Enabled = False
    Me.CboBanco.SetFocus
End Sub

Private Sub Option2_Click()
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    Me.cboEstado.Enabled = False
    Me.fecha1.Enabled = False
    Me.Fecha2.Enabled = False
    Me.Fecha3.Enabled = True
    Me.FechaDesde.Enabled = False
    Me.FechaHasta.Enabled = False
    Me.Refresh
    Me.Fecha3.SetFocus
End Sub

Private Sub Option3_Click()
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = True
    Me.TxtFecIngresoH.Enabled = True
    Me.cboEstado.Enabled = False
    Me.fecha1.Enabled = False
    Me.Fecha2.Enabled = False
    Me.Fecha3.Enabled = False
    Me.FechaDesde.Enabled = False
    Me.FechaHasta.Enabled = False
    Me.TxtFecIngresoD.SetFocus
End Sub

Private Sub Option4_Click()
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    Me.cboEstado.ListIndex = 0
    Me.cboEstado.Enabled = True
    Me.fecha1.Enabled = True
    Me.Fecha2.Enabled = True
    Me.Fecha3.Enabled = False
    Me.FechaDesde.Enabled = False
    Me.FechaHasta.Enabled = False
    Me.cboEstado.SetFocus
End Sub

Private Sub Option5_Click()
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    Me.cboEstado.Enabled = False
    Me.fecha1.Enabled = False
    Me.Fecha2.Enabled = False
    Me.Fecha3.Enabled = True
    Me.FechaDesde.Enabled = False
    Me.FechaHasta.Enabled = False
    Me.Refresh
    Me.Fecha3.SetFocus
End Sub

Private Sub Option6_Click()
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    Me.cboEstado.Enabled = False
    Me.fecha1.Enabled = False
    Me.Fecha2.Enabled = False
    Me.Fecha3.Enabled = False
    Me.FechaDesde.Enabled = True
    Me.FechaHasta.Enabled = True
    Me.Refresh
    Me.FechaDesde.SetFocus
End Sub

Private Sub TxtFecIngresoD_LostFocus()
   'If Me.Option3.Value = True And TxtFecIngresoD.value = "" Then TxtFecIngresoD.value = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub TxtFecIngresoH_LostFocus()
'If Me.Option3.Value = True And TxtFecIngresoH.value = "" Then TxtFecIngresoH.value = Format(Date, "dd/mm/yyyy")

  If IsDate(TxtFecIngresoD.Value) And IsDate(TxtFecIngresoH.Value) Then
    
    If CVDate(TxtFecIngresoD.Value) > CVDate(TxtFecIngresoH.Value) Then
      MsgBox "La Fecha Hasta no puede ser inferior a la Fecha Desde. Verifique!", 16, TIT_MSGBOX
      TxtFecIngresoH.Value = ""
      TxtFecIngresoH.SetFocus
      Exit Sub
    Else
      If Not IsDate(TxtFecIngresoD.Value) Then TxtFecIngresoD.Value = ""
      If Not IsDate(TxtFecIngresoH.Value) Then TxtFecIngresoH.Value = ""
    End If
    
 End If
End Sub

Private Sub TxtFecVtoD_LostFocus()
  'If Me.Option0.Value = True And TxtFecVtoD.value = "" Then TxtFecVtoD.value = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub TxtFecVtoH_LostFocus()

  'If Me.Option0.Value = True And TxtFecVtoh.value = "" Then TxtFecVtoh.value = Format(Date, "dd/mm/yyyy")
  
  If IsDate(TxtFecVtoD.Value) And IsDate(TxtFecVtoH.Value) Then
  
    If CVDate(TxtFecVtoD.Value) > CVDate(TxtFecVtoH.Value) Then
      MsgBox "La Fecha Hasta no puede ser inferior a la Fecha Desde. Verifique!", 16, TIT_MSGBOX
      TxtFecVtoH.Value = ""
      TxtFecVtoD.SetFocus
      Exit Sub
    Else
      CmdAgregar.SetFocus
    End If
 End If
End Sub

Private Sub TxtNroCheque_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroTE(KeyAscii)
End Sub

Private Sub TxtNroCheque_LostFocus()
    If Me.Option1.Value = True And Me.TxtNroCheque.Text = "" Then
        MsgBox "El Número de Cheque no puede ser Nulo. Verifique!", 16, TIT_MSGBOX
        TxtNroCheque.SetFocus
    Else
        If Len(TxtNroCheque.Text) < 10 Then TxtNroCheque.Text = CompletarConCeros(TxtNroCheque.Text, 10)
        CmdAgregar.SetFocus
    End If
End Sub

