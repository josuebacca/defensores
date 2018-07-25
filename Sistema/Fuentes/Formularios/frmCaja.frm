VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre Caja"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraImpresion 
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   120
      TabIndex        =   30
      Top             =   7095
      Width           =   2175
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "frmCaja.frx":0000
         Left            =   450
         List            =   "frmCaja.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   270
         Width           =   1635
      End
      Begin VB.PictureBox picSalida 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   135
         Picture         =   "frmCaja.frx":002F
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   33
         Top             =   315
         Width           =   240
      End
      Begin VB.PictureBox picSalida 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   135
         Picture         =   "frmCaja.frx":0131
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   32
         Top             =   315
         Width           =   240
      End
      Begin VB.PictureBox picSalida 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   135
         Picture         =   "frmCaja.frx":0233
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   31
         Top             =   315
         Width           =   240
      End
   End
   Begin VB.CommandButton CBCancelar 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   8595
      TabIndex        =   29
      Top             =   7515
      Width           =   1065
   End
   Begin VB.CommandButton CBCerrar 
      Caption         =   "&Cierre"
      Height          =   405
      Left            =   7500
      TabIndex        =   28
      Top             =   7095
      Width           =   2145
   End
   Begin VB.CommandButton CBVolver 
      Caption         =   "&Nuevo"
      Height          =   405
      Left            =   7500
      TabIndex        =   27
      Top             =   7515
      Width           =   1065
   End
   Begin VB.Frame FraTipoReporte 
      Caption         =   "Tipos de Reportes a Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   2325
      TabIndex        =   23
      Top             =   7095
      Width           =   4230
      Begin VB.CheckBox ChkResumen 
         Caption         =   "Resumen y Saldo Final de Caja"
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   480
         Width           =   2550
      End
      Begin VB.CheckBox ChkCaja 
         Caption         =   "Caja"
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   270
         Width           =   735
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   405
         Left            =   2820
         TabIndex        =   24
         Top             =   285
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ver..."
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
      Left            =   6660
      TabIndex        =   13
      Top             =   75
      Width           =   3045
      Begin VB.OptionButton oResumido 
         Caption         =   "Listado Resumido"
         Height          =   195
         Left            =   3765
         TabIndex        =   17
         Top             =   120
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton CBImprimir 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   1695
         TabIndex        =   16
         Top             =   450
         Width           =   1095
      End
      Begin VB.OptionButton oDetallado 
         Caption         =   "Listado Detallado"
         Height          =   195
         Left            =   3855
         TabIndex        =   14
         Top             =   210
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker Fecha1 
         Height          =   315
         Left            =   255
         TabIndex        =   15
         Top             =   495
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         Format          =   48431105
         CurrentDate     =   38651
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Caja del Dia.:"
         Height          =   195
         Left            =   270
         TabIndex        =   18
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.Frame FrameUltimoCierre 
      Caption         =   "Datos último cierre de Caja Registrado"
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
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   6600
      Begin VB.TextBox txtSaldoF 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1005
         TabIndex        =   6
         Top             =   570
         Width           =   1200
      End
      Begin VB.TextBox txtSaldoI 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1005
         TabIndex        =   5
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox txtSaldoFEft 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   3090
         TabIndex        =   4
         Top             =   570
         Width           =   1200
      End
      Begin VB.TextBox txtSaldoIEft 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   3090
         TabIndex        =   3
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox txtSaldoFChe 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   5265
         TabIndex        =   2
         Top             =   570
         Width           =   1200
      End
      Begin VB.TextBox txtSaldoIChe 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   5265
         TabIndex        =   1
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Final:"
         Height          =   195
         Left            =   75
         TabIndex        =   12
         Top             =   615
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Inicial:"
         Height          =   195
         Left            =   75
         TabIndex        =   11
         Top             =   285
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "SaldoF Eft:"
         Height          =   195
         Left            =   2280
         TabIndex        =   10
         Top             =   615
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "SaldoI Eft:"
         Height          =   195
         Left            =   2280
         TabIndex        =   9
         Top             =   285
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "SaldoF Che:"
         Height          =   195
         Left            =   4380
         TabIndex        =   8
         Top             =   615
         Width           =   870
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "SaldoI Che:"
         Height          =   195
         Left            =   4380
         TabIndex        =   7
         Top             =   285
         Width           =   840
      End
   End
   Begin TabDlg.SSTab SSTabCaja 
      Height          =   5970
      Left            =   60
      TabIndex        =   19
      Top             =   1050
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   10530
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Caja"
      TabPicture(0)   =   "frmCaja.frx":0335
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GrdCaja"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Resumen y Saldo Final de Caja"
      TabPicture(1)   =   "frmCaja.frx":0351
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdResumen"
      Tab(1).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid GrdCaja 
         Height          =   5520
         Left            =   60
         TabIndex        =   20
         Top             =   360
         Width           =   9480
         _ExtentX        =   16722
         _ExtentY        =   9737
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   280
         BackColor       =   14737632
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid GrdResumen 
         Height          =   5520
         Left            =   -74925
         TabIndex        =   21
         Top             =   360
         Width           =   9480
         _ExtentX        =   16722
         _ExtentY        =   9737
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   280
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "EFECTIVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74850
         TabIndex        =   22
         Top             =   2430
         Width           =   915
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   6825
      Top             =   7245
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim mSaldo As Double
'Dim AplicoImpuesto As Boolean
'Dim ValorImpuesto As Double
Dim mTotMoneda As Double
Dim mTotCheque As Double
Dim mTotMoneda1 As Double
Dim mTotCheque1 As Double
Dim mTotIng As Double
Dim mTotIngResumen As Double
Dim mTotIngInicial As Double
Dim mTotEgre As Double
Dim mTotEgreResumen As Double
Dim TOTAL As Double

Dim FechaBusqueda As String

Private Sub CBCancelar_Click()
    Set frmCaja = Nothing
    Unload Me
End Sub

Private Sub CBCerrar_Click()
    sql = "INSERT INTO CAJA (CAJA_FECHA,"
    sql = sql & " CAJA_SALDOI,CAJA_SALDO_PESOSI,CAJA_SALDO_CHEQUESI,"
    sql = sql & " CAJA_SALDOF,CAJA_SALDO_PESOSF,CAJA_SALDO_CHEQUESF)"
    sql = sql & " VALUES ("
    sql = sql & XDQ(fecha1.Value) & ","
    sql = sql & XN(txtSaldoF.Text) & ","
    sql = sql & XN(txtSaldoFEft.Text) & ","
    sql = sql & XN(txtSaldoFChe.Text) & ","
    sql = sql & XN(CStr(mTotIngResumen - mTotEgreResumen)) & ","
    sql = sql & XN(CStr(mTotMoneda)) & ","
    sql = sql & XN("0") & ")"
    DBConn.Execute sql
    fecha1.Value = CDate(fecha1.Value) + 1
    CBVolver_Click
End Sub

Private Sub CBImprimir_Click()
    FechaBusqueda = ""
    mTotMoneda = 0
    mTotCheque = 0
    mTotIng = 0
    mTotIngResumen = 0
    mTotIngInicial = 0
    mTotEgre = 0
    mTotEgreResumen = 0
    TOTAL = 0
    GrdCaja.Rows = 1
    GrdResumen.Rows = 1
    
    cmdImprimir.Enabled = True
    CBCerrar.Enabled = True
    
    Screen.MousePointer = vbHourglass
    'CmdNuevo_Click
    'lblEstado.Caption = "Cargando datos de caja..."
    
    'Set rec = New ADODB.Recordset
    sql = "SELECT * FROM CAJA"
    sql = sql & " WHERE"
    sql = sql & " CAJA_FECHA = " & XDQ(fecha1.Value)
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
'        cmdGrabar.Enabled = False
'        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No puede cerrar la caja, ya que la fecha seleccionada es igual al de la última caja cerada", vbCritical, TIT_MSGBOX
        Rec.Close
        Exit Sub
    Else
        Rec.Close
'        If Valido_Caja = False Then
'            lblEstado.Caption = ""
'            Screen.MousePointer = vbNormal
'            Exit Sub
'        End If
'        cmdGrabar.Enabled = True
        sql = "SELECT *"
        sql = sql & " FROM CAJA"
        sql = sql & " WHERE"
        sql = sql & " CAJA_FECHA= (SELECT MAX(CAJA_FECHA) AS FECHA FROM CAJA)"
        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    End If
    'SALDO INICIAL
    If Rec.EOF = False Then
      FechaBusqueda = Rec!CAJA_FECHA
      txtSaldoI.Text = Valido_Importe(Chk0(Rec!CAJA_SALDOI))
      txtSaldoF.Text = Valido_Importe(Chk0(Rec!CAJA_SALDOF))
      txtSaldoIEft.Text = Valido_Importe(Chk0(Rec!CAJA_SALDO_PESOSI))
      txtSaldoFEft.Text = Valido_Importe(Chk0(Rec!CAJA_SALDO_PESOSF))
      txtSaldoIChe.Text = Valido_Importe(Chk0(Rec!CAJA_SALDO_CHEQUESI))
      txtSaldoFChe.Text = Valido_Importe(Chk0(Rec!CAJA_SALDO_CHEQUESF))
      mTotCheque = Chk0(Rec!CAJA_SALDO_CHEQUESF)
      mTotMoneda = Chk0(Rec!CAJA_SALDO_PESOSF)
      'mTotIngInicial = Chk0(rec!CAJA_SALDOF)
      mTotIngInicial = Chk0(Rec!CAJA_SALDO_PESOSF)
    End If
    Rec.Close
    
    RESUMEN_CAJA_INICIAL
    
    GrdCaja.AddItem "INGRESOS"
    GrdCaja.Col = 0
    GrdCaja.row = GrdCaja.Rows - 1
    For I = 0 To 3
        GrdCaja.Col = I
        'GrdCaja.CellBackColor = &HC0FFC0
        GrdCaja.CellFontBold = True
    Next
   
    
    'INGRESOS
    Call BUSCO_RECIBOS(FechaBusqueda)
    Call BUSCO_INGRESO_CAJA(FechaBusqueda)
    GrdCaja.AddItem "" & Chr(9) & "TOTAL INGRESOS" & Chr(9) & Valido_Importe(CStr(mTotIng))
    
    mTotIng = mTotIng + mTotIngInicial
    mTotIngResumen = mTotIngResumen + mTotIngInicial
    
    GrdCaja.row = GrdCaja.Rows - 1
    For I = 0 To 3
        GrdCaja.Col = I
        GrdCaja.CellBackColor = &HFF8080
        GrdCaja.CellForeColor = &HFFFFFF    'FUENTE COLOR BLANCO
        GrdCaja.CellFontBold = True
    Next
    
    'EGRESOS
    GrdCaja.AddItem "EGRESOS"
    GrdCaja.Col = 0
    GrdCaja.row = GrdCaja.Rows - 1
    For I = 0 To 3
        GrdCaja.Col = I
        'GrdCaja.CellBackColor = &HC0FFC0
        GrdCaja.CellFontBold = True
    Next
    GrdResumen.AddItem "EGRESOS"
    GrdResumen.Col = 0
    GrdResumen.row = GrdResumen.Rows - 1
    For I = 0 To 3
        GrdResumen.Col = I
        'GrdResumen.CellBackColor = &HC0FFC0
        GrdResumen.CellFontBold = True
    Next
    
    Call BUSCO_EGRESO_CAJA(FechaBusqueda)
    
    GrdCaja.AddItem "" & Chr(9) & "TOTAL EGRESOS" & Chr(9) & "" & Chr(9) & Valido_Importe(CStr(mTotEgre))
    'GrdCaja.Col = 0
    GrdCaja.row = GrdCaja.Rows - 1
    For I = 0 To 3
        GrdCaja.Col = I
        GrdCaja.CellBackColor = &HFF8080 '&HFF8080
        GrdCaja.CellForeColor = &HFFFFFF    'FUENTE COLOR BLANCO
        GrdCaja.CellFontBold = True
    Next
    
    RESUMEN_CAJA_FINAL
    Screen.MousePointer = vbNormal
    
    
End Sub

Private Sub RESUMEN_CAJA_INICIAL()
    GrdResumen.AddItem "SALDO INICIAL"
    GrdResumen.Col = 0
    GrdResumen.row = GrdResumen.Rows - 1
    For I = 0 To 3
        GrdResumen.Col = I
        'GrdResumen.CellBackColor = &HC0FFC0
        GrdResumen.CellFontBold = True
    Next
    GrdResumen.AddItem "    CHEQUES" & Chr(9) & "" & Chr(9) & txtSaldoFChe.Text
    PintoBlancoResumenIng
    GrdResumen.AddItem "    EFECTIVO" & Chr(9) & "" & Chr(9) & txtSaldoFEft.Text
    PintoBlancoResumenIng
    GrdResumen.AddItem "" & Chr(9) & "TOTAL SALDO INICIAL" & Chr(9) & txtSaldoFEft.Text 'txtSaldoF.Text
    PintoBlancoResumenIng
    
    GrdResumen.row = GrdResumen.Rows - 1
    For I = 0 To 3
        GrdResumen.Col = I
        GrdResumen.CellBackColor = &HFF8080
        GrdResumen.CellForeColor = &HFFFFFF    'FUENTE COLOR BLANCO
        GrdResumen.CellFontBold = True
    Next
    
    GrdResumen.AddItem "INGRESOS"
    GrdResumen.Col = 0
    GrdResumen.row = GrdResumen.Rows - 1
    For I = 0 To 3
        GrdResumen.Col = I
        'GrdResumen.CellBackColor = &HC0FFC0
        GrdResumen.CellFontBold = True
    Next
End Sub

Private Sub RESUMEN_CAJA_FINAL()
    GrdResumen.AddItem "SALDO FINAL"
    GrdResumen.Col = 0
    GrdResumen.row = GrdResumen.Rows - 1
    For I = 0 To 3
        GrdResumen.Col = I
        'GrdResumen.CellBackColor = &HC0FFC0
        GrdResumen.CellFontBold = True
    Next
    'GrdResumen.AddItem "    CHEQUES" & Chr(9) & "" & Chr(9) & Valido_Importe(CStr(mTotCheque))
    GrdResumen.AddItem "    CHEQUES" & Chr(9) & "" & Chr(9) & Valido_Importe("0")
    PintoBlancoResumenIng
    GrdResumen.AddItem "    EFECTIVO" & Chr(9) & "" & Chr(9) & Valido_Importe(CStr(mTotMoneda))
    PintoBlancoResumenIng
    GrdResumen.AddItem "" & Chr(9) & "TOTAL SALDO FINAL" & Chr(9) & Valido_Importe(CStr(mTotMoneda)) 'Valido_Importe(CStr(mTotIngResumen - mTotEgreResumen))
    
    GrdResumen.row = GrdResumen.Rows - 1
    For I = 0 To 3
        GrdResumen.Col = I
        GrdResumen.CellBackColor = &HFF8080
        GrdResumen.CellForeColor = &HFFFFFF    'FUENTE COLOR BLANCO
        GrdResumen.CellFontBold = True
    Next
End Sub

Private Sub BUSCO_RECIBOS(FechaCaja As String)
    mTotMoneda1 = 0
    mTotCheque1 = 0
    
    GrdCaja.AddItem "RECIBOS COBRANZA"
    GrdCaja.Col = 0
    GrdCaja.row = GrdCaja.Rows - 1
    For I = 0 To 3
        GrdCaja.Col = I
        'GrdCaja.CellBackColor = &HC0FFFF 'AMARILIO
        GrdCaja.CellFontBold = True
    Next
    GrdResumen.AddItem "RECIBOS COBRANZA"
    GrdResumen.Col = 0
    GrdResumen.row = GrdResumen.Rows - 1
    For I = 0 To 3
        GrdResumen.Col = I
        'GrdResumen.CellBackColor = &HC0FFFF 'AMARILIO
        GrdResumen.CellFontBold = True
    Next
    TOTAL = 0
    
    'BUSCO EN LOS RECIBOS
    sql = "SELECT R.REC_IMPORTE,R.REC_FECHA,REC_NROTXT, SOC_NOMBRE, REC_NUMERO"
    sql = sql & " FROM RECIBO R, SOCIOS S"
    sql = sql & " WHERE R.SOC_CODIGO=S.SOC_CODIGO"
    sql = sql & " AND R.REC_ESTADO=1" 'SOLO RECIBOS DEFINITIVOS
    If FechaCaja <> "" Then
        sql = sql & " AND R.REC_FECHA >" & XDQ(FechaCaja)
    End If
    sql = sql & " AND R.REC_FECHA <=" & XDQ(fecha1.Value)
    sql = sql & " ORDER BY R.REC_NUMERO"
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Do While Rec.EOF = False
            GrdCaja.AddItem "Socio: " & Rec!SOC_NOMBRE & Chr(9) & "Rec Nro: " & Rec!REC_NROTXT & Chr(9) & Valido_Importe(Rec!REC_IMPORTE)
            CambiaColorAFilaDeGrilla GrdCaja, GrdCaja.Rows - 1, vbBlack, , True
            TOTAL = TOTAL + CDbl(Rec!REC_IMPORTE)
            PintoBlancoIng
            
            'BUSCO EN LOS RECIBOS
            sql = "SELECT F.FPG_DESCRI, D.CHE_NUMERO, D.REC_PAGO,D.FPG_CODIGO"
            sql = sql & " FROM RECIBO_PAGOS D, FORMA_PAGO F"
            sql = sql & " WHERE D.REC_NUMERO=" & XN(Rec!REC_NUMERO)
            sql = sql & " AND D.FPG_CODIGO=F.FPG_CODIGO"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                Do While Rec1.EOF = False
                    Select Case Rec1!FPG_CODIGO
                        Case 1
                            GrdCaja.AddItem "" & Chr(9) & Rec1!FPG_DESCRI & Chr(9) & Valido_Importe(Rec1!REC_PAGO)
                            mTotMoneda = mTotMoneda + CDbl(Rec1!REC_PAGO)
                            mTotMoneda1 = mTotMoneda1 + CDbl(Rec1!REC_PAGO)
                            PintoBlancoIng
                        Case 2, 3
                            GrdCaja.AddItem "" & Chr(9) & "CHEQUE T - Che Nro: " & Rec1!CHE_NUMERO & Chr(9) & Valido_Importe(Rec1!REC_PAGO)
                            mTotCheque = mTotCheque + CDbl(Rec1!REC_PAGO)
                            mTotCheque1 = mTotCheque1 + CDbl(Rec1!REC_PAGO)
                            PintoBlancoIng
                    End Select
                    mTotIng = mTotIng + CDbl(Rec1!REC_PAGO)
                    mTotIngResumen = mTotMoneda1
                    Rec1.MoveNext
                Loop
            End If
            Rec1.Close
            Rec.MoveNext
        Loop
    End If
    Rec.Close
    
    GrdCaja.AddItem "" & Chr(9) & "SUB TOTAL" & Chr(9) & Valido_Importe(CStr(TOTAL))
    GrdCaja.row = GrdCaja.Rows - 1
    For I = 0 To 3
        GrdCaja.Col = I
        GrdCaja.CellForeColor = &H0& 'FUENTE COLOR BLANCO
        'GrdCaja.CellBackColor = &HE0E0E0 'GRIS OSCURO
        GrdCaja.CellFontBold = True
    Next
    
    'GrdResumen.AddItem "    CHEQUES" & Chr(9) & "" & Chr(9) & Valido_Importe(CStr(mTotCheque1))
    GrdResumen.AddItem "    CHEQUES" & Chr(9) & "" & Chr(9) & Valido_Importe("0")
    PintoBlancoResumenIng
    GrdResumen.AddItem "    EFECTIVO" & Chr(9) & "" & Chr(9) & Valido_Importe(CStr(mTotMoneda1))
    PintoBlancoResumenIng
End Sub

Private Sub PintoBlancoResumenIng()
    GrdResumen.Col = 0
    GrdResumen.row = GrdResumen.Rows - 1
    For I = 0 To 0
        GrdResumen.Col = I
        GrdResumen.CellBackColor = &H80000005 '&HC0FFFF 'AMARILIO
        'GrdCaja.CellFontBold = True
    Next
    GrdResumen.Col = 2
    GrdResumen.CellBackColor = &HC0FFFF 'AMARILIO
End Sub

Private Sub PintoBlancoResumenEgr()
    GrdResumen.Col = 0
    GrdResumen.row = GrdResumen.Rows - 1
    For I = 0 To 0
        GrdResumen.Col = I
        GrdResumen.CellBackColor = &H80000005 '&HC0FFFF 'AMARILIO
        'GrdCaja.CellFontBold = True
    Next
    GrdResumen.Col = 3
    GrdResumen.CellBackColor = &HC0FFFF 'AMARILIO
End Sub

Private Sub PintoBlancoIng()
    GrdCaja.Col = 0
    GrdCaja.row = GrdCaja.Rows - 1
    For I = 0 To 1
        GrdCaja.Col = I
        GrdCaja.CellBackColor = &H80000005 '&HC0FFFF 'AMARILIO
        'GrdCaja.CellFontBold = True
    Next
    GrdCaja.Col = 2
    GrdCaja.CellBackColor = &HC0FFFF 'AMARILIO
End Sub

Private Sub PintoBlancoEgr()
    GrdCaja.Col = 0
    GrdCaja.row = GrdCaja.Rows - 1
    For I = 0 To 1
        GrdCaja.Col = I
        GrdCaja.CellBackColor = &H80000005 '&HC0FFFF 'AMARILIO
        'GrdCaja.CellFontBold = True
    Next
    GrdCaja.Col = 3
    GrdCaja.CellBackColor = &HC0FFFF 'AMARILIO
End Sub

Private Sub BUSCO_INGRESO_CAJA(FechaCaja As String)
    mTotMoneda1 = 0
    mTotCheque1 = 0
    
    GrdCaja.AddItem "CAJA INGRESO"
    GrdCaja.Col = 0
    GrdCaja.row = GrdCaja.Rows - 1
    For I = 0 To 3
        GrdCaja.Col = I
        'GrdCaja.CellBackColor = &HC0FFFF 'AMARILIO
        GrdCaja.CellFontBold = True
    Next
    GrdResumen.AddItem "CAJA INGRESO"
    GrdResumen.Col = 0
    GrdResumen.row = GrdResumen.Rows - 1
    For I = 0 To 3
        GrdResumen.Col = I
        'GrdResumen.CellBackColor = &HC0FFFF 'AMARILIO
        GrdResumen.CellFontBold = True
    Next
    
    TOTAL = 0
    
    'BUSCO EN CAJA INGRESO
    sql = "SELECT C.CIGR_IMPORTE,C.CIGR_FECHA,C.CIGR_NUMERO, T.TIG_DESCRI"
    sql = sql & " FROM CAJA_INGRESO C, TIPO_INGRESO T"
    sql = sql & " WHERE T.TIG_CODIGO=C.TIG_CODIGO"
    sql = sql & " AND C.EST_CODIGO=3" 'SOLO LO DEFINITIVO
    If FechaCaja <> "" Then
        sql = sql & " AND C.CIGR_FECHA >" & XDQ(FechaCaja)
    End If
    sql = sql & " AND C.CIGR_FECHA <=" & XDQ(fecha1.Value)
    sql = sql & " ORDER BY C.CIGR_NUMERO"
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Do While Rec.EOF = False
            GrdCaja.AddItem Rec!TIG_DESCRI & Chr(9) & "Nro: " & Format(Rec!CIGR_NUMERO, "00000000") & Chr(9) & Valido_Importe(Rec!CIGR_IMPORTE)
            CambiaColorAFilaDeGrilla GrdCaja, GrdCaja.Rows - 1, vbBlack, , True
            TOTAL = TOTAL + CDbl(Rec!CIGR_IMPORTE)
            PintoBlancoIng
            
            'BUSCO EN LOS INGRESOS A CAJA
            sql = "SELECT F.FPG_DESCRI, D.CHE_NUMERO, D.CIGR_IMPORTE,D.FPG_CODIGO"
            sql = sql & " FROM DETALLE_CAJA_INGRESO D, FORMA_PAGO F"
            sql = sql & " WHERE D.CIGR_NUMERO=" & XN(Rec!CIGR_NUMERO)
            sql = sql & " AND D.FPG_CODIGO=F.FPG_CODIGO"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                Do While Rec1.EOF = False
                    Select Case Rec1!FPG_CODIGO
                        Case 1
                            GrdCaja.AddItem "" & Chr(9) & Rec1!FPG_DESCRI & Chr(9) & Valido_Importe(Rec1!CIGR_IMPORTE)
                            mTotMoneda = mTotMoneda + CDbl(Rec1!CIGR_IMPORTE)
                            mTotMoneda1 = mTotMoneda1 + CDbl(Rec1!CIGR_IMPORTE)
                            PintoBlancoIng
                            
                        Case 2, 3
                            GrdCaja.AddItem "" & Chr(9) & "CHEQUE T - Che Nro: " & Rec1!CHE_NUMERO & Chr(9) & Valido_Importe(Rec1!CIGR_IMPORTE)
                            mTotCheque = mTotCheque + CDbl(Rec1!CIGR_IMPORTE)
                            mTotCheque1 = mTotCheque1 + CDbl(Rec1!CIGR_IMPORTE)
                            PintoBlancoIng
                    End Select
                    mTotIng = mTotIng + CDbl(Rec1!CIGR_IMPORTE)
                    Rec1.MoveNext
                Loop
            End If
            Rec1.Close
            Rec.MoveNext
        Loop
        mTotIngResumen = mTotIngResumen + mTotMoneda1
    End If
    Rec.Close
    
    GrdCaja.AddItem "" & Chr(9) & "SUB TOTAL" & Chr(9) & Valido_Importe(CStr(TOTAL))
    GrdCaja.row = GrdCaja.Rows - 1
    For I = 0 To 3
        GrdCaja.Col = I
        GrdCaja.CellForeColor = &H0& 'FUENTE COLOR BLANCO
        'GrdCaja.CellBackColor = &HE0E0E0 'GRIS OSCURO
        GrdCaja.CellFontBold = True
    Next
    
    'GrdResumen.AddItem "    CHEQUES" & Chr(9) & "" & Chr(9) & Valido_Importe(CStr(mTotCheque1))
    GrdResumen.AddItem "    CHEQUES" & Chr(9) & "" & Chr(9) & Valido_Importe("0")
    PintoBlancoResumenIng
    GrdResumen.AddItem "    EFECTIVO" & Chr(9) & "" & Chr(9) & Valido_Importe(CStr(mTotMoneda1))
    PintoBlancoResumenIng
    GrdResumen.AddItem "" & Chr(9) & "TOTAL INGRESOS" & Chr(9) & Valido_Importe(CStr(mTotIng)) 'Valido_Importe(CStr(mTotIngResumen))
    
    GrdResumen.row = GrdResumen.Rows - 1
    For I = 0 To 3
        GrdResumen.Col = I
        GrdResumen.CellBackColor = &HFF8080
        GrdResumen.CellForeColor = &HFFFFFF    'FUENTE COLOR BLANCO
        GrdResumen.CellFontBold = True
    Next
End Sub

Private Sub BUSCO_EGRESO_CAJA(FechaCaja As String)
    mTotCheque1 = 0
    mTotMoneda1 = 0
    GrdCaja.AddItem "CAJA EGRESO"
    GrdCaja.Col = 0
    GrdCaja.row = GrdCaja.Rows - 1
    For I = 0 To 3
        GrdCaja.Col = I
        'GrdCaja.CellBackColor = &HC0FFFF 'AMARILIO
        GrdCaja.CellFontBold = True
    Next
    GrdResumen.AddItem "CAJA EGRESO"
    GrdResumen.Col = 0
    GrdResumen.row = GrdResumen.Rows - 1
    For I = 0 To 3
        GrdResumen.Col = I
        'GrdResumen.CellBackColor = &HC0FFFF 'AMARILIO
        GrdResumen.CellFontBold = True
    Next
    
    TOTAL = 0
    
    'BUSCO EN CAJA EGRESO
    sql = "SELECT C.CEGR_IMPORTE,C.CEGR_FECHA,C.CEGR_NUMERO, T.TEG_DESCRI"
    sql = sql & " FROM CAJA_EGRESO C, TIPO_EGRESO T"
    sql = sql & " WHERE T.TEG_CODIGO=C.TEG_CODIGO"
    sql = sql & " AND C.EST_CODIGO=3" 'SOLO LO DEFINITIVO
    If FechaCaja <> "" Then
        sql = sql & " AND C.CEGR_FECHA >" & XDQ(FechaCaja)
    End If
    sql = sql & " AND C.CEGR_FECHA <=" & XDQ(fecha1.Value)
    sql = sql & " ORDER BY C.CEGR_NUMERO"
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Do While Rec.EOF = False
            GrdCaja.AddItem Rec!TEG_DESCRI & Chr(9) & "Nro: " & Format(Rec!CEGR_NUMERO, "00000000") & Chr(9) & "" & Chr(9) & Valido_Importe(Rec!CEGR_IMPORTE)
            CambiaColorAFilaDeGrilla GrdCaja, GrdCaja.Rows - 1, vbBlack, , True
            TOTAL = TOTAL + CDbl(Rec!CEGR_IMPORTE)
            PintoBlancoEgr
            
            'BUSCO EN LOS EGRESOS A CAJA
            sql = "SELECT F.FPG_DESCRI, D.CHE_NUMERO, D.CEGR_IMPORTE,D.FPG_CODIGO"
            sql = sql & " FROM DETALLE_CAJA_EGRESO D, FORMA_PAGO F"
            sql = sql & " WHERE D.CEGR_NUMERO=" & XN(Rec!CEGR_NUMERO)
            sql = sql & " AND D.FPG_CODIGO=F.FPG_CODIGO"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                Do While Rec1.EOF = False
                    Select Case Rec1!FPG_CODIGO
                        Case 1
                            GrdCaja.AddItem "" & Chr(9) & Rec1!FPG_DESCRI & Chr(9) & "" & Chr(9) & Valido_Importe(Rec1!CEGR_IMPORTE)
                            mTotMoneda = mTotMoneda - CDbl(Rec1!CEGR_IMPORTE)
                            mTotMoneda1 = mTotMoneda1 + CDbl(Rec1!CEGR_IMPORTE)
                            PintoBlancoEgr
                        Case 2
                            GrdCaja.AddItem "" & Chr(9) & "CHEQUE T - Che Nro: " & Rec1!CHE_NUMERO & Chr(9) & "" & Chr(9) & Valido_Importe(Rec1!CEGR_IMPORTE)
                            mTotCheque = mTotCheque - CDbl(Rec1!CEGR_IMPORTE)
                            mTotCheque1 = mTotCheque1 + CDbl(Rec1!CEGR_IMPORTE)
                            PintoBlancoEgr
                        Case 3
                            GrdCaja.AddItem "" & Chr(9) & "CHEQUE P - Che Nro: " & Rec1!CHE_NUMERO & Chr(9) & "" & Chr(9) & Valido_Importe(Rec1!CEGR_IMPORTE)
                            'mTotCheque = mTotCheque - CDbl(Rec1!CEGR_IMPORTE)
                            PintoBlancoEgr
                    End Select
                    mTotEgre = mTotEgre + CDbl(Rec1!CEGR_IMPORTE)
                    Rec1.MoveNext
                Loop
            End If
            Rec1.Close
            Rec.MoveNext
        Loop
        mTotEgreResumen = mTotMoneda1
    End If
    Rec.Close
    
    GrdCaja.AddItem "" & Chr(9) & "SUB TOTAL" & Chr(9) & "" & Chr(9) & Valido_Importe(CStr(TOTAL))
    GrdCaja.row = GrdCaja.Rows - 1
    For I = 0 To 3
        GrdCaja.Col = I
        GrdCaja.CellForeColor = &H0& 'FUENTE COLOR BLANCO
        GrdCaja.CellBackColor = &HE0E0E0 'GRIS OSCURO
        GrdCaja.CellFontBold = True
    Next
    
    'GrdResumen.AddItem "    CHEQUES" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Valido_Importe(CStr(mTotCheque1))
    GrdResumen.AddItem "    CHEQUES" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Valido_Importe("0")
    PintoBlancoResumenEgr
    GrdResumen.AddItem "    EFECTIVO" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Valido_Importe(CStr(mTotMoneda1))
    PintoBlancoResumenEgr
    GrdResumen.AddItem "" & Chr(9) & "TOTAL EGRESOS" & Chr(9) & "" & Chr(9) & Valido_Importe(CStr(mTotEgreResumen))
    
    GrdResumen.row = GrdResumen.Rows - 1
    For I = 0 To 3
        GrdResumen.Col = I
        GrdResumen.CellBackColor = &HFF8080
        GrdResumen.CellForeColor = &HFFFFFF    'FUENTE COLOR BLANCO
        GrdResumen.CellFontBold = True
    Next
End Sub

Private Sub cboDestino_Click()
    picSalida(0).Visible = False
    picSalida(1).Visible = False
    picSalida(2).Visible = False
    picSalida(cboDestino.ListIndex).Visible = True
End Sub

Private Sub CBVolver_Click()
    cmdImprimir.Enabled = False
    CBCerrar.Enabled = False
    GrdCaja.Rows = 1
    GrdResumen.Rows = 1
    BuscarUltimoCierre
End Sub

Private Sub cmdImprimir_Click()
    If ChkCaja.Value = Checked Then
        ReporteCaja
    End If
    
    If ChkResumen.Value = Checked Then
        ReporteCajaResumen
    End If
End Sub

Private Sub ReporteCaja()
    sql = "DELETE FROM TMP_CAJA"
    DBConn.Execute sql
    
    For I = 1 To GrdCaja.Rows - 1
        sql = "INSERT INTO TMP_CAJA (CONCEPTO,"
        sql = sql & " COMPROBANTE, INGRESO, EGRESO, ORDEN )"
        sql = sql & " VALUES ("
        sql = sql & XS(GrdCaja.TextMatrix(I, 0)) & ","
        sql = sql & XS(GrdCaja.TextMatrix(I, 1)) & ","
        sql = sql & XN(GrdCaja.TextMatrix(I, 2)) & ","
        sql = sql & XN(GrdCaja.TextMatrix(I, 3)) & "," & I & ")"
        DBConn.Execute sql
    Next
    sql = "DELETE FROM TMP_RESUMEN_CUENTA_BANCO"
    DBConn.Execute sql
    
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    
    Rep.Formulas(0) = "TITULO='LISTADO DE CAJA'"
    Rep.Formulas(1) = "FECHA='FECHA CAJA:  " & fecha1.Value & "'"
    Select Case cboDestino.ListIndex
        Case 0
            Rep.Destination = crptToWindow
        Case 1
            Rep.Destination = crptToPrinter
        Case 2
            Rep.Destination = crptToFile
    End Select
    
    Rep.WindowTitle = "Listado de Caja"
    Rep.ReportFileName = DRIVE & DirReport & "Caja.rpt"
    Rep.Action = 1
End Sub

Private Sub ReporteCajaResumen()
    sql = "DELETE FROM TMP_CAJA"
    DBConn.Execute sql
    
    For I = 1 To GrdResumen.Rows - 1
        sql = "INSERT INTO TMP_CAJA (CONCEPTO,"
        sql = sql & " COMPROBANTE, INGRESO, EGRESO, ORDEN )"
        sql = sql & " VALUES ("
        sql = sql & XS(GrdResumen.TextMatrix(I, 0)) & ","
        sql = sql & XS(GrdResumen.TextMatrix(I, 1)) & ","
        sql = sql & XN(GrdResumen.TextMatrix(I, 2)) & ","
        sql = sql & XN(GrdResumen.TextMatrix(I, 3)) & "," & I & ")"
        DBConn.Execute sql
    Next
    sql = "DELETE FROM TMP_RESUMEN_CUENTA_BANCO"
    DBConn.Execute sql
    
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    
    Rep.Formulas(0) = "TITULO='LISTADO DE CAJA - RESUMEN'"
    Rep.Formulas(1) = "FECHA='FECHA CAJA:  " & fecha1.Value & "'"
    Select Case cboDestino.ListIndex
        Case 0
            Rep.Destination = crptToWindow
        Case 1
            Rep.Destination = crptToPrinter
        Case 2
            Rep.Destination = crptToFile
    End Select
    
    Rep.WindowTitle = "Listado de Caja Resumen"
    Rep.ReportFileName = DRIVE & DirReport & "Caja.rpt"
    Rep.Action = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
            MySendKeys Chr(9)
            KeyAscii = 0
    End If
End Sub

Private Sub BuscarUltimoCierre()
    sql = "SELECT *"
    sql = sql & " FROM CAJA"
    sql = sql & " WHERE"
    sql = sql & " CAJA_FECHA= (SELECT MAX(CAJA_FECHA) AS FECHA FROM CAJA)"
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
      FechaBusqueda = Rec!CAJA_FECHA
      txtSaldoI.Text = Valido_Importe(Chk0(Rec!CAJA_SALDOI))
      txtSaldoF.Text = Valido_Importe(Chk0(Rec!CAJA_SALDOF))
      txtSaldoIEft.Text = Valido_Importe(Chk0(Rec!CAJA_SALDO_PESOSI))
      txtSaldoFEft.Text = Valido_Importe(Chk0(Rec!CAJA_SALDO_PESOSF))
      txtSaldoIChe.Text = Valido_Importe(Chk0(Rec!CAJA_SALDO_CHEQUESI))
      txtSaldoFChe.Text = Valido_Importe(Chk0(Rec!CAJA_SALDO_CHEQUESF))
      fecha1.Value = CDate(Rec!CAJA_FECHA) + 1
      FrameUltimoCierre.Caption = "Datos último cierre de Caja Registrado (" & FechaBusqueda & ")"
    End If
    Rec.Close
End Sub

Private Sub Form_Load()
    Set Rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    fecha1.Value = Date
    cboDestino.ListIndex = 0
    SSTabCaja.Tab = 0
    cmdImprimir.Enabled = False
    CBCerrar.Enabled = False
    
    Centrar_pantalla Me
    
    GrdCaja.FormatString = "<Concepto|<Comprobante|>Ingreso|>Egreso"
    GrdCaja.ColWidth(0) = 4200
    GrdCaja.ColWidth(1) = 3000
    GrdCaja.ColWidth(2) = 1000
    GrdCaja.ColWidth(3) = 1000
    GrdCaja.Rows = 1
    GrdCaja.HighLight = flexHighlightNever
    GrdCaja.BorderStyle = flexBorderNone
    GrdCaja.row = 0
    For I = 0 To GrdCaja.Cols - 1
        GrdCaja.Col = I
        GrdCaja.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdCaja.CellBackColor = &H808080    'GRIS OSCURO
        GrdCaja.CellFontBold = True
    Next
    
    GrdResumen.FormatString = "<Concepto|<Comprobante|>Ingreso|>Egreso"
    GrdResumen.ColWidth(0) = 4200
    GrdResumen.ColWidth(1) = 3000
    GrdResumen.ColWidth(2) = 1000
    GrdResumen.ColWidth(3) = 1000
    GrdResumen.Rows = 1
    GrdResumen.HighLight = flexHighlightNever
    GrdResumen.BorderStyle = flexBorderNone
    GrdResumen.row = 0
    For I = 0 To GrdResumen.Cols - 1
        GrdResumen.Col = I
        GrdResumen.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdResumen.CellBackColor = &H808080    'GRIS OSCURO
        GrdResumen.CellFontBold = True
    Next
    
    BuscarUltimoCierre
End Sub
