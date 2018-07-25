VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmResumenCuentaBanco 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen de Cuenta - Banco"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
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
      Height          =   735
      Left            =   60
      TabIndex        =   23
      Top             =   2835
      Width           =   2175
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
         Picture         =   "frmResumenCuentaBanco.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   27
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
         Picture         =   "frmResumenCuentaBanco.frx":0102
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   26
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
         Index           =   0
         Left            =   135
         Picture         =   "frmResumenCuentaBanco.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   25
         Top             =   315
         Width           =   240
      End
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "frmResumenCuentaBanco.frx":0306
         Left            =   450
         List            =   "frmResumenCuentaBanco.frx":0313
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   270
         Width           =   1635
      End
   End
   Begin VB.Frame FrameCierre 
      Caption         =   "Cierre Mensual"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   2985
      TabIndex        =   15
      Top             =   1425
      Width           =   2940
      Begin MSComCtl2.DTPicker Fecha 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   645
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   52625409
         CurrentDate     =   42925
      End
      Begin VB.Label lblPeriodo1 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   17
         Top             =   645
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   16
         Top             =   450
         Width           =   600
      End
   End
   Begin VB.Frame FrameLibro 
      Caption         =   "Libro Banco"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   45
      TabIndex        =   18
      Top             =   1425
      Width           =   2940
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   52625409
         CurrentDate     =   42925
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   990
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   52625409
         CurrentDate     =   42925
      End
      Begin VB.Label lblPeriodoH 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1350
         TabIndex        =   22
         Top             =   990
         Width           =   1530
      End
      Begin VB.Label lblPeriodoD 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1350
         TabIndex        =   21
         Top             =   465
         Width           =   1530
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Periodo Hasta:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   780
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Periodo Desde:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   255
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   4170
      TabIndex        =   8
      Top             =   3120
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   5055
      TabIndex        =   9
      Top             =   3120
      Width           =   870
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   450
      Left            =   3300
      TabIndex        =   7
      Top             =   3120
      Width           =   855
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   2850
      Top             =   2955
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2370
      Top             =   2955
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frameResumen 
      Caption         =   "Banco y Cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1425
      Left            =   45
      TabIndex        =   10
      Top             =   15
      Width           =   5880
      Begin VB.OptionButton optCierre 
         Caption         =   "Cierre Mensual"
         Height          =   195
         Left            =   2460
         TabIndex        =   3
         Top             =   1110
         Width           =   1575
      End
      Begin VB.OptionButton optLibroBanco 
         Caption         =   "Libro Banco"
         Height          =   195
         Left            =   930
         TabIndex        =   2
         Top             =   1095
         Width           =   1230
      End
      Begin VB.ComboBox CboCuentas 
         Height          =   315
         Left            =   945
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   660
         Width           =   1965
      End
      Begin VB.ComboBox CboBancoBoleta 
         Height          =   315
         ItemData        =   "frmResumenCuentaBanco.frx":0335
         Left            =   945
         List            =   "frmResumenCuentaBanco.frx":0337
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   330
         Width           =   3810
      End
      Begin VB.TextBox TxtBanCodInt 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4290
         TabIndex        =   11
         Top             =   675
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta:"
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   14
         Top             =   690
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Index           =   5
         Left            =   405
         TabIndex        =   13
         Top             =   390
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Index           =   7
         Left            =   3705
         TabIndex        =   12
         Top             =   705
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmResumenCuentaBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AplicoImpuesto As Boolean
Dim ValorImpuesto As Double
Dim Saldo As Double
Dim I As Integer
Dim FechaSaldo As String
Dim UltimoSaldo As Double
Dim RegistroSaldo As Boolean
Dim MES As String
Dim ano As String
Dim mImpCheque As Double
Dim mIvaBase As Double
Dim mOtrosImp As Double
Dim mGravamen As Double

Private Sub CboBancoBoleta_LostFocus()
     If CboBancoBoleta.ListIndex <> -1 Then
        Me.TxtBanCodInt.Text = CStr(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
     End If
End Sub

Private Sub CboCuentas_GotFocus()
    If Trim(CboBancoBoleta.Text) <> "" Then
        CboCuentas.Clear
        Call CargoCtaBancaria(CStr(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)))
        CboCuentas.ListIndex = 0
    End If
End Sub

Private Sub CargoCtaBancaria(Banco As String)
    Set Rec = New ADODB.Recordset
    sql = "SELECT CTA_NROCTA FROM CTA_BANCARIA"
    sql = sql & " WHERE BAN_CODINT=" & XN(Banco)
    sql = sql & " AND CTA_FECCIE IS NULL"
    sql = sql & " ORDER BY CTA_NROCTA DESC"
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
     Do While Rec.EOF = False
         CboCuentas.AddItem Trim(Rec!CTA_NROCTA)
         Rec.MoveNext
     Loop
    End If
    Rec.Close
End Sub

Private Sub CboCuentas_LostFocus()
    If CboBancoBoleta.ListIndex <> -1 And CboCuentas.ListIndex <> -1 Then
'        sql = " SELECT RCB_FECHA,RCB_SALDO"
'        sql = sql & " FROM RESUMEN_CUENTA_BANCO"
'        sql = sql & " WHERE RCB_FECHA="
'        sql = sql & " (SELECT MAX(RCB_FECHA) AS FECHA"
'        sql = sql & " FROM RESUMEN_CUENTA_BANCO"
'        sql = sql & " WHERE BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
'        sql = sql & " AND CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex)) & ")"
'        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If rec.EOF = False Then
'            lblUltimoSaldo.Caption = "ÚLTIMO SALDO: " & rec!RCB_FECHA & " - (" & Valido_Importe(rec!RCB_SALDO) & ")"
'            cmdVerResumen.Enabled = True
'        Else
'            rec.Close
'            sql = "SELECT CTA_FECAPE,CTA_SALINI"
'            sql = sql & " FROM CTA_BANCARIA"
'            sql = sql & " WHERE BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
'            sql = sql & " AND CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
'            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'            If Not IsNull(rec!CTA_SALINI) Then
'                lblUltimoSaldo.Caption = "ÚLTIMO SALDO: " & rec!CTA_FECAPE & " - (" & Valido_Importe(rec!CTA_SALINI) & ")"
'                cmdVerResumen.Enabled = True
'            Else
'                MsgBox "NO se encuentra registrado el saldo de la cuenta", vbCritical, TIT_MSGBOX
'                cmdVerResumen.Enabled = False
'            End If
'        End If
        If Rec.State = 1 Then Rec.Close
    End If
End Sub

Private Sub cboDestino_Click()
    picSalida(0).Visible = False
    picSalida(1).Visible = False
    picSalida(2).Visible = False
    picSalida(cboDestino.ListIndex).Visible = True
End Sub

Private Sub cmdListar_Click()
    If CboCuentas.ListIndex = -1 Then
        MsgBox "Debe elegir una Cuenta Bancaria", vbCritical, TIT_MSGBOX
        CboBancoBoleta.SetFocus
        Exit Sub
    End If
    If optLibroBanco.Value = True And FechaDesde.Value = "" Then
        MsgBox "Debe ingresar el Periodo Desde", vbCritical, TIT_MSGBOX
        FechaDesde.SetFocus
        Exit Sub
    End If
    If optLibroBanco.Value = True And FechaHasta.Value = "" Then
        MsgBox "Debe ingresar la Fecha Hasta", vbCritical, TIT_MSGBOX
        FechaHasta.SetFocus
        Exit Sub
    End If
    If optCierre.Value = True And Fecha.Value = "" Then
        MsgBox "Debe ingresar el Periodo de Cierre", vbCritical, TIT_MSGBOX
        Fecha.SetFocus
        Exit Sub
    End If
    
    On Error GoTo HayErorr
    DBConn.BeginTrans
    Screen.MousePointer = vbHourglass
    
    'BUSCO LOS DATOS DEL RESUMEN
    CALCULA_RESUMEN
    
    DBConn.CommitTrans
    DBConn.Execute "DELETE FROM TMP_ORDEN_PAGO"
    Screen.MousePointer = vbNormal
    
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SISPECARI"
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
    Rep.Formulas(3) = ""
    Rep.Formulas(4) = ""
    Rep.Formulas(5) = ""
    Rep.Formulas(6) = ""
    
    Rep.SelectionFormula = ""
    
    
    Rep.Formulas(0) = "SALDO='" & Valido_Importe(CStr(Saldo)) & "'"
    If optCierre.Value = True Then
        Rep.Formulas(1) = "PERIODO='" & lblPeriodo1.Caption & "'"
        Rep.Formulas(2) = "TITULO='Resumen de Cuenta'"
    Else
        Rep.Formulas(1) = "PERIODO='" & FechaDesde.Value & " al " & FechaHasta.Value & "'"
        Rep.Formulas(2) = "TITULO='Libro Banco'"
    End If
    Rep.SelectionFormula = "{TMP_RESUMEN_CUENTA_BANCO.BANCO}=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
    Rep.SelectionFormula = Rep.SelectionFormula & " AND {TMP_RESUMEN_CUENTA_BANCO.CUENTA}=" & XS(CboCuentas.List(CboCuentas.ListIndex))
    'Rep.SelectionFormula = Rep.SelectionFormula & " AND {TMP_RESUMEN_CUENTA_BANCO.PERIODO}=" & XS(Format(Fecha.Value, "mm/yyyy"))
    
    'LE PONGO LOS GASTOS BANCARIOS
    Rep.Formulas(3) = "IMP_CHEQUE='" & Valido_Importe(CStr(mImpCheque)) & "'"
    Rep.Formulas(4) = "IVA_BASE='" & Valido_Importe(CStr(mIvaBase)) & "'"
    Rep.Formulas(5) = "OTRO_IMP='" & Valido_Importe(CStr(mOtrosImp)) & "'"
    Rep.Formulas(6) = "GRAVAMEN='" & Valido_Importe(CStr(mGravamen)) & "'"
    
    Rep.WindowTitle = "Resumen Cuenta - Banco..."
    Rep.ReportFileName = DRIVE & DirReport & "ResumenCuentaBanco.rpt"
    
    Select Case cboDestino.ListIndex
        Case 0
            Rep.Destination = crptToWindow
        Case 1
            Rep.Destination = crptToPrinter
        Case 2
            Rep.Destination = crptToFile
    End Select
    
    Rep.Action = 1
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
    Rep.Formulas(3) = ""
    Rep.Formulas(4) = ""
    Rep.Formulas(5) = ""
    Exit Sub
    
HayErorr:
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
    If Rec.State = 1 Then Rec.Close
    If Rec2.State = 1 Then Rec2.Close
    MsgBox Err.Description, vbCritical, TIT_MSGBOX

End Sub

Private Sub CmdNuevo_Click()
    CboBancoBoleta.ListIndex = 0
    CboCuentas.Clear
    TxtBanCodInt.Text = ""
    TxtBanCodInt.Text = ""
    Fecha.Value = ""
    FechaDesde.Value = ""
    FechaHasta.Value = ""
    optLibroBanco.Value = True
    RegistroSaldo = False
    CboBancoBoleta.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Set frmResumenCuentaBanco = Nothing
    Unload Me
End Sub

Private Sub CALCULA_RESUMEN()
    I = 0
    Saldo = 0
    mImpCheque = 0
    mIvaBase = 0
    mOtrosImp = 0
    mGravamen = 0
    
    sql = "DELETE FROM TMP_RESUMEN_CUENTA_BANCO"
    DBConn.Execute sql
   
    'BUSCO EL ULTIMO SALDO
    BUSCO_ULTIMO_SALDO
    'BUSCO LAS BOLETAS DE DEPOSITO
    BUSCO_BOLETAS_DEPOSITO
    'BUSCO LOS GASTOS BANCARIOS
    BUSCO_GASTOS_BANCARIOS
    'BUSCO DEBITOS CREDITOS
    BUSCO_DEBITOS_CREDITOS
    'BUSCO LOS CHEQUES LIBRADOS
    BUSCO_CHEQUES_LIBRADOS
    'CALCULO SALDO DEL RESUMEN
    CALCULO_SALDO
    If optCierre.Value = True And RegistroSaldo = True Then
        If MsgBox("Esta acción Registra un nuevo Resumen de Cuenta Bancaria," & Chr(13) & _
                  "Confirma Nuevo Resumen?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            'REGISTRO EL SALDO DEL RESUMEN
            REGISTRO_SALDO_RESUMEN
        End If
    End If
End Sub

'Private Sub BUSCO_SALDO_ANTERIOR_MITAD(FechaD As String, FechaH As String, SaldoI As Double)
'    'Set Rec2 = New ADODB.Recordset
'    Dim SaldoMitad As Double
'    SaldoMitad = SaldoI
'    'DEPOSITOS
'    sql = "SELECT BOL_FECHA,BOL_NUMERO,BOL_TOTAL"
'    sql = sql & " FROM BOL_DEPOSITO"
'    sql = sql & " WHERE BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
'    sql = sql & " AND CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
'    sql = sql & " AND EBO_CODIGO<> 2" 'BOLETAS NO ANULADAS
'    sql = sql & " AND BOL_FECHA>=" & XDQ(FechaD)
'    sql = sql & " AND BOL_FECHA<" & XDQ(FechaH)
'    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
'    If Rec2.EOF = False Then
'        Do While Rec2.EOF = False
'            'INSERTO DEPOSITO
'            SaldoMitad = SaldoMitad + CDbl(Rec2!BOL_TOTAL)
'            If AplicoImpuesto = True Then
'                SaldoMitad = SaldoMitad - (CDbl(Rec2!BOL_TOTAL) * (ValorImpuesto))
'            End If
'            Rec2.MoveNext
'        Loop
'    End If
'    Rec2.Close
'
'    'GASTOS BANCARIOS
'    sql = "SELECT GB.GBA_NUMERO,GB.GBA_FECHA,GB.GBA_IMPORTE,TG.TGB_DESCRI,GB.GBA_IMPUESTO"
'    sql = sql & " FROM GASTOS_BANCARIOS GB,TIPO_GASTO_BANCARIO TG"
'    sql = sql & " WHERE GB.BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
'    sql = sql & " AND GB.CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
'    sql = sql & " AND GB.TGB_CODIGO=TG.TGB_CODIGO"
'    sql = sql & " AND GB.GBA_FECHA>=" & XDQ(FechaD)
'    sql = sql & " AND GB.GBA_FECHA<" & XDQ(FechaH)
'    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
'    If Rec2.EOF = False Then
'        Do While Rec2.EOF = False
'            SaldoMitad = SaldoMitad - CDbl(Rec2!GBA_IMPORTE)
'            If AplicoImpuesto = True And Rec2!GBA_IMPUESTO = "S" Then
'                SaldoMitad = SaldoMitad - (CDbl(Rec2!GBA_IMPORTE) * (ValorImpuesto))
'            End If
'            Rec2.MoveNext
'        Loop
'    End If
'    Rec2.Close
'
'    'DEBITOS CREDITOS BANCARIOA
'    sql = "SELECT GB.DCB_NUMERO,GB.DCB_FECHA,GB.DCB_IMPORTE,TG.TDCB_DESCRI,"
'    sql = sql & " GB.DCB_IMPUESTO, GB.DCB_TIPO"
'    sql = sql & " FROM DEBCRE_BANCARIOS GB, TIPO_DEBCRE_BANCARIO TG"
'    sql = sql & " WHERE GB.BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
'    sql = sql & " AND GB.CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
'    sql = sql & " AND GB.TDCB_CODIGO=TG.TDCB_CODIGO"
'    sql = sql & " AND GB.DCB_FECHA>=" & XDQ(FechaD)
'    sql = sql & " AND GB.DCB_FECHA<" & XDQ(FechaH)
'    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
'    If Rec2.EOF = False Then
'        Do While Rec2.EOF = False
'            'INSERTO GASTOS BANCARIOS
'            If Trim(Rec2!DCB_TIPO) = "D" Then
'                SaldoMitad = SaldoMitad - CDbl(Rec2!DCB_IMPORTE)
'            ElseIf Trim(Rec2!DCB_TIPO) = "C" Then
'                SaldoMitad = SaldoMitad + CDbl(Rec2!DCB_IMPORTE)
'            End If
'            'VERIFICO SI LE APLICO EL IMPUESTO AL CHEQUE
'            If AplicoImpuesto = True And Rec2!DCB_IMPUESTO = "S" Then
'                SaldoMitad = SaldoMitad - (CDbl(Rec2!DCB_IMPORTE) * (ValorImpuesto))
'            End If
'            Rec2.MoveNext
'        Loop
'    End If
'    Rec2.Close
'
'    'CHEQUES LIBRADOS
'    sql = "SELECT CHEP_FECVTO,CHEP_NUMERO,CHEP_IMPORT,CHEP_NOMBRE"
'    sql = sql & " FROM ChequePropioEstadoVigente"
'    sql = sql & " WHERE BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
'    sql = sql & " AND CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
'    sql = sql & " AND ECH_CODIGO IN (7,8)" 'CHEQUES LIBRADOS O RESTITUIDOS
'    sql = sql & " AND CHEP_FECVTO>=" & XDQ(FechaD)
'    sql = sql & " AND CHEP_FECVTO<" & XDQ(FechaH)
'    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
'    If Rec2.EOF = False Then
'        Do While Rec2.EOF = False
'            SaldoMitad = SaldoMitad - CDbl(Rec2!CHEP_IMPORT)
'            'VERIFICO SI LE APLICO EL IMPUESTO AL CHEQUE
'            If AplicoImpuesto = True Then
'                SaldoMitad = SaldoMitad - (CDbl(Rec2!CHEP_IMPORT) * (ValorImpuesto))
'            End If
'            Rec2.MoveNext
'        Loop
'    End If
'    Rec2.Close
'
'    FechaSaldo = Format(FechaH, "mmmm/yyyy")
'    UltimoSaldo = SaldoMitad
'End Sub

Private Function RESUMEN_ANTERIOR() As Double
    sql = "SELECT * FROM RESUMEN_CUENTA_BANCO"
    sql = sql & " WHERE BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
    sql = sql & " AND CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
    sql = sql & " AND MONTH(RCB_FECHA)=" & XN(Month(Fecha.Value))
    sql = sql & " AND YEAR(RCB_FECHA)=" & XN(Year(Fecha.Value))
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = True Then
        RESUMEN_ANTERIOR = False
        RegistroSaldo = True
        Rec.Close
        Exit Function
    End If
    Rec.Close
    
    sql = "SELECT *"
    sql = sql & " FROM RESUMEN_CUENTA_BANCO"
    sql = sql & " WHERE RCB_FECHA="
    sql = sql & " (SELECT MAX(RCB_FECHA) AS FECHA"
    sql = sql & " FROM RESUMEN_CUENTA_BANCO"
    sql = sql & " WHERE BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
    sql = sql & " AND CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
    sql = sql & " AND MONTH(RCB_FECHA)<" & XN(Month(Fecha.Value))
    sql = sql & " AND YEAR(RCB_FECHA)=" & XN(Year(Fecha.Value)) & ")"
    sql = sql & " AND BAN_CODINT = " & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
    sql = sql & " AND CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        FechaSaldo = Format(Rec!RCB_FECHA, "mmmm/yyyy")
        UltimoSaldo = CDbl(Rec!RCB_SALDO)
        RESUMEN_ANTERIOR = True
    Else
        RESUMEN_ANTERIOR = False
    End If
    Rec.Close
End Function

Private Sub BUSCO_ULTIMO_SALDO()
    FechaSaldo = ""
    UltimoSaldo = 0
    RegistroSaldo = False
    
    If optCierre.Value = True Then
        If RESUMEN_ANTERIOR = False Then
            sql = " SELECT B.RCB_FECHA,B.RCB_SALDO"
            sql = sql & " FROM RESUMEN_CUENTA_BANCO B"
            sql = sql & " WHERE B.RCB_FECHA="
            sql = sql & " (SELECT MAX(R.RCB_FECHA) AS FECHA"
            sql = sql & " FROM RESUMEN_CUENTA_BANCO R"
            sql = sql & " WHERE R.BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
            sql = sql & " AND R.CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex)) & ")"
            sql = sql & " AND B.BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
            sql = sql & " AND B.CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
            Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec.EOF = False Then
                FechaSaldo = Format(Rec!RCB_FECHA, "mmmm/yyyy")
                UltimoSaldo = CDbl(Rec!RCB_SALDO)
            Else
                Rec.Close
                sql = "SELECT CTA_FECAPE,CTA_SALINI"
                sql = sql & " FROM CTA_BANCARIA"
                sql = sql & " WHERE BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
                sql = sql & " AND CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
                Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If Rec.EOF = False Then
                    FechaSaldo = Format(Rec!CTA_FECAPE, "mmmm/yyyy")
                    UltimoSaldo = CDbl(Rec!CTA_SALINI)
                End If
            End If
        End If
    End If
    If optLibroBanco.Value = True Then
        Select Case Month(FechaDesde.Value)
        Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
            MES = Month(FechaDesde.Value) - 1
            ano = Year(FechaDesde.Value)
        Case 1
            MES = 12
            ano = Year(FechaDesde.Value) - 1
        End Select
        
        sql = " SELECT B.RCB_FECHA,B.RCB_SALDO"
        sql = sql & " FROM RESUMEN_CUENTA_BANCO B"
        sql = sql & " WHERE RCB_FECHA="
        sql = sql & " (SELECT MAX(R.RCB_FECHA) AS FECHA"
        sql = sql & " FROM RESUMEN_CUENTA_BANCO R"
        sql = sql & " WHERE R.BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
        sql = sql & " AND R.CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
        sql = sql & " AND MONTH(R.RCB_FECHA)=" & MES
        sql = sql & " AND YEAR(R.RCB_FECHA)=" & ano
        sql = sql & ")"
        sql = sql & " AND B.BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
        sql = sql & " AND B.CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
        
        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec.EOF = False Then
            FechaSaldo = Format(Rec!RCB_FECHA, "mmmm/yyyy")
            UltimoSaldo = CDbl(Rec!RCB_SALDO)
        End If
    End If
    
    If Rec.State = 1 Then Rec.Close
    
    I = I + 1
    'INSERTO EN LA TEMPORAL
    sql = "INSERT INTO TMP_RESUMEN_CUENTA_BANCO"
    sql = sql & " (BANCO,CUENTA,PERIODO,FECHA,DESCRIPCION,COMPROBANTE,"
    sql = sql & " DEBITO,CREDITO,SALDO,ORDEN) VALUES ("
    sql = sql & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)) & ","
    sql = sql & XS(CboCuentas.List(CboCuentas.ListIndex)) & ","
    sql = sql & XS(Format(Fecha.Value, "mm/yyyy")) & ","
    sql = sql & XDQ("") & ","
    sql = sql & XS("Saldo Anterior - " & FechaSaldo) & ","
    sql = sql & XS("") & ","
    sql = sql & XN("0") & ","
    sql = sql & XN("0") & ","
    sql = sql & XN(CStr(UltimoSaldo)) & ","
    sql = sql & XN(CStr(I)) & ")"
    DBConn.Execute sql
End Sub

Private Sub BUSCO_BOLETAS_DEPOSITO()
    sql = "SELECT BOL_FECHA,BOL_NUMERO,BOL_TOTAL,BOL_EFECVO"
    sql = sql & " FROM BOL_DEPOSITO"
    sql = sql & " WHERE BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
    sql = sql & " AND CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
    sql = sql & " AND EBO_CODIGO<> 2" 'BOLETAS NO ANULADAS
    If optCierre.Value = True Then
        sql = sql & " AND MONTH(BOL_FECHA)=" & XN(Month(Fecha.Value))
        sql = sql & " AND YEAR(BOL_FECHA)=" & XN(Year(Fecha.Value))
    End If
    If optLibroBanco.Value = True Then
        sql = sql & " AND BOL_FECHA>=" & XDQ("01/" & Month(FechaDesde.Value) & "/" & Year(FechaDesde.Value))
        sql = sql & " AND BOL_FECHA<=" & XDQ(Day(UltimoDiadelMes(FechaHasta.Value)) & "/" & Month(FechaHasta.Value) & "/" & Year(FechaHasta.Value))
    End If
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If Rec.EOF = False Then
        Do While Rec.EOF = False
            'INSERTO DEPOSITO
            I = I + 1
            sql = "INSERT INTO TMP_RESUMEN_CUENTA_BANCO"
            sql = sql & " (BANCO,CUENTA,PERIODO,FECHA,DESCRIPCION,COMPROBANTE,"
            sql = sql & " DEBITO,CREDITO,SALDO,ORDEN) VALUES ("
            sql = sql & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)) & ","
            sql = sql & XS(CboCuentas.List(CboCuentas.ListIndex)) & ","
            sql = sql & XS(Format(Fecha.Value, "mm/yyyy")) & ","
            sql = sql & XDQ(Rec!BOL_FECHA) & ","
            sql = sql & "'DEPOSITO NRO " & Trim(Rec!BOL_NUMERO) & "',"
            sql = sql & XS(Rec!BOL_NUMERO) & ","
            sql = sql & XN("0") & ","
            sql = sql & XN(Rec!BOL_TOTAL) & ","
            sql = sql & XN("0") & ","
            sql = sql & XN(CStr(I)) & ")"
            DBConn.Execute sql
            'VERIFICO SI LE APLICO EL IMPUESTO AL CHEQUE
            If AplicoImpuesto = True Then
                If (CDbl(Rec!BOL_TOTAL) - CDbl(Chk0(Rec!BOL_EFECVO))) <> 0 Then
                    I = I + 1
                    sql = "INSERT INTO TMP_RESUMEN_CUENTA_BANCO"
                    sql = sql & " (BANCO,CUENTA,PERIODO,FECHA,DESCRIPCION,COMPROBANTE,"
                    sql = sql & " DEBITO,CREDITO,SALDO,ORDEN) VALUES ("
                    sql = sql & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)) & ","
                    sql = sql & XS(CboCuentas.List(CboCuentas.ListIndex)) & ","
                    sql = sql & XS(Format(Fecha.Value, "mm/yyyy")) & ","
                    sql = sql & XDQ(Rec!BOL_FECHA) & ","
                    sql = sql & XS("GRAVAMEN LEY 25413 S/CRE") & ","
                    sql = sql & XS(Rec!BOL_NUMERO) & ","
                    sql = sql & XN(CStr((CDbl(Rec!BOL_TOTAL) - CDbl(Chk0(Rec!BOL_EFECVO))) * (ValorImpuesto))) & ","
                    sql = sql & XN("0") & ","
                    sql = sql & XN("0") & ","
                    sql = sql & XN(CStr(I)) & ")"
                    DBConn.Execute sql
                    
                    mImpCheque = mImpCheque + ((CDbl(Rec!BOL_TOTAL) - CDbl(Chk0(Rec!BOL_EFECVO))) * (ValorImpuesto))
                End If
            End If
            Rec.MoveNext
        Loop
    End If
    Rec.Close
End Sub

Private Sub BUSCO_GASTOS_BANCARIOS()
    sql = "SELECT GB.GBA_NUMERO,GB.GBA_FECHA,GB.GBA_IMPORTE,TG.TGB_DESCRI,"
    sql = sql & " GB.GBA_IMPUESTO, GB.TGB_CODIGO"
    sql = sql & " FROM GASTOS_BANCARIOS GB,TIPO_GASTO_BANCARIO TG"
    sql = sql & " WHERE GB.BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
    sql = sql & " AND GB.CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
    sql = sql & " AND GB.TGB_CODIGO=TG.TGB_CODIGO"
    If optCierre.Value = True Then
        sql = sql & " AND MONTH(GB.GBA_FECHA)=" & XN(Month(Fecha.Value))
        sql = sql & " AND YEAR(GB.GBA_FECHA)=" & XN(Year(Fecha.Value))
    End If
    If optLibroBanco.Value = True Then
        sql = sql & " AND GB.GBA_FECHA>=" & XDQ("01/" & Month(FechaDesde.Value) & "/" & Year(FechaDesde.Value))
        sql = sql & " AND GB.GBA_FECHA<=" & XDQ(Day(UltimoDiadelMes(FechaHasta.Value)) & "/" & Month(FechaHasta.Value) & "/" & Year(FechaHasta.Value))
    End If
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If Rec.EOF = False Then
        Do While Rec.EOF = False
            'Saldo = Saldo - CDbl(rec!GBA_IMPORTE)
            I = I + 1
            'INSERTO GASTOS BANCARIOS
            sql = "INSERT INTO TMP_RESUMEN_CUENTA_BANCO"
            sql = sql & " (BANCO,CUENTA,PERIODO,FECHA,DESCRIPCION,COMPROBANTE,"
            sql = sql & " DEBITO,CREDITO,SALDO,ORDEN) VALUES ("
            sql = sql & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)) & ","
            sql = sql & XS(CboCuentas.List(CboCuentas.ListIndex)) & ","
            sql = sql & XS(Format(Fecha.Value, "mm/yyyy")) & ","
            sql = sql & XDQ(Rec!GBA_FECHA) & ","
            sql = sql & XS(Rec!TGB_DESCRI) & ","
            sql = sql & XS(Rec!GBA_NUMERO) & ","
            sql = sql & XN(Rec!GBA_IMPORTE) & ","
            sql = sql & XN("0") & ","
            sql = sql & XN("0") & ","
            sql = sql & XN(CStr(I)) & ")"
            DBConn.Execute sql
            Select Case Rec!TGB_CODIGO
                Case 4 'IVA BASE
                    mIvaBase = mIvaBase + CDbl(Rec!GBA_IMPORTE)
                Case 10 'GRAVAMEN IBCD S/CRED
                    mGravamen = mGravamen + CDbl(Rec!GBA_IMPORTE)
                Case Else
                    mOtrosImp = mOtrosImp + CDbl(Rec!GBA_IMPORTE)
            End Select
                
            'VERIFICO SI LE APLICO EL IMPUESTO AL CHEQUE
            If AplicoImpuesto = True And Rec!GBA_IMPUESTO = "S" Then
                I = I + 1
                sql = "INSERT INTO TMP_RESUMEN_CUENTA_BANCO"
                sql = sql & " (BANCO,CUENTA,PERIODO,FECHA,DESCRIPCION,COMPROBANTE,"
                sql = sql & " DEBITO,CREDITO,SALDO,ORDEN) VALUES ("
                sql = sql & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)) & ","
                sql = sql & XS(CboCuentas.List(CboCuentas.ListIndex)) & ","
                sql = sql & XS(Format(Fecha.Value, "mm/yyyy")) & ","
                sql = sql & XDQ(Rec!GBA_FECHA) & ","
                sql = sql & XS("GRAVAMEN LEY 25413 S/DEB") & ","
                sql = sql & XS(Rec!GBA_NUMERO) & ","
                sql = sql & XN(CStr(CDbl(Rec!GBA_IMPORTE) * (ValorImpuesto))) & ","
                sql = sql & XN("0") & ","
                sql = sql & XN("0") & ","
                sql = sql & XN(CStr(I)) & ")"
                DBConn.Execute sql
                
                mImpCheque = mImpCheque + (CDbl(Rec!GBA_IMPORTE) * (ValorImpuesto))
            End If
            Rec.MoveNext
        Loop
    End If
    Rec.Close
End Sub

Private Sub BUSCO_DEBITOS_CREDITOS()
    sql = "SELECT GB.DCB_NUMERO,GB.DCB_FECHA,GB.DCB_IMPORTE,TG.TDCB_DESCRI,"
    sql = sql & " GB.DCB_IMPUESTO, GB.DCB_TIPO"
    sql = sql & " FROM DEBCRE_BANCARIOS GB, TIPO_DEBCRE_BANCARIO TG"
    sql = sql & " WHERE GB.BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
    sql = sql & " AND GB.CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
    sql = sql & " AND GB.TDCB_CODIGO=TG.TDCB_CODIGO"
    If optCierre.Value = True Then
        sql = sql & " AND MONTH(GB.DCB_FECHA)=" & XN(Month(Fecha.Value))
        sql = sql & " AND YEAR(GB.DCB_FECHA)=" & XN(Year(Fecha.Value))
    End If
    If optLibroBanco.Value = True Then
        sql = sql & " AND GB.DCB_FECHA>=" & XDQ("01/" & Month(FechaDesde.Value) & "/" & Year(FechaDesde.Value))
        sql = sql & " AND GB.DCB_FECHA<=" & XDQ(Day(UltimoDiadelMes(FechaHasta.Value)) & "/" & Month(FechaHasta.Value) & "/" & Year(FechaHasta.Value))
    End If
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Do While Rec.EOF = False
            'Saldo = Saldo - CDbl(rec!DCB_IMPORTE)
            I = I + 1
            'INSERTO GASTOS BANCARIOS
            sql = "INSERT INTO TMP_RESUMEN_CUENTA_BANCO"
            sql = sql & " (BANCO,CUENTA,PERIODO,FECHA,DESCRIPCION,COMPROBANTE,"
            sql = sql & " DEBITO,CREDITO,SALDO,ORDEN) VALUES ("
            sql = sql & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)) & ","
            sql = sql & XS(CboCuentas.List(CboCuentas.ListIndex)) & ","
            sql = sql & XS(Format(Fecha.Value, "mm/yyyy")) & ","
            sql = sql & XDQ(Rec!DCB_FECHA) & ","
            sql = sql & XS(Rec!TDCB_DESCRI) & ","
            sql = sql & XS(Rec!DCB_NUMERO) & ","
            If Trim(Rec!DCB_TIPO) = "D" Then
                sql = sql & XN(Rec!DCB_IMPORTE) & ","
                sql = sql & XN("0") & ","
            ElseIf Trim(Rec!DCB_TIPO) = "C" Then
                sql = sql & XN("0") & ","
                sql = sql & XN(Rec!DCB_IMPORTE) & ","
            End If
            sql = sql & XN("0") & ","
            sql = sql & XN(CStr(I)) & ")"
            DBConn.Execute sql
            'VERIFICO SI LE APLICO EL IMPUESTO AL CHEQUE
            If AplicoImpuesto = True And Rec!DCB_IMPUESTO = "S" Then
                I = I + 1
                sql = "INSERT INTO TMP_RESUMEN_CUENTA_BANCO"
                sql = sql & " (BANCO,CUENTA,PERIODO,FECHA,DESCRIPCION,COMPROBANTE,"
                sql = sql & " DEBITO,CREDITO,SALDO,ORDEN) VALUES ("
                sql = sql & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)) & ","
                sql = sql & XS(CboCuentas.List(CboCuentas.ListIndex)) & ","
                sql = sql & XS(Format(Fecha.Value, "mm/yyyy")) & ","
                sql = sql & XDQ(Rec!DCB_FECHA) & ","
                sql = sql & XS("GRAVAMEN LEY 25413 S/DEB") & ","
                sql = sql & XS(Rec!DCB_NUMERO) & ","
                sql = sql & XN(CStr(CDbl(Rec!DCB_IMPORTE) * (ValorImpuesto))) & ","
                sql = sql & XN("0") & ","
                sql = sql & XN("0") & ","
                sql = sql & XN(CStr(I)) & ")"
                DBConn.Execute sql
                
                mImpCheque = mImpCheque + (CDbl(Rec!DCB_IMPORTE) * (ValorImpuesto))
            End If
            Rec.MoveNext
        Loop
    End If
    Rec.Close
End Sub

Private Sub BUSCO_CHEQUES_LIBRADOS()
    sql = "SELECT CHEP_FECVTO,CHEP_NUMERO,CHEP_IMPORT,CHEP_NOMBRE"
    sql = sql & " FROM ChequePropioEstadoVigente"
    sql = sql & " WHERE BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
    sql = sql & " AND CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
    sql = sql & " AND ECH_CODIGO IN (7,8)" 'CHEQUES LIBRADOS O RESTITUIDOS
    If optCierre.Value = True Then
        sql = sql & " AND MONTH(CHEP_FECVTO)=" & XN(Month(Fecha.Value))
        sql = sql & " AND YEAR(CHEP_FECVTO)=" & XN(Year(Fecha.Value))
    End If
    If optLibroBanco.Value = True Then
        sql = sql & " AND CHEP_FECVTO>=" & XDQ("01/" & Month(FechaDesde.Value) & "/" & Year(FechaDesde.Value))
        sql = sql & " AND CHEP_FECVTO<=" & XDQ(Day(UltimoDiadelMes(FechaHasta.Value)) & "/" & Month(FechaHasta.Value) & "/" & Year(FechaHasta.Value))
    End If
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If Rec.EOF = False Then
        Do While Rec.EOF = False
            I = I + 1
            'INSERTO CHEQUES LIBRADOS
            sql = "INSERT INTO TMP_RESUMEN_CUENTA_BANCO"
            sql = sql & " (BANCO,CUENTA,PERIODO,FECHA,DESCRIPCION,COMPROBANTE,"
            sql = sql & " DEBITO,CREDITO,SALDO,ORDEN) VALUES ("
            sql = sql & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)) & ","
            sql = sql & XS(CboCuentas.List(CboCuentas.ListIndex)) & ","
            sql = sql & XS(Format(Fecha.Value, "mm/yyyy")) & ","
            sql = sql & XDQ(Rec!CHEP_FECVTO) & "," 'FECHA DE PAGO
            sql = sql & XS(Rec!CHEP_NOMBRE) & ","
            sql = sql & XS(Rec!CHEP_NUMERO) & ","
            sql = sql & XN(Rec!CHEP_IMPORT) & ","
            sql = sql & XN("0") & ","
            sql = sql & XN("0") & ","
            sql = sql & XN(CStr(I)) & ")"
            DBConn.Execute sql
            'VERIFICO SI LE APLICO EL IMPUESTO AL CHEQUE
            If AplicoImpuesto = True Then
                I = I + 1
                sql = "INSERT INTO TMP_RESUMEN_CUENTA_BANCO"
                sql = sql & " (BANCO,CUENTA,PERIODO,FECHA,DESCRIPCION,COMPROBANTE,"
                sql = sql & " DEBITO,CREDITO,SALDO,ORDEN) VALUES ("
                sql = sql & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)) & ","
                sql = sql & XS(CboCuentas.List(CboCuentas.ListIndex)) & ","
                sql = sql & XS(Format(Fecha.Value, "mm/yyyy")) & ","
                sql = sql & XDQ(Rec!CHEP_FECVTO) & ","
                sql = sql & XS("GRAVAMEN LEY 25413 S/DEB") & ","
                sql = sql & XS(Rec!CHEP_NUMERO) & ","
                sql = sql & XN(CStr(CDbl(Rec!CHEP_IMPORT) * (ValorImpuesto))) & ","
                sql = sql & XN("0") & ","
                sql = sql & XN("0") & ","
                sql = sql & XN(CStr(I)) & ")"
                DBConn.Execute sql
                
                mImpCheque = mImpCheque + (CDbl(Rec!CHEP_IMPORT) * (ValorImpuesto))
            End If
            Rec.MoveNext
        Loop
    End If
    Rec.Close
End Sub

Private Sub CALCULO_SALDO()
    sql = "SELECT * FROM TMP_RESUMEN_CUENTA_BANCO"
    sql = sql & " WHERE BANCO=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
    sql = sql & " AND CUENTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
    'sql = sql & " AND PERIODO=" & XS(Format(Fecha.Value, "mm/yyyy"))
    sql = sql & " ORDER BY FECHA,ORDEN"
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Do While Rec.EOF = False
            If Rec!DEBITO <> 0 Then
                Saldo = Saldo - CDbl(Rec!DEBITO)
            ElseIf Rec!CREDITO <> 0 Then
                Saldo = Saldo + CDbl(Rec!CREDITO)
            Else
                Saldo = Rec!Saldo
            End If
            'ACTUALIZO EL SALDO
            sql = "UPDATE TMP_RESUMEN_CUENTA_BANCO"
            sql = sql & " SET SALDO=" & XN(CStr(Saldo))
            sql = sql & " WHERE BANCO=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
            sql = sql & " AND CUENTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
            'sql = sql & " AND PERIODO=" & XS(Format(Fecha.Value, "mm/yyyy"))
            sql = sql & " and ORDEN=" & XN(Rec!Orden)
            DBConn.Execute sql
            Rec.MoveNext
        Loop
    End If
    Rec.Close
End Sub

Private Sub REGISTRO_SALDO_RESUMEN()
    'ACTUALIZO EL SALDO EN RESUMEN CUENTA BANCO
    sql = "INSERT INTO RESUMEN_CUENTA_BANCO"
    sql = sql & " (BAN_CODINT,CTA_NROCTA,RCB_FECHA,RCB_SALDO)"
    sql = sql & " VALUES ("
    sql = sql & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex)) & ","
    sql = sql & XS(CboCuentas.List(CboCuentas.ListIndex)) & ","
    sql = sql & XDQ(Fecha.Value) & ","
    sql = sql & XN(CStr(Saldo)) & ")"
    DBConn.Execute sql
    'ACTUALIZO EL SALDO EN LA CUENTA BANCARIA
    sql = "UPDATE CTA_BANCARIA"
    sql = sql & " SET CTA_SALACT=" & XN(CStr(Saldo))
    sql = sql & " WHERE BAN_CODINT=" & XN(CboBancoBoleta.ItemData(CboBancoBoleta.ListIndex))
    sql = sql & " AND CTA_NROCTA=" & XS(CboCuentas.List(CboCuentas.ListIndex))
    DBConn.Execute sql
End Sub

Private Sub Fecha_Change()
    If Fecha.Value = "" Then lblPeriodo1.Caption = ""
    FechaDesde.Value = ""
    FechaHasta.Value = ""
End Sub

Private Sub Fecha_LostFocus()
    If Trim(Fecha.Value) <> "" Then
        lblPeriodo1.Caption = UCase(Format(Fecha.Value, "mmmm/yyyy"))
    Else
        lblPeriodo1.Caption = ""
    End If
End Sub

Private Sub FechaDesde_Change()
    Fecha.Value = ""
    If FechaDesde.Value = "" Then lblPeriodoD.Caption = ""
End Sub

Private Sub FechaDesde_LostFocus()
    If Trim(FechaDesde.Value) <> "" Then
        lblPeriodoD.Caption = UCase(Format(FechaDesde.Value, "mmmm/yyyy"))
    Else
        lblPeriodoD.Caption = ""
    End If
End Sub

Private Sub FechaHasta_Change()
    Fecha.Value = ""
    If FechaHasta.Value = "" Then lblPeriodoH.Caption = ""
End Sub

Private Sub FechaHasta_LostFocus()
    If Trim(FechaHasta.Value) <> "" Then
        lblPeriodoH.Caption = UCase(Format(FechaHasta.Value, "mmmm/yyyy"))
    Else
        lblPeriodoH.Caption = ""
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    Set Rec = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    cboDestino.ListIndex = 0
    
    Me.Left = 0
    Me.Top = 0
    
    'CARGO COMBO BANCO
    CargoBanco
    cmdListar.Enabled = True
    sql = "SELECT APLICA_IMPUESTO,VALOR_IMPUESTO FROM PARAMETROS"
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If Rec.EOF = False Then
        AplicoImpuesto = True 'APLICO IMPUESTO
        ValorImpuesto = CDbl(Rec!VALOR_IMPUESTO)
    Else
        AplicoImpuesto = False 'NO APLICO IMPUESTO
        ValorImpuesto = 0
    End If
    Rec.Close
    optLibroBanco.Value = True
    FrameCierre.Enabled = False
    RegistroSaldo = False
End Sub

Private Sub CargoBanco()
    sql = "SELECT DISTINCT B.BAN_CODINT, B.BAN_DESCRI"
    sql = sql & " FROM BANCO B, CTA_BANCARIA CB"
    sql = sql & " WHERE B.BAN_CODINT=CB.BAN_CODINT"
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Do While Rec.EOF = False
            CboBancoBoleta.AddItem Trim(Rec!BAN_DESCRI)
            CboBancoBoleta.ItemData(CboBancoBoleta.NewIndex) = Trim(Rec!BAN_CODINT)
            Rec.MoveNext
        Loop
        CboBancoBoleta.ListIndex = 0
    End If
    Rec.Close
End Sub

Private Sub optCierre_Click()
    If optCierre.Value = True Then
        Fecha.Value = ""
        FechaDesde.Value = ""
        FechaHasta.Value = ""
        FrameCierre.Enabled = True
        FrameLibro.Enabled = False
    End If
End Sub

Private Sub optLibroBanco_Click()
    If optLibroBanco.Value = True Then
        Fecha.Value = ""
        FechaDesde.Value = ""
        FechaHasta.Value = ""
        FrameCierre.Enabled = False
        FrameLibro.Enabled = True
    End If
End Sub
