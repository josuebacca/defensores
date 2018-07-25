VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmListChequesPropios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Cheques Propios"
   ClientHeight    =   5415
   ClientLeft      =   1365
   ClientTop       =   975
   ClientWidth     =   8670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmListChequesPropios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   8670
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
      Height          =   5370
      Left            =   30
      TabIndex        =   21
      Top             =   0
      Width           =   8595
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
         TabIndex        =   22
         Top             =   3570
         Width           =   8265
         Begin VB.CommandButton CmdCambiarImp 
            Caption         =   "&Configurar Impresora"
            Height          =   435
            Left            =   195
            TabIndex        =   36
            Top             =   660
            Width           =   1890
         End
         Begin VB.OptionButton oImpresora 
            Caption         =   "Impresora"
            Height          =   255
            Left            =   2280
            TabIndex        =   17
            Top             =   270
            Width           =   1095
         End
         Begin VB.OptionButton oPantalla 
            Caption         =   "Pantalla"
            Height          =   255
            Left            =   1230
            TabIndex        =   16
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
            TabIndex        =   37
            Top             =   840
            Width           =   1485
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Destino:"
            Height          =   195
            Left            =   480
            TabIndex        =   35
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
         Height          =   3300
         Left            =   3825
         TabIndex        =   25
         Top             =   285
         Width           =   4590
         Begin MSComCtl2.DTPicker TxtFecVtoD 
            Height          =   315
            Left            =   1560
            TabIndex        =   4
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   52428801
            CurrentDate     =   42925
         End
         Begin VB.ComboBox cboCtaBancaria 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1140
            Width           =   2100
         End
         Begin VB.ComboBox CboEstado 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2385
            Width           =   2985
         End
         Begin VB.TextBox TxtNroCheque 
            Enabled         =   0   'False
            Height          =   330
            Left            =   1530
            TabIndex        =   8
            Top             =   1545
            Width           =   1080
         End
         Begin VB.ComboBox CboBanco 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   735
            Width           =   2985
         End
         Begin MSComCtl2.DTPicker TxtFecVtoH 
            Height          =   315
            Left            =   3240
            TabIndex        =   5
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   52428801
            CurrentDate     =   42925
         End
         Begin MSComCtl2.DTPicker TxtFecIngresoD 
            Height          =   315
            Left            =   1530
            TabIndex        =   9
            Top             =   1950
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   52428801
            CurrentDate     =   42925
         End
         Begin MSComCtl2.DTPicker TxtFecIngresoH 
            Height          =   315
            Left            =   3240
            TabIndex        =   10
            Top             =   1950
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   52428801
            CurrentDate     =   42925
         End
         Begin MSComCtl2.DTPicker Fecha1 
            Height          =   315
            Left            =   1560
            TabIndex        =   12
            Top             =   2800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   52428801
            CurrentDate     =   42925
         End
         Begin MSComCtl2.DTPicker Fecha2 
            Height          =   315
            Left            =   3270
            TabIndex        =   13
            Top             =   2800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   52428801
            CurrentDate     =   42925
         End
         Begin VB.Label Label9 
            Caption         =   "(Fecha Cambio Estado)"
            Height          =   570
            Left            =   165
            TabIndex        =   39
            Top             =   2490
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro Cuenta:"
            Height          =   195
            Index           =   4
            Left            =   525
            TabIndex        =   38
            Top             =   1170
            Width           =   885
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            Height          =   195
            Left            =   915
            TabIndex        =   34
            Top             =   2820
            Width           =   510
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Index           =   2
            Left            =   3015
            TabIndex        =   33
            Top             =   2820
            Width           =   120
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Index           =   1
            Left            =   3030
            TabIndex        =   32
            Top             =   390
            Width           =   120
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Index           =   0
            Left            =   3030
            TabIndex        =   31
            Top             =   2010
            Width           =   120
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   840
            TabIndex        =   30
            Top             =   2460
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   870
            TabIndex        =   29
            Top             =   795
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nro de Cheque:"
            Height          =   195
            Left            =   255
            TabIndex        =   28
            Top             =   1575
            Width           =   1140
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emisión:"
            Height          =   195
            Left            =   75
            TabIndex        =   27
            Top             =   2010
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Pago:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   375
            Width           =   1125
         End
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         DisabledPicture =   "FrmListChequesPropios.frx":27A2
         Height          =   420
         Left            =   7350
         TabIndex        =   20
         Top             =   4815
         Width           =   1080
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Aceptar"
         DisabledPicture =   "FrmListChequesPropios.frx":2BEC
         Height          =   420
         Left            =   5160
         TabIndex        =   18
         Top             =   4815
         Width           =   1080
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         DisabledPicture =   "FrmListChequesPropios.frx":34B6
         Height          =   420
         Left            =   6255
         TabIndex        =   19
         Top             =   4815
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
         TabIndex        =   24
         Top             =   2625
         Width           =   3660
         Begin VB.OptionButton oDescendente 
            Caption         =   "Descendente"
            Height          =   255
            Left            =   1965
            TabIndex        =   15
            Top             =   435
            Width           =   1335
         End
         Begin VB.OptionButton oAscendente 
            Caption         =   "Ascendente"
            Height          =   255
            Left            =   210
            TabIndex        =   14
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
         Height          =   2355
         Left            =   165
         TabIndex        =   23
         Top             =   285
         Width           =   3660
         Begin VB.OptionButton Option0 
            Caption         =   "... por Fecha de Pago"
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   225
            Value           =   -1  'True
            Width           =   2910
         End
         Begin VB.OptionButton Option4 
            Caption         =   "... por Estado"
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1875
            Width           =   2910
         End
         Begin VB.OptionButton Option1 
            Caption         =   "... por Banco y Nro de Cheque"
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   765
            Width           =   2910
         End
         Begin VB.OptionButton Option3 
            Caption         =   "... por Fecha de Emisión"
            Height          =   330
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1320
            Width           =   2910
         End
      End
      Begin Crystal.CrystalReport Rep 
         Left            =   4005
         Top             =   4830
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
         Left            =   4485
         Top             =   4800
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Flags           =   64
      End
   End
End
Attribute VB_Name = "FrmListChequesPropios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Limpio_Campos()
   Me.TxtFecVtoD.Value = ""
   Me.TxtFecVtoH.Value = ""
   Me.CboBanco.ListIndex = -1
   Me.cboCtaBancaria.ListIndex = -1
   Me.TxtNroCheque.Text = ""
   Me.TxtFecIngresoD.Value = ""
   Me.TxtFecIngresoH.Value = ""
   Me.cboEstado.ListIndex = -1
   Me.fecha1.Value = ""
   Me.Fecha2.Value = ""
End Sub

Private Sub CboBanco_LostFocus()
    If CboBanco.ListIndex <> -1 Then
        Set Rec1 = New ADODB.Recordset
        cboCtaBancaria.Clear
        sql = "SELECT CTA_NROCTA FROM CTA_BANCARIA"
        sql = sql & " WHERE BAN_CODINT=" & XN(CboBanco.ItemData(CboBanco.ListIndex))
        sql = sql & " AND CTA_FECCIE IS NULL"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
         Do While Rec1.EOF = False
             cboCtaBancaria.AddItem Trim(Rec1!CTA_NROCTA)
             Rec1.MoveNext
         Loop
         cboCtaBancaria.ListIndex = 0
        End If
        Rec1.Close
    End If
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
    
    'VALIDO LAS FECHAS
    If Option0.Value = True Then
        If TxtFecVtoD.Value = "" Then
            MsgBox "Falta ingresar la fecha desde la cual quiere consultar", vbExclamation, TIT_MSGBOX
            TxtFecVtoD.SetFocus
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
    End If
   
   On Error GoTo ErrorTrans
   
   Screen.MousePointer = 11
   
   'Sentido del Orden
   If oAscendente.Value = True Then
      wSentido = "+"
      Rep.Formulas(1) = "sentido ='Sentido: ASCENDENTE'"
   Else
      wSentido = "-"
      Rep.Formulas(1) = "sentido ='Sentido: DESCENDENTE '"
   End If
   
   If Me.Option0.Value = True Then 'Por Fecha de Vencimiento (fec pago)
       
       If Me.TxtFecVtoD.Value = "" Or Me.TxtFecVtoH.Value = "" Then
          If Me.TxtFecVtoD.Value = "" Then
            Me.TxtFecVtoD.Value = Format(Date, "dd/mm/yyyy")
          ElseIf Me.TxtFecVtoH.Value = "" Then
            Me.TxtFecVtoH.Value = Format(Date, "dd/mm/yyyy")
          End If
       End If
       Call MuestroFechaCrystal(Me.TxtFecVtoD, Me.TxtFecVtoH, "Fecha de pago  ")
       
       '{ChequePropioEstadoVigente.ECH_CODIGO} = 1 Unicamente Cheques en Cartera
        sql = sql & " {ChequePropioEstadoVigente.CHEP_FECVTO} >= DATE(" & Mid(TxtFecVtoD.Value, 7, 4) & "," & _
                                                            Mid(TxtFecVtoD.Value, 4, 2) & "," & _
                                                            Mid(TxtFecVtoD.Value, 1, 2) & ") and " & _
                      "{ChequePropioEstadoVigente.CHEP_FECVTO} <= DATE(" & Mid(TxtFecVtoH.Value, 7, 4) & "," & _
                                                                    Mid(TxtFecVtoH.Value, 4, 2) & "," & _
                                                                    Mid(TxtFecVtoH.Value, 1, 2) & ")"
       wCondicion = wSentido & " {ChequePropioEstadoVigente.CHEP_FECVTO}"
       wCondicion1 = wSentido & " {ChequePropioEstadoVigente.CHEP_NUMERO}"
       Rep.Formulas(0) = "orden ='Ordenado por: FECHA DE PAGO. Y NRO DE CHEQUE'"
       
   ElseIf Me.Option1.Value = True Then 'por Banco y Nº de Cheque
       
        sql = sql & " {ChequePropioEstadoVigente.BAN_CODINT} =  " & XN(CboBanco.ItemData(CboBanco.ListIndex)) _
                 & " AND {ChequePropioEstadoVigente.CTA_NROCTA} = " & XS(cboCtaBancaria.List(cboCtaBancaria.ListIndex))
        If TxtNroCheque.Text <> "" Then
            sql = sql & " AND {ChequePropioEstadoVigente.CHEP_NUMERO} =  " & XS(TxtNroCheque.Text)
        End If
       
       wCondicion = wSentido & " {ChequePropioEstadoVigente.CHEP_NUMERO}"
       wCondicion1 = ""
       Rep.Formulas(0) = "orden ='Ordenado por: NÚMERO DE CHEQUE'"
          
   ElseIf Me.Option3.Value = True Then 'por Fecha de Ingreso (fec emi)
   
       If Me.TxtFecIngresoD.Value = "" Or Me.TxtFecIngresoH.Value = "" Then
          If Me.TxtFecIngresoD.Value = "" Then
            Me.TxtFecIngresoD.Value = Format(Date, "dd/mm/yyyy")
          ElseIf Me.TxtFecIngresoH.Value = "" Then
            Me.TxtFecIngresoH.Value = Format(Date, "dd/mm/yyyy")
          End If
       End If
       Call MuestroFechaCrystal(Me.TxtFecIngresoD, Me.TxtFecIngresoH, "Fecha de Emisión  ")
       
       sql = sql & "{ChequePropioEstadoVigente.CHEP_FECEMI} >= DATE(" & Mid(TxtFecIngresoD.Value, 7, 4) & _
                                                      "," & Mid(TxtFecIngresoD.Value, 4, 2) & _
                                                      "," & Mid(TxtFecIngresoD.Value, 1, 2) & ")and " & _
                   "{ChequePropioEstadoVigente.CHEP_FECEMI} <= DATE(" & Mid(TxtFecIngresoH.Value, 7, 4) & "," & _
                                                            Mid(TxtFecIngresoH.Value, 4, 2) & "," & _
                                                            Mid(TxtFecIngresoH.Value, 1, 2) & ")"
       
       wCondicion = wSentido & " {ChequePropioEstadoVigente.CHEP_FECENT}"
       wCondicion1 = wSentido & " {ChequePropioEstadoVigente.CHEP_FECVTO}"
       
       Rep.Formulas(0) = "orden ='Ordenado por: FECHA DE INGRESO y FECHA DE PAGO.'"
   
   ElseIf Me.Option4.Value = True Then 'por Estado y Fecha de Cambio de estado
   
       If fecha1.Value = "" Or Fecha2.Value = "" Then
          If fecha1.Value = "" Then
            fecha1.Value = Format(Date, "dd/mm/yyyy")
          ElseIf Fecha2.Value = "" Then
            Fecha2.Value = Format(Date, "dd/mm/yyyy")
          End If
       End If
       Call MuestroFechaCrystal(Me.fecha1, Me.Fecha2, "Fecha de Estado  ")
    
       sql = sql & " {ChequePropioEstadoVigente.CPES_FECHA} >= DATE(" & Mid(fecha1.Value, 7, 4) & "," & _
                                                                    Mid(fecha1.Value, 4, 2) & "," & _
                                                                    Mid(fecha1.Value, 1, 2) & ") and " & _
                   "{ChequePropioEstadoVigente.CPES_FECHA} <= DATE(" & Mid(Fecha2.Value, 7, 4) & "," & _
                                                                    Mid(Fecha2.Value, 4, 2) & "," & _
                                                                    Mid(Fecha2.Value, 1, 2) & ")"
       'por Estado
       If Me.cboEstado.List(Me.cboEstado.ListIndex) <> "(Todos)" Then
           If Me.cboEstado.List(Me.cboEstado.ListIndex) = "RECHAZADOS TODOS" Then
              sql = sql & " AND {ChequePropioEstadoVigente.ECH_CODIGO} >= 8 " & _
                            " AND {ChequePropioEstadoVigente.ECH_CODIGO} <= 24 "
           Else
              sql = sql & " AND {ChequePropioEstadoVigente.ECH_CODIGO} =  " & XN(cboEstado.ItemData(cboEstado.ListIndex))
           End If
       End If
       wCondicion = wSentido & " {ChequePropioEstadoVigente.CHEP_FECVTO}"
       wCondicion1 = wSentido & " {ChequePropioEstadoVigente.CHEP_NUMERO}"
       Rep.Formulas(0) = "orden ='Ordenado por: FECHA DE PAGO. Y NRO. DE CHEQUE'"
   
   End If
   
   If oImpresora = True Then
       Rep.Destination = 1
   Else
       Rep.Destination = 0
       Rep.WindowMinButton = 0
       Rep.WindowTitle = "Consulta de Cheques Propios"
       Rep.WindowBorderStyle = 2
   End If
   
   
   Rep.SortFields(0) = wCondicion
   Rep.SortFields(1) = wCondicion1
   
   Rep.SelectionFormula = sql
   Rep.WindowState = crptMaximized
   Rep.WindowBorderStyle = crptNoBorder
   Rep.Connect = "Provider=MSDASQL.1;Persst Security Info=False;Data Source=" & SERVIDOR
   
   Rep.ReportFileName = DRIVE & DirReport & "chequepropio.rpt"
   Rep.Action = 1
   
   Rep.Formulas(0) = ""
   Rep.Formulas(1) = ""
   Rep.Formulas(2) = ""
   Rep.Formulas(3) = ""
   Rep.Formulas(4) = ""
   
   Screen.MousePointer = 1
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
    Set FrmListChequesPropios = Nothing
    Unload Me
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
    Me.cboCtaBancaria.Enabled = False
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
    Me.cboCtaBancaria.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    Me.cboEstado.Enabled = False
    Me.fecha1.Enabled = False
    Me.Fecha2.Enabled = False
    If Me.TxtFecVtoD.Visible = True Then Me.TxtFecVtoD.SetFocus
End Sub

Private Sub Option1_Click()
    Me.CboBanco.Clear
    Set Rec = New ADODB.Recordset
    sql = "SELECT DISTINCT B.BAN_CODINT, B.BAN_DESCRI"
    sql = sql & " FROM BANCO B, CTA_BANCARIA CB"
    sql = sql & " WHERE B.BAN_CODINT=CB.BAN_CODINT"
    sql = sql & " ORDER BY B.BAN_DESCRI"
    
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Rec.MoveFirst
        Do While Not Rec.EOF
            Me.CboBanco.AddItem Trim(Rec!BAN_DESCRI)
            Me.CboBanco.ItemData(Me.CboBanco.NewIndex) = Rec!BAN_CODINT
            Rec.MoveNext
        Loop
        Me.CboBanco.ListIndex = 0
    End If
    Rec.Close
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.ListIndex = 0
    Me.CboBanco.Enabled = True
    Me.cboCtaBancaria.Enabled = True
    Me.TxtNroCheque.Enabled = True
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    Me.cboEstado.Enabled = False
    Me.fecha1.Enabled = False
    Me.Fecha2.Enabled = False
    Me.CboBanco.SetFocus
End Sub

Private Sub Option3_Click()
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.Enabled = False
    Me.cboCtaBancaria.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = True
    Me.TxtFecIngresoH.Enabled = True
    Me.cboEstado.Enabled = False
    Me.fecha1.Enabled = False
    Me.Fecha2.Enabled = False
    Me.TxtFecIngresoD.SetFocus
End Sub

Private Sub Option4_Click()
    Limpio_Campos
    Me.TxtFecVtoD.Enabled = False
    Me.TxtFecVtoH.Enabled = False
    Me.CboBanco.Enabled = False
    Me.cboCtaBancaria.Enabled = False
    Me.TxtNroCheque.Enabled = False
    Me.TxtFecIngresoD.Enabled = False
    Me.TxtFecIngresoH.Enabled = False
    Me.cboEstado.ListIndex = 0
    Me.cboEstado.Enabled = True
    Me.fecha1.Enabled = True
    Me.Fecha2.Enabled = True
    Me.cboEstado.SetFocus
End Sub

Private Sub TxtFecIngresoD_LostFocus()
   'If Me.Option3.Value = True And TxtFecIngresoD.value = "" Then TxtFecIngresoD.value = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub TxtFecIngresoH_LostFocus()
'If Me.Option3.Value = True And TxtFecIngresoh.value = "" Then TxtFecIngresoh.value = Format(Date, "dd/mm/yyyy")

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

  'If Me.Option0.Value = True And TxtFecVtoH.value = "" Then TxtFecVtoH.value = Format(Date, "dd/mm/yyyy")
  
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
    If TxtNroCheque.Text <> "" Then
        If Len(TxtNroCheque.Text) < 10 Then TxtNroCheque.Text = CompletarConCeros(TxtNroCheque.Text, 10)
    End If
    CmdAgregar.SetFocus
End Sub
