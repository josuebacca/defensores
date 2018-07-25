VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmChPro_Emitidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Cheques"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   11595
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOrden 
      Height          =   315
      Left            =   60
      TabIndex        =   11
      Text            =   "A"
      Top             =   4065
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Frame Frame5 
      Caption         =   "Referencia (Con respecto a la Fec Vto)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   105
      TabIndex        =   12
      Top             =   4770
      Width           =   11355
      Begin VB.Shape Shape3 
         BackColor       =   &H00FF00FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   5910
         Top             =   315
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   2865
         Top             =   315
         Width           =   1215
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   8700
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cheque en Fecha."
         Height          =   195
         Left            =   7185
         TabIndex        =   16
         Top             =   330
         Width           =   1320
      End
      Begin VB.Label Label8 
         Caption         =   "Cheque con margen de 1 a 2 días."
         Height          =   390
         Left            =   4170
         TabIndex        =   15
         Top             =   270
         Width           =   1515
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cheque Vencido."
         Height          =   195
         Left            =   9990
         TabIndex        =   14
         Top             =   330
         Width           =   1230
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H0080FF80&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   90
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Valores al Cobro"
         Height          =   195
         Left            =   1395
         TabIndex        =   13
         Top             =   330
         Width           =   1155
      End
   End
   Begin Crystal.CrystalReport Rpt 
      Left            =   6210
      Top             =   6135
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdCambImpresora 
      Caption         =   "Cambiar Impresora"
      Height          =   375
      Left            =   1530
      TabIndex        =   8
      Top             =   6255
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton CmdReporte 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3930
      TabIndex        =   7
      Top             =   6255
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox mTotImpCh 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   5655
      TabIndex        =   5
      Top             =   4335
      Width           =   1395
   End
   Begin VB.TextBox TxtCant_ch 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2670
      TabIndex        =   3
      Top             =   4335
      Width           =   975
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10320
      TabIndex        =   2
      Top             =   4350
      Width           =   1095
   End
   Begin VB.CommandButton CmdConsultarCh 
      Caption         =   "&Consulta"
      Height          =   375
      Left            =   9045
      TabIndex        =   0
      Top             =   4350
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid GrdCheques 
      Height          =   4215
      Left            =   105
      TabIndex        =   1
      Top             =   30
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedCols       =   0
      RowHeightMin    =   300
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   510
      Top             =   6195
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   64
   End
   Begin VB.TextBox TxtVieneOtroForm 
      Height          =   285
      Left            =   10110
      TabIndex        =   10
      Top             =   30
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "CHEQUES NO A LA ORDEN"
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
      Left            =   7965
      TabIndex        =   17
      Top             =   4410
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   150
      Left            =   7290
      Top             =   4440
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label LBImpActual 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Impresora Actual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   855
      TabIndex        =   9
      Top             =   5955
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Importe Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4095
      TabIndex        =   6
      Top             =   4365
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad de Cheques "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   4
      Top             =   4365
      Width           =   2310
   End
End
Attribute VB_Name = "FrmChPro_Emitidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REC_PARAM As ADODB.Recordset
Dim CANT_CH As Double
Public mConsulta As String
Public mQuienLlama As String
Dim i As Integer

Private Sub CmdCambImpresora_Click()
  CDImpresora.PrinterDefault = True
  CDImpresora.ShowPrinter
  LBImpActual.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub CmdConsultarCh_Click()
    Set rec = New ADODB.Recordset
    mTotImpCh.Text = ""
    GrdCheques.Rows = 1
    Set Rec1 = New ADODB.Recordset
    If Rec1.State = 1 Then Rec1.Close
    
    cSQL = "SELECT  CHE_FECVTO,CHE_NUMERO,"
    cSQL = cSQL & " CHE_FECEMI,CHE_IMPORT,BAN_DESCRI,BAN_CODINT, CHE_NOMCTA,"
    cSQL = cSQL & " BAN_BANCO, BAN_LOCALIDAD, BAN_SUCURSAL, BAN_CODIGO, ECH_CODIGO"
    cSQL = cSQL & " FROM ChequeEstadoVigente"
    cSQL = cSQL & " WHERE ECH_CODIGO = 1" 'CHEQUES TERCERO EN CARTERA
    cSQL = cSQL & " ORDER BY CHE_FECVTO ASC "
    Rec1.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    CANT_CH = 0
    If Not Rec1.EOF Then
        Do While Not Rec1.EOF
            
            CANT_CH = CANT_CH + 1
            
            GrdCheques.AddItem Rec1!CHE_FECEMI & Chr(9) & Rec1!CHE_NUMERO & Chr(9) & Rec1!CHE_FECVTO & Chr(9) & _
                            Format(Rec1!CHE_IMPORT, "0.00") & Chr(9) & Trim(Rec1!BAN_DESCRI) & Chr(9) & _
                            Rec1!BAN_CODINT & Chr(9) & ChkNull(Rec1!CHE_NOMCTA) & Chr(9) & _
                            Rec1!BAN_BANCO & Chr(9) & Rec1!BAN_LOCALIDAD & Chr(9) & _
                            Rec1!BAN_SUCURSAL & Chr(9) & Rec1!BAN_CODIGO
            
            i = GrdCheques.Rows - 1
            If (Int(i / 2) * 2) = i Then
                CambiaColorAFilaDeGrilla GrdCheques, i, , &HE0E0E0
            Else
                CambiaColorAFilaDeGrilla GrdCheques, i, , vbWhite
            End If
'            i = GrdCheques.Rows - 1
'            If (Int(i / 2) * 2) = i Then
'                CambiaColorAFilaDeGrilla GrdCheques, i, , &HE0E0E0
'            Else
'                CambiaColorAFilaDeGrilla GrdCheques, i, , vbWhite
'            End If
            
            'ESTA OPCION ES PARA CHEQUES EN CARTERA
            If ChkNull(Rec1!CHE_FECVTO) <> "" Then
                If CDate(Rec1!CHE_FECVTO) <= CDate(Date) Then
                    GrdCheques.row = GrdCheques.Rows - 1
                    GrdCheques.Col = 3
                    GrdCheques.CellBackColor = &H80FF80      'VERDE PARA CUANDO ESTA EN FECHA
                End If
                CNT_DIAS = (CDate(Rec1!CHE_FECVTO) + 30) - CDate(Date)
                If CNT_DIAS <= 2 And CNT_DIAS > 0 Then
                    GrdCheques.row = GrdCheques.Rows - 1
                    GrdCheques.Col = 3
                    GrdCheques.CellBackColor = &H8080FF      'ROSA VIEJO
                    'GrdCheques.ToolTipText = "La Fecha de Cobro de este cheque vencera en " & CNT_DIAS & " Días "
                ElseIf CNT_DIAS = 0 Then
                    GrdCheques.row = GrdCheques.Rows - 1
                    GrdCheques.Col = 3
                    GrdCheques.CellBackColor = &HFF00FF      'LILA
                    'GrdCheques.ToolTipText = "La Fecha de Cobro de este cheque vencera HOY "
                ElseIf CNT_DIAS < 0 Then
                    GrdCheques.row = GrdCheques.Rows - 1
                    GrdCheques.Col = 3
                    GrdCheques.CellBackColor = &HFF&         'ROJO
                    'GrdCheques.ToolTipText = "La Fecha de Cobro de este cheque vencio hace " & Abs(CNT_DIAS) & " Días "
                End If
            End If

            mTotImpCh.Text = CDbl(Chk0(mTotImpCh.Text)) + CDbl(Rec1!CHE_IMPORT)
            mTotImpCh.Text = Valido_Importe(mTotImpCh.Text)
            Rec1.MoveNext
        Loop
    Else
        MsgBox "No se Encontraron Cheques", vbExclamation, TIT_MSGBOX
    End If
    Rec1.Close
    TxtCant_ch.Text = CANT_CH
End Sub

Private Sub CmdReporte_Click()
    Dim cnt_ch As Integer
    If GrdCheques.Rows > 1 Then
        DBConn.Execute "DELETE FROM tmp_plache "
        DBConn.BeginTrans
        cnt_ch = 1
        Do While cnt_ch < GrdCheques.Rows
            ISQL = " insert into tmp_plache (ID,Cta,Fec_Dep,Hs_Acre,Nro_Cheque,Fecha,Importe,Banco,Suc_Rec,Recibo,Suc_O_Pago,Ord_Pgo) values ( "
            ISQL = ISQL & XN(ChkNull(GrdCheques.TextMatrix(cnt_ch, 0))) & ", "
            ISQL = ISQL & XS(ChkNull(GrdCheques.TextMatrix(cnt_ch, 1))) & ", "
            ISQL = ISQL & XD(ChkNull(GrdCheques.TextMatrix(cnt_ch, 2))) & ", "
            ISQL = ISQL & XS(ChkNull(GrdCheques.TextMatrix(cnt_ch, 3))) & ", "   'hs. acreditacion
            ISQL = ISQL & XS(ChkNull(GrdCheques.TextMatrix(cnt_ch, 4))) & ", "   'nro de cheque
            ISQL = ISQL & XD(ChkNull(GrdCheques.TextMatrix(cnt_ch, 5))) & ", "   'fecha
            ISQL = ISQL & XN(ChkNull(GrdCheques.TextMatrix(cnt_ch, 6))) & ", "   'importe
            ISQL = ISQL & XS(ChkNull(GrdCheques.TextMatrix(cnt_ch, 7))) & ", "   'banco
            ISQL = ISQL & XN(ChkNull(GrdCheques.TextMatrix(cnt_ch, 8))) & ", "   'suc recibo
            ISQL = ISQL & XN(ChkNull(GrdCheques.TextMatrix(cnt_ch, 9))) & ", "   'nro recibo
            ISQL = ISQL & XN(ChkNull(GrdCheques.TextMatrix(cnt_ch, 10))) & ", "  'suc orden pago
            ISQL = ISQL & XN(ChkNull(GrdCheques.TextMatrix(cnt_ch, 11)))         'orden de pago
            ISQL = ISQL & ")"
            cnt_ch = cnt_ch + 1
            DBConn.Execute ISQL
        Loop
        DBConn.CommitTrans
        
        Dim wCondicion, wSentido As String
        Rpt.Connect = RPTCONNECT
        Set REC_PARAM = New ADODB.Recordset
        csql_param = "SELECT * FROM PARAM "
        If REC_PARAM.State = 1 Then REC_PARAM.Close
        REC_PARAM.Open csql_param, DBConn, adOpenStatic, adLockOptimistic
        mDir_inmecar = ChkNull(REC_PARAM!direccion)
        mTel_inmecar = ChkNull(REC_PARAM!TELEFONOS)
        mCp_inmecar = ChkNull(REC_PARAM!CP)
        mCiudad_inmecar = ChkNull(REC_PARAM!CIUDAD)
        mEmail_inmecar = ChkNull(REC_PARAM!email)
        mCuit_inmecar = Mid(REC_PARAM!cuit, 1, 2) & "-" & Mid(REC_PARAM!cuit, 3, 8) & "-" & Mid(REC_PARAM!cuit, 11, 1)
        REC_PARAM.Close
        Rpt.Formulas(0) = "DIR_INMECAR = " & XS(mDir_inmecar)
        Rpt.Formulas(1) = "CIUDAD_INMECAR = " & XS(mCiudad_inmecar)
        Rpt.Formulas(2) = "TEL_INMECAR = " & XS(mTel_inmecar)
        Rpt.Formulas(3) = "CP_INMECAR = " & XS(mCp_inmecar)
        Rpt.Formulas(4) = "EMAIL = " & XS(mEmail_inmecar)
        Rpt.Formulas(5) = "CUIT_EMP = " & XS(mCuit_inmecar)
        
'        If VIENE_DE_CAJA.Value = 1 And frmListCaja.ChkResumen.Value = 1 Then
'            Rpt.Formulas(6) = "FECHAD = " & XS(frmListCaja.mFecha.Text)
'            Rpt.Formulas(7) = "FECHAH = " & XS(frmListCaja.mFechaH.Text)
'
'            If frmListCaja.OptCajaDiaria.Value = True Then
'              Rpt.Formulas(8) = "TITULO_CAJA = " & XS("CIERRE DE CAJA DIARIA AL  " & frmListCaja.mFechaH.Text)
'            Else
'              Rpt.Formulas(8) = "TITULO_CAJA = " & XS("CIERRE DE CAJA DESDE FECHA  " & frmListCaja.mFecha.Text & "  HASTA FECHA " & frmListCaja.mFechaH.Text)
'            End If
'
'            Rpt.ReportFileName = RPTPATH + "CH_CARTERA.RPT"
'            Rpt.Action = 1
'            'CmdSalir_Click
'        ElseIf VIENE_DE_CAJA.Value = 0 Then
'            Rpt.Formulas(6) = ""
'            Rpt.Formulas(7) = ""
'            Rpt.Formulas(8) = ""
'
'            Rpt.ReportFileName = RPTPATH + "CH_CARTERA.RPT"
'            Rpt.Action = 1
'        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
    Set FrmChPro_Emitidos = Nothing
End Sub

Private Sub Form_Activate()
    If Me.TxtVieneOtroForm <> "" And ActiveControl.Name <> "CmdSalir" And ActiveControl.Name <> "GrdCheques" Then
        CmdConsultarCh.Enabled = False
        CmdReporte.Enabled = False
        
        CmdConsultarCh_Click
        CmdReporte_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9): KeyAscii = 0
End Sub

Private Sub Form_Load()
    CentrarW Me
    
    GrdCheques.FormatString = "^Fec.Emisión|^Nro.Cheque" _
                            & "|^Fec.Cobro|>Importe|<Banco" _
                            & "|BAN_CODINT|^Nro Cuenta" _
                            & "|BAN_BANCO|BAN_LOCALIDAD|BAN_SUCURSAL|BAN_CODIGO"
                            
    GrdCheques.ColWidth(0) = 1300   'Fec.Emisión
    GrdCheques.ColWidth(1) = 1300  'Nro.Cheque
    GrdCheques.ColWidth(2) = 1300  'Fec.Vencimineto
    GrdCheques.ColWidth(3) = 1100  'Importe
    GrdCheques.ColWidth(4) = 4500  'Banco
    GrdCheques.ColWidth(5) = 0     'BAN_CODINT
    GrdCheques.ColWidth(6) = 1300  'CHE_NROCTA
    GrdCheques.ColWidth(7) = 0     'BAN_BANCO
    GrdCheques.ColWidth(8) = 0     'BAN_LOCALIDAD
    GrdCheques.ColWidth(9) = 0     'BAN_SUCURSAL
    GrdCheques.ColWidth(10) = 0    'BAN_CODIGO
    GrdCheques.Cols = 11
    GrdCheques.Rows = 1
    GrdCheques.row = 0
    For i = 0 To GrdCheques.Cols - 1
        GrdCheques.Col = i
        GrdCheques.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdCheques.CellBackColor = &H808080    'GRIS OSCURO
        GrdCheques.CellFontBold = True
    Next
    
    CmdConsultarCh_Click
End Sub

Private Sub GrdCheques_Click()
    If GrdCheques.MouseRow = 0 Then
        GrdCheques.Col = GrdCheques.MouseCol
        GrdCheques.ColSel = GrdCheques.MouseCol
        
        If txtOrden.Text = "A" Then
            GrdCheques.Sort = 2
            txtOrden.Text = "B"
        Else
            GrdCheques.Sort = 1
            txtOrden.Text = "A"
        End If
    End If
End Sub

Private Sub GrdCheques_DblClick()
    If GrdCheques.Rows > 1 Then
        For i = 1 To ABMEgresos.grdPagos.Rows - 1
            If ABMEgresos.grdPagos.TextMatrix(i, 4) = GrdCheques.TextMatrix(GrdCheques.RowSel, 0) Then
                MsgBox "El Cheque ya fue Seleccionado", vbExclamation, TIT_MSGBOX
                Exit Sub
            End If
        Next
        ABMEgresos.grdPagos.AddItem "CHEQUE TERCERO" & Chr(9) & GrdCheques.TextMatrix(GrdCheques.RowSel, 3) & Chr(9) & "2" & Chr(9) & _
                     GrdCheques.TextMatrix(GrdCheques.RowSel, 5) & Chr(9) & GrdCheques.TextMatrix(GrdCheques.RowSel, 1) & Chr(9) & _
                     GrdCheques.TextMatrix(GrdCheques.RowSel, 0)
        If GrdCheques.Rows > 2 Then
            GrdCheques.RemoveItem GrdCheques.RowSel
        Else
            GrdCheques.Rows = 1
        End If
    End If
End Sub

Private Sub GrdCheques_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        If GrdCheques.ColSel <> 8 And GrdCheques.ColSel <> 9 And GrdCheques.ColSel <> 10 And GrdCheques.ColSel <> 11 Then
'            FrmCargaCheques.mNro.Text = GrdCheques.TextMatrix(GrdCheques.RowSel, 4)
'            FrmCargaCheques.mID_Cheque.Caption = GrdCheques.TextMatrix(GrdCheques.RowSel, 0)
'            'FrmCargaCheques.mSucu.Text = GrdCheques.TextMatrix(GrdCheques.Rows, 4)
'            FrmCargaCheques.Show vbModal
'        ElseIf (GrdCheques.ColSel = 10 Or GrdCheques.ColSel = 11) And GrdCheques.TextMatrix(GrdCheques.RowSel, 11) <> "" Then
'            'frmConsPagos.mCodigo.Text = Me.TxtCodigoProv.Text
'            frmConsPagos.mIniPag.Text = GrdCheques.TextMatrix(GrdCheques.RowSel, 10)
'            frmConsPagos.mNroFactu.Text = GrdCheques.TextMatrix(GrdCheques.RowSel, 11)
'            frmConsPagos.Show vbModal
'        ElseIf (GrdCheques.ColSel = 8 Or GrdCheques.ColSel = 9) And GrdCheques.TextMatrix(GrdCheques.RowSel, 9) <> "" Then
'            frmConsCobro.mNroFactu.Text = GrdCheques.TextMatrix(GrdCheques.RowSel, 9)
'            'If auxDllActiva.FormBase.lstvLista.SelectedItem.SubItems(12) = "A" Then
'            '    frmConsCobro.Otipo_A.Value = True
'            'Else
'            '    frmConsCobro.Otipo_B.Value = True
'            'End If
'            frmConsCobro.Show vbModal
'        End If
'    End If
End Sub

Private Sub OptBuscaID_Click()
    txtID.Enabled = True
    txtID.SetFocus
End Sub

Private Sub OptChPropio_Click()
'    TxtID.Enabled = False
'    TxtID.Text = ""
End Sub

Private Sub OptChTercero_Click()
'    TxtID.Enabled = False
'    TxtID.Text = ""
End Sub

Private Sub TxtID_GotFocus()
    seltxt
End Sub

Private Sub txtID_LostFocus()
    If txtID.Text <> "" And ActiveControl.Name <> "CmdSalir" Then
        CmdConsultarCh.SetFocus
    End If
End Sub

