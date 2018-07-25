VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "Fecha32.ocx"
Begin VB.Form frmIngresoGastosBancarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Gastos Bancarios"
   ClientHeight    =   4020
   ClientLeft      =   1620
   ClientTop       =   1950
   ClientWidth     =   6975
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6975
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4185
      TabIndex        =   8
      Top             =   3615
      Width           =   915
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3255
      TabIndex        =   7
      Top             =   3615
      Width           =   915
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6045
      TabIndex        =   10
      Top             =   3615
      Width           =   915
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   5115
      TabIndex        =   9
      Top             =   3615
      Width           =   915
   End
   Begin TabDlg.SSTab TabTB 
      Height          =   3555
      Left            =   30
      TabIndex        =   17
      Top             =   15
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   6271
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Datos"
      TabPicture(0)   =   "frmIngresoGastosBancarios.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "B&uscar"
      TabPicture(1)   =   "frmIngresoGastosBancarios.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
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
         Height          =   1260
         Left            =   195
         TabIndex        =   20
         Top             =   360
         Width           =   6570
         Begin VB.ComboBox cboBanco1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   540
            Width           =   3915
         End
         Begin VB.ComboBox cboGasto1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   195
            Width           =   3915
         End
         Begin FechaCtl.Fecha mFechaD 
            Height          =   315
            Left            =   1230
            TabIndex        =   13
            Top             =   900
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.CommandButton CmdBuscAprox 
            Caption         =   "Buscar"
            Height          =   330
            Left            =   5355
            MaskColor       =   &H00000000&
            TabIndex        =   15
            ToolTipText     =   "Buscar"
            Top             =   825
            UseMaskColor    =   -1  'True
            Width           =   1110
         End
         Begin FechaCtl.Fecha mFechaH 
            Height          =   315
            Left            =   3990
            TabIndex        =   14
            Top             =   900
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   165
            TabIndex        =   30
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gasto:"
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   29
            Top             =   255
            Width           =   480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecah Hasta:"
            Height          =   195
            Left            =   2925
            TabIndex        =   25
            Top             =   930
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   165
            TabIndex        =   24
            Top             =   945
            Width           =   990
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Datos del Ingreso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2925
         Left            =   -74865
         TabIndex        =   18
         Top             =   510
         Width           =   6615
         Begin VB.CheckBox chkAplicoImpuesto 
            Caption         =   "Aplicar impuesto transacciones financieras"
            Height          =   240
            Left            =   1260
            TabIndex        =   6
            Top             =   2565
            Width           =   4125
         End
         Begin VB.ComboBox CboBanco 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1455
            Width           =   4395
         End
         Begin VB.ComboBox CboGasto 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1095
            Width           =   4395
         End
         Begin VB.ComboBox cboCtaBancaria 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1815
            Width           =   2100
         End
         Begin VB.TextBox txtImporte 
            Height          =   330
            Left            =   1260
            TabIndex        =   5
            Top             =   2175
            Width           =   1125
         End
         Begin FechaCtl.Fecha FechaGasto 
            Height          =   315
            Left            =   1260
            TabIndex        =   1
            Top             =   765
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   315
            Left            =   1260
            TabIndex        =   0
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "<F1> Buscar Gasto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4275
            TabIndex        =   31
            Top             =   15
            Width           =   1875
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   210
            TabIndex        =   28
            Top             =   1515
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro Cuenta:"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   27
            Top             =   1845
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Index           =   4
            Left            =   210
            TabIndex        =   23
            Top             =   2220
            Width           =   630
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   210
            TabIndex        =   22
            Top             =   810
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Ingreso:"
            Height          =   195
            Index           =   3
            Left            =   210
            TabIndex        =   21
            Top             =   435
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gasto:"
            Height          =   195
            Index           =   2
            Left            =   210
            TabIndex        =   19
            Top             =   1155
            Width           =   480
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   1770
         Left            =   180
         TabIndex        =   16
         Top             =   1650
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   3122
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorSel    =   16761024
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
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
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   26
      Top             =   3630
      Width           =   660
   End
End
Attribute VB_Name = "frmIngresoGastosBancarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer

Private Sub BuscoDatos()
    Set Rec = New ADODB.Recordset
    sql = "SELECT * FROM GASTOS_BANCARIOS"
    sql = sql & " WHERE GBA_NUMERO = " & XN(TxtCodigo.Text)
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then ' si existe
        FechaGasto.Text = ChkNull(Rec!GBA_FECHA)
        Call BuscaCodigoProxItemData(CInt(Rec!TGB_CODIGO), CboGasto)
        Call BuscaCodigoProxItemData(CInt(Rec!BAN_CODINT), CboBanco)
        CboBanco_LostFocus
        Call BuscaProx(Rec!CTA_NROCTA, cboCtaBancaria)
        txtImporte.Text = Valido_Importe(ChkNull(Rec!GBA_IMPORTE))
        If Rec!GBA_IMPUESTO = "S" Then
            chkAplicoImpuesto.Value = Checked
        ElseIf Rec!GBA_IMPUESTO = "N" Then
            chkAplicoImpuesto.Value = Unchecked
        End If
'    Else
'        MsgBox "Gasto Inexistente", vbCritical
'        TxtCODIGO.Text = ""
'        TxtCODIGO.SetFocus
'        Rec.Close
'        Exit Sub
    End If
    Rec.Close
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

Private Sub cmdBorrar_Click()
    On Error GoTo CLAVOSE
    If Trim(TxtCodigo.Text) <> "" Then
        If MsgBox("Seguro desea eliminar el Gasto?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            Screen.MousePointer = vbHourglass
            lblEstado.Caption = "Eliminando ..."
            DBConn.BeginTrans
            DBConn.Execute "DELETE FROM GASTOS_BANCARIOS WHERE GBA_NUMERO = " & XN(TxtCodigo.Text)
            DBConn.CommitTrans
            FechaGasto.SetFocus
            
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            CmdNuevo_Click
        End If
    End If
    Exit Sub
    
CLAVOSE:
    If Rec.State = 1 Then Rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
End Sub

Private Sub CmdBuscAprox_Click()
    Set Rec = New ADODB.Recordset
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    Me.Refresh
    sql = "SELECT GB.GBA_NUMERO, GB.GBA_FECHA, GB.GBA_IMPORTE,"
    sql = sql & " TG.TGB_DESCRI, B.BAN_DESCRI"
    sql = sql & " FROM GASTOS_BANCARIOS GB, TIPO_GASTO_BANCARIO TG,"
    sql = sql & " BANCO B"
    sql = sql & " WHERE"
    sql = sql & " GB.TGB_CODIGO=TG.TGB_CODIGO"
    sql = sql & " AND GB.BAN_CODINT=B.BAN_CODINT"
    If cboBanco1.List(cboBanco1.ListIndex) <> "<Todos>" Then
        sql = sql & " AND GB.BAN_CODINT=" & XN(cboBanco1.ItemData(cboBanco1.ListIndex))
    End If
    If cboGasto1.List(cboGasto1.ListIndex) <> "<Todos>" Then
        sql = sql & " AND GB.TGB_CODIGO=" & XN(cboGasto1.ItemData(cboGasto1.ListIndex))
    End If
    If mFechaD.Text <> "" Then sql = sql & " AND GBA_FECHA >= " & XDQ(mFechaD.Text)
    If mFechaH.Text <> "" Then sql = sql & " AND GBA_FECHA <= " & XDQ(mFechaH.Text)
    
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Rec.MoveFirst
        '^Fecha|importe|gasto|banco|numero
        Do While Not Rec.EOF
            GrdModulos.AddItem Rec!GBA_FECHA & Chr(9) & Valido_Importe(Rec!GBA_IMPORTE) & Chr(9) & _
                        Rec!TGB_DESCRI & Chr(9) & Rec!BAN_DESCRI & Chr(9) & Rec!GBA_NUMERO
            Rec.MoveNext
        Loop
        GrdModulos.HighLight = flexHighlightAlways
        If GrdModulos.Enabled Then GrdModulos.SetFocus
        lblEstado.Caption = ""
    Else
        lblEstado.Caption = ""
        MsgBox "No se encontraron gastos Bancarios", vbExclamation, TIT_MSGBOX
        If cboGasto1.Enabled Then cboGasto1.SetFocus
    End If
    lblEstado.Caption = ""
    Rec.Close
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdGrabar_Click()
    On Error GoTo CLAVOSE
    If TxtCodigo.Text = "" Then
        MsgBox "No ha ingresado el Número de Gasto", vbCritical, TIT_MSGBOX
        TxtCodigo.SetFocus
        Exit Sub
    End If
    If FechaGasto.Text = "" Then
        MsgBox "No ha ingresado la Fecha del Gasto", vbExclamation, TIT_MSGBOX
        FechaGasto.SetFocus
        Exit Sub
    End If
    If cboCtaBancaria.List(cboCtaBancaria.ListIndex) = "" Then
        MsgBox "Debe elegir una Cuenta Bancaria", vbExclamation, TIT_MSGBOX
        CboBanco.SetFocus
        Exit Sub
    End If
    If txtImporte.Text = "" Then
        MsgBox "No ha ingresado el Importe del Gasto", vbExclamation, TIT_MSGBOX
        If txtImporte.Enabled Then txtImporte.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando ..."
    
'    Set Rec = New ADODB.Recordset
    DBConn.BeginTrans
    
    sql = "SELECT GBA_FECHA FROM GASTOS_BANCARIOS WHERE GBA_NUMERO = " & XN(TxtCodigo.Text)
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If Rec.EOF = False Then 'MODIFICO UN GASTO
        sql = "UPDATE GASTOS_BANCARIOS"
        sql = sql & " SET GBA_FECHA = " & XDQ(FechaGasto.Text)
        sql = sql & " ,TGB_CODIGO = " & XN(CboGasto.ItemData(CboGasto.ListIndex))
        sql = sql & " ,BAN_CODINT = " & XN(CboBanco.ItemData(CboBanco.ListIndex))
        sql = sql & " ,CTA_NROCTA = " & XS(cboCtaBancaria.List(cboCtaBancaria.ListIndex))
        sql = sql & " ,GBA_IMPORTE = " & XN(txtImporte.Text)
        If chkAplicoImpuesto.Value = Checked Then
            sql = sql & " ,GBA_IMPUESTO=" & XS("S")
        Else
            sql = sql & " ,GBA_IMPUESTO=" & XS("N")
        End If
        sql = sql & " WHERE GBA_NUMERO = " & XN(TxtCodigo.Text)
        
        DBConn.Execute sql
        
    Else 'INSERTO UN NUEVO GASTO
    
        sql = "INSERT INTO GASTOS_BANCARIOS"
        sql = sql & " (GBA_NUMERO, GBA_FECHA, TGB_CODIGO, BAN_CODINT,"
        sql = sql & " CTA_NROCTA, GBA_IMPORTE, GBA_IMPUESTO)"
        sql = sql & " VALUES ("
        sql = sql & XN(TxtCodigo.Text) & ","
        sql = sql & XDQ(FechaGasto.Text) & ","
        sql = sql & XN(CboGasto.ItemData(CboGasto.ListIndex)) & ","
        sql = sql & XN(CboBanco.ItemData(CboBanco.ListIndex)) & ","
        sql = sql & XS(cboCtaBancaria.List(cboCtaBancaria.ListIndex)) & ","
        sql = sql & XN(txtImporte.Text) & ","
        If chkAplicoImpuesto.Value = Checked Then
            sql = sql & XS("S") & ")"
        Else
            sql = sql & XS("N") & ")"
        End If
        DBConn.Execute sql
    End If
    Rec.Close
    DBConn.CommitTrans
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    CmdNuevo_Click
    Exit Sub
    
CLAVOSE:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
End Sub

Private Sub CmdNuevo_Click()
    TabTB.Tab = 0
    TxtCodigo.Text = ""
    CboGasto.ListIndex = 0
    CboBanco.ListIndex = 0
    cboCtaBancaria.Clear
    chkAplicoImpuesto.Value = Unchecked
    txtImporte.Text = ""
    lblEstado.Caption = ""
    FechaGasto.Text = ""
    GrdModulos.Rows = 1
    TxtCodigo.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
    Set frmIngresoGastosBancarios = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'si preciona f1 voy a la busqueda
    If KeyCode = vbKeyF1 Then TabTB.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'si presiono ESCAPE salgo del form
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
    'If KeyAscii = vbKeyReturn And _
        Me.ActiveControl.Name <> "TxtDescriB" And _
        Me.ActiveControl.Name <> "GrdContactos" Then  'avanza de campo
    If KeyAscii = vbKeyReturn Then
            MySendKeys Chr(9)
            KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Set Rec = New ADODB.Recordset
    'CARGO COMBO BANCO
    CargoComboBanco
    'CARGO COMBO GASTOS
    CargoComboGasto
    
    lblEstado.Caption = ""
    cmdGrabar.Enabled = True
    cmdNuevo.Enabled = True
    cmdBorrar.Enabled = False
    
    GrdModulos.FormatString = "^Fecha|>Importe|Gasto|Banco|numero"
                    
    GrdModulos.ColWidth(0) = 1100
    GrdModulos.ColWidth(1) = 1000
    GrdModulos.ColWidth(2) = 3000
    GrdModulos.ColWidth(3) = 3000
    GrdModulos.ColWidth(4) = 0
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
    GrdModulos.BorderStyle = flexBorderNone
    GrdModulos.row = 0
    For I = 0 To GrdModulos.Cols - 1
        GrdModulos.Col = I
        GrdModulos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdModulos.CellBackColor = &H808080    'GRIS OSCURO
        GrdModulos.CellFontBold = True
    Next
    
    TabTB.Tab = 0
    
    Screen.MousePointer = vbNormal

    Me.Left = 0
    Me.Top = 0
End Sub

Private Sub CargoComboBanco()
    sql = "SELECT B.BAN_DESCRI, B.BAN_CODINT"
    sql = sql & " FROM BANCO B, CTA_BANCARIA CB"
    sql = sql & " WHERE B.BAN_CODINT=CB.BAN_CODINT"
    sql = sql & " ORDER BY B.BAN_DESCRI"
    
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Rec.MoveFirst
            Me.cboBanco1.AddItem "<Todos>"
        Do While Not Rec.EOF
            Me.CboBanco.AddItem Trim(Rec!BAN_DESCRI)
            Me.CboBanco.ItemData(Me.CboBanco.NewIndex) = Rec!BAN_CODINT
            Me.cboBanco1.AddItem Trim(Rec!BAN_DESCRI)
            Me.cboBanco1.ItemData(Me.cboBanco1.NewIndex) = Rec!BAN_CODINT
            Rec.MoveNext
        Loop
        Me.CboBanco.ListIndex = 0
        Me.cboBanco1.ListIndex = 0
    End If
    Rec.Close
End Sub

Private Sub CargoComboGasto()
    sql = "SELECT TGB_CODIGO, TGB_DESCRI"
    sql = sql & " FROM TIPO_GASTO_BANCARIO"
    sql = sql & " ORDER BY TGB_DESCRI"
    
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Rec.MoveFirst
            Me.cboGasto1.AddItem "<Todos>"
        Do While Not Rec.EOF
            Me.CboGasto.AddItem Trim(Rec!TGB_DESCRI)
            Me.CboGasto.ItemData(Me.CboGasto.NewIndex) = Rec!TGB_CODIGO
            Me.cboGasto1.AddItem Trim(Rec!TGB_DESCRI)
            Me.cboGasto1.ItemData(Me.cboGasto1.NewIndex) = Rec!TGB_CODIGO
            Rec.MoveNext
        Loop
        Me.CboGasto.ListIndex = 0
        Me.CboGasto.ListIndex = 0
    End If
    Rec.Close
End Sub

Private Sub GrdModulos_dblClick()
    If GrdModulos.Rows > 1 Then
        'paso el item seleccionado al tab 'DATOS'
        TxtCodigo.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 4)
        TxtCodigo_LostFocus
        TabTB.Tab = 0
    End If
End Sub

Private Sub GrdModulos_GotFocus()
    GrdModulos.Col = 0
    GrdModulos.ColSel = 1
    GrdModulos.HighLight = flexHighlightAlways
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cmdBorrar_Click
    If KeyCode = vbKeyReturn Then GrdModulos_dblClick
End Sub

Private Sub GrdModulos_LostFocus()
    GrdModulos.HighLight = flexHighlightNever
End Sub

Private Sub mFechaD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        MySendKeys Chr(9)
    End If
End Sub

Private Sub mFechaH_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        MySendKeys Chr(9)
    End If
End Sub

Private Sub tabTB_Click(PreviousTab As Integer)
    'Si cambio de 'Pestaña' en el tab
    'pongo el foco en el primer campo de la misma
    'If TabTB.Tab = 0 And Me.Visible Then TxtDescrip.SetFocus
    If TabTB.Tab = 1 Then
        GrdModulos.Rows = 1
        GrdModulos.HighLight = flexHighlightNever
        cboGasto1.ListIndex = 0
        cboBanco1.ListIndex = 0
        mFechaD.Text = ""
        mFechaH.Text = ""
        If cboGasto1.Enabled Then cboGasto1.SetFocus
    Else
        If FechaGasto.Visible = True Then FechaGasto.SetFocus
    End If
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto TxtCodigo
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    If Trim(TxtCodigo.Text) <> "" Then ' si no viene vacio
        BuscoDatos
    Else
        cmdGrabar.Enabled = True
        cmdNuevo.Enabled = True
        cmdBorrar.Enabled = False
    End If
End Sub

Private Sub TxtCodigo_Change()
    If Trim(TxtCodigo.Text) = "" And cmdBorrar.Enabled Then
        cmdBorrar.Enabled = False
    ElseIf Trim(TxtCodigo.Text) <> "" Then
        cmdBorrar.Enabled = True
    End If
End Sub

Private Sub txtImporte_GotFocus()
    SelecTexto txtImporte
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImporte, KeyAscii)
End Sub

Private Sub txtImporte_LostFocus()
    If txtImporte.Text = "" Then
        txtImporte.Text = "0,00"
    Else
        txtImporte.Text = Valido_Importe(txtImporte.Text)
    End If
End Sub
