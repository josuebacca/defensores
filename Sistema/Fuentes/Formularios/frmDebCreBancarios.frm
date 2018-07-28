VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDebCreBancarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Débitos y Créditos Bancarios"
   ClientHeight    =   4500
   ClientLeft      =   1620
   ClientTop       =   1950
   ClientWidth     =   7005
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
   ScaleHeight     =   4500
   ScaleWidth      =   7005
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   4155
      TabIndex        =   8
      Top             =   4035
      Width           =   915
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   3225
      TabIndex        =   7
      Top             =   4035
      Width           =   915
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   6015
      TabIndex        =   10
      Top             =   4035
      Width           =   915
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      Height          =   435
      Left            =   5085
      TabIndex        =   9
      Top             =   4035
      Width           =   915
   End
   Begin TabDlg.SSTab TabTB 
      Height          =   3960
      Left            =   30
      TabIndex        =   15
      Top             =   30
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   6985
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
      TabPicture(0)   =   "frmDebCreBancarios.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "B&uscar"
      TabPicture(1)   =   "frmDebCreBancarios.frx":001C
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
         TabIndex        =   18
         Top             =   360
         Width           =   6570
         Begin VB.ComboBox cboBanco1 
            Height          =   315
            Left            =   1185
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   540
            Width           =   4005
         End
         Begin VB.ComboBox cboGasto1 
            Height          =   315
            Left            =   1185
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   195
            Width           =   4005
         End
         Begin VB.CommandButton CmdBuscAprox 
            Caption         =   "Buscar"
            Height          =   375
            Left            =   5400
            MaskColor       =   &H00404040&
            TabIndex        =   13
            ToolTipText     =   "Buscar"
            Top             =   825
            UseMaskColor    =   -1  'True
            Width           =   1065
         End
         Begin MSComCtl2.DTPicker mFechaD 
            Height          =   315
            Left            =   1185
            TabIndex        =   32
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852993
            CurrentDate     =   43307
         End
         Begin MSComCtl2.DTPicker mFechaH 
            Height          =   315
            Left            =   3960
            TabIndex        =   33
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852993
            CurrentDate     =   43307
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   75
            TabIndex        =   28
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Deb./Cre.:"
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   27
            Top             =   255
            Width           =   780
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   2955
            TabIndex        =   23
            Top             =   930
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   75
            TabIndex        =   22
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
         Height          =   3420
         Left            =   -74865
         TabIndex        =   16
         Top             =   420
         Width           =   6615
         Begin VB.ComboBox cboDebCre 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1215
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   810
            Width           =   1410
         End
         Begin VB.CheckBox chkAplicoImpuesto 
            Caption         =   "Aplicar impuesto transacciones financieras"
            Height          =   240
            Left            =   1215
            TabIndex        =   6
            Top             =   2970
            Width           =   4005
         End
         Begin VB.ComboBox CboBanco 
            Height          =   315
            Left            =   1215
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1860
            Width           =   4395
         End
         Begin VB.ComboBox CboGasto 
            Height          =   315
            Left            =   1215
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1500
            Width           =   4395
         End
         Begin VB.ComboBox cboCtaBancaria 
            Height          =   315
            Left            =   1215
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2220
            Width           =   2100
         End
         Begin VB.TextBox txtImporte 
            Height          =   330
            Left            =   1215
            TabIndex        =   5
            Top             =   2580
            Width           =   1125
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   315
            Left            =   1215
            TabIndex        =   0
            Top             =   435
            Width           =   1125
         End
         Begin MSComCtl2.DTPicker FechaGasto 
            Height          =   315
            Left            =   1215
            TabIndex        =   31
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   112852993
            CurrentDate     =   43307
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Deb/Cre:"
            Height          =   195
            Index           =   5
            Left            =   165
            TabIndex        =   30
            Top             =   855
            Width           =   660
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
            TabIndex        =   29
            Top             =   15
            Width           =   1875
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   165
            TabIndex        =   26
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro Cuenta:"
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   25
            Top             =   2250
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Index           =   4
            Left            =   165
            TabIndex        =   21
            Top             =   2625
            Width           =   630
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   165
            TabIndex        =   20
            Top             =   1215
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Ingreso:"
            Height          =   195
            Index           =   3
            Left            =   165
            TabIndex        =   19
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Index           =   2
            Left            =   165
            TabIndex        =   17
            Top             =   1545
            Width           =   360
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   2100
         Left            =   180
         TabIndex        =   14
         Top             =   1650
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   3704
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
      Left            =   120
      TabIndex        =   24
      Top             =   4095
      Width           =   660
   End
End
Attribute VB_Name = "frmDebCreBancarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer

Private Sub BuscoDatos()
    'Set rec = New ADODB.Recordset
    sql = "SELECT * FROM DEBCRE_BANCARIOS"
    sql = sql & " WHERE DCB_NUMERO = " & XN(TxtCodigo.Text)
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then ' si existe
        If Trim(Rec!DCB_TIPO) = "D" Then
            cboDebCre.ListIndex = 1
            cboDebCre_LostFocus
        Else
            cboDebCre.ListIndex = 0
            cboDebCre_LostFocus
        End If
        FechaGasto.Value = ChkNull(Rec!DCB_FECHA)
        Call BuscaCodigoProxItemData(CInt(Rec!TDCB_CODIGO), CboGasto)
        Call BuscaCodigoProxItemData(CInt(Rec!BAN_CODINT), CboBanco)
        CboBanco_LostFocus
        Call BuscaProx(Rec!CTA_NROCTA, cboCtaBancaria)
        txtImporte.Text = Valido_Importe(ChkNull(Rec!DCB_IMPORTE))
        If Rec!DCB_IMPUESTO = "S" Then
            chkAplicoImpuesto.Value = Checked
        ElseIf Rec!DCB_IMPUESTO = "N" Then
            chkAplicoImpuesto.Value = Unchecked
        End If
'    Else
'        MsgBox "Débito o Crédito Inexistente", vbCritical
'        TxtCodigo.Text = ""
'        TxtCodigo.SetFocus
'        rec.Close
'        Exit Sub
    End If
    Rec.Close
End Sub

Private Sub CboBanco_LostFocus()
    If CboBanco.ListIndex <> -1 Then
        'Set Rec1 = New ADODB.Recordset
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

Private Sub cboDebCre_LostFocus()
    If Left(cboDebCre.List(cboDebCre.ListIndex), 1) = "D" Then
        CboGasto.Clear
        sql = "SELECT * FROM TIPO_DEBCRE_BANCARIO"
        sql = sql & " WHERE TDCB_TIPO='D'"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            Do While Rec1.EOF = False
                CboGasto.AddItem Rec1!TDCB_DESCRI
                CboGasto.ItemData(CboGasto.NewIndex) = Rec1!TDCB_CODIGO
                Rec1.MoveNext
            Loop
            CboGasto.ListIndex = 0
        End If
        Rec1.Close
    Else
        CboGasto.Clear
        sql = "SELECT * FROM TIPO_DEBCRE_BANCARIO"
        sql = sql & " WHERE TDCB_TIPO='C'"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            Do While Rec1.EOF = False
                CboGasto.AddItem Rec1!TDCB_DESCRI
                CboGasto.ItemData(CboGasto.NewIndex) = Rec1!TDCB_CODIGO
                Rec1.MoveNext
            Loop
            CboGasto.ListIndex = 0
        End If
        Rec1.Close
'        If Viene_Form = True Then
'            Call BuscaCodigoProxItemData(5, CboGasto)
'        End If
    End If
End Sub

Private Sub cmdBorrar_Click()
    On Error GoTo CLAVOSE
    If Trim(TxtCodigo.Text) <> "" Then
        If MsgBox("¿Seguro desea eliminar?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            Screen.MousePointer = vbHourglass
            lblEstado.Caption = "Eliminando ..."
            DBConn.BeginTrans
            DBConn.Execute "DELETE FROM DEBCRE_BANCARIOS WHERE DCB_NUMERO = " & XN(TxtCodigo.Text)
            DBConn.CommitTrans
            FechaGasto.SetFocus
            
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            CmdNuevo_Click
            TxtCodigo.SetFocus
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
    'Set rec = New ADODB.Recordset
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    Me.Refresh
    sql = "SELECT GB.DCB_NUMERO, GB.DCB_FECHA, GB.DCB_IMPORTE,"
    sql = sql & " TG.TDCB_DESCRI, B.BAN_DESCRI"
    sql = sql & " FROM DEBCRE_BANCARIOS GB, TIPO_DEBCRE_BANCARIO TG,"
    sql = sql & " BANCO B"
    sql = sql & " WHERE"
    sql = sql & " GB.TDCB_CODIGO=TG.TDCB_CODIGO"
    sql = sql & " AND GB.BAN_CODINT=B.BAN_CODINT"
    If cboBanco1.List(cboBanco1.ListIndex) <> "(Todos)" Then
        sql = sql & " AND GB.BAN_CODINT=" & XN(cboBanco1.ItemData(cboBanco1.ListIndex))
    End If
    If cboGasto1.List(cboGasto1.ListIndex) <> "(Todos)" Then
        sql = sql & " AND GB.TDCB_CODIGO=" & XN(cboGasto1.ItemData(cboGasto1.ListIndex))
    End If
    If mFechaD.Value <> "" Then sql = sql & " AND DCB_FECHA >= " & XDQ(mFechaD.Value)
    If mFechaH.Value <> "" Then sql = sql & " AND DCB_FECHA <= " & XDQ(mFechaH.Value)
    
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Rec.MoveFirst
        Do While Not Rec.EOF
            GrdModulos.AddItem Rec!DCB_FECHA & Chr(9) & Valido_Importe(Rec!DCB_IMPORTE) & Chr(9) & _
                        Rec!TDCB_DESCRI & Chr(9) & Rec!BAN_DESCRI & Chr(9) & Rec!DCB_NUMERO
            Rec.MoveNext
        Loop
        GrdModulos.HighLight = flexHighlightAlways
        If GrdModulos.Enabled Then GrdModulos.SetFocus
        lblEstado.Caption = ""
    Else
        lblEstado.Caption = ""
        MsgBox "No se encontraron items con este Criterio", vbExclamation, TIT_MSGBOX
        If mFechaD.Enabled Then mFechaD.SetFocus
    End If
    lblEstado.Caption = ""
    Rec.Close
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdGrabar_Click()
    On Error GoTo CLAVOSE
    
    If TxtCodigo.Text = "" Then
        MsgBox "No ha ingresado el Número del Débito y/o Crédito", vbCritical, TIT_MSGBOX
        TxtCodigo.SetFocus
        Exit Sub
    End If
    If FechaGasto.Value = "" Then
        MsgBox "No ha ingresado la Fecha", vbExclamation, TIT_MSGBOX
        FechaGasto.SetFocus
        Exit Sub
    End If
    If cboCtaBancaria.List(cboCtaBancaria.ListIndex) = "" Then
        MsgBox "Debe elegir una Cuenta Bancaria", vbExclamation, TIT_MSGBOX
        CboBanco.SetFocus
        Exit Sub
    End If
    If txtImporte.Text = "" Then
        MsgBox "No ha ingresado el Importe", vbExclamation, TIT_MSGBOX
        If txtImporte.Enabled Then txtImporte.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando ..."
    
    'Set rec = New ADODB.Recordset
    
    DBConn.BeginTrans
    
    sql = "SELECT DCB_FECHA FROM DEBCRE_BANCARIOS WHERE DCB_NUMERO = " & XN(TxtCodigo.Text)
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If Rec.EOF = False Then
        sql = "UPDATE DEBCRE_BANCARIOS"
        sql = sql & " SET DCB_FECHA = " & XDQ(FechaGasto.Value)
        sql = sql & " ,TDCB_CODIGO = " & XN(CboGasto.ItemData(CboGasto.ListIndex))
        sql = sql & " ,BAN_CODINT = " & XN(CboBanco.ItemData(CboBanco.ListIndex))
        sql = sql & " ,CTA_NROCTA = " & XS(cboCtaBancaria.List(cboCtaBancaria.ListIndex))
        sql = sql & " ,DCB_IMPORTE = " & XN(txtImporte.Text)
        sql = sql & " ,DCB_TIPO=" & XS(Left(cboDebCre.List(cboDebCre.ListIndex), 1))
'        If optDebito.Value = True Then
'            sql = sql & " ,DCB_TIPO='D'"
'        Else
'            sql = sql & " ,DCB_TIPO='C'"
'        End If
        If chkAplicoImpuesto.Value = Checked Then
            sql = sql & " ,DCB_IMPUESTO=" & XS("S")
        Else
            sql = sql & " ,DCB_IMPUESTO=" & XS("N")
        End If
        sql = sql & " WHERE DCB_NUMERO = " & XN(TxtCodigo.Text)
        
        DBConn.Execute sql
    Else
        
        sql = "INSERT INTO DEBCRE_BANCARIOS"
        sql = sql & " (DCB_NUMERO, DCB_FECHA, TDCB_CODIGO, BAN_CODINT,"
        sql = sql & " CTA_NROCTA, DCB_IMPORTE, DCB_IMPUESTO,DCB_TIPO)"
        sql = sql & " VALUES ("
        sql = sql & XN(TxtCodigo.Text) & ","
        sql = sql & XDQ(FechaGasto.Value) & ","
        sql = sql & XN(CboGasto.ItemData(CboGasto.ListIndex)) & ","
        sql = sql & XN(CboBanco.ItemData(CboBanco.ListIndex)) & ","
        sql = sql & XS(cboCtaBancaria.List(cboCtaBancaria.ListIndex)) & ","
        sql = sql & XN(txtImporte.Text) & ","
        If chkAplicoImpuesto.Value = Checked Then
            sql = sql & XS("S") & ","
        Else
            sql = sql & XS("N") & ","
        End If
        sql = sql & XS(Left(cboDebCre.List(cboDebCre.ListIndex), 1)) & ")"
'        If optDebito.Value = True Then
'            sql = sql & "'D')"
'        Else
'            sql = sql & "'C')"
'        End If
        DBConn.Execute sql
    End If
    Rec.Close
    DBConn.CommitTrans
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    'If Viene_Form = False Then
        CmdNuevo_Click
    'Else
    '    CmdSalir_Click
    '    frmReciboCliente.GrillaComp.HighLight = flexHighlightAlways
    '    For i = 1 To frmReciboCliente.GrillaComp.Rows - 1
    '        frmReciboCliente.txtTotalComprobante.Text = CDbl(Chk0(frmReciboCliente.txtTotalComprobante.Text)) + CDbl(frmReciboCliente.GrillaComp.TextMatrix(i, 3))
    '    Next
    '    frmReciboCliente.txtTotalComprobante.Text = Valido_Importe(frmReciboCliente.txtTotalComprobante.Text)
    '    frmReciboCliente.txtNroComprobantes.Text = ""
    '    frmReciboCliente.txtNroCompSuc.Text = ""
    '    frmReciboCliente.fechaComprobantes.Text = ""
    '    frmReciboCliente.cboComprobantes.SetFocus
    'End If
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
    FechaGasto.Value = Date
    GrdModulos.Rows = 1
    cboDebCre.ListIndex = 0
    TxtCodigo.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Set frmDebCreBancarios = Nothing
    Unload Me
End Sub

Private Sub FechaGasto_LostFocus()
    If FechaGasto.Value = "" Then FechaGasto.Value = Date
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
    Set Rec1 = New ADODB.Recordset
    
    'COMBO DEB/CRE
    cboDebCre.AddItem "Crédito"
    cboDebCre.AddItem "Débito"
    cboDebCre.ListIndex = 0
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
    
'    If Viene_Form = False Then
'        TabTB.Tab = 0
'        Me.Left = 0
'        Me.Top = 0
'    Else 'VIENE DEL RECIBO
'        TabTB.Tab = 0
'        optCredito.Value = True
'        optDebito.Enabled = False
'        optCredito.Enabled = False
'        optCredito_LostFocus
'        cmdNuevo.Enabled = False
'        cmdBorrar.Enabled = False
'        TxtCodigo.Text = frmReciboCliente.txtNroComprobantes.Text
'        TxtCodigo.Enabled = False
'        FechaGasto.Text = frmReciboCliente.fechaComprobantes.Text
'        FechaGasto.Enabled = False
'        Call BuscaCodigoProxItemData(5, CboGasto)
'        CboGasto.Enabled = False
'        txtImporte.Text = Valido_Importe(frmReciboCliente.txtImporteComprobante.Text)
'        txtImporte.Enabled = False
'    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub CargoComboBanco()
    sql = "SELECT DISTINCT B.BAN_CODINT, B.BAN_DESCRI"
    sql = sql & " FROM BANCO B, CTA_BANCARIA CB"
    sql = sql & " WHERE B.BAN_CODINT=CB.BAN_CODINT"
    sql = sql & " ORDER BY B.BAN_DESCRI"
    
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Rec.MoveFirst
            Me.cboBanco1.AddItem "(Todos)"
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
    sql = "SELECT TDCB_CODIGO, TDCB_DESCRI"
    sql = sql & " FROM TIPO_DEBCRE_BANCARIO"
    sql = sql & " ORDER BY TDCB_DESCRI"
    
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Rec.MoveFirst
            Me.cboGasto1.AddItem "(Todos)"
        Do While Not Rec.EOF
            Me.cboGasto1.AddItem Trim(Rec!TDCB_DESCRI)
            Me.cboGasto1.ItemData(Me.cboGasto1.NewIndex) = Rec!TDCB_CODIGO
            Rec.MoveNext
        Loop
        Me.cboGasto1.ListIndex = 0
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

Private Sub optCredito_LostFocus()
    
End Sub

Private Sub optDebito_LostFocus()
    
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
        mFechaD.Value = ""
        mFechaH.Value = ""
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
