VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDebitosManuales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Débitos Manuales"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6885
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
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker fecha1 
      Height          =   315
      Left            =   1000
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   52297729
      CurrentDate     =   42839
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   390
      Left            =   4140
      TabIndex        =   10
      Top             =   3270
      Width           =   1300
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   60
      TabIndex        =   26
      Top             =   3105
      Width           =   6795
   End
   Begin VB.OptionButton optDeporte 
      Caption         =   "Genera Deporte"
      Height          =   225
      Left            =   4260
      TabIndex        =   8
      Top             =   2670
      Width           =   2085
   End
   Begin VB.OptionButton optCuota 
      Caption         =   "Genera Cuota"
      Height          =   225
      Left            =   4260
      TabIndex        =   7
      Top             =   2340
      Width           =   2085
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   60
      TabIndex        =   25
      Top             =   2145
      Width           =   6780
   End
   Begin VB.ComboBox cboDeporte 
      Height          =   315
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2640
      Width           =   2790
   End
   Begin VB.ComboBox cboTipoCuota 
      Height          =   315
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2295
      Width           =   2790
   End
   Begin VB.TextBox txtMes 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   5130
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1680
      Width           =   645
   End
   Begin VB.TextBox txtAno 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   5805
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   2805
      TabIndex        =   9
      Top             =   3270
      Width           =   1300
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   5475
      TabIndex        =   11
      Top             =   3270
      Width           =   1300
   End
   Begin VB.Frame FrameSocio 
      Caption         =   "Socio..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   105
      TabIndex        =   12
      Top             =   15
      Width           =   6720
      Begin VB.TextBox txtDomici 
         Enabled         =   0   'False
         Height          =   330
         Left            =   915
         TabIndex        =   16
         Top             =   690
         Width           =   5520
      End
      Begin VB.TextBox txtNroDoc 
         Enabled         =   0   'False
         Height          =   330
         Left            =   915
         TabIndex        =   15
         Top             =   1035
         Width           =   1770
      End
      Begin VB.TextBox txtTelefono 
         Enabled         =   0   'False
         Height          =   330
         Left            =   4665
         TabIndex        =   14
         Top             =   1035
         Width           =   1770
      End
      Begin VB.CommandButton cmdBuscaSocio 
         Height          =   315
         Left            =   6105
         Picture         =   "frmDebitosManuales.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Buscar Socio"
         Top             =   345
         Width           =   330
      End
      Begin VB.TextBox txtNomSoc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1785
         TabIndex        =   2
         Top             =   345
         Width           =   4275
      End
      Begin VB.TextBox txtCodSoc 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   915
         TabIndex        =   0
         Top             =   345
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio:"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   735
         Width           =   660
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Nro Doc:"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono:"
         Height          =   195
         Left            =   3915
         TabIndex        =   18
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Socio:"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   375
         Width           =   435
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Deporte:"
      Height          =   195
      Left            =   105
      TabIndex        =   24
      Top             =   2685
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Cuota:"
      Height          =   195
      Left            =   105
      TabIndex        =   23
      Top             =   2340
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   195
      Left            =   105
      TabIndex        =   22
      Top             =   1725
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Periodo (mm/yyyy):"
      Height          =   195
      Left            =   3660
      TabIndex        =   21
      Top             =   1725
      Width           =   1425
   End
End
Attribute VB_Name = "frmDebitosManuales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim Rec As ADODB.Recordset
Dim Rec1 As ADODB.Recordset

Private Sub cmdAceptar_Click()
    If txtCodSoc.Text = "" Then
        MsgBox "Falta Ingresar el Socios", vbCritical, TIT_MSGBOX
        txtCodSoc.SetFocus
        Exit Sub
    End If
    If fecha1.Value = "" Then
        MsgBox "Falta Ingresar la Fecha", vbCritical, TIT_MSGBOX
        fecha1.SetFocus
        Exit Sub
    End If
    If txtMes.Text = "" Then
        MsgBox "Falta Ingresar el Mes del Débito", vbCritical, TIT_MSGBOX
        txtMes.SetFocus
        Exit Sub
    End If
    If txtAno.Text = "" Then
        MsgBox "Falta Ingresar el Año del Débito", vbCritical, TIT_MSGBOX
        txtAno.SetFocus
        Exit Sub
    End If
    If optCuota.Value = False And optDeporte.Value = False Then
        MsgBox "Falta Ingresar el Débito que quiere Generar", vbCritical, TIT_MSGBOX
        optCuota.SetFocus
        Exit Sub
    End If
    'CONTROLO QUE NO SE HAYA GENERADO YA EL DEBITO
    sql = "SELECT DEB_MES"
    sql = sql & " FROM DEBITOS"
    sql = sql & " WHERE DEB_MES = " & XN(txtMes.Text)
    sql = sql & " AND DEB_ANO = " & XN(txtAno.Text)
    sql = sql & " AND SOC_CODIGO = " & XN(txtCodSoc.Text)
    If optCuota.Value = True Then
        sql = sql & " AND TIC_CODIGO = " & cboTipoCuota.ItemData(cboTipoCuota.ListIndex)
    ElseIf optDeporte.Value = True Then
        sql = sql & " AND DEP_CODIGO = " & cboDeporte.ItemData(cboDeporte.ListIndex)
    End If
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        MsgBox "El Periodo ingresado ya fue Generado", vbExclamation, TIT_MSGBOX
        Rec.Close
        txtMes.SetFocus
        Exit Sub
    End If
    Rec.Close
    
    If MsgBox("¿Genera Débito?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        
        sql = "SELECT MAX(DEB_ITEM) AS NUMERITO FROM DEBITOS"
        sql = sql & " WHERE DEB_MES = " & XN(txtMes.Text)
        sql = sql & " AND DEB_ANO = " & XN(txtAno.Text)
        sql = sql & " AND SOC_CODIGO = " & XN(txtCodSoc.Text)
        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec.EOF = False Then
            I = CInt(Chk0(Rec!NUMERITO)) + 1
        Else
            I = 1
        End If
        Rec.Close
        
        If optCuota.Value = True Then
            'BUSCO LOS TITULARES POR CUOTA SOCIAL
            sql = "SELECT SOC_CODIGO, T.TIC_DESCRI, T.TIC_CUOTA, T.TIC_CODIGO"
            sql = sql & " FROM SOCIOS S, TIPO_CUOTA T"
            sql = sql & " WHERE "
            sql = sql & " T.TIC_CODIGO=S.TIC_CODIGO"
            sql = sql & " AND T.TIC_CODIGO=" & cboTipoCuota.ItemData(cboTipoCuota.ListIndex)
            sql = sql & " AND SOC_CODIGO=" & XN(txtCodSoc.Text)
            Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec.EOF = False Then
                sql = "INSERT INTO DEBITOS (DEB_MES,DEB_ANO,SOC_CODIGO,"
                sql = sql & " DEB_ITEM,DEB_DETALLE,DEB_IMPORTE,DEB_SALDO,DEB_FECHA,TIC_CODIGO)"
                sql = sql & " VALUES ("
                sql = sql & XN(txtMes.Text) & ","
                sql = sql & XN(txtAno.Text) & ","
                sql = sql & XN(Rec!SOC_CODIGO) & ","
                sql = sql & I & "," & XS("CUOTA " & Rec!TIC_DESCRI) & ","
                sql = sql & XN(Rec!TIC_CUOTA) & ","
                sql = sql & XN(Rec!TIC_CUOTA) & ","
                sql = sql & XDQ(fecha1.Value) & ","
                sql = sql & XN(Rec!TIC_CODIGO) & ")"
                DBConn.Execute sql
            Else
                MsgBox "El Socio No genera el Tipo de Cuota Seleccionado", vbCritical, TIT_MSGBOX
                Rec.Close
                Exit Sub
            End If
            Rec.Close
            
        ElseIf optDeporte.Value = True Then
               
            'BUSCO EN LOS SOCIOS LOS DEPORTES QUE REALIZAN PARA COBRAR LA CUOTA DE C/DEPORTE
            sql = "SELECT S.SOC_CODIGO, D.DEP_DESCRI, D.DEP_CUOTA, D.DEP_CODIGO"
            sql = sql & " FROM SOCIOS S, SOCIOS_DEPORTES SD, DEPORTE D"
            sql = sql & " WHERE "
            sql = sql & " D.DEP_CODIGO=SD.DEP_CODIGO"
            sql = sql & " AND D.DEP_CODIGO=SD.DEP_CODIGO"
            sql = sql & " AND D.DEP_CODIGO=" & cboDeporte.ItemData(cboDeporte.ListIndex)
            sql = sql & " AND S.SOC_CODIGO=" & XN(txtCodSoc.Text)
            Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec.EOF = False Then
                sql = "INSERT INTO DEBITOS (DEB_MES,DEB_ANO,SOC_CODIGO,"
                sql = sql & " DEB_ITEM,DEB_DETALLE,DEB_IMPORTE,DEB_SALDO,DEB_FECHA,DEP_CODIGO)"
                sql = sql & " VALUES ("
                sql = sql & XN(txtMes.Text) & ","
                sql = sql & XN(txtAno.Text) & ","
                sql = sql & XN(Rec!SOC_CODIGO) & ","
                sql = sql & I & "," & XS(Rec!DEP_DESCRI) & ","
                sql = sql & XN(Rec!DEP_CUOTA) & ","
                sql = sql & XN(Rec!DEP_CUOTA) & ","
                sql = sql & XDQ(fecha1.Value) & ","
                sql = sql & XN(Rec!DEP_CODIGO) & ")"
                DBConn.Execute sql
            Else
                MsgBox "El Socio no esta Registrado en el Deporte Seleccionado", vbCritical, TIT_MSGBOX
                Rec.Close
                Exit Sub
            End If
            Rec.Close
        End If
        MsgBox "Débito Generado", vbExclamation, TIT_MSGBOX
        CmdNuevo_Click
    End If
End Sub

Private Sub cmdBuscaSocio_Click()
    txtCodSoc.Text = ""
    BuscarSocios "txtCodSoc", "CODIGO"
    txtNomSoc.SetFocus
End Sub

Private Sub cmdCerrar_Click()
    Set frmDebitosManuales = Nothing
    Unload Me
End Sub

Private Sub CmdNuevo_Click()
    txtCodSoc.Text = ""
    fecha1.Value = Date
    txtMes.Text = ""
    txtAno.Text = ""
    cboTipoCuota.ListIndex = 0
    cboDeporte.ListIndex = 0
    optCuota.Value = False
    optDeporte.Value = False
    txtCodSoc.SetFocus
End Sub

Private Sub fecha1_LostFocus()
    If fecha1.Value = "" Then
        fecha1.Value = Date
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    Set Rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Centrar_pantalla Me
    
    cboTipoCuota.Clear
    Call CargoComboBox(cboTipoCuota, "TIPO_CUOTA", "TIC_CODIGO", "TIC_DESCRI")
    cboTipoCuota.ListIndex = 0
    
    cboDeporte.Clear
    Call CargoComboBox(cboDeporte, "DEPORTE", "DEP_CODIGO", "DEP_DESCRI")
    cboDeporte.ListIndex = 0
End Sub

Private Sub txtCodSoc_Change()
    If txtCodSoc.Text = "" Then
        txtNomSoc.Text = ""
        txtDomici.Text = ""
        txtNroDoc.Text = ""
        txtTelefono.Text = ""
    End If
End Sub

Private Sub txtCodSoc_GotFocus()
    SelecTexto txtCodSoc
End Sub

Private Sub txtCodSoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        txtCodSoc.Text = ""
        BuscarSocios "txtCodSoc", "CODIGO"
    End If
End Sub

Private Sub txtCodSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodSoc_LostFocus()
    If txtCodSoc.Text <> "" Then
        sql = "SELECT SOC_CODIGO, SOC_NOMBRE, SOC_DOMICI, SOC_NRODOC, SOC_TELEFONO"
        sql = sql & " FROM SOCIOS"
        sql = sql & " WHERE "
        sql = sql & " SOC_CODIGO =" & XN(txtCodSoc.Text)
        'sql = sql & " AND TIS_CODIGO=1"
        If Rec.State = 1 Then Rec.Close
        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec.EOF = False Then
            txtNomSoc.Text = ChkNull(Rec!SOC_NOMBRE)
            txtDomici.Text = ChkNull(Rec!SOC_DOMICI)
            txtNroDoc.Text = ChkNull(Rec!SOC_NRODOC)
            txtTelefono.Text = ChkNull(Rec!SOC_TELEFONO)
        Else
            MsgBox "El Código no existe", vbInformation
            txtNomSoc.Text = ""
            txtCodSoc.Text = ""
            txtCodSoc.SetFocus
        End If
        If Rec.State = 1 Then Rec.Close
    End If
End Sub

Private Sub txtAno_GotFocus()
    SelecTexto txtAno
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtAno_LostFocus()
    If txtAno.Text <> "" Then
        If Len(txtAno.Text) = 2 Then
            txtAno.Text = Mid(Year(Date), 1, 2) & txtAno.Text
        Else
            txtAno.Text = Format(txtAno.Text, "0000")
        End If
    End If
End Sub

Private Sub txtMes_GotFocus()
    SelecTexto txtMes
End Sub

Private Sub txtMes_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtMes_LostFocus()
    If txtMes.Text <> "" Then
        txtMes.Text = Format(txtMes.Text, "00")
    End If
End Sub

Private Sub txtNomSoc_Change()
    If txtNomSoc.Text = "" Then
        txtCodSoc.Text = ""
        txtDomici.Text = ""
        txtNroDoc.Text = ""
        txtTelefono.Text = ""
    End If
End Sub

Private Sub txtNomSoc_GotFocus()
    SelecTexto txtNomSoc
End Sub

Private Sub txtNomSoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarSocios "txtcodCli", "CODIGO"
    End If
End Sub

Private Sub txtNomSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtNomSoc_LostFocus()
    If txtCodSoc.Text = "" And txtNomSoc.Text <> "" Then
        sql = "SELECT SOC_CODIGO,SOC_NOMBRE, SOC_DOMICI, SOC_NRODOC, SOC_TELEFONO"
        sql = sql & " FROM SOCIOS"
        sql = sql & " WHERE "
        sql = sql & " SOC_NOMBRE LIKE '" & XN(Trim(txtNomSoc.Text)) & "%'"
        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec.EOF = False Then
            If Rec.RecordCount > 1 Then
                BuscarSocios "txtCodSoc", "CADENA", Trim(txtNomSoc.Text)
                If Rec.State = 1 Then Rec.Close
                txtNomSoc.SetFocus
            Else
                txtCodSoc.Text = Rec!SOC_CODIGO
                txtNomSoc.Text = Rec!SOC_NOMBRE
                txtDomici.Text = ChkNull(Rec!SOC_DOMICI)
                txtNroDoc.Text = ChkNull(Rec!SOC_NRODOC)
                txtTelefono.Text = ChkNull(Rec!SOC_TELEFONO)
            End If
        Else
            MsgBox "El Socio no existe", vbExclamation, TIT_MSGBOX
            txtCodSoc.Text = ""
            txtNomSoc.SetFocus
        End If
        If Rec.State = 1 Then Rec.Close
    End If
End Sub

Public Sub BuscarSocios(Txt As String, mQuien As String, Optional mCadena As String)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim I, posicion As Integer
    Dim cadena As String
    
    Set B = New CBusqueda
    With B
        cSQL = "SELECT SOC_NOMBRE, SOC_DOMICI, SOC_CODIGO"
        cSQL = cSQL & " FROM SOCIOS"
        If mQuien = "CADENA" Then
            cSQL = cSQL & " WHERE SOC_NOMBRE LIKE '" & Trim(mCadena) & "%'"
            'cSQL = cSQL & " AND TIS_CODIGO=1"
        'Else
            'cSQL = cSQL & " WHERE TIS_CODIGO=1"
        End If
        
        
        hSQL = "Nombre, Domicilio, Código"
        .sql = cSQL
        .Headers = hSQL
        .Field = "SOC_NOMBRE"
        campo1 = .Field
        .Field = "SOC_DOMICI"
        campo2 = .Field
        .Field = "SOC_CODIGO"
        campo3 = .Field
        .OrderBy = "SOC_NOMBRE"
        camponumerico = False
        .Titulo = "Busqueda de Socios :"
        .MaxRecords = 1
        .Show

        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
            If Txt = "txtCodSoc" Then
                txtCodSoc.Text = .ResultFields(3)
                txtCodSoc_LostFocus
            End If
        End If
    End With
    
    Set B = Nothing
End Sub

