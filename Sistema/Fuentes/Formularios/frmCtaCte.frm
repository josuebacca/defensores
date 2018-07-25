VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCtaCte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cta Cte - Socios"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
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
   ScaleHeight     =   1485
   ScaleWidth      =   6270
   Begin VB.Frame Frame1 
      Caption         =   "Ver"
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
      Left            =   2310
      TabIndex        =   11
      Top             =   675
      Width           =   2175
      Begin VB.ComboBox cboVer 
         Height          =   315
         ItemData        =   "frmCtaCte.frx":0000
         Left            =   135
         List            =   "frmCtaCte.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   270
         Width           =   1950
      End
   End
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
      Left            =   90
      TabIndex        =   6
      Top             =   675
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
         Picture         =   "frmCtaCte.frx":0004
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   10
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
         Picture         =   "frmCtaCte.frx":0106
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   9
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
         Picture         =   "frmCtaCte.frx":0208
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   8
         Top             =   315
         Width           =   240
      End
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "frmCtaCte.frx":030A
         Left            =   450
         List            =   "frmCtaCte.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   270
         Width           =   1635
      End
   End
   Begin VB.TextBox txtCodSocB 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   645
      TabIndex        =   4
      Top             =   165
      Width           =   855
   End
   Begin VB.TextBox txtNomSocB 
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
      Left            =   1515
      TabIndex        =   3
      Top             =   165
      Width           =   4275
   End
   Begin VB.CommandButton cmdBuscaSocioB 
      Height          =   315
      Left            =   5835
      Picture         =   "frmCtaCte.frx":0339
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Buscar Socio"
      Top             =   165
      Width           =   330
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   4830
      TabIndex        =   0
      Top             =   615
      Width           =   1300
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   4830
      TabIndex        =   1
      Top             =   1020
      Width           =   1300
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   4365
      Top             =   900
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Socio:"
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   210
      Width           =   435
   End
End
Attribute VB_Name = "frmCtaCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim Rec As ADODB.Recordset

'PARA EL PROGRES BAR
Dim Registro As Long
Dim Tamanio As Long
Dim I As Integer

Private Sub cboDestino_Click()
    picSalida(0).Visible = False
    picSalida(1).Visible = False
    picSalida(2).Visible = False
    picSalida(cboDestino.ListIndex).Visible = True
End Sub

Private Sub cmdAceptar_Click()
    sql = "DELETE FROM CTA_CTE_SOCIOS"
    DBConn.Execute sql
    
    If txtCodSocB.Text <> "" Then
        Call BuscarDeuda(txtCodSocB.Text, txtNomSocB.Text)
    Else
        sql = "SELECT SOC_CODIGO, SOC_NOMBRE"
        sql = sql & " FROM SOCIOS"
        sql = sql & " WHERE "
        sql = sql & " TIS_CODIGO = 1" 'SOLO TITULARES
        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec.EOF = False Then
            Do While Rec.EOF = False
                Call BuscarDeuda(Rec!SOC_CODIGO, Rec!SOC_NOMBRE)
                Rec.MoveNext
            Loop
        End If
        Rec.Close
    End If
    
    sql = "DELETE FROM TMP_RESUMEN_CUENTA_BANCO"
    DBConn.Execute sql
    
    
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Select Case cboDestino.ListIndex
        Case 0
            Rep.Destination = crptToWindow
        Case 1
            Rep.Destination = crptToPrinter
        Case 2
            Rep.Destination = crptToFile
    End Select
    
    Select Case cboVer.ListIndex
        Case 0
            Rep.WindowTitle = "Listado de Cta-Cte General"
            Rep.ReportFileName = DRIVE & DirReport & "CtaCteGeneral.rpt"
        Case 1
            Rep.WindowTitle = "Listado de Cta-Cte Detalle"
            Rep.ReportFileName = DRIVE & DirReport & "CtaCte.rpt"
    End Select
    
    Rep.Action = 1
End Sub

Private Sub cmdBuscaSocioB_Click()
    txtCodSocB.Text = ""
    BuscarSocios "txtCodSocB", "CODIGO"
    txtNomSocB.SetFocus
End Sub

Private Sub cmdCerrar_Click()
    Set frmCtaCte = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    If KeyAscii = 27 Then
        cmdCerrar_Click
    End If
End Sub

Private Sub Form_Load()
    Set Rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    
    cboVer.AddItem "General"
    cboVer.AddItem "Detalle"
    cboVer.ListIndex = 0
    
    cboDestino.ListIndex = 0
End Sub

Private Sub txtCodSocB_Change()
    If txtCodSocB.Text = "" Then
        txtNomSocB.Text = ""
    End If
End Sub

Private Sub txtCodSocB_GotFocus()
    SelecTexto txtCodSocB
End Sub

Private Sub txtCodSocB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        txtCodSocB.Text = ""
        BuscarSocios "txtCodSocB", "CODIGO"
    End If
End Sub

Private Sub txtCodSocB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodSocB_LostFocus()
    If txtCodSocB.Text <> "" Then
        sql = "SELECT SOC_CODIGO, SOC_NOMBRE"
        sql = sql & " FROM SOCIOS"
        sql = sql & " WHERE "
        sql = sql & " SOC_CODIGO =" & XN(txtCodSocB.Text)
        If Rec.State = 1 Then Rec.Close
        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec.EOF = False Then
            txtNomSocB.Text = ChkNull(Rec!SOC_NOMBRE)
        Else
            MsgBox "El Código no existe", vbInformation
            txtNomSocB.Text = ""
            txtCodSocB.Text = ""
            txtCodSocB.SetFocus
        End If
        If Rec.State = 1 Then Rec.Close
    End If
End Sub

Private Sub txtNomSocB_Change()
    If txtNomSocB.Text = "" Then
        txtCodSocB.Text = ""
    End If
End Sub

Private Sub txtNomSocB_GotFocus()
    SelecTexto txtNomSocB
End Sub

Private Sub txtNomSocB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarSocios "txtcodCliB", "CODIGO"
    End If
End Sub

Private Sub txtNomSocB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtNomSocB_LostFocus()
    If txtCodSocB.Text = "" And txtNomSocB.Text <> "" Then
        sql = "SELECT SOC_CODIGO,SOC_NOMBRE"
        sql = sql & " FROM SOCIOS"
        sql = sql & " WHERE "
        sql = sql & " SOC_NOMBRE LIKE '" & XN(Trim(txtNomSocB.Text)) & "%'"
        sql = sql & " AND TIS_CODIGO=1"
        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec.EOF = False Then
            If Rec.RecordCount > 1 Then
                BuscarSocios "txtCodSocB", "CADENA", Trim(txtNomSocB.Text)
                If Rec.State = 1 Then Rec.Close
                txtNomSocB.SetFocus
            Else
                txtCodSocB.Text = Rec!SOC_CODIGO
                txtNomSocB.Text = Rec!SOC_NOMBRE
            End If
        Else
            MsgBox "El Socio no existe", vbExclamation, TIT_MSGBOX
            txtCodSocB.Text = ""
            txtNomSocB.SetFocus
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
            cSQL = cSQL & " AND TIS_CODIGO=1"
        Else
            cSQL = cSQL & " WHERE TIS_CODIGO=1"
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
            txtCodSocB.Text = .ResultFields(3)
            txtCodSocB_LostFocus
        End If
    End With
    
    Set B = Nothing
End Sub

Private Sub BuscarDeuda(mCodigo As String, mNombre As String)
    'TITULARES
    sql = "SELECT S.SOC_NOMBRE, D.DEB_MES, D.DEB_ANO, S.SOC_CODIGO,"
    sql = sql & " D.DEB_ITEM, D.DEB_DETALLE, D.DEB_IMPORTE, D.DEB_FECHA"
    sql = sql & " FROM DEBITOS D, SOCIOS S"
    sql = sql & " WHERE "
    sql = sql & " S.SOC_CODIGO=D.SOC_CODIGO"
    sql = sql & " AND S.SOC_CODIGO=" & XN(mCodigo)
    sql = sql & " AND D.DEB_SALDO > 0"
    sql = sql & " AND D.DEB_PAGADO IS NULL"
    
    sql = sql & " UNION ALL"
    
    'OPTATIVOS
    sql = sql & " SELECT S1.SOC_NOMBRE, D1.DEB_MES, D1.DEB_ANO, S1.SOC_CODIGO,"
    sql = sql & " D1.DEB_ITEM, D1.DEB_DETALLE, D1.DEB_IMPORTE, D1.DEB_FECHA"
    sql = sql & " FROM DEBITOS D1, SOCIOS S1"
    sql = sql & " WHERE"
    sql = sql & " S1.SOC_CODIGO=D1.SOC_CODIGO"
    sql = sql & " AND S1.SOC_TITULAR=" & XN(mCodigo)
    sql = sql & " AND D1.DEB_SALDO > 0"
    sql = sql & " AND D1.DEB_PAGADO IS NULL"
    sql = sql & " ORDER BY D.DEB_FECHA, S.SOC_CODIGO, D.DEB_ITEM"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            sql = "INSERT INTO CTA_CTE_SOCIOS (SOC_CODIGO,SOC_NOMBRE,"
            sql = sql & " DEB_MES,DEB_ANO,DEB_DETALLE,DEB_IMPORTE,"
            sql = sql & " DEB_ITEM,DEB_FECHA,CTACTE_TITULAR,CTACTE_CODIGO)"
            sql = sql & " VALUES ("
            sql = sql & XN(Rec1!SOC_CODIGO) & ","
            sql = sql & XS(Rec1!SOC_NOMBRE) & ","
            sql = sql & XS(Format(Rec1!DEB_MES, "00")) & ","
            sql = sql & XS(Format(Rec1!DEB_ANO, "0000")) & ","
            sql = sql & XS(Rec1!DEB_DETALLE) & ","
            sql = sql & XN(Rec1!DEB_IMPORTE) & ","
            sql = sql & XN(Rec1!DEB_ITEM) & ","
            sql = sql & XDQ(Rec1!DEB_FECHA) & ","
            sql = sql & XS(mNombre) & ","
            sql = sql & XN(mCodigo) & ")"
            DBConn.Execute sql
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

