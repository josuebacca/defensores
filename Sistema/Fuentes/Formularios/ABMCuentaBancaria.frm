VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ABMCuentaBancaria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos Cuenta Bancaria..."
   ClientHeight    =   2745
   ClientLeft      =   2700
   ClientTop       =   2625
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ABMCuentaBancaria.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBuscarBanco 
      Height          =   315
      Left            =   5925
      Picture         =   "ABMCuentaBancaria.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Buscar Socio"
      Top             =   90
      Width           =   330
   End
   Begin VB.TextBox txtNomBanco 
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
      Left            =   2115
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   90
      Width           =   3780
   End
   Begin VB.TextBox txtCodBanco 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   1335
      TabIndex        =   0
      Top             =   90
      Width           =   750
   End
   Begin VB.TextBox txtdescri 
      Height          =   315
      Left            =   1335
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1890
      Width           =   4920
   End
   Begin VB.TextBox txtSaldoActual 
      Height          =   315
      Left            =   3930
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1530
      Width           =   1245
   End
   Begin VB.TextBox txtSaldoInicial 
      Height          =   315
      Left            =   3930
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1170
      Width           =   1245
   End
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   90
      Picture         =   "ABMCuentaBancaria.frx":0396
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2295
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txtID 
      Height          =   315
      Left            =   1335
      MaxLength       =   10
      TabIndex        =   1
      Top             =   450
      Width           =   1245
   End
   Begin VB.ComboBox cboTipoCuenta 
      Height          =   315
      ItemData        =   "ABMCuentaBancaria.frx":04E0
      Left            =   1335
      List            =   "ABMCuentaBancaria.frx":04E2
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   810
      Width           =   3855
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   4935
      TabIndex        =   7
      Top             =   2295
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   3555
      TabIndex        =   6
      Top             =   2295
      Width           =   1300
   End
   Begin MSComCtl2.DTPicker FechaApertura 
      Height          =   315
      Left            =   1335
      TabIndex        =   19
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   110428161
      CurrentDate     =   43307
   End
   Begin MSComCtl2.DTPicker FechaCierre 
      Height          =   315
      Left            =   1335
      TabIndex        =   20
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   110428161
      CurrentDate     =   43307
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      Height          =   195
      Index           =   5
      Left            =   75
      TabIndex        =   16
      Top             =   1920
      Width           =   870
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Actual:"
      Height          =   195
      Index           =   4
      Left            =   2850
      TabIndex        =   15
      Top             =   1575
      Width           =   945
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Cierre:"
      Height          =   195
      Index           =   3
      Left            =   75
      TabIndex        =   14
      Top             =   1575
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Inicial:"
      Height          =   195
      Index           =   2
      Left            =   2850
      TabIndex        =   13
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Cuenta:"
      Height          =   195
      Index           =   1
      Left            =   75
      TabIndex        =   11
      Top             =   480
      Width           =   885
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Banco:"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Apertura:"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   9
      Top             =   1200
      Width           =   1185
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Cuenta:"
      Height          =   195
      Index           =   1
      Left            =   75
      TabIndex        =   8
      Top             =   840
      Width           =   930
   End
End
Attribute VB_Name = "ABMCuentaBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'parametros para la configuración de la ventana de datos
Dim vFieldID As String
Dim vFieldID1 As String
Dim vStringSQL As String
Dim vFormLlama As Form
Dim vMode As Integer
Dim vListView As ListView
Dim vDesFieldID As String
Dim Rec1 As ADODB.Recordset

'constantes para funcionalidad de uso del formulario
Const cSugerirID = True 'si es True si sugiere un identificador cuando deja el campo en blanco
Const cTabla = "CTA_BANCARIA"
Const cCampoID = "BAN_CODINT"
Const cDesRegistro = "Cuenta Bancaria"

Function ActualizarListaBase(pMode As Integer)
    On Error GoTo moco
    Dim Rec As ADODB.Recordset
    Dim cSQL As String
    Dim I As Integer
    Dim auxListItem As ListItem
    Dim IndiceCampoID As Integer
    Dim OrdenCampo As Integer
    Dim f As ADODB.Field
    Set Rec = New ADODB.Recordset
    
    'armo la cadena a ejecutar
    If InStr(1, vStringSQL, "WHERE") = 0 Then
        cSQL = vStringSQL & " WHERE " & cCampoID & " = " & txtID.Text
    Else
        cSQL = vStringSQL & " AND " & cCampoID & " = " & txtID.Text
    End If
    
    If pMode = 4 Then
        vListView.ListItems.Remove vListView.SelectedItem.Index
        Exit Function
    End If
    
    Rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (Rec.BOF And Rec.EOF) = 0 Then
        If Rec.EOF = False Then
        
'            'busco el indce del campo identificador
            OrdenCampo = 0
            IndiceCampoID = 0
            For Each f In Rec.Fields
                OrdenCampo = OrdenCampo + 1
                If UCase(f.Name) = UCase(vDesFieldID) Then
                    IndiceCampoID = OrdenCampo - 1
                End If
            Next f
        
            'recorro la coleción de campos a actualizar
            For I = 0 To Rec.Fields.Count - 1
                If I = 0 Then
                    Select Case pMode
                        Case 1
                            Set auxListItem = vListView.ListItems.Add(, "'" & Rec.Fields(IndiceCampoID) & "'", CStr(IIf(IsNull(Rec.Fields(I)), "", Rec.Fields(I))), 1)
                            auxListItem.Icon = 1
                            auxListItem.SmallIcon = 1
                            
                        Case 2
                            Set auxListItem = vListView.SelectedItem
                            auxListItem.Text = Rec.Fields(I)
                    End Select
                Else
                    auxListItem.SubItems(I) = IIf(IsNull(Rec.Fields(I)), "", Rec.Fields(I))
                End If
            Next I
        End If
    End If
    Exit Function
moco:
    If Err.Number = 35613 Then
        Call Menu.mnuContextABM_Click(4)
    End If
End Function

Function SetMode(pMode As Integer)

    'Configura los controles del form segun el parametro pMode
    'Parametro: pMode indica el modo en que se utilizará este form
    '  pMode  =             1> Indica nuevo registro
    '                       2> Editar registro existente
    '                       3> Mostrar dato del registro existente
    '                       4> Eliminar registro existente
    
    
    Select Case pMode
        Case 1, 2
            AcCtrl cboTipoCuenta
            'DesacCtrl txtNomBanco
            'AcCtrl txtfechaApertura
            'AcCtrl txtSaldoActual
            'AcCtrl txtSaldoInicial
            'AcCtrl txtFechaCierre
            'AcCtrl txtdescri
        Case 3, 4
            DesacCtrl cboTipoCuenta
            'DesacCtrl txtNomBanco
            'DesacCtrl txtfechaApertura
            'DesacCtrl txtSaldoActual
            'DesacCtrl txtSaldoInicial
            'DesacCtrl txtFechaCierre
            'DesacCtrl txtdescri
    End Select
    
    Select Case pMode
        Case 1
            cmdAceptar.Enabled = False
            Me.Caption = "Nueva Cuenta Bancaria..."
            AcCtrl txtID
            AcCtrl txtCodBanco
        Case 2
            cmdAceptar.Enabled = False
            Me.Caption = "Editando Cuenta Bancaria..."
            DesacCtrl txtID
            DesacCtrl txtCodBanco
        Case 3
            cmdAceptar.Visible = False
            Me.Caption = "Datos de la Cuenta Bancaria..."
            DesacCtrl txtID
            DesacCtrl txtCodBanco
        Case 4
            cmdAceptar.Enabled = True
            Me.Caption = "Eliminando Empresa..."
            DesacCtrl txtID
            DesacCtrl txtCodBanco
    End Select
    
End Function

Public Function SetWindow(pWindow As Form, pSQL As String, pMode As Integer, pListview As ListView, pDesID As String)
    
    Set vFormLlama = pWindow 'Objeto ventana que que llama a la ventana de datos
    vStringSQL = pSQL 'string utilizado para argar la lista base
    vMode = pMode  'modo en que se utilizará la ventana de datos
    Set vListView = pListview 'objeto listview que se está editando
    vDesFieldID = pDesID 'nombre del campo identificador
    
    'valor del campo identificador de registro seleccionado (0 si es un reg. nuevo)
    If vMode <> 1 Then
        If vListView.SelectedItem.Selected = True Then
            vFieldID = vListView.SelectedItem.Key
            vFieldID1 = vListView.SelectedItem.SubItems(3)
        Else
            vFieldID = 0
        End If
    Else
        vFieldID = 0
    End If

End Function


Function Validar(pMode As Integer) As Boolean

    If pMode <> 4 Then
        Validar = False
        If txtID.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Número de " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtID.SetFocus
            Exit Function
            
        ElseIf txtCodBanco.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Banco antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtCodBanco.SetFocus
            Exit Function
            
        ElseIf FechaApertura.Value = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la Fecha de Apertura de " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            FechaApertura.SetFocus
            Exit Function
        End If
    End If
    
    Validar = True
    
End Function

Private Sub cmdAceptar_Click()

    Dim cSQL As String
    
    If Validar(vMode) = True Then
    
        Screen.MousePointer = vbHourglass
    
        DBConn.BeginTrans
        Select Case vMode
            Case 1 'nuevo
            
                cSQL = "INSERT INTO " & cTabla
                cSQL = cSQL & "     (CTA_NROCTA, BAN_CODINT, CTA_FECAPE, CTA_SALINI, CTA_SALACT, CTA_FECCIE, TCU_CODIGO, CTA_DESCRI) "
                cSQL = cSQL & "VALUES "
                cSQL = cSQL & "     (" & XS(txtID.Text) & ", " & XN(txtCodBanco.Text) & ", "
                cSQL = cSQL & XD(FechaApertura.Value) & ", " & XN(txtSaldoInicial.Text) & ", "
                cSQL = cSQL & XN(txtSaldoActual.Text) & ", " & XD(FechaCierre.Value) & ", "
                cSQL = cSQL & XN(cboTipoCuenta.ItemData(cboTipoCuenta.ListIndex)) & ", " & XS(txtDescri.Text) & ")"
            
            Case 2 'editar
                
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & "     CTA_FECAPE=" & XD(FechaApertura.Value) & ", CTA_SALINI=" & XN(txtSaldoInicial.Text)
                cSQL = cSQL & ", CTA_SALACT=" & XN(txtSaldoActual.Text) & ", CTA_FECCIE=" & XD(FechaCierre.Value)
                cSQL = cSQL & ", TCU_CODIGO=" & XN(cboTipoCuenta.ItemData(cboTipoCuenta.ListIndex)) & ", CTA_DESCRI = " & XS(txtDescri.Text)
                cSQL = cSQL & "  WHERE CTA_NROCTA =" & XN(txtID.Text)
                cSQL = cSQL & " AND BAN_CODINT=" & XN(txtCodBanco.Text)
            
            Case 4 'eliminar
            
                cSQL = "DELETE FROM " & cTabla
                cSQL = cSQL & "  WHERE CTA_NROCTA =" & XN(txtID.Text)
                cSQL = cSQL & " AND BAN_CODINT=" & XN(txtCodBanco.Text)
            
        End Select
        
        DBConn.Execute cSQL
        DBConn.CommitTrans
        On Error GoTo 0
        
        'actualizo la lista base
        ActualizarListaBase vMode
        
        Screen.MousePointer = vbDefault
        Unload Me
    End If
    Exit Sub
    
ErrorTran:
    
    DBConn.RollbackTrans
    Screen.MousePointer = vbDefault
    
    'manejo el error
    ManejoDeErrores DBConn.ErrorNative
    
End Sub


Private Sub cmdAyuda_Click()
    Call WinHelp(Me.hWnd, App.Path & "\help\AYUDA.HLP", cdlHelpContext, 5)
End Sub

Private Sub cmdBuscarBanco_Click()
    txtCodBanco.Text = ""
    BuscarBanco
    'txtNomBanco.SetFocus
End Sub

Private Sub cmdCerrar_Click()

    Unload Me
    
End Sub

Private Sub Form_Activate()
    'hizo click en una columna no correcta
    If vMode = 2 And vFieldID = "0" Then
        Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
    If KeyAscii = vbKeyEscape Then Unload Me
    
End Sub

Private Sub Form_Load()

    Dim cSQL As String
    Dim hSQL As String
    Dim Rec As ADODB.Recordset
    Set Rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    
    'cargo el combo de tipo CUENTAS
    cboTipoCuenta.Clear
    cSQL = "SELECT * FROM TIPO_CUENTA order by TCU_DESCRI"
    Rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (Rec.BOF And Rec.EOF) = 0 Then
       Do While Rec.EOF = False
          cboTipoCuenta.AddItem Trim(Rec!TCU_DESCRI)
          cboTipoCuenta.ItemData(cboTipoCuenta.NewIndex) = Rec!TCU_CODIGO
          Rec.MoveNext
       Loop
       cboTipoCuenta.ListIndex = cboTipoCuenta.ListIndex + 1
    End If
    Rec.Close
       
    If vMode <> 1 Then
        If vFieldID <> "0" Then
            cSQL = "SELECT * FROM " & cTabla & "  WHERE BAN_CODINT = " & Trim(Mid(vFieldID1, 1, 10))
            cSQL = cSQL & " AND CTA_NROCTA = " & Trim(Mid(vFieldID, 1, 15)) '& frmCListaBaseABM.lstvLista.ListItems(2)
            Rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (Rec.BOF And Rec.EOF) = 0 Then
                'si encontró el registro muestro los datos
                txtID.Text = Rec!CTA_NROCTA
                'Call BuscaCodigoProxItemData(CLng(Rec!BAN_CODINT), cboBanco)
                txtCodBanco.Text = Rec!BAN_CODINT
                txtCodBanco_LostFocus
                
                Call BuscaCodigoProxItemData(CLng(Rec!TCU_CODIGO), cboTipoCuenta)
                FechaApertura.Value = Rec!CTA_FECAPE
                txtSaldoInicial.Text = ChkNull(Rec!CTA_SALINI)
                txtSaldoActual.Text = ChkNull(Rec!CTA_SALACT)
                FechaCierre.Value = ChkNull(Rec!CTA_FECCIE)
                txtDescri.Text = ChkNull(Rec!CTA_DESCRI)
            Else
                Beep
                MsgBox "Imposible encontrar el registro seleccionado.", vbCritical + vbOKOnly, App.Title
            End If
        End If
    End If
    
    'establesco funcionalidad del form de datos
    SetMode vMode
    
End Sub

Private Sub txtCodBanco_GotFocus()
    SelecTexto txtCodBanco
End Sub

Private Sub txtCodBanco_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodBanco_LostFocus()
    If txtCodBanco.Text <> "" Then
        sql = "SELECT BAN_DESCRI, BAN_CODINT"
        sql = sql & " FROM BANCO"
        sql = sql & " WHERE"
        sql = sql & " BAN_CODINT=" & XN(txtCodBanco.Text)
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtCodBanco.Text = Rec1!BAN_CODINT
            txtNomBanco.Text = Trim(Rec1!BAN_DESCRI)
        Else
            MsgBox "El Código ingresado no Existe", vbExclamation, TIT_MSGBOX
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtID_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub TxtID_GotFocus()
    seltxt
End Sub

Private Sub txtID_LostFocus()

    Dim cSQL As String
    Dim Rec As ADODB.Recordset
    Set Rec = New ADODB.Recordset
    
    If vMode = 1 Then ' si se esta usando en modo de nuevo registro
        If txtID.Text = "" Then
            If cSugerirID = True Then
                cSQL = "SELECT MAX(" & cCampoID & ") FROM " & cTabla
                Rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
                If (Rec.BOF And Rec.EOF) = 0 Then
                    If Rec.Fields(0) > 0 Then
                        txtID.Text = Rec.Fields(0) + 1
                    Else
                        txtID.Text = 1
                    End If
                End If
            End If
        Else
            'verifico que no sea clave repetida
            cSQL = "SELECT COUNT(*) FROM " & cTabla & " WHERE " & cCampoID & " = " & txtID.Text
            Rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (Rec.BOF And Rec.EOF) = 0 Then
                If Rec.Fields(0) > 0 Then
                    Beep
                    MsgBox "Código de " & cDesRegistro & " repetido." & Chr(13) & _
                                     "El código ingresado Pertenece a otro registro de " & cDesRegistro & ".", vbCritical + vbOKOnly, App.Title
                    txtID.Text = ""
                    txtID.SetFocus
                End If
            End If
        End If
    End If
    
End Sub

Private Sub txtTele_Enti_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtTele_Enti_GotFocus()
    seltxt
End Sub

Private Sub txtSaldoActual_GotFocus()
    seltxt
End Sub

Private Sub txtSaldoActual_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtSaldoActual, KeyAscii)
End Sub

Private Sub txtSaldoInicial_GotFocus()
    seltxt
End Sub

Private Sub txtSaldoInicial_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtSaldoInicial, KeyAscii)
End Sub

Public Sub BuscarBanco()
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim I, posicion As Integer
    Dim cadena As String
    
    Set B = New CBusqueda
    With B
        cSQL = "SELECT BAN_DESCRI, BAN_CODINT"
        cSQL = cSQL & " FROM BANCO"
        hSQL = "Nombre, Código"
        .sql = cSQL
        .Headers = hSQL
        .Field = "BAN_DESCRI"
        campo1 = .Field
        .Field = "BAN_CODINT"
        campo2 = .Field
        .OrderBy = "BAN_DESCRI"
        camponumerico = False
        .Titulo = "Busqueda de Bancos: "
        .MaxRecords = 1
        .Show

        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
            txtCodBanco.Text = .ResultFields(2)
            txtCodBanco_LostFocus
        End If
    End With
    
    Set B = Nothing
End Sub

