VERSION 5.00
Begin VB.Form ABMEmpleado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos del Empleado..."
   ClientHeight    =   3660
   ClientLeft      =   2700
   ClientTop       =   2625
   ClientWidth     =   4605
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ABMEmpleado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkVenEstado 
      Caption         =   "Dar de Baja"
      Height          =   285
      Left            =   1065
      TabIndex        =   7
      Top             =   2865
      Width           =   1140
   End
   Begin VB.TextBox txtDomicilio 
      Height          =   315
      Left            =   1065
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1380
      Width           =   3375
   End
   Begin VB.TextBox txtMail 
      Height          =   315
      Left            =   1065
      MaxLength       =   50
      TabIndex        =   6
      Top             =   2490
      Width           =   3375
   End
   Begin VB.TextBox txtFax 
      Height          =   315
      Left            =   1065
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox txtTelefono 
      Height          =   315
      Left            =   1065
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1830
      Width           =   3375
   End
   Begin VB.ComboBox cboLocalidad 
      Height          =   315
      ItemData        =   "ABMEmpleado.frx":000C
      Left            =   1065
      List            =   "ABMEmpleado.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1035
      Width           =   3375
   End
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   240
      Picture         =   "ABMEmpleado.frx":0010
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   1065
      MaxLength       =   50
      TabIndex        =   1
      Top             =   630
      Width           =   3375
   End
   Begin VB.TextBox txtID 
      Height          =   315
      Left            =   1065
      TabIndex        =   0
      Top             =   285
      Width           =   720
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   3150
      TabIndex        =   9
      Top             =   3240
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   1800
      TabIndex        =   8
      Top             =   3240
      Width           =   1300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Domicilio:"
      Height          =   195
      Index           =   8
      Left            =   135
      TabIndex        =   17
      Top             =   1425
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "e-mail:"
      Height          =   195
      Index           =   7
      Left            =   135
      TabIndex        =   16
      Top             =   2535
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fax:"
      Height          =   195
      Index           =   6
      Left            =   135
      TabIndex        =   15
      Top             =   2205
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tel�fono:"
      Height          =   195
      Index           =   5
      Left            =   135
      TabIndex        =   14
      Top             =   1875
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Localidad:"
      Height          =   195
      Index           =   4
      Left            =   135
      TabIndex        =   13
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   11
      Top             =   675
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Id.:"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   10
      Top             =   315
      Width           =   270
   End
End
Attribute VB_Name = "ABMEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'parametros para la configuraci�n de la ventana de datos
Dim vFieldID As String
Dim vStringSQL As String
Dim vFormLlama As Form
Dim vMode As Integer
Dim vListView As ListView
Dim vDesFieldID As String
Dim Pais As String
Dim Provincia As String


'constantes para funcionalidad de uso del formulario
Const cSugerirID = True 'si es True si sugiere un identificador cuando deja el campo en blanco
Const cTabla = "EMPLEADOS"
Const cCampoID = "EMP_CODIGO"
Const cDesRegistro = "Empleado"

Function ActualizarListaBase(pMode As Integer)
    On Error GoTo moco
    Dim rec As ADODB.Recordset
    Dim cSQL As String
    Dim i As Integer
    Dim auxListItem As ListItem
    Dim IndiceCampoID As Integer
    Dim OrdenCampo As Integer
    Dim f As ADODB.Field
    Set rec = New ADODB.Recordset
    
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
    
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
        If rec.EOF = False Then
        
'            'busco el indce del campo identificador
            OrdenCampo = 0
            IndiceCampoID = 0
            For Each f In rec.Fields
                OrdenCampo = OrdenCampo + 1
                If UCase(f.Name) = UCase(vDesFieldID) Then
                    IndiceCampoID = OrdenCampo - 1
                End If
            Next f
        
            'recorro la coleci�n de campos a actualizar
            For i = 0 To rec.Fields.Count - 1
                If i = 0 Then
                    Select Case pMode
                        Case 1
                            Set auxListItem = vListView.ListItems.Add(, "'" & rec.Fields(IndiceCampoID) & "'", CStr(IIf(IsNull(rec.Fields(i)), "", rec.Fields(i))), 1)
                            auxListItem.Icon = 1
                            auxListItem.SmallIcon = 1
                            
                        Case 2
                            Set auxListItem = vListView.SelectedItem
                            auxListItem.Text = rec.Fields(i)
                    End Select
                Else
                    auxListItem.SubItems(i) = IIf(IsNull(rec.Fields(i)), "", rec.Fields(i))
                End If
            Next i
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
    'Parametro: pMode indica el modo en que se utilizar� este form
    '  pMode  =             1> Indica nuevo registro
    '                       2> Editar registro existente
    '                       3> Mostrar dato del registro existente
    '                       4> Eliminar registro existente
    
    
    Select Case pMode
        Case 1, 2
            AcCtrl txtNombre
            AcCtrl cboLocalidad
            AcCtrl txtDomicilio
            AcCtrl txtTelefono
            AcCtrl txtFax
            AcCtrl txtMail
            AcCtrl chkVenEstado
        Case 3, 4
            DesacCtrl txtNombre
            DesacCtrl cboLocalidad
            DesacCtrl txtDomicilio
            DesacCtrl txtTelefono
            DesacCtrl txtFax
            DesacCtrl txtMail
            DesacCtrl chkVenEstado
    End Select
    
    Select Case pMode
        Case 1
            cmdAceptar.Enabled = False
            Me.Caption = "Nuevo " & cDesRegistro
            txtID_LostFocus
            DesacCtrl txtID
        Case 2
            cmdAceptar.Enabled = False
            Me.Caption = "Editando " & cDesRegistro
            DesacCtrl txtID
        Case 3
            cmdAceptar.Visible = False
            Me.Caption = "Datos del " & cDesRegistro
            DesacCtrl txtID
        Case 4
            cmdAceptar.Enabled = True
            Me.Caption = "Eliminando " & cDesRegistro
            DesacCtrl txtID
    End Select
    
End Function

Public Function SetWindow(pWindow As Form, pSQL As String, pMode As Integer, pListview As ListView, pDesID As String)
    
    Set vFormLlama = pWindow 'Objeto ventana que que llama a la ventana de datos
    vStringSQL = pSQL 'string utilizado para argar la lista base
    vMode = pMode  'modo en que se utilizar� la ventana de datos
    Set vListView = pListview 'objeto listview que se est� editando
    vDesFieldID = pDesID 'nombre del campo identificador
    
    'valor del campo identificador de registro seleccionado (0 si es un reg. nuevo)
    If vMode <> 1 Then
        If vListView.SelectedItem.Selected = True Then
            vFieldID = vListView.SelectedItem.Key
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
            MsgBox "Falta informaci�n." & Chr(13) & _
                             "Ingrese la Identificaci�n del  " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtID.SetFocus
            Exit Function
        ElseIf txtNombre.Text = "" Then
            Beep
            MsgBox "Falta informaci�n." & Chr(13) & _
                             "Ingrese el Nombre del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtNombre.SetFocus
            Exit Function
        
        ElseIf cboLocalidad.ListIndex = -1 Then
            Beep
            MsgBox "Falta informaci�n." & Chr(13) & _
                             "Ingrese la Localidad del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboLocalidad.SetFocus
            Exit Function
        End If
    End If
    
    Validar = True
    
End Function

Private Sub cboLocalidad_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub chkVenEstado_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cmdAceptar_Click()

    Dim cSQL As String
    
    If Validar(vMode) = True Then
        
        On Error GoTo ErrorTran
        
        Screen.MousePointer = vbHourglass
    
        DBConn.BeginTrans
        Select Case vMode
            Case 1 'nuevo
            
                cSQL = "INSERT INTO " & cTabla
                cSQL = cSQL & "     (EMP_CODIGO, EMP_NOMBRE, EMP_DOMICI, EMP_TELEFONO,"
                cSQL = cSQL & " EMP_MAIL, EMP_FAX, LOC_CODIGO, EMP_ESTADO) "
                cSQL = cSQL & " VALUES "
                cSQL = cSQL & "     (" & XN(txtID.Text) & ", " & XS(txtNombre.Text, True) & ", "
                cSQL = cSQL & XS(txtDomicilio.Text, True) & ", " & XS(txtTelefono.Text, True) & ", "
                cSQL = cSQL & XS(txtMail.Text, True) & ", " & XS(txtFax.Text, True) & ", "
                cSQL = cSQL & cboLocalidad.ItemData(cboLocalidad.ListIndex) & ", "
                If chkVenEstado.Value = Checked Then
                    cSQL = cSQL & "'S')"
                Else
                    cSQL = cSQL & "'N')"
                End If
            Case 2 'editar
                
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & "  EMP_NOMBRE=" & XS(txtNombre.Text, True)
                cSQL = cSQL & " ,EMP_DOMICI=" & XS(txtDomicilio.Text, True)
                cSQL = cSQL & " ,EMP_TELEFONO=" & XS(txtTelefono.Text, True)
                cSQL = cSQL & " ,EMP_MAIL=" & XS(txtMail.Text, True)
                cSQL = cSQL & " ,EMP_FAX=" & XS(txtFax.Text, True)
                cSQL = cSQL & " ,LOC_CODIGO=" & cboLocalidad.ItemData(cboLocalidad.ListIndex)
                If chkVenEstado.Value = Checked Then
                    cSQL = cSQL & " ,EMP_ESTADO = 'S'"
                Else
                    cSQL = cSQL & " ,EMP_ESTADO = 'N'"
                End If
                cSQL = cSQL & " WHERE EMP_CODIGO  = " & XN(txtID.Text)
            
            Case 4 'eliminar
            
                cSQL = "DELETE FROM " & cTabla & " WHERE EMP_CODIGO  = " & XN(txtID.Text)
        End Select
        
        DBConn.Execute cSQL
        DBConn.CommitTrans
        'On Error GoTo 0
        
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
    'ManejoDeErrores DBConn.ErrorNative
    MsgBox Err.Description, vbCritical
    
End Sub


Private Sub cmdAyuda_Click()
    Call WinHelp(Me.hWnd, App.Path & "\help\AYUDA.HLP", cdlHelpContext, 12)
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
    
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()

    Dim cSQL As String
    Dim hSQL As String
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    
    'Me.Top = vFormLlama.Top + 1500
    'Me.Left = vFormLlama.Left + 1000

    Call CargoComboBox(cboLocalidad, "LOCALIDAD", "LOC_CODIGO", "LOC_DESCRI")
    If cboLocalidad.ListCount > 0 Then
        cboLocalidad.ListIndex = -1
    End If
    
    If vMode <> 1 Then
        If vFieldID <> "0" Then
            cSQL = "SELECT * FROM " & cTabla & "  WHERE EMP_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                txtID.Text = rec!EMP_CODIGO
                txtNombre.Text = rec!EMP_NOMBRE
                'si encontr� el registro muestro los datos
                     
                Call BuscaCodigoProxItemData(CInt(rec!LOC_CODIGO), cboLocalidad)
                txtDomicilio.Text = ChkNull(rec!EMP_DOMICI)
                txtTelefono.Text = ChkNull(rec!EMP_TELEFONO)
                txtFax.Text = ChkNull(rec!EMP_FAX)
                txtMail.Text = ChkNull(rec!EMP_MAIL)
                
                If ChkNull(rec!EMP_ESTADO) = "N" Or ChkNull(rec!EMP_ESTADO) = "" Then
                    chkVenEstado.Value = Unchecked
                Else
                    chkVenEstado.Value = Checked
                End If
            Else
                Beep
                MsgBox "Imposible encontrar el registro seleccionado.", vbCritical + vbOKOnly, App.Title
            End If
        End If
    End If
    
    'establesco funcionalidad del form de datos
    SetMode vMode
End Sub

Private Sub txtDomicilio_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtDomicilio_GotFocus()
    SelecTexto txtDomicilio
End Sub

Private Sub txtDomicilio_KeyPress(KeyAscii As Integer)
    'KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtFax_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtFax_GotFocus()
    SelecTexto txtFax
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    'KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtMail_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtMail_GotFocus()
    SelecTexto txtMail
End Sub

Private Sub txtNombre_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtNombre_GotFocus()
    seltxt
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    'KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtID_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtID_GotFocus()
    seltxt
End Sub

Private Sub txtID_LostFocus()

    Dim cSQL As String
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    
    If vMode = 1 Then ' si se esta usando en modo de nuevo registro
        If txtID.Text = "" Then
            If cSugerirID = True Then
                cSQL = "SELECT MAX(" & cCampoID & ") FROM " & cTabla
                'cSQL = cSQL & " WHERE PAI_CODIGO = " & cboPais.ItemData(cboPais.ListIndex)
                rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
                If (rec.BOF And rec.EOF) = 0 Then
                    If rec.Fields(0) > 0 Then
                        txtID.Text = rec.Fields(0) + 1
                    Else
                        txtID.Text = 1
                    End If
                End If
            End If
        Else
            'verifico que no sea clave repetida
            cSQL = "SELECT COUNT(*) FROM " & cTabla & " WHERE " & cCampoID & " = " & XN(txtID.Text)
            'cSQL = cSQL & " AND PAI_CODIGO = " & cboPais.ItemData(cboPais.ListIndex)
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                If rec.Fields(0) > 0 Then
                    Beep
                    MsgBox "C�digo de " & cDesRegistro & " repetido." & Chr(13) & _
                                     "El c�digo ingresado Pertenece a otro registro de " & cDesRegistro & ".", vbCritical + vbOKOnly, App.Title
                    txtID.Text = ""
                    txtID.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub txtTelefono_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txttelefono_GotFocus()
    SelecTexto txtTelefono
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
    'KeyAscii = CarTexto(KeyAscii)
End Sub
