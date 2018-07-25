VERSION 5.00
Begin VB.Form ABMBancos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos de Bancos..."
   ClientHeight    =   3060
   ClientLeft      =   2700
   ClientTop       =   2625
   ClientWidth     =   5610
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ABMBancos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtBanCodigo 
      Height          =   300
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1530
      Width           =   720
   End
   Begin VB.TextBox txtBanSucursal 
      Height          =   300
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1200
      Width           =   720
   End
   Begin VB.TextBox txtBanTipo 
      Height          =   300
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   2
      Top             =   870
      Width           =   720
   End
   Begin VB.TextBox txtBanBanco 
      Height          =   300
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   1
      Top             =   540
      Width           =   720
   End
   Begin VB.TextBox txtBanNomCorto 
      Height          =   300
      Left            =   1200
      MaxLength       =   25
      TabIndex        =   6
      Top             =   2190
      Width           =   4275
   End
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   210
      Picture         =   "ABMBancos.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2625
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txtDescri 
      Height          =   300
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   5
      Top             =   1860
      Width           =   4275
   End
   Begin VB.TextBox txtID 
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   210
      Width           =   720
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   4185
      TabIndex        =   8
      Top             =   2625
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   2805
      TabIndex        =   7
      Top             =   2625
      Width           =   1300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sucursal:"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   1230
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   900
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Banco:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   585
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre corto:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   2205
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1875
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Id.:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   255
      Width           =   270
   End
End
Attribute VB_Name = "ABMBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'parametros para la configuración de la ventana de datos
Dim vFieldID As String
Dim vStringSQL As String
Dim vFormLlama As Form
Dim vMode As Integer
Dim vListView As ListView
Dim vDesFieldID As String


'constantes para funcionalidad de uso del formulario
Const cSugerirID = True 'si es True si sugiere un identificador cuando deja el campo en blanco
Const cTabla = "BANCO"
Const cCampoID = "BAN_CODINT"
Const cDesRegistro = "Banco"

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
            AcCtrl txtBanBanco
            AcCtrl txtBanTipo
            AcCtrl txtBanSucursal
            AcCtrl txtBanCodigo
            AcCtrl txtDescri
            AcCtrl txtBanNomCorto
        Case 3, 4
            DesacCtrl txtBanBanco
            DesacCtrl txtBanTipo
            DesacCtrl txtBanSucursal
            DesacCtrl txtBanCodigo
            DesacCtrl txtDescri
            DesacCtrl txtBanNomCorto
    End Select
    
    
    Select Case pMode
        Case 1
            cmdAceptar.Enabled = False
            Me.Caption = "Nuevo Banco..."
            AcCtrl txtID
        Case 2
            cmdAceptar.Enabled = False
            Me.Caption = "Editando Banco..."
            DesacCtrl txtID
        Case 3
            cmdAceptar.Visible = False
            Me.Caption = "Datos de Banco..."
            DesacCtrl txtID
        Case 4
            cmdAceptar.Enabled = True
            Me.Caption = "Eliminando Banco..."
            DesacCtrl txtID
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
                             "Ingrese la Identificación del Banco antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtID.SetFocus
            Exit Function
        ElseIf txtBanBanco.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Banco antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtBanBanco.SetFocus
            Exit Function
        ElseIf txtBanTipo.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Tipo de Entidad antes de aceptar.(2- entidad bancaria)", vbCritical + vbOKOnly, App.Title
            txtBanTipo.SetFocus
            Exit Function
        ElseIf txtBanSucursal.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la Sucursal del Banco antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtBanSucursal.SetFocus
            Exit Function
        ElseIf txtBanCodigo.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Código del Banco antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtBanCodigo.SetFocus
            Exit Function
        ElseIf txtDescri.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la descripción del Banco antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtDescri.SetFocus
            Exit Function
        ElseIf txtBanNomCorto.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Nombre corto del Banco antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtBanNomCorto.SetFocus
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
                cSQL = cSQL & " (BAN_CODINT, BAN_BANCO,BAN_LOCALIDAD,BAN_SUCURSAL,"
                cSQL = cSQL & " BAN_CODIGO,BAN_DESCRI,BAN_NOMCOR) "
                cSQL = cSQL & " VALUES ("
                cSQL = cSQL & XN(txtID.Text) & ", "
                cSQL = cSQL & XS(txtBanBanco.Text) & ", "
                cSQL = cSQL & XS(txtBanTipo.Text) & ", "
                cSQL = cSQL & XS(txtBanSucursal.Text) & ", "
                cSQL = cSQL & XS(txtBanCodigo.Text) & ", "
                cSQL = cSQL & XS(txtDescri.Text) & ", "
                cSQL = cSQL & XS(txtBanNomCorto.Text) & ")"
            
            Case 2 'editar
                
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & " BAN_DESCRI = " & XS(txtDescri.Text)
                cSQL = cSQL & " ,BAN_NOMCOR = " & XS(txtBanNomCorto.Text)
                cSQL = cSQL & " ,BAN_BANCO=" & XS(txtBanBanco.Text)
                cSQL = cSQL & " ,BAN_LOCALIDAD=" & XS(txtBanTipo.Text)
                cSQL = cSQL & " ,BAN_SUCURSAL=" & XS(txtBanSucursal.Text)
                cSQL = cSQL & " ,BAN_CODIGO=" & XS(txtBanCodigo.Text)
                cSQL = cSQL & " WHERE BAN_CODINT  = " & XN(txtID.Text)
            
            Case 4 'eliminar
            
                cSQL = "DELETE FROM " & cTabla & " WHERE BAN_CODINT  = " & XN(txtID.Text)
            
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
    Dim Rec As ADODB.Recordset
    Set Rec = New ADODB.Recordset
    
    'Me.Top = vFormLlama.Top + 1500
    'Me.Left = vFormLlama.Left + 1000
    
    'txtID.MaxLength = 4
    'txtDescri.MaxLength = 30
    
    If vMode <> 1 Then
        If vFieldID <> "0" Then
            cSQL = "SELECT * FROM " & cTabla & "  WHERE BAN_CODINT = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
            Rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (Rec.BOF And Rec.EOF) = 0 Then
                'si encontró el registro muestro los datos
                txtID.Text = Rec!BAN_CODINT
                txtBanBanco.Text = Rec!BAN_BANCO
                txtBanTipo.Text = Rec!BAN_LOCALIDAD
                txtBanSucursal.Text = Rec!BAN_SUCURSAL
                txtBanCodigo.Text = Rec!BAN_CODIGO
                txtDescri.Text = ChkNull(Trim(Rec!BAN_DESCRI))
                txtBanNomCorto.Text = ChkNull(Trim(Rec!BAN_NOMCOR))
            Else
                Beep
                MsgBox "Imposible encontrar el registro seleccionado.", vbCritical + vbOKOnly, App.Title
            End If
        End If
    End If
    
    'establesco funcionalidad del form de datos
    SetMode vMode
    
End Sub

Private Sub txtBanBanco_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtBanBanco_GotFocus()
    SelecTexto txtBanBanco
End Sub

Private Sub txtBanBanco_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtBanBanco_LostFocus()
    If txtBanBanco.Text <> "" Then
        txtBanBanco.Text = Format(txtBanBanco.Text, "000")
    End If
End Sub

Private Sub txtBanCodigo_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtBanCodigo_GotFocus()
    SelecTexto txtBanCodigo
End Sub

Private Sub txtBanCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtBanCodigo_LostFocus()
    If txtBanCodigo.Text <> "" Then
        txtBanCodigo.Text = Format(txtBanCodigo.Text, "000000")
    End If
End Sub

Private Sub txtBanNomCorto_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtBanNomCorto_GotFocus()
    SelecTexto txtBanNomCorto
End Sub

Private Sub txtBanNomCorto_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtBanSucursal_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtBanSucursal_GotFocus()
    SelecTexto txtBanSucursal
End Sub

Private Sub txtBanSucursal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtBanSucursal_LostFocus()
    If txtBanSucursal.Text <> "" Then
        txtBanSucursal.Text = Format(txtBanSucursal.Text, "000")
    End If
End Sub

Private Sub txtBanTipo_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtBanTipo_GotFocus()
    SelecTexto txtBanTipo
End Sub

Private Sub txtBanTipo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtBanTipo_LostFocus()
    If txtBanTipo.Text <> "" Then
        txtBanTipo.Text = Format(txtBanTipo.Text, "000")
    End If
End Sub

Private Sub txtdescri_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtdescri_GotFocus()
    SelecTexto txtDescri
End Sub

Private Sub txtDescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtID_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub TxtID_GotFocus()
    SelecTexto txtID
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
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


