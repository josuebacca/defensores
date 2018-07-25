VERSION 5.00
Begin VB.Form ABMDeporte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos del Deporte..."
   ClientHeight    =   2685
   ClientLeft      =   2700
   ClientTop       =   2625
   ClientWidth     =   4530
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ABMDeporte.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkCobroCuota 
      Alignment       =   1  'Right Justify
      Caption         =   "No Generar Débitos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2550
      TabIndex        =   11
      Top             =   510
      Width           =   1905
   End
   Begin VB.TextBox txtRecargo 
      Height          =   330
      Left            =   2970
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1770
      Width           =   1515
   End
   Begin VB.TextBox txtCuota 
      Height          =   330
      Left            =   195
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1770
      Width           =   1515
   End
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   210
      Picture         =   "ABMDeporte.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   330
   End
   Begin VB.TextBox txtDescri 
      Height          =   330
      Left            =   210
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1125
      Width           =   4275
   End
   Begin VB.TextBox txtID 
      Height          =   330
      Left            =   210
      TabIndex        =   0
      Top             =   465
      Width           =   720
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   3120
      TabIndex        =   5
      Top             =   2280
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   1740
      TabIndex        =   4
      Top             =   2280
      Width           =   1300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Valor Recargo:"
      Height          =   195
      Index           =   3
      Left            =   2970
      TabIndex        =   10
      Top             =   1545
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Valor Cuota:"
      Height          =   195
      Index           =   2
      Left            =   225
      TabIndex        =   9
      Top             =   1545
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   900
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Id.:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   270
   End
End
Attribute VB_Name = "ABMDeporte"
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
Const cTabla = "DEPORTE"
Const cCampoID = "DEP_CODIGO"
Const cDesRegistro = "Deporte"

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
            AcCtrl txtDescri
            AcCtrl txtCuota
            AcCtrl txtRecargo
        Case 3, 4
            DesacCtrl txtDescri
            DesacCtrl txtCuota
            DesacCtrl txtRecargo
    End Select
    
    
    Select Case pMode
        Case 1
            cmdAceptar.Enabled = False
            Me.Caption = "Nuevo Deporte..."
            txtID_LostFocus
            DesacCtrl txtID
        Case 2
            cmdAceptar.Enabled = False
            Me.Caption = "Editando Deporte..."
            DesacCtrl txtID
        Case 3
            cmdAceptar.Visible = False
            Me.Caption = "Datos del Deporte..."
            DesacCtrl txtID
        Case 4
            cmdAceptar.Enabled = True
            Me.Caption = "Eliminando Deporte..."
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
                             "Ingrese la Identificación del Deporte antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtID.SetFocus
            Exit Function
        ElseIf txtDescri.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la descripción del Deporte antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtDescri.SetFocus
            Exit Function
            
        ElseIf txtCuota.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el valor de la Cuota antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtCuota.SetFocus
            Exit Function
        End If
    End If
    
    Validar = True
    
End Function

Private Sub chkCobroCuota_Click()
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
                cSQL = cSQL & " (DEP_CODIGO, DEP_DESCRI, DEP_DEBITO, DEP_CUOTA, DEP_RECARGO) "
                cSQL = cSQL & "VALUES "
                cSQL = cSQL & " (" & XN(txtID.Text) & ", " & XS(txtDescri.Text, True) & ", "
                cSQL = cSQL & IIf(chkCobroCuota.Value = 0, XS(""), XS("S")) & ","
                cSQL = cSQL & XN(txtCuota.Text) & "," & XN(txtRecargo.Text) & ")"
            
            Case 2 'editar
                
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & "  DEP_DESCRI = " & XS(txtDescri.Text, True)
                cSQL = cSQL & " ,DEP_CUOTA = " & XN(txtCuota.Text)
                cSQL = cSQL & " ,DEP_RECARGO = " & XN(txtRecargo.Text)
                cSQL = cSQL & " ,DEP_DEBITO=" & IIf(chkCobroCuota.Value = 0, XS(""), XS("S"))
                cSQL = cSQL & " WHERE DEP_CODIGO  = " & XN(txtID.Text)
            
            Case 4 'eliminar
            
                cSQL = "DELETE FROM " & cTabla & " WHERE DEP_CODIGO  = " & XN(txtID.Text)
            
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
    Dim Rec As ADODB.Recordset
    Set Rec = New ADODB.Recordset
    
    'Me.Top = vFormLlama.Top + 1500
    'Me.Left = vFormLlama.Left + 1000
    
    'txtID.MaxLength = 4
    'txtDescri.MaxLength = 30
    
    If vMode <> 1 Then
        If vFieldID <> "0" Then
            cSQL = "SELECT * FROM " & cTabla & "  WHERE DEP_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
            Rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (Rec.BOF And Rec.EOF) = 0 Then
                'si encontró el registro muestro los datos
                txtID.Text = Rec!DEP_CODIGO
                txtDescri.Text = Trim(ChkNull(Rec!DEP_DESCRI))
                txtCuota.Text = Valido_Importe(Chk0(Rec!DEP_CUOTA))
                txtRecargo.Text = Valido_Importe(Chk0(Rec!DEP_RECARGO))
                If IsNull(Rec!DEP_DEBITO) Then
                    chkCobroCuota.Value = 0
                Else
                    chkCobroCuota.Value = 1
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

Private Sub txtCuota_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtCuota_GotFocus()
    SelecTexto txtCuota
End Sub

Private Sub txtCuota_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtCuota, KeyAscii)
End Sub

Private Sub txtCuota_LostFocus()
    If txtCuota.Text <> "" Then
        txtCuota.Text = Valido_Importe(txtCuota.Text)
    End If
End Sub

Private Sub txtdescri_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtdescri_GotFocus()
    seltxt
End Sub

Private Sub txtDescri_KeyPress(KeyAscii As Integer)
    'KeyAscii = CarTexto(KeyAscii)
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

Private Sub txtRecargo_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtRecargo_GotFocus()
    SelecTexto txtRecargo
End Sub

Private Sub txtRecargo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtRecargo, KeyAscii)
End Sub

Private Sub txtRecargo_LostFocus()
    If txtRecargo.Text <> "" Then
        txtRecargo.Text = Valido_Importe(txtRecargo.Text)
    End If
End Sub
