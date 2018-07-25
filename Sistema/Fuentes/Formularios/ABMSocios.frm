VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ABMSocios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos del Socio..."
   ClientHeight    =   6450
   ClientLeft      =   2700
   ClientTop       =   2625
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ABMSocios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker Fecha 
      Height          =   315
      Left            =   1845
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   52363265
      CurrentDate     =   42839
   End
   Begin VB.ComboBox cboEstado 
      Height          =   315
      ItemData        =   "ABMSocios.frx":000C
      Left            =   6105
      List            =   "ABMSocios.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   285
      Width           =   1890
   End
   Begin VB.TextBox txtMail 
      Height          =   330
      Left            =   1845
      MaxLength       =   80
      TabIndex        =   16
      Top             =   3420
      Width           =   6165
   End
   Begin VB.ComboBox cboTipo 
      Height          =   315
      ItemData        =   "ABMSocios.frx":0010
      Left            =   5670
      List            =   "ABMSocios.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   4020
      Width           =   2340
   End
   Begin VB.TextBox txtCodSoc 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1830
      TabIndex        =   19
      Top             =   4350
      Width           =   990
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
      Height          =   315
      Left            =   2835
      TabIndex        =   20
      Top             =   4350
      Width           =   4800
   End
   Begin VB.CommandButton cmdBuscaSocio 
      Height          =   315
      Left            =   7680
      Picture         =   "ABMSocios.frx":0014
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Buscar Socio"
      Top             =   4350
      Width           =   330
   End
   Begin VB.ListBox lstDeportes 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   1830
      Style           =   1  'Checkbox
      TabIndex        =   21
      Top             =   4710
      Width           =   5820
   End
   Begin VB.ComboBox cboTipoCuota 
      Height          =   315
      ItemData        =   "ABMSocios.frx":039E
      Left            =   1845
      List            =   "ABMSocios.frx":03A0
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   4020
      Width           =   2340
   End
   Begin VB.TextBox txtCobroCuota 
      Height          =   330
      Left            =   1845
      MaxLength       =   80
      TabIndex        =   15
      Top             =   3075
      Width           =   6165
   End
   Begin VB.TextBox txtProfesion 
      Height          =   330
      Left            =   1845
      MaxLength       =   100
      TabIndex        =   14
      Top             =   2715
      Width           =   6165
   End
   Begin VB.TextBox txtHijos 
      Height          =   315
      Left            =   6570
      MaxLength       =   30
      TabIndex        =   13
      Top             =   2370
      Width           =   1470
   End
   Begin VB.ComboBox cboEstadoCivil 
      Height          =   315
      ItemData        =   "ABMSocios.frx":03A2
      Left            =   4395
      List            =   "ABMSocios.frx":03A4
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2370
      Width           =   1470
   End
   Begin VB.ComboBox cboSexo 
      Height          =   315
      ItemData        =   "ABMSocios.frx":03A6
      Left            =   1845
      List            =   "ABMSocios.frx":03A8
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2370
      Width           =   1515
   End
   Begin VB.ComboBox cboTipoDoc 
      Height          =   315
      ItemData        =   "ABMSocios.frx":03AA
      Left            =   1845
      List            =   "ABMSocios.frx":03AC
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2010
      Width           =   1140
   End
   Begin VB.TextBox txtNacionalidad 
      Height          =   315
      Left            =   6165
      MaxLength       =   50
      TabIndex        =   8
      Top             =   1665
      Width           =   1860
   End
   Begin VB.TextBox txtNroDoc 
      Height          =   315
      Left            =   6165
      MaxLength       =   9
      TabIndex        =   10
      Top             =   2010
      Width           =   1860
   End
   Begin VB.TextBox txtCodPostal 
      Height          =   315
      Left            =   6165
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1320
      Width           =   1860
   End
   Begin VB.TextBox txtDomicilio 
      Height          =   315
      Left            =   1845
      MaxLength       =   80
      TabIndex        =   3
      Top             =   975
      Width           =   2700
   End
   Begin VB.TextBox txtTelefono 
      Height          =   315
      Left            =   5325
      MaxLength       =   50
      TabIndex        =   4
      Top             =   975
      Width           =   2700
   End
   Begin VB.ComboBox cboLocalidad 
      Height          =   315
      ItemData        =   "ABMSocios.frx":03AE
      Left            =   1845
      List            =   "ABMSocios.frx":03B0
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1320
      Width           =   2670
   End
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   240
      Picture         =   "ABMSocios.frx":03B2
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6015
      Width           =   330
   End
   Begin VB.TextBox txtNombre 
      Height          =   330
      Left            =   1845
      MaxLength       =   80
      TabIndex        =   2
      Top             =   630
      Width           =   6165
   End
   Begin VB.TextBox txtID 
      Height          =   315
      Left            =   1845
      TabIndex        =   0
      Top             =   285
      Width           =   1095
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   6645
      TabIndex        =   23
      Top             =   6030
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   5295
      TabIndex        =   22
      Top             =   6030
      Width           =   1300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Estado:"
      Height          =   195
      Index           =   18
      Left            =   5490
      TabIndex        =   46
      Top             =   315
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail:"
      Height          =   195
      Index           =   17
      Left            =   120
      TabIndex        =   45
      Top             =   3450
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      Height          =   195
      Index           =   12
      Left            =   5205
      TabIndex        =   44
      Top             =   4065
      Width           =   360
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Socio Titular:"
      Height          =   195
      Left            =   120
      TabIndex        =   43
      Top             =   4380
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Deportes:"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   41
      Top             =   4725
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Cuota:"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   40
      Top             =   4065
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Domicilio de Cobro:"
      Height          =   195
      Index           =   16
      Left            =   120
      TabIndex        =   39
      Top             =   3105
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Profesion u Ocupación:"
      Height          =   195
      Index           =   15
      Left            =   135
      TabIndex        =   38
      Top             =   2760
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hijos:"
      Height          =   195
      Index           =   14
      Left            =   6030
      TabIndex        =   37
      Top             =   2415
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Estado Civil:"
      Height          =   195
      Index           =   13
      Left            =   3465
      TabIndex        =   36
      Top             =   2415
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sexo:"
      Height          =   195
      Index           =   11
      Left            =   135
      TabIndex        =   35
      Top             =   2415
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tip. Doc::"
      Height          =   195
      Index           =   10
      Left            =   135
      TabIndex        =   34
      Top             =   2055
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nacionalidad:"
      Height          =   195
      Index           =   9
      Left            =   5115
      TabIndex        =   33
      Top             =   1710
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nro. Doc.:"
      Height          =   195
      Index           =   3
      Left            =   5115
      TabIndex        =   32
      Top             =   2055
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código Postal:"
      Height          =   195
      Index           =   2
      Left            =   5115
      TabIndex        =   31
      Top             =   1380
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "F. Nacimiento:"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   30
      Top             =   1710
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Dirección:"
      Height          =   195
      Index           =   8
      Left            =   135
      TabIndex        =   29
      Top             =   1020
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono:"
      Height          =   195
      Index           =   5
      Left            =   4620
      TabIndex        =   28
      Top             =   1005
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Localidad:"
      Height          =   195
      Index           =   4
      Left            =   135
      TabIndex        =   27
      Top             =   1365
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ape y Nom:"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   25
      Top             =   675
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Id.:"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   24
      Top             =   315
      Width           =   270
   End
End
Attribute VB_Name = "ABMSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'parametros para la configuración de la ventana de datos
Public vFieldID As String
Dim vStringSQL As String
Dim vFormLlama As Form
Public vMode As Integer
Dim vListView As ListView
Dim vDesFieldID As String
Dim Pais As String
Dim Provincia As String
Dim I As Integer
Dim Rec As ADODB.Recordset

'constantes para funcionalidad de uso del formulario
Const cSugerirID = True 'si es True si sugiere un identificador cuando deja el campo en blanco
Const cTabla = "SOCIOS"
Const cCampoID = "SOC_CODIGO"
Const cDesRegistro = "Socio"

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
            AcCtrl txtNombre
            AcCtrl txtNroDoc
            AcCtrl Fecha
            AcCtrl cboLocalidad
            AcCtrl txtDomicilio
            AcCtrl txtTelefono
            AcCtrl txtCodPostal
        Case 3, 4
            DesacCtrl txtNombre
            DesacCtrl txtNroDoc
            DesacCtrl Fecha
            DesacCtrl cboLocalidad
            DesacCtrl txtDomicilio
            DesacCtrl txtTelefono
            DesacCtrl txtCodPostal
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
                             "Ingrese la Identificación del  " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtID.SetFocus
            Exit Function
        ElseIf txtNombre.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Nombre del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtNombre.SetFocus
            Exit Function
            
        ElseIf txtNroDoc.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Nro de Documento del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtNroDoc.SetFocus
            Exit Function
        
        ElseIf txtDomicilio.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Domicilio del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtDomicilio.SetFocus
            Exit Function
            
        ElseIf cboLocalidad.ListIndex = -1 Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la Localidad del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboLocalidad.SetFocus
            Exit Function
            
        ElseIf Fecha.Value = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la Fecha de Nacimiento del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            Fecha.SetFocus
            Exit Function
            
        ElseIf cboTipoDoc.ListIndex = -1 Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Tipo de Documento del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboTipoDoc.SetFocus
            Exit Function
            
        ElseIf txtNroDoc.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Número de Documento del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtNroDoc.SetFocus
            Exit Function
            
        ElseIf cboSexo.ListIndex = -1 Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Sexo del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboSexo.SetFocus
            Exit Function
            
        ElseIf cboEstadoCivil.ListIndex = -1 Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Estado Civil del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboEstadoCivil.SetFocus
            Exit Function
            
        ElseIf txtCobroCuota.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Domicilio de Cobro del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtCobroCuota.SetFocus
            Exit Function
            
        ElseIf cboTipoCuota.ListIndex = -1 Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Tipo de Cuota del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboTipoCuota.SetFocus
            Exit Function
            
        ElseIf cboTipo.ListIndex = -1 Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Tipo de " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboTipo.SetFocus
            Exit Function
            
        ElseIf cboEstado.ListIndex = -1 Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Estado del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboEstado.SetFocus
            Exit Function
            
        ElseIf cboTipoCuota.ItemData(cboTipoCuota.ListIndex) = 1 And cboTipo.ItemData(cboTipo.ListIndex) = 2 And txtCodSoc.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Socio Titular del Grupo Familair antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtCodSoc.SetFocus
            Exit Function
        End If
    End If
    
    Validar = True
    
End Function

Private Sub cboCanal_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboIva_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboEstado_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboEstadoCivil_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboLocalidad_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboSexo_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboTipo_Click()
    cmdAceptar.Enabled = True
    If cboTipoCuota.ItemData(cboTipoCuota.ListIndex) = 1 Then 'grupo_familiar
        If cboTipo.ItemData(cboTipo.ListIndex) = 1 Then
            'TITULAR
            txtCodSoc.Text = ""
            txtNomSoc.Text = ""
            txtCodSoc.Enabled = False
            txtNomSoc.Enabled = False
        Else
            'OPATTIVO
            txtCodSoc.Enabled = True
            txtNomSoc.Enabled = True
            txtCodSoc.Text = ""
            txtNomSoc.Text = ""
        End If
    End If
End Sub

Private Sub cboTipoCuota_Click()
    cmdAceptar.Enabled = True
    If cboTipoCuota.ItemData(cboTipoCuota.ListIndex) = 2 Then 'INDIVIDUAL
        BuscaCodigoProxItemData 1, cboTipo
        cboTipo.Enabled = False
    Else
        cboTipo.Enabled = True
    End If
End Sub

Private Sub cboTipoDoc_Click()
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
                cSQL = cSQL & "  (SOC_CODIGO, SOC_NOMBRE, SOC_DOMICI, TIP_TIPDOC, SOC_NRODOC,"
                cSQL = cSQL & " SOC_TELEFONO, SOC_MAIL, SOC_FECNAC, LOC_CODIGO,"
                cSQL = cSQL & " SOC_CODPOS,SEX_CODIGO,EST_CODIGO,SOC_OCUPACION,"
                cSQL = cSQL & " SOC_DOMCOB,SOC_NACIONA,TIS_CODIGO,TIC_CODIGO,SOC_TITULAR,ESS_CODIGO, SOC_HIJOS)"
                cSQL = cSQL & " VALUES ("
                cSQL = cSQL & XN(txtID.Text) & ", " & XS(txtNombre.Text) & ", "
                cSQL = cSQL & XS(txtDomicilio.Text) & ", " & XS(cboTipoDoc.List(cboTipoDoc.ListIndex)) & ", " & XN(txtNroDoc.Text) & ", "
                cSQL = cSQL & XS(txtTelefono.Text) & ", " & XS(txtMail.Text, True) & ", "
                cSQL = cSQL & XDQ(Fecha.Value) & ", " & cboLocalidad.ItemData(cboLocalidad.ListIndex) & ", "
                cSQL = cSQL & XS(txtCodPostal.Text) & ", " & cboSexo.ItemData(cboSexo.ListIndex) & ", "
                cSQL = cSQL & cboEstadoCivil.ItemData(cboEstadoCivil.ListIndex) & ", "
                cSQL = cSQL & XS(txtProfesion.Text) & ", " & XS(txtCobroCuota.Text) & ", "
                cSQL = cSQL & XS(txtNacionalidad.Text) & ", " & cboTipo.ItemData(cboTipo.ListIndex) & ", "
                cSQL = cSQL & cboTipoCuota.ItemData(cboTipoCuota.ListIndex) & ", " & XN(txtCodSoc.Text) & ","
                cSQL = cSQL & cboEstado.ItemData(cboEstado.ListIndex) & "," & XS(txtHijos.Text) & ")"
                DBConn.Execute cSQL
                
                'DOY DE ALTA LOS DEPORTES
                For I = 0 To lstDeportes.ListCount - 1
                    If lstDeportes.Selected(I) = True Then
                        cSQL = "INSERT INTO SOCIOS_DEPORTES (SOC_CODIGO,DEP_CODIGO)"
                        cSQL = cSQL & " VALUES ("
                        cSQL = cSQL & XN(txtID.Text) & ","
                        cSQL = cSQL & lstDeportes.ItemData(I) & ")"
                        DBConn.Execute cSQL
                    End If
                Next
                
            Case 2 'editar
                
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & "  SOC_NOMBRE=" & XS(txtNombre.Text)
                cSQL = cSQL & " ,SOC_DOMICI=" & XS(txtDomicilio.Text)
                cSQL = cSQL & " ,SOC_NRODOC=" & XN(txtNroDoc.Text)
                cSQL = cSQL & " ,TIP_TIPDOC=" & XS(cboTipoDoc.List(cboTipoDoc.ListIndex))
                cSQL = cSQL & " ,SOC_TELEFONO=" & XS(txtTelefono.Text)
                cSQL = cSQL & " ,SOC_MAIL=" & XS(txtMail.Text, True)
                cSQL = cSQL & " ,SOC_FECNAC=" & XDQ(Fecha.Value)
                cSQL = cSQL & " ,LOC_CODIGO=" & cboLocalidad.ItemData(cboLocalidad.ListIndex)
                cSQL = cSQL & " ,SOC_CODPOS=" & XS(txtCodPostal.Text)
                cSQL = cSQL & " ,SEX_CODIGO=" & cboSexo.ItemData(cboSexo.ListIndex)
                cSQL = cSQL & " ,EST_CODIGO=" & cboEstadoCivil.ItemData(cboEstadoCivil.ListIndex)
                cSQL = cSQL & " ,SOC_OCUPACION=" & XS(txtProfesion.Text)
                cSQL = cSQL & " ,SOC_DOMCOB=" & XS(txtCobroCuota.Text)
                cSQL = cSQL & " ,SOC_NACIONA=" & XS(txtNacionalidad.Text)
                cSQL = cSQL & " ,TIS_CODIGO=" & cboTipo.ItemData(cboTipo.ListIndex)
                cSQL = cSQL & " ,TIC_CODIGO=" & cboTipoCuota.ItemData(cboTipoCuota.ListIndex)
                cSQL = cSQL & " ,SOC_TITULAR=" & XN(txtCodSoc.Text)
                cSQL = cSQL & " ,ESS_CODIGO=" & cboEstado.ItemData(cboEstado.ListIndex)
                cSQL = cSQL & " ,SOC_HIJOS=" & XS(txtHijos.Text)
                cSQL = cSQL & " WHERE SOC_CODIGO  = " & XN(txtID.Text)
                DBConn.Execute cSQL
                        
                'DOY DE ALTA LOS DEPORTES
                cSQL = "DELETE FROM SOCIOS_DEPORTES WHERE SOC_CODIGO=" & XN(txtID.Text)
                DBConn.Execute cSQL
                
                For I = 0 To lstDeportes.ListCount - 1
                    If lstDeportes.Selected(I) = True Then
                        cSQL = "INSERT INTO SOCIOS_DEPORTES (SOC_CODIGO,DEP_CODIGO)"
                        cSQL = cSQL & " VALUES ("
                        cSQL = cSQL & XN(txtID.Text) & ","
                        cSQL = cSQL & lstDeportes.ItemData(I) & ")"
                        DBConn.Execute cSQL
                    End If
                Next
            
            Case 4 'eliminar
                cSQL = "SELECT SOC_CODIGO FROM " & cTabla & " WHERE SOC_TITULAR= " & XN(txtID.Text)
                Rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
                If Rec.EOF = True Then
                    cSQL = "DELETE FROM SOCIOS_DEPORTES WHERE SOC_CODIGO=" & XN(txtID.Text)
                    DBConn.Execute cSQL
                    
                    cSQL = "DELETE FROM " & cTabla & " WHERE SOC_CODIGO  = " & XN(txtID.Text)
                    DBConn.Execute cSQL
                Else
                    MsgBox "El Socio No se puede Borrar ya que es Titular de un Grupo Familiar", vbExclamation, TIT_MSGBOX
                End If
                Rec.Close
        End Select
        
        
        DBConn.CommitTrans
        'On Error GoTo 0
        
        If FormLlamado = "" Then
            'actualizo la lista base
            ActualizarListaBase vMode
        End If
        
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

Private Sub cmdbuscaComp_Click()

End Sub

Private Sub cmdBuscaSocio_Click()
    txtCodSoc.Text = ""
    BuscarSocios "txtCodSoc", "CODIGO"
    txtNomSoc.SetFocus
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
    Dim Rec1 As ADODB.Recordset
    Dim Rec2 As ADODB.Recordset
    Set Rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    
    Centrar_pantalla Me
    'Me.Top = vFormLlama.Top + 1500
    'Me.Left = vFormLlama.Left + 1000
    
    'COMBO LOCALIDAD
    Call CargoComboBox(cboLocalidad, "LOCALIDAD", "LOC_CODIGO", "LOC_DESCRI")
    If cboLocalidad.ListCount > 0 Then
        cboLocalidad.ListIndex = -1
    End If
    
    'COMBO ESTADO SOCIO
    cSQL = "SELECT ESS_CODIGO, ESS_DESCRI FROM ESTADO_SOCIO ORDER BY ESS_DESCRI"
    Call CargoComboBoxItemData(cboEstado, cSQL)
    If cboEstado.ListCount > 0 Then
        cboEstado.ListIndex = -1
    End If
    
    'COMBO ESTADO CIVIL
    cSQL = "SELECT EST_CODIGO, EST_DESCRI FROM ESTADO_CIVIL ORDER BY EST_DESCRI"
    Call CargoComboBoxItemData(cboEstadoCivil, cSQL)
    If cboEstadoCivil.ListCount > 0 Then
        cboEstadoCivil.ListIndex = -1
    End If
    
    'COMBO SEXO
    cSQL = "SELECT SEX_CODIGO, SEX_DESCRI FROM SEXO ORDER BY SEX_DESCRI"
    Call CargoComboBoxItemData(cboSexo, cSQL)
    If cboSexo.ListCount > 0 Then
        cboSexo.ListIndex = -1
    End If
    
    'COMBO TIPO CUOTA
    cSQL = "SELECT TIC_CODIGO, TIC_DESCRI FROM TIPO_CUOTA ORDER BY TIC_DESCRI"
    Call CargoComboBoxItemData(cboTipoCuota, cSQL)
    If cboTipoCuota.ListCount > 0 Then
        cboTipoCuota.ListIndex = -1
    End If
    
    'COMBO TIPO CUOTA
    cSQL = "SELECT TIS_CODIGO, TIS_DESCRI FROM TIPO_SOCIO ORDER BY TIS_DESCRI"
    Call CargoComboBoxItemData(cboTipo, cSQL)
    If cboTipo.ListCount > 0 Then
        cboTipo.ListIndex = -1
    End If
    
    'CARGO COMBO TIPO DE DOCUMENTO
    cSQL = "SELECT TIP_TIPDOC FROM TIPO_DOCUMENTO"
    Rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Do While Rec.EOF = False
            cboTipoDoc.AddItem Trim(Rec!Tip_TipDoc)
            Rec.MoveNext
        Loop
    End If
    Rec.Close
    If cboTipoDoc.ListCount > 0 Then
        cboTipoDoc.ListIndex = -1
    End If
    
    'CARGO LISTA DE DEPORTES
    cSQL = "SELECT DEP_CODIGO, DEP_DESCRI FROM DEPORTE"
    Rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Do While Rec.EOF = False
            lstDeportes.AddItem Trim(Rec!DEP_DESCRI)
            lstDeportes.ItemData(lstDeportes.NewIndex) = Rec!DEP_CODIGO
            Rec.MoveNext
        Loop
    End If
    Rec.Close
    
    If vMode <> 1 Then
        If vFieldID <> "0" Then
            cSQL = "SELECT * FROM " & cTabla & "  WHERE SOC_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
            Rec1.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (Rec1.BOF And Rec1.EOF) = 0 Then
                'si encontró el registro muestro los datos
                txtID.Text = Rec1!SOC_CODIGO
                txtNombre.Text = Rec1!SOC_NOMBRE
                Fecha.Value = ChkNull(Rec1!SOC_FECNAC)
                Call BuscaCodigoProxItemData(CInt(Rec1!LOC_CODIGO), cboLocalidad)
                txtNroDoc.Text = ChkNull(Rec1!SOC_NRODOC)
                txtDomicilio.Text = ChkNull(Rec1!SOC_DOMICI)
                txtTelefono.Text = ChkNull(Rec1!SOC_TELEFONO)
                txtCodPostal.Text = ChkNull(Rec1!SOC_CODPOS)
                txtMail.Text = ChkNull(Rec1!SOC_MAIL)
                Call BuscaProx(Rec1!Tip_TipDoc, cboTipoDoc)
                txtNacionalidad.Text = ChkNull(Rec1!SOC_NACIONA)
                Call BuscaCodigoProxItemData(Rec1!SEX_CODIGO, cboSexo)
                Call BuscaCodigoProxItemData(Rec1!EST_CODIGO, cboEstadoCivil)
                txtHijos.Text = ChkNull(Rec1!SOC_HIJOS)
                txtProfesion.Text = ChkNull(Rec1!SOC_OCUPACION)
                txtCobroCuota.Text = ChkNull(Rec1!SOC_DOMCOB)
                Call BuscaCodigoProxItemData(Rec1!TIC_CODIGO, cboTipoCuota)
                Call BuscaCodigoProxItemData(Rec1!TIS_CODIGO, cboTipo)
                txtCodSoc.Text = ChkNull(Rec1!SOC_TITULAR)
                txtCodSoc_LostFocus
                Call BuscaCodigoProxItemData(Rec1!ESS_CODIGO, cboEstado)
                
                cSQL = "SELECT D.DEP_DESCRI, S.DEP_CODIGO"
                cSQL = cSQL & " FROM DEPORTE D, SOCIOS_DEPORTES S"
                cSQL = cSQL & " WHERE "
                cSQL = cSQL & " D.DEP_CODIGO=S.DEP_CODIGO"
                cSQL = cSQL & " AND S.SOC_CODIGO=" & XN(Rec1!SOC_CODIGO)
                Rec2.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
                If Rec2.EOF = False Then
                Do While Rec2.EOF = False
                    For I = 0 To lstDeportes.ListCount - 1
                        If Rec2!DEP_CODIGO = lstDeportes.ItemData(I) Then
                            lstDeportes.Selected(I) = True
                            Exit For
                        End If
                    Next
                    Rec2.MoveNext
                Loop
                End If
                Rec2.Close
                
            Else
                Beep
                MsgBox "Imposible encontrar el registro seleccionado.", vbCritical + vbOKOnly, App.Title
            End If
            Rec1.Close
        End If
    End If
    
    'establesco funcionalidad del form de datos
    SetMode vMode
End Sub

Private Sub lstDeportes_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtCobroCuota_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtCobroCuota_GotFocus()
    SelecTexto txtCobroCuota
End Sub

Private Sub txtCobroCuota_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtCodPostal_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtCodPostal_GotFocus()
    SelecTexto txtCodPostal
End Sub

Private Sub txtCodPostal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtCodSoc_Change()
    cmdAceptar.Enabled = True
    If txtCodSoc.Text = "" Then
        txtNomSoc.Text = ""
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
        sql = "SELECT SOC_CODIGO, SOC_NOMBRE "
        sql = sql & " FROM SOCIOS"
        sql = sql & " WHERE "
        sql = sql & " SOC_CODIGO =" & XN(txtCodSoc.Text)
        If Rec.State = 1 Then Rec.Close
        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec.EOF = False Then
            txtNomSoc.Text = ChkNull(Rec!SOC_NOMBRE)
        Else
            MsgBox "El Código no existe", vbInformation
            txtNomSoc.Text = ""
            txtCodSoc.Text = ""
            txtCodSoc.SetFocus
        End If
        If Rec.State = 1 Then Rec.Close
    End If
End Sub

Private Sub txtDomicilio_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtDomicilio_GotFocus()
    SelecTexto txtDomicilio
End Sub

Private Sub txtDomicilio_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtHijos_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtHijos_GotFocus()
    SelecTexto txtHijos
End Sub

Private Sub txtHijos_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtMail_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtMail_GotFocus()
    SelecTexto txtMail
End Sub

Private Sub txtNacionalidad_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtNacionalidad_GotFocus()
    SelecTexto txtNacionalidad
End Sub

Private Sub txtNacionalidad_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtNombre_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtNombre_GotFocus()
    seltxt
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
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
                'cSQL = cSQL & " WHERE PAI_CODIGO = " & cboPais.ItemData(cboPais.ListIndex)
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
            cSQL = "SELECT COUNT(*) FROM " & cTabla & " WHERE " & cCampoID & " = " & XN(txtID.Text)
            'cSQL = cSQL & " AND PAI_CODIGO = " & cboPais.ItemData(cboPais.ListIndex)
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

Private Sub txtNomSoc_Change()
    cmdAceptar.Enabled = True
    If txtNomSoc.Text = "" Then
        txtCodSoc.Text = ""
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

Private Sub txtNomSoc_LostFocus()
    If txtCodSoc.Text = "" And txtNomSoc.Text <> "" Then
        sql = "SELECT SOC_CODIGO,SOC_NOMBRE"
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

Private Sub txtNroDoc_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtNroDoc_GotFocus()
    SelecTexto txtNroDoc
End Sub

Private Sub txtNroDoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtProfesion_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtProfesion_GotFocus()
    SelecTexto txtProfesion
End Sub

Private Sub txtProfesion_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub
