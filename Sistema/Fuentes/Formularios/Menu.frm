VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm Menu 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema Pilar Sport Club"
   ClientHeight    =   6270
   ClientLeft      =   105
   ClientTop       =   2085
   ClientWidth     =   11100
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "Menu.frx":0CCA
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbrPrincipal 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   12
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Socios"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Deportes"
            Object.Tag             =   ""
            ImageIndex      =   23
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Registro de Pagos"
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Organizar Ventanas"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar stbPrincipal 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   5955
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   556
      SimpleText      =   "Listo."
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   6526
            MinWidth        =   6526
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7673
            MinWidth        =   7673
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "NÚM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            Bevel           =   2
            Enabled         =   0   'False
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "MAYÚS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "12:21"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "28/07/2018"
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   23
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":49D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":4CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":5004
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":531E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":5638
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":5952
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":5C6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":5F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":62A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":65BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":68D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":6BEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":6F08
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":7222
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":753C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":7856
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":7B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":7E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":81A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":84BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":87D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":8AF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":8E0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArc 
      Caption         =   "&Archivos"
      Begin VB.Menu mnuArchivoActualizaciones 
         Caption         =   "Actualizaciones"
         Begin VB.Menu mnuSexo 
            Caption         =   "Sexo"
         End
         Begin VB.Menu mnuParentesco 
            Caption         =   "Parentesco"
         End
         Begin VB.Menu mnuEstadoCivil 
            Caption         =   "Estado Civil"
         End
         Begin VB.Menu mnuEmpleados 
            Caption         =   "Empleados"
         End
         Begin VB.Menu mnuABMLocalidades 
            Caption         =   "Localidades"
         End
         Begin VB.Menu mnuRaya1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTipoCuota 
            Caption         =   "Tipo Cuota"
         End
         Begin VB.Menu mnuDeportes 
            Caption         =   "Deportes"
         End
         Begin VB.Menu mnuClientes 
            Caption         =   "Socios"
         End
         Begin VB.Menu mnuABMEstadoDocumento 
            Caption         =   "Estado de Documentos"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuABMFormaPago 
            Caption         =   "Forma de Pago"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuClave 
         Caption         =   "Clave"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRaya111 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConectar 
         Caption         =   "Conectar"
      End
      Begin VB.Menu mnuDesconectar 
         Caption         =   "Desconectar"
      End
      Begin VB.Menu mnuRaya4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUsuarios 
         Caption         =   "&Usuarios"
      End
      Begin VB.Menu mnuPermisos 
         Caption         =   "Permi&sos"
      End
      Begin VB.Menu mnuRaya16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParametros 
         Caption         =   "Parametros"
      End
      Begin VB.Menu mnuDebitosManuales 
         Caption         =   "Débitos Manuales"
      End
      Begin VB.Menu mnuRayaSalir 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArcSal 
         Caption         =   "&Salir"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuFondos 
      Caption         =   "&Fondos"
      Begin VB.Menu mnuFondosActualizaciones 
         Caption         =   "Actualizaciones"
         Begin VB.Menu mnuTipoIngreso 
            Caption         =   "Tipo Ingreso"
         End
         Begin VB.Menu mnuTipoEgreso 
            Caption         =   "Tipo Egreso"
         End
         Begin VB.Menu mnuRaya10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuABMBancos 
            Caption         =   "Bancos"
         End
         Begin VB.Menu mnuTipoCuentas 
            Caption         =   "Tipos de &Cuentas"
         End
         Begin VB.Menu mnuTiposMoneda 
            Caption         =   "Tipos de &Moneda"
         End
         Begin VB.Menu mnuTipoGastosBancarios 
            Caption         =   "Tipos de Gastos &Bancarios"
         End
         Begin VB.Menu mnuABMDebitosCreditosBancarios 
            Caption         =   "Tipos de Débitos y Créditos Bancarios"
         End
         Begin VB.Menu mnuEstadosCheques 
            Caption         =   "Tipos de Estados Cheque"
         End
      End
      Begin VB.Menu mnuRaya11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFondosGestonBancaria 
         Caption         =   "Gestión &Bancaria"
         Visible         =   0   'False
         Begin VB.Menu mnuABMCuentasBancarias 
            Caption         =   "Cuentas Bancarias"
         End
         Begin VB.Menu mnuRaya15 
            Caption         =   "-"
         End
         Begin VB.Menu mnuChequesPropios 
            Caption         =   "Cheques &Propios"
            Begin VB.Menu mnuCargaChequesPropios 
               Caption         =   "Carga de Cheques"
            End
            Begin VB.Menu mnuCambioEstadoChequesPropios 
               Caption         =   "Cambio de Estado Cheques"
            End
            Begin VB.Menu mnuListadoChequespropios 
               Caption         =   "Listado de Cheques"
            End
         End
         Begin VB.Menu mnuChequesDeTerceros 
            Caption         =   "Cheques de &Terceros"
            Begin VB.Menu mnuCargaChequesTerceros 
               Caption         =   "Carga de Cheques"
            End
            Begin VB.Menu mnuCambioEstadoChequesTerceros 
               Caption         =   "Cambio de Estados Cheques"
            End
            Begin VB.Menu mnuListadoChequesTerceros 
               Caption         =   "Listado Cheques"
            End
         End
         Begin VB.Menu mnuRaya12 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBoletaDeposito 
            Caption         =   "&Boleta Déposito"
         End
         Begin VB.Menu mnuIngresoGastosBancarios 
            Caption         =   "&Ingreso de Gastos Bancarios"
         End
         Begin VB.Menu mnuIngresoDebCreBancarios 
            Caption         =   "Ingreso de Débitos y Créditos Bancarios"
         End
         Begin VB.Menu mnuRaya13 
            Caption         =   "-"
         End
         Begin VB.Menu mnuResumenCuenta 
            Caption         =   "&Resumen de Cuenta"
         End
      End
      Begin VB.Menu mnuCaja 
         Caption         =   "Gestión &Caja"
         Begin VB.Menu mnuIngresos 
            Caption         =   "Carga de Ingresos"
         End
         Begin VB.Menu mnuEgresos 
            Caption         =   "Carga de Egresos"
         End
         Begin VB.Menu mnuRaya14 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCierreCaja 
            Caption         =   "Cierre de Caja"
         End
         Begin VB.Menu mnuListadoIngresosEgresos 
            Caption         =   "Listado de Ingresos y Egresos"
         End
      End
      Begin VB.Menu mnuRaya30 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDebitos 
         Caption         =   "Generar Débitos"
      End
      Begin VB.Menu mnuCtaCte 
         Caption         =   "Cta Cte x Cliente"
      End
      Begin VB.Menu mnuCtaCteDeporte 
         Caption         =   "Cta Cte x Deporte"
      End
      Begin VB.Menu mnuRaya2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPagos 
         Caption         =   "Registro Pagos"
      End
      Begin VB.Menu mnuListadoCobranza 
         Caption         =   "Listado de Cobranza"
      End
      Begin VB.Menu mnuListadoCobranzaDeporte 
         Caption         =   "Listado de Cobranza x Deporte"
      End
      Begin VB.Menu mnuRaya19 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComisiones 
         Caption         =   "Comisiones"
      End
   End
   Begin VB.Menu mnuListados 
      Caption         =   "&Listados"
      Begin VB.Menu mnuLstDeportesSocios 
         Caption         =   "Socios x Deportes"
      End
   End
   Begin VB.Menu mnuVentana 
      Caption         =   "&Ventana"
      WindowList      =   -1  'True
      Begin VB.Menu mnuMosHoriz 
         Caption         =   "Mosaico &horizontal"
      End
      Begin VB.Menu mnuMosVert 
         Caption         =   "Mosaico &vertical"
      End
      Begin VB.Menu mnuCascada 
         Caption         =   "&Cascada"
      End
      Begin VB.Menu mnuIconos 
         Caption         =   "Organizar &Iconos"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "A&yuda"
      Begin VB.Menu mnuContenido 
         Caption         =   "&Contenido"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAAcerca 
         Caption         =   "&Acerca de..."
      End
   End
   Begin VB.Menu ContextBaseABM 
      Caption         =   "ContextBaseABM"
      Visible         =   0   'False
      Begin VB.Menu mnuContextABM 
         Caption         =   "Nuevo"
         Index           =   0
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "Editar"
         Index           =   1
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "Eliminar"
         Index           =   2
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "Refrescar"
         Index           =   4
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "Buscar"
         Index           =   6
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "Imprimir"
         Index           =   7
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "Ver Datos"
         Index           =   9
      End
   End
   Begin VB.Menu ContextABMCta 
      Caption         =   "ContextABMCta"
      Visible         =   0   'False
      Begin VB.Menu mnuContextABMCta 
         Caption         =   "Nuevo"
         Index           =   0
      End
      Begin VB.Menu mnuContextABMCta 
         Caption         =   "Editar"
         Index           =   1
      End
      Begin VB.Menu mnuContextABMCta 
         Caption         =   "Eliminar"
         Index           =   2
      End
      Begin VB.Menu mnuContextABMCta 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuContextABMCta 
         Caption         =   "Refrescar"
         Index           =   4
      End
      Begin VB.Menu mnuContextABMCta 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuContextABMCta 
         Caption         =   "Ver Datos"
         Index           =   6
      End
   End
   Begin VB.Menu ContextABMPresu 
      Caption         =   "ContextABMPresu"
      Visible         =   0   'False
      Begin VB.Menu mnuContextABMPresu 
         Caption         =   "Nuevo"
         Index           =   0
      End
      Begin VB.Menu mnuContextABMPresu 
         Caption         =   "Editar"
         Index           =   1
      End
      Begin VB.Menu mnuContextABMPresu 
         Caption         =   "Eliminar"
         Index           =   2
      End
      Begin VB.Menu mnuContextABMPresu 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuContextABMPresu 
         Caption         =   "Refrescar"
         Index           =   4
      End
      Begin VB.Menu mnuContextABMPresu 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuContextABMPresu 
         Caption         =   "Ver Datos"
         Index           =   6
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public USUARIO_LOCAL As String
Dim TituloPrincipal As String
Dim mBlak As Boolean

Private Declare Function ShellAbout Lib "shell32.dll" Alias _
"ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, _
ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

Private Sub MDIForm_Load()
'    If Dir("c:\windows\cpce.ini") = "" Then
'        Menu.Picture = LoadPicture(App.Path & "\fotos\Demaría.bmp")
'    End If
    
    TituloPrincipal = TIT_MSGBOX '"Sistema de Gestión y Administración"
    Me.Caption = TituloPrincipal
    
    If VeoConfiguracionRegional = False Then
        MsgBox "Problemas en la CONFIGURACION REGIONAL !!!" & Chr(13) & "El Sistema se CERRARÁ para configurar correctamente su PC ", vbCritical, "ERROR !"
        End
    End If
    
    Me.Show
    FrmInicio.Show vbModal
    
    'Me.Caption = TituloPrincipal & " - (Usuario " & UCase(mNomUser) & " conectado a " & UCase(SERVIDOR) & " - " & UCase(BASEDATO) & ")"
    Me.Caption = TituloPrincipal & "    V. " & App.Major & "." & App.Minor & "." & App.Revision & "          - (Usuario " & UCase(mNomUser) & " conectado a " & UCase(SERVIDOR) & " - " & UCase(BASEDATO) & ")"
    Menu.mnuConectar.Enabled = False
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Call mnuArcSal_Click
End Sub

Private Sub mnuAAcerca_Click()
    Call ShellAbout(Me.hWnd, "Sistema de Gestión Administrativo", "Copyright 2008", Me.Icon)
End Sub

Private Sub mnuABMBancos_Click()
    Dim cSQL As String
    
    mOrigen = True
    
    Set vABMBancos = New CListaBaseABM
    
    With vABMBancos
        .Caption = "Actualización de Bancos"
        .sql = "SELECT BAN_DESCRI, BAN_BANCO, BAN_LOCALIDAD, BAN_SUCURSAL, BAN_CODIGO," & _
                "BAN_CODINT FROM BANCO"
        .HeaderSQL = "Descripción, Banco, Tipo , Sucursal, Código, Código Int."
        .FieldID = "BAN_CODINT"
        '.Report = RptPath & "tipoiva.rpt"
        Set .FormBase = vFormBancos
        Set .FormDatos = ABMBancos
    End With
    
    Set auxDllActiva = vABMBancos
    
    vABMBancos.Show
End Sub

Private Sub mnuABMCuentasBancarias_Click()
    Dim cSQL As String
    
    mOrigen = True
    
    Set vABMCuentaBancaria = New CListaBaseABM
    
    With vABMCuentaBancaria
        .Caption = "Actualización de Cuentas Bancarias"
        .sql = "SELECT B.BAN_DESCRI, TC.TCU_DESCRI, C.CTA_NROCTA, B.BAN_CODINT" & _
                " FROM BANCO B, TIPO_CUENTA TC, CTA_BANCARIA C" & _
                " WHERE C.BAN_CODINT=B.BAN_CODINT AND C.TCU_CODIGO=TC.TCU_CODIGO"
        .HeaderSQL = "Banco, Tipo Cuenta, Nro. Cuenta, Código Banco"
        .FieldID = "CTA_NROCTA"
        '.Report = RptPath & "tipoiva.rpt"
        Set .FormBase = vFormCuentaBancaria
        Set .FormDatos = ABMCuentaBancaria
    End With
    
    Set auxDllActiva = vABMCuentaBancaria
    
    vABMCuentaBancaria.Show
End Sub

Private Sub mnuABMDebitosCreditosBancarios_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMDebCreBancario = New CListaBaseABM
    
    With vABMDebCreBancario
        .Caption = "Actualización de Débitos y Créditos Bancarios"
        .sql = "SELECT TDCB_DESCRI, TDCB_CODIGO, TDCB_TIPO FROM TIPO_DEBCRE_BANCARIO"
        .HeaderSQL = "Descripción, Código, Deb-Cre"
        .FieldID = "TDCB_CODIGO"
        '.Report = RptPath & "cuenta.rpt"
        Set .FormBase = vFormDebCreBancario
        Set .FormDatos = ABMDebCreBancario
    End With
    
    Set auxDllActiva = vABMDebCreBancario
    
    vABMDebCreBancario.Show
End Sub

Private Sub mnuABMEstadoDocumento_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMEstadoDocumento = New CListaBaseABM
    
    With vABMEstadoDocumento
        .Caption = "Actualización de Estado Documento"
        .sql = "SELECT EST_DESCRI, EST_CODIGO FROM ESTADO_DOCUMENTO"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "EST_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormEstadoDocumento
        Set .FormDatos = ABMEstadoDocumento
    End With
    
    Set auxDllActiva = vABMEstadoDocumento
    
    vABMEstadoDocumento.Show
End Sub

Private Sub mnuABMFormaPago_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMFormaPago = New CListaBaseABM
    
    With vABMFormaPago
        .Caption = "Actualización de Estado Documento"
        .sql = "SELECT FPG_DESCRI, FPG_CODIGO FROM FORMA_PAGO"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "FPG_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormFormaPago
        Set .FormDatos = ABMFormaPago
    End With
    
    Set auxDllActiva = vABMFormaPago
    
    vABMFormaPago.Show
End Sub

Private Sub mnuABMLocalidades_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMLocalidad = New CListaBaseABM
    
    With vABMLocalidad
        .Caption = "Actualización de Localidad"
        .sql = "SELECT LOC_DESCRI, LOC_CODIGO FROM LOCALIDAD"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "LOC_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormLocalidad
        Set .FormDatos = ABMLocalidad
    End With
    
    Set auxDllActiva = vABMLocalidad
    
    vABMLocalidad.Show
End Sub

Private Sub mnuCategoria_Click()

End Sub

Private Sub mnuBoletaDeposito_Click()
    FrmBoletaDeposito.Show vbModal
End Sub

Private Sub mnuCierreCaja_Click()
    frmCaja.Show vbModal
End Sub

Private Sub mnuClientes_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMSocios = New CListaBaseABM
    
    With vABMSocios
        .Caption = "Actualización de Socios"
        .sql = "SELECT SOC_NOMBRE, SOC_CODIGO, SOC_DOMICI, SOC_TELEFONO, SOC_MAIL" & _
               " FROM SOCIOS"
        .HeaderSQL = "Nombre, Código, Domicilio, Teléfono, e-mail"
        .FieldID = "SOC_CODIGO"
        .Report = DRIVE & DirReport & "Socios.rpt"
        Set .FormBase = vFormSocios
        Set .FormDatos = ABMSocios
    End With
    
    Set auxDllActiva = vABMSocios
    FormLlamado = ""
    
    vABMSocios.Show
End Sub

Private Sub mnuArcSal_Click()
    On Error Resume Next
    DBConn.CloseConnection
    Set DBConn = Nothing
    Set Menu = Nothing
    End
    'verifico si la conexión esta abierta antes de salir
    'If Me.mnuConexion.Enabled = False Then
End Sub

Private Sub mnuCambioEstadoChequesPropios_Click()
    ABMCambioEstadoChPropio.Show vbModal
End Sub

Private Sub mnuCambioEstadoChequesTerceros_Click()
    ABMCambioEstado.Show vbModal
End Sub

Private Sub mnuCargaChequesPropios_Click()
    FrmCargaChequesPropios.Show vbModal
End Sub

Private Sub mnuCargaChequesTerceros_Click()
    FrmCargaCheques.Show vbModal
End Sub

Private Sub mnuCascada_Click()
    Me.Arrange 0
End Sub

Private Sub mnuComisiones_Click()
    frmComisiones.Show
End Sub

Private Sub mnuConectar_Click()
    FrmInicio.Show vbModal
    Me.Caption = TituloPrincipal & " - (Usuario " & UCase(mNomUser) & " conectado a " & UCase(SERVIDOR) & ")"
    Me.mnuConectar.Enabled = False
End Sub

Private Sub mnuContenido_Click()
    Call WinHelp(Me.hWnd, App.Path & "\help\AYUDA.HLP", HelpFinder, 0&)
End Sub

Public Sub mnuContextABM_Click(Index As Integer)

Dim auxListView As ListView
Dim auxModo As Integer
    
    auxModo = 0
    Select Case Index
        Case 0 'nuevo
            auxModo = 1
        Case 1 'editar
            auxModo = 2
        Case 2 'eliminar
            auxModo = 4
        Case 9 ' ver datos
            auxModo = 3
        'Case 7 ' imprimir
        '   auxModo = 7
    End Select
    
    If auxModo > 0 Then
        Set auxListView = auxDllActiva.FormBase.lstvLista
        auxDllActiva.FormDatos.SetWindow auxDllActiva.FormBase, auxDllActiva.sql, auxModo, auxListView, auxDllActiva.FieldID
        auxDllActiva.FormDatos.Show vbModal
    Else
        'si es una acción de edición de datos
        Select Case Index
            Case 4 'refresh
                Screen.MousePointer = vbHourglass
                With auxDllActiva
                    Set auxListView = .FormBase.lstvLista
                    CargarListView .FormBase, auxListView, .sql, .FieldID, .HeaderSQL, .FormBase.ImgLstLista
                    .FormBase.sBarEstado.Panels(1).Text = auxListView.ListItems.Count & " Registro(s)"
                End With
                Screen.MousePointer = vbDefault

            Case 5 'refresh
                'auxDllActiva.FormBase.txtBusqueda.Text = ""
                'auxDllActiva.FormBase.fraFiltro.Visible = True
                'auxDllActiva.FormBase.txtBusqueda.SetFocus
                With auxDllActiva
'                    If .Caption = "Actualización de Productos" Then
'                        frmFiltroProducto.Show
'                    Else
                        frmFiltro.Show
'                    End If
                End With

            Case 6 'Buscar
                    auxDllActiva.Find
                
            Case 7 'imprimir
                'Select Case mQuienLlamo
'                    Case "ABMProducto"
'                        frmImprimeProducto.Show vbModal
'                    Case Else
                        On Error GoTo ErrorReport
                        auxDllActiva.FormBase.rptListado.Action = 1
                        On Error GoTo 0
'                End Select
        End Select
    End If
    Exit Sub
    
ErrorReport:
    
    Beep
    MsgBox "Error " & Err.Number & Chr(13) & Err.Description, vbCritical + vbOKOnly, App.Title
    
End Sub


Private Sub mnuCtaCte_Click()
    frmCtaCte.Show
End Sub

Private Sub mnuCtaCteDeporte_Click()
    frmCtaCteDeporte.Show
End Sub

Private Sub mnuDebitos_Click()
    frmGenerarDebitos.Show
End Sub

Private Sub mnuDebitosManuales_Click()
    frmDebitosManuales.Show vbModal
End Sub

Private Sub mnuDeportes_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMDeporte = New CListaBaseABM
    
    With vABMDeporte
        .Caption = "Actualización de Deporte"
        .sql = "SELECT DEP_DESCRI, DEP_CODIGO, DEP_CUOTA FROM DEPORTE"
        .HeaderSQL = "Descripción, Código, Cuota"
        .FieldID = "DEP_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormDeporte
        Set .FormDatos = ABMDeporte
    End With
    
    Set auxDllActiva = vABMDeporte
    
    vABMDeporte.Show
End Sub

Private Sub mnuDesconectar_Click()
    If DBConn.State = adStateOpen Then
        DBConn.Close
        
        DeshabilitarMenu Me
        
        Me.mnuArc.Enabled = True
        Me.mnuConectar.Enabled = True
        Me.mnuArcSal.Enabled = True
        Me.mnuDesconectar.Enabled = False
        
        Me.Caption = TituloPrincipal & " - (No conectado)"
    End If
End Sub

Private Sub mnuEgresos_Click()
    ABMEgresos.Show
End Sub

Private Sub mnuEmpleados_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMEmpleados = New CListaBaseABM
    
    With vABMEmpleados
        .Caption = "Actualización de Empleados"
        .sql = "SELECT EMP_NOMBRE, EMP_CODIGO, EMP_DOMICI, EMP_TELEFONO, EMP_MAIL" & _
               " FROM EMPLEADOS"
        .HeaderSQL = "Nombre, Código, Domicilio, Teléfono, e-mail"
        .FieldID = "EMP_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormEmpleados
        Set .FormDatos = ABMEmpleado
    End With
    
    Set auxDllActiva = vABMEmpleados
    
    vABMEmpleados.Show
End Sub


Private Sub mnuEstadoCivil_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMEstadoCivil = New CListaBaseABM
    
    With vABMEstadoCivil
        .Caption = "Actualización de Estado Civil"
        .sql = "SELECT EST_DESCRI, EST_CODIGO FROM ESTADO_CIVIL"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "EST_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormEstadoCivil
        Set .FormDatos = ABMEstadoCivil
    End With
    
    Set auxDllActiva = vABMEstadoCivil
    
    vABMEstadoCivil.Show
End Sub

Private Sub mnuEstadosCheques_Click()
    Dim cSQL As String
    
    mOrigen = True
    
    Set vABMEstadosCheques = New CListaBaseABM
    
    With vABMEstadosCheques
        .Caption = "Actualización de Tipos de Estados Cheques"
        .sql = "SELECT ECH_DESCRI, ECH_CODIGO FROM ESTADO_CHEQUE"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "ECH_CODIGO"
        '.Report = RptPath & "tipoiva.rpt"
        Set .FormBase = vFormEstadosCheques
        Set .FormDatos = ABMEstadoCheque
    End With
    
    Set auxDllActiva = vABMEstadosCheques
    
    vABMEstadosCheques.Show
End Sub

Private Sub mnuIconos_Click()
    Me.Arrange 3
End Sub

Private Sub mnuIngresoDebCreBancarios_Click()
    frmDebCreBancarios.Show
End Sub

Private Sub mnuIngresoGastosBancarios_Click()
    frmIngresoGastosBancarios.Show
End Sub

Private Sub mnuIngresos_Click()
    ABMIngresos.Show 'vbModal
End Sub

Private Sub mnuListadoChequespropios_Click()
    FrmListChequesPropios.Show
End Sub

Private Sub mnuListadoChequesTerceros_Click()
    FrmListCheques.Show
End Sub

Private Sub mnuListadoCobranza_Click()
    frmListadoCobranza.Show
End Sub

Private Sub mnuListadoCobranzaDeporte_Click()
    frmListadoCobranzaDeporte.Show
End Sub

Private Sub mnuListadoIngresosEgresos_Click()
    frmListadoIngersosEgersos.Show
End Sub

Private Sub mnuLstDeportesSocios_Click()
    frmListadoSociosxDeporte.Show
End Sub

Private Sub mnuMosHoriz_Click()
    Me.Arrange 1
End Sub

Private Sub mnuMosVert_Click()
    Me.Arrange 2
End Sub

Private Sub mnuPagos_Click()
    frmRecibo.Show vbModal
End Sub

Private Sub mnuParametros_Click()
    frmParametros.Show
End Sub

Private Sub mnuParentesco_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMParentesco = New CListaBaseABM
    
    With vABMParentesco
        .Caption = "Actualización de Parentesco"
        .sql = "SELECT PAR_DESCRI, PAR_CODIGO FROM PARENTESCO"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "PAR_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormParentesco
        Set .FormDatos = ABMParentesco
    End With
    
    Set auxDllActiva = vABMParentesco
    
    vABMParentesco.Show
End Sub

Private Sub mnuPermisos_Click()
    FrmPermisos.Show vbModal
End Sub

Private Sub mnuResumenCuenta_Click()
    frmResumenCuentaBanco.Show
End Sub

Private Sub mnuSexo_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMSexo = New CListaBaseABM
    
    With vABMSexo
        .Caption = "Actualización de Sexo"
        .sql = "SELECT SEX_DESCRI, SEX_CODIGO FROM SEXO"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "SEX_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormSexo
        Set .FormDatos = ABMSexo
    End With
    
    Set auxDllActiva = vABMSexo
    
    vABMSexo.Show
End Sub

Private Sub mnuTipoCuentas_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMTCuenta = New CListaBaseABM
    
    With vABMTCuenta
        .Caption = "Actualización de Tipos de Cuentas"
        .sql = "SELECT TCU_DESCRI, TCU_CODIGO FROM TIPO_CUENTA"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "TCU_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormTCuenta
        Set .FormDatos = ABMTipoCuentas
    End With
    
    Set auxDllActiva = vABMTCuenta
    
    vABMTCuenta.Show
End Sub

Private Sub mnuTipoCuota_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMTipoCuota = New CListaBaseABM
    
    With vABMTipoCuota
        .Caption = "Actualización de Tipo Cuota"
        .sql = "SELECT TIC_DESCRI, TIC_CODIGO, TIC_CUOTA FROM TIPO_CUOTA"
        .HeaderSQL = "Descripción, Código, Cuota"
        .FieldID = "TIC_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormTipoCuota
        Set .FormDatos = ABMTipoCuota
    End With
    
    Set auxDllActiva = vABMTipoCuota
    
    vABMTipoCuota.Show
End Sub

Private Sub mnuTipoEgreso_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMTipoEgreso = New CListaBaseABM
    
    With vABMTipoEgreso
        .Caption = "Actualización de Tipo Egreso"
        .sql = "SELECT TEG_DESCRI, TEG_CODIGO FROM TIPO_EGRESO"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "TEG_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormTipoEgreso
        Set .FormDatos = ABMTipoEgreso
    End With
    
    Set auxDllActiva = vABMTipoEgreso
    
    vABMTipoEgreso.Show
End Sub

Private Sub mnuTipoGastosBancarios_Click()
    Dim cSQL As String
    
    mOrigen = True
       
    Set vABMGastoBancario = New CListaBaseABM
    
    With vABMGastoBancario
        .Caption = "Actualización de Gastos Bancarios"
        .sql = "SELECT TGB_DESCRI, TGB_CODIGO FROM TIPO_GASTO_BANCARIO"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "TGB_CODIGO"
        '.Report = RptPath & "cuenta.rpt"
        Set .FormBase = vFormGastoBancario
        Set .FormDatos = ABMTipoGastosBancarios
    End With
    
    Set auxDllActiva = vABMGastoBancario
    
    vABMGastoBancario.Show
End Sub

Private Sub mnuTipoIngreso_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMTipoIngreso = New CListaBaseABM
    
    With vABMTipoIngreso
        .Caption = "Actualización de Tipo Ingreso"
        .sql = "SELECT TIG_DESCRI, TIG_CODIGO FROM TIPO_INGRESO"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "TIG_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormTipoIngreso
        Set .FormDatos = ABMTipoIngreso
    End With
    
    Set auxDllActiva = vABMTipoIngreso
    
    vABMTipoIngreso.Show
End Sub

Private Sub mnuTiposMoneda_Click()
    Dim cSQL As String
    
    mOrigen = True
    
    Set vABMTMONEDA = New CListaBaseABM
    
    With vABMTMONEDA
        .Caption = "Actualización de Tipos de Moneda"
        .sql = "SELECT MON_DESCRI, MON_CODIGO FROM MONEDA"
        .HeaderSQL = "Descripción, Código"
        .FieldID = "MON_CODIGO"
        '.Report = RptPath & "tipoiva.rpt"
        Set .FormBase = vFormTMONEDA
        Set .FormDatos = ABMMoneda
    End With
    
    Set auxDllActiva = vABMTMONEDA
    
    vABMTMONEDA.Show
End Sub

Private Sub mnuUsuarios_Click()
    FrmUsuarios.Show vbModal
End Sub


Private Sub stbPrincipal_PanelClick(ByVal Panel As ComctlLib.Panel)
    If Panel.Index = 6 Then
        'stbPrincipal_MouseUp
        mBlak = True
    Else
        mBlak = False
    End If
End Sub

Private Sub tbrPrincipal_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Index
        Case 2: Call mnuClientes_Click
        Case 3: Call mnuDeportes_Click
'        Case 4: Call mnuFacturacionFacturacion_Click
'        Case 6: Call mnuPedidosProveedores_Click
'        Case 7: Call mnuEnvioComposturas_Click
'        Case 8: Call mnuConsignaciones_Click
'        'Case 10: Call mnuReciboClientes_Click
'        Case 10: Call mnuCtaCteVieja_Click
'        Case 12: frmAdmListados.Show
'        Case 13: Call mnuABMClientes_Click
        Case 5: Call mnuPagos_Click
        Case 6: Call mnuCascada_Click
    End Select
End Sub

Public Function VeoConfiguracionRegional() As Boolean
    VeoConfiguracionRegional = True

    If (ActualConfgRegional(LOCALE_SDECIMAL) <> ",") Then
        PonerConfgRegional LOCALE_SDECIMAL, ","
        VeoConfiguracionRegional = False
    End If
    If (ActualConfgRegional(LOCALE_STHOUSAND) <> ".") Then
        PonerConfgRegional LOCALE_STHOUSAND, "."
        VeoConfiguracionRegional = False
    End If
    If (ActualConfgRegional(LOCALE_SMONDECIMALSEP) <> ",") Then
        PonerConfgRegional LOCALE_SMONDECIMALSEP, ","
        VeoConfiguracionRegional = False
    End If
    If (ActualConfgRegional(LOCALE_SMONTHOUSANDSEP) <> ".") Then
        PonerConfgRegional LOCALE_SMONTHOUSANDSEP, "."
        VeoConfiguracionRegional = False
    End If
    If (ActualConfgRegional(LOCALE_SSHORTDATE) <> "dd/MM/yyyy") Then
        PonerConfgRegional LOCALE_SSHORTDATE, "dd/MM/yyyy"
        VeoConfiguracionRegional = False
    End If
End Function

Public Function ActualConfgRegional(lngTipo As Long) As String
    Dim lngBufferLen As Long
    Dim intRetorno As Integer
    Dim strBuffer As String
    On Error GoTo ActualConfgRegional_err

    lngBufferLen = 50
    strBuffer = Space$(lngBufferLen)

    intRetorno = GetLocaleInfo(LOCALE_USER_DEFAULT, lngTipo, strBuffer, lngBufferLen)
    'intRetorno = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, lngTipo, strBuffer, lngBufferLen)
    
    ActualConfgRegional = Left$(strBuffer, InStr(strBuffer, Chr(0)) - 1)

    Exit Function
ActualConfgRegional_err:
    'MensajeError "ActualConfgRegional", " Editando valor " & CStr(lngTipo)
End Function

Public Sub PonerConfgRegional(lngTipo As Long, strNuevoValor As String)
    Dim intRetorno As Integer
    On Error GoTo PonerConfgRegional_err

    intRetorno = SetLocaleInfo(LOCALE_USER_DEFAULT, lngTipo, strNuevoValor)

    Exit Sub
PonerConfgRegional_err:
    'MensajeError "PonerConfgRegional", " Estableciendo valor " & CStr(lngTipo)
End Sub

Public Function ConfgRegionalCorrecta() As Boolean
    On Error GoTo ConfgRegionalCorrecta_err

    ConfgRegionalCorrecta = True

    If (ActualConfgRegional(LOCALE_ICURRDIGITS) <> "3") Then
        ConfgRegionalCorrecta = False
    Else
        If (ActualConfgRegional(LOCALE_SSHORTDATE) <> "dd/MM/yyyy") Then
            ConfgRegionalCorrecta = False
        Else
            If (ActualConfgRegional(LOCALE_SCURRENCY) <> "pts") Then
                ConfgRegionalCorrecta = False
            Else
                If (ActualConfgRegional(LOCALE_SDATE) <> "/") Then
                    ConfgRegionalCorrecta = False
                Else
                    If (ActualConfgRegional(LOCALE_SDECIMAL) <> ",") Then
                        ConfgRegionalCorrecta = False
                    Else
                        If (ActualConfgRegional(LOCALE_STHOUSAND) <> ".") Then
                            ConfgRegionalCorrecta = False
                        End If
                    End If
                End If
            End If
        End If
    End If

    Exit Function
ConfgRegionalCorrecta_err:
    'MensajeError "ConfgRegionalCorrecta"
End Function

Public Sub AjustarConfgReg()
    On Error GoTo AjustarConfgReg_err

    'PonerConfgRegional LOCALE_ICURRDIGITS, "3"
    PonerConfgRegional LOCALE_SSHORTDATE, "dd/MM/yyyy"
    'PonerConfgRegional LOCALE_SCURRENCY, "pts"
    'PonerConfgRegional LOCALE_SDATE, "/"
    PonerConfgRegional LOCALE_SDECIMAL, ","
    PonerConfgRegional LOCALE_STHOUSAND, "."

    Exit Sub
AjustarConfgReg_err:
    'MensajeError "AjustarConfgReg"
End Sub


