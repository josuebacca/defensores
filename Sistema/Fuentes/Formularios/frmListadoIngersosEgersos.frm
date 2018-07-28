VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListadoIngersosEgersos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Ingersos y Egresos"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
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
   ScaleHeight     =   2310
   ScaleWidth      =   6255
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
      TabIndex        =   8
      Top             =   1515
      Width           =   2175
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "frmListadoIngersosEgersos.frx":0000
         Left            =   450
         List            =   "frmListadoIngersosEgersos.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   270
         Width           =   1635
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
         Picture         =   "frmListadoIngersosEgersos.frx":002F
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   11
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
         Picture         =   "frmListadoIngersosEgersos.frx":0131
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
         Index           =   2
         Left            =   135
         Picture         =   "frmListadoIngersosEgersos.frx":0233
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   9
         Top             =   315
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   435
      Left            =   3540
      TabIndex        =   0
      Top             =   1815
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   5310
      TabIndex        =   2
      Top             =   1815
      Width           =   840
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoIngersosEgersos.frx":0335
      Height          =   435
      Left            =   4425
      TabIndex        =   1
      Top             =   1815
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listar por..."
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
      Left            =   60
      TabIndex        =   3
      Top             =   -15
      Width           =   6105
      Begin VB.ComboBox CboGastos 
         Height          =   315
         Left            =   1245
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   690
         Width           =   4650
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   1245
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   345
         Width           =   2700
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   1245
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   111017985
         CurrentDate     =   42925
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   4590
         TabIndex        =   16
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   111017985
         CurrentDate     =   42925
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   405
         Width           =   360
      End
      Begin VB.Label lblFechaDesde 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   225
         TabIndex        =   6
         Top             =   1110
         Width           =   990
      End
      Begin VB.Label lblFechaHasta 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   3615
         TabIndex        =   5
         Top             =   1110
         Width           =   960
      End
      Begin VB.Label lblxx 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Egreso:"
         Height          =   195
         Left            =   225
         TabIndex        =   4
         Top             =   765
         Width           =   900
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   2550
      Top             =   1605
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmListadoIngersosEgersos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboDestino_Click()
    picSalida(0).Visible = False
    picSalida(1).Visible = False
    picSalida(2).Visible = False
    picSalida(cboDestino.ListIndex).Visible = True
End Sub

Private Sub cboTipo_Click()
    Select Case cboTipo.ListIndex
        Case 0 'EGRESOS
            CboGastos.Clear
            CboGastos.AddItem "(TODOS)"
            Call CargoComboBox(CboGastos, "TIPO_EGRESO", "TEG_CODIGO", "TEG_DESCRI")
            CboGastos.ListIndex = 0
        Case 1 'INGRESOS
            CboGastos.Clear
            CboGastos.AddItem "(TODOS)"
            Call CargoComboBox(CboGastos, "TIPO_INGRESO", "TIG_CODIGO", "TIG_DESCRI")
            CboGastos.ListIndex = 0
    End Select
End Sub

Private Sub cmdListar_Click()
    If FechaDesde.Value = "" Then
        MsgBox "Falta Ingresar la Fecha Desde", vbExclamation, TIT_MSGBOX
        FechaDesde.SetFocus
        Exit Sub
    End If
    If FechaHasta.Value = "" Then
        MsgBox "Falta Ingresar la Fecha Hasta", vbExclamation, TIT_MSGBOX
        FechaHasta.SetFocus
        Exit Sub
    End If
    
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    
    Select Case cboTipo.ListIndex
        Case 0 'EGRESOS
            Rep.SelectionFormula = " {CAJA_EGRESO.EST_CODIGO}=3 "
            If CboGastos.List(CboGastos.ListIndex) <> "(TODOS)" Then
                If Rep.SelectionFormula = "" Then
                    Rep.SelectionFormula = " {CAJA_EGRESO.TEG_CODIGO}=" & XN(CboGastos.ItemData(CboGastos.ListIndex))
                Else
                    Rep.SelectionFormula = Rep.SelectionFormula & " AND {CAJA_EGRESO.TEG_CODIGO}=" & XN(CboGastos.ItemData(CboGastos.ListIndex))
                End If
            End If
            If FechaDesde.Value <> "" Then
                If Rep.SelectionFormula = "" Then
                    Rep.SelectionFormula = " {CAJA_EGRESO.CEGR_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
                Else
                    Rep.SelectionFormula = Rep.SelectionFormula & " AND {CAJA_EGRESO.CEGR_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
                End If
            End If
            If FechaHasta.Value <> "" Then
                If Rep.SelectionFormula = "" Then
                    Rep.SelectionFormula = " {CAJA_EGRESO.CEGR_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                                   
                Else
                    Rep.SelectionFormula = Rep.SelectionFormula & " AND {CAJA_EGRESO.CEGR_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                End If
            End If
            Rep.WindowTitle = "Listado de Egresos"
            Rep.ReportFileName = DRIVE & DirReport & "ListadoEgresos.rpt"
            
        Case 1 'INGRESOS
            Rep.SelectionFormula = " {CAJA_INGRESO.EST_CODIGO}=3 "
            If CboGastos.List(CboGastos.ListIndex) <> "(TODOS)" Then
                If Rep.SelectionFormula = "" Then
                    Rep.SelectionFormula = " {CAJA_INGRESO.TIG_CODIGO}=" & XN(CboGastos.ItemData(CboGastos.ListIndex))
                Else
                    Rep.SelectionFormula = Rep.SelectionFormula & " AND {CAJA_INGRESO.TIG_CODIGO}=" & XN(CboGastos.ItemData(CboGastos.ListIndex))
                End If
            End If
            If FechaDesde.Value <> "" Then
                If Rep.SelectionFormula = "" Then
                    Rep.SelectionFormula = " {CAJA_INGRESO.CIGR_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
                Else
                    Rep.SelectionFormula = Rep.SelectionFormula & " AND {CAJA_INGRESO.CIGR_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
                End If
            End If
            If FechaHasta.Value <> "" Then
                If Rep.SelectionFormula = "" Then
                    Rep.SelectionFormula = " {CAJA_INGRESO.CIGR_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                                   
                Else
                    Rep.SelectionFormula = Rep.SelectionFormula & " AND {CAJA_INGRESO.CIGR_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                End If
            End If
            Rep.WindowTitle = "Listado de Ingresos"
            Rep.ReportFileName = DRIVE & DirReport & "ListadoIngresos.rpt"
    End Select
    
    If FechaDesde.Value <> "" Then
        Rep.Formulas(0) = "FECHAD='" & "DESDE: " & FechaDesde.Value & "'"
    End If
    If FechaHasta.Value <> "" Then
        Rep.Formulas(1) = "FECHAH='" & "HASTA: " & FechaHasta.Value & "'"
    End If
    
    Select Case cboDestino.ListIndex
        Case 0
            Rep.Destination = crptToWindow
        Case 1
            Rep.Destination = crptToPrinter
        Case 2
            Rep.Destination = crptToFile
    End Select
    Rep.Action = 1
    
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
End Sub

Private Sub CmdNuevo_Click()
    CboGastos.ListIndex = 0
    FechaDesde.Value = Date
    FechaHasta.Value = Date
    cboTipo.ListIndex = 0
    cboTipo.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Set frmListadoIngersosEgersos = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    Set Rec = New ADODB.Recordset
    cboDestino.ListIndex = 0
    
    Me.Left = 0
    Me.Top = 0
    
    cboTipo.AddItem "EGRESOS"
    cboTipo.AddItem "INGRESOS"
    cboTipo.ListIndex = 0
    
'    CboGastos.AddItem "(TODOS)"
'    Call CargoComboBox(CboGastos, "TIPO_EGRESO", "TEG_CODIGO", "TEG_DESCRI")
'    CboGastos.ListIndex = 0

End Sub

Private Sub optEgresos_Click()

End Sub

Private Sub OptIngresos_Click()

End Sub
