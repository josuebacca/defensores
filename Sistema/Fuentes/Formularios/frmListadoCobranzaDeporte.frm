VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListadoCobranzaDeporte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Cobranza x Deporte"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5400
   Begin VB.Frame Frame2 
      Caption         =   "Ordenado por.."
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
      Left            =   2220
      TabIndex        =   16
      Top             =   2340
      Width           =   3090
      Begin VB.ComboBox cboOrden 
         Height          =   315
         ItemData        =   "frmListadoCobranzaDeporte.frx":0000
         Left            =   675
         List            =   "frmListadoCobranzaDeporte.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   270
         Width           =   2070
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
      Left            =   45
      TabIndex        =   9
      Top             =   2340
      Width           =   2175
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "frmListadoCobranzaDeporte.frx":0004
         Left            =   450
         List            =   "frmListadoCobranzaDeporte.frx":0011
         Style           =   2  'Dropdown List
         TabIndex        =   13
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
         Picture         =   "frmListadoCobranzaDeporte.frx":0033
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   12
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
         Picture         =   "frmListadoCobranzaDeporte.frx":0135
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
         Index           =   2
         Left            =   135
         Picture         =   "frmListadoCobranzaDeporte.frx":0237
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   10
         Top             =   315
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   435
      Left            =   2685
      TabIndex        =   3
      Top             =   3150
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   4455
      TabIndex        =   5
      Top             =   3150
      Width           =   840
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoCobranzaDeporte.frx":0339
      Height          =   435
      Left            =   3570
      TabIndex        =   4
      Top             =   3150
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
      Height          =   2310
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   5265
      Begin VB.CheckBox chkRecargos 
         Alignment       =   1  'Right Justify
         Caption         =   "Incluir Recargos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3150
         TabIndex        =   15
         Top             =   720
         Width           =   1695
      End
      Begin VB.ListBox lstDeportes 
         Height          =   1185
         Left            =   225
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   975
         Width           =   4635
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20840449
         CurrentDate     =   42925
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   3585
         TabIndex        =   1
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20840449
         CurrentDate     =   42925
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Deportes:"
         Height          =   195
         Index           =   7
         Left            =   225
         TabIndex        =   14
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lblFechaDesde 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   225
         TabIndex        =   8
         Top             =   450
         Width           =   990
      End
      Begin VB.Label lblFechaHasta 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   2565
         TabIndex        =   7
         Top             =   450
         Width           =   960
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   1770
      Top             =   2985
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmListadoCobranzaDeporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRecargo As Double
Dim I As Integer

Private Sub cboDestino_Click()
    picSalida(0).Visible = False
    picSalida(1).Visible = False
    picSalida(2).Visible = False
    picSalida(cboDestino.ListIndex).Visible = True
End Sub

Private Sub cmdListar_Click()
    Dim mDep As String
    Dim mDep1 As String
    Dim J As Integer
    J = 1
    mDep = ""
    mDep1 = ""
    
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
    mRecargo = 0
    
    For I = 0 To lstDeportes.ListCount - 1
        If lstDeportes.Selected(I) = True Then
            If J = 1 Then
                mDep = "{DEPORTE.DEP_CODIGO}=" & lstDeportes.ItemData(I)
                mDep1 = lstDeportes.ItemData(I)
            Else
                mDep = mDep & " OR {DEPORTE.DEP_CODIGO}=" & lstDeportes.ItemData(I)
                mDep1 = mDep1 & "," & lstDeportes.ItemData(I)
            End If
            J = J + 1
        End If
    Next
    
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
    
    
    Rep.SelectionFormula = " {RECIBO.REC_ESTADO}=1 " 'SOLO RECIBOS DEFINITIVOS
    
    If FechaDesde.Value <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {RECIBO.REC_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {RECIBO.REC_FECHA}>= DATE (" & Mid(FechaDesde.Value, 7, 4) & "," & Mid(FechaDesde.Value, 4, 2) & "," & Mid(FechaDesde.Value, 1, 2) & ")"
        End If
    End If
    If FechaHasta.Value <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {RECIBO.REC_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
                           
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {RECIBO.REC_FECHA}<= DATE (" & Mid(FechaHasta.Value, 7, 4) & "," & Mid(FechaHasta.Value, 4, 2) & "," & Mid(FechaHasta.Value, 1, 2) & ")"
        End If
    End If
    
    If mDep <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = mDep
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND (" & mDep & ")"
        End If
    End If
           
    If chkRecargos.Value = Checked Then
        sql = "SELECT SUM(REC_RECARGO) AS MONTO FROM RECIBO"
        sql = sql & " WHERE"
        sql = sql & " REC_NUMERO IN (SELECT DISTINCT R.REC_NUMERO "
        sql = sql & " FROM RECIBO R, DEBITOS D"
        sql = sql & " WHERE R.REC_NUMERO=D.REC_NUMERO"
        sql = sql & " AND R.REC_ESTADO=1" 'SOLO RECIBOS DEFINITIVOS
        sql = sql & " AND R.REC_FECHA>=" & XDQ(FechaDesde.Value)
        sql = sql & " AND R.REC_FECHA<=" & XDQ(FechaHasta.Value)
        sql = sql & " AND D.DEP_CODIGO IN(" & mDep1 & "))"
        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec.EOF = False Then
            mRecargo = Valido_Importe(Chk0(Rec!MONTO))
        End If
        Rec.Close
    End If
        
    Rep.Formulas(2) = "RECARGO=" & XN(CStr(mRecargo))
    
    Rep.WindowTitle = "Listado de Cobranza - Detallado"
    Rep.ReportFileName = DRIVE & DirReport & "CobranzaDetalladoDeporte.rpt"
    
    Select Case cboOrden.ListIndex
        Case 0 'POR SOCIO
            Rep.SortFields(0) = "+{SOCIOS.SOC_NOMBRE}"
        Case 1 'POR NEO DE RECIBO
            Rep.SortFields(0) = "+{RECIBO.REC_NUMERO}"
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
    Rep.Formulas(2) = ""
End Sub

Private Sub CmdNuevo_Click()
    FechaDesde.Value = ""
    FechaHasta.Value = ""
    chkRecargos.Value = Unchecked
    For I = 0 To lstDeportes.ListCount - 1
        lstDeportes.Selected(I) = False
    Next
    FechaDesde.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Set frmListadoCobranzaDeporte = Nothing
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
    
    cboOrden.AddItem "SOCIO"
    cboOrden.AddItem "NRO RECIBO"
    cboOrden.ListIndex = 0
    
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
End Sub

Private Sub optEgresos_Click()

End Sub

Private Sub OptIngresos_Click()

End Sub
