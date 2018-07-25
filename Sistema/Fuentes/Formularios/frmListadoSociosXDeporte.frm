VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListadoSociosxDeporte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Socios x Deporte"
   ClientHeight    =   3135
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5400
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
      TabIndex        =   5
      Top             =   2340
      Width           =   2175
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "frmListadoSociosXDeporte.frx":0000
         Left            =   450
         List            =   "frmListadoSociosXDeporte.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   9
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
         Picture         =   "frmListadoSociosXDeporte.frx":002F
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   8
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
         Picture         =   "frmListadoSociosXDeporte.frx":0131
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   7
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
         Picture         =   "frmListadoSociosXDeporte.frx":0233
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   6
         Top             =   315
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   435
      Left            =   2670
      TabIndex        =   1
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   4440
      TabIndex        =   3
      Top             =   2640
      Width           =   840
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoSociosXDeporte.frx":0335
      Height          =   435
      Left            =   3555
      TabIndex        =   2
      Top             =   2640
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
      TabIndex        =   4
      Top             =   0
      Width           =   5265
      Begin VB.ListBox lstDeportes 
         Height          =   1635
         Left            =   225
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   525
         Width           =   4635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Deportes:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   225
         TabIndex        =   10
         Top             =   300
         Width           =   900
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   2550
      Top             =   2430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmListadoSociosxDeporte"
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
    
    If mDep <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = mDep
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND (" & mDep & ")"
        End If
    End If

    
    Rep.WindowTitle = "Listado de Socios x Deporte"
    Rep.ReportFileName = DRIVE & DirReport & "SociosxDeporte.rpt"
    
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
    For I = 0 To lstDeportes.ListCount - 1
        lstDeportes.Selected(I) = False
    Next
    lstDeportes.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Set frmListadoSociosxDeporte = Nothing
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
