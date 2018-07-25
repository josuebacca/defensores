VERSION 5.00
Begin VB.Form frmParametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametros"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
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
   ScaleHeight     =   2970
   ScaleWidth      =   6165
   Begin VB.Frame Frame3 
      Caption         =   "Recargos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   75
      TabIndex        =   15
      Top             =   1740
      Width           =   6015
      Begin VB.TextBox txtRecargo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1500
         MaxLength       =   5
         TabIndex        =   4
         Top             =   255
         Width           =   825
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dias Recargo: "
         Height          =   195
         Left            =   405
         TabIndex        =   16
         Top             =   300
         Width           =   1050
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comisiones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   75
      TabIndex        =   11
      Top             =   930
      Width           =   6015
      Begin VB.TextBox txtComDep 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4875
         MaxLength       =   5
         TabIndex        =   3
         Top             =   360
         Width           =   825
      End
      Begin VB.TextBox txtComPor 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1500
         MaxLength       =   5
         TabIndex        =   2
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Por Cuota Deportiva ($): "
         Height          =   195
         Left            =   3015
         TabIndex        =   14
         Top             =   390
         Width           =   1830
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Por Cuota Social: "
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   390
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   2385
         TabIndex        =   12
         Top             =   435
         Width           =   165
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impuesto Transacciones Financieras (Impuesto al Cheque)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   75
      TabIndex        =   8
      Top             =   75
      Width           =   6015
      Begin VB.CheckBox chkAplicoImpuesto 
         Caption         =   "Aplicar Impuesto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2340
         TabIndex        =   1
         Top             =   375
         Width           =   1815
      End
      Begin VB.TextBox txtImpuestoCheque 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   705
         MaxLength       =   5
         TabIndex        =   0
         Top             =   345
         Width           =   825
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   1590
         TabIndex        =   10
         Top             =   420
         Width           =   165
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor: "
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   375
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4050
      Picture         =   "frmParametros.frx":0000
      TabIndex        =   5
      Top             =   2535
      Width           =   1000
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5085
      Picture         =   "frmParametros.frx":030A
      TabIndex        =   6
      Top             =   2535
      Width           =   1000
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   2535
      Width           =   660
   End
End
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdGrabar_Click()
    If Validar_Parametros = False Then Exit Sub
    
    On Error GoTo HayError
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Actualizando..."
    
    DBConn.BeginTrans
    sql = "UPDATE PARAMETROS"
    sql = sql & " SET VALOR_IMPUESTO=" & XN(txtImpuestoCheque.Text)
    If chkAplicoImpuesto.Value = Checked Then
        sql = sql & " ,APLICA_IMPUESTO=" & XS("S")
    Else
        sql = sql & " ,APLICA_IMPUESTO=" & XS("N")
    End If
    sql = sql & " ,COM_CUTSOC=" & XN(txtComPor.Text)
    sql = sql & " ,COM_CUTDEP=" & XN(txtComDep.Text)
    sql = sql & " ,DIAS_RECARGO=" & XN(txtRecargo.Text)
    DBConn.Execute sql
    DBConn.CommitTrans
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    Exit Sub
    
HayError:
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Function Validar_Parametros() As Boolean
      Validar_Parametros = True
End Function

Private Sub cmdSalir_Click()
    Set frmParametros = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    Set Rec = New ADODB.Recordset
    'Centrar_pantalla Me
    Me.Top = 0
    Me.Left = 0
    'busco datos
    BuscarDatos
    lblEstado.Caption = ""
End Sub

Private Sub BuscarDatos()
    sql = "SELECT * FROM PARAMETROS"
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        'DATOS PECARI
        txtImpuestoCheque.Text = IIf(IsNull(Rec!VALOR_IMPUESTO), 0, Rec!VALOR_IMPUESTO)
        If Rec!APLICA_IMPUESTO = "S" Then
            chkAplicoImpuesto.Value = Checked
        Else
            chkAplicoImpuesto.Value = Unchecked
        End If
        txtComPor.Text = Valido_Importe(Chk0(Rec!COM_CUTSOC))
        txtComDep.Text = Valido_Importe(Chk0(Rec!COM_CUTDEP))
        txtRecargo.Text = Chk0(Rec!DIAS_RECARGO)
    End If
    Rec.Close
End Sub

Private Sub txtComDep_GotFocus()
    SelecTexto txtComDep
End Sub

Private Sub txtComDep_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtComDep, KeyAscii)
End Sub

Private Sub txtComDep_LostFocus()
    If txtComDep.Text <> "" Then
        txtComDep.Text = Valido_Importe(txtComDep.Text)
    End If
End Sub

Private Sub txtComPor_GotFocus()
    SelecTexto txtComPor
End Sub

Private Sub txtComPor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtComPor, KeyAscii)
End Sub

Private Sub txtComPor_LostFocus()
    If txtComPor.Text <> "" Then
        txtComPor.Text = Valido_Importe(txtComPor.Text)
    End If
End Sub

Private Sub txtImpuestoCheque_GotFocus()
    SelecTexto txtImpuestoCheque
End Sub

Private Sub txtImpuestoCheque_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImpuestoCheque, KeyAscii)
End Sub

Private Sub txtRecargo_GotFocus()
    SelecTexto txtRecargo
End Sub

Private Sub txtRecargo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub
