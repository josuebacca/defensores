VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCtaCteDeporte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cta Cte - Deportes"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4890
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
   ScaleHeight     =   2505
   ScaleWidth      =   4890
   Begin VB.ListBox lstDeportes 
      Height          =   1185
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   315
      Width           =   4635
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
      Left            =   90
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
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
         Picture         =   "frmCtaCteDeporte.frx":0000
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
         Index           =   1
         Left            =   135
         Picture         =   "frmCtaCteDeporte.frx":0102
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   6
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
         Index           =   0
         Left            =   135
         Picture         =   "frmCtaCteDeporte.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   5
         Top             =   315
         Width           =   240
      End
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "frmCtaCteDeporte.frx":0306
         Left            =   450
         List            =   "frmCtaCteDeporte.frx":0313
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   270
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   3450
      TabIndex        =   1
      Top             =   1605
      Width           =   1300
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   3455
      TabIndex        =   2
      Top             =   2025
      Width           =   1300
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   2370
      Top             =   1590
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Deportes:"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   8
      Top             =   90
      Width           =   720
   End
End
Attribute VB_Name = "frmCtaCteDeporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim Rec As ADODB.Recordset

'PARA EL PROGRES BAR
Dim Registro As Long
Dim Tamanio As Long
Dim I As Integer
Dim J As Integer
Dim mDep1 As String

Private Sub cboDestino_Click()
    picSalida(0).Visible = False
    picSalida(1).Visible = False
    picSalida(2).Visible = False
    picSalida(cboDestino.ListIndex).Visible = True
End Sub

Private Sub cmdAceptar_Click()
    sql = "DELETE FROM CTA_CTE_SOCIOS"
    DBConn.Execute sql
    
    mDep1 = ""
    J = 1
    For I = 0 To lstDeportes.ListCount - 1
        If lstDeportes.Selected(I) = True Then
            If J = 1 Then
                mDep1 = lstDeportes.ItemData(I)
            Else
                mDep1 = mDep1 & "," & lstDeportes.ItemData(I)
            End If
            J = J + 1
        End If
    Next
        
    BuscarDeuda
    
    sql = "DELETE FROM TMP_RESUMEN_CUENTA_BANCO"
    DBConn.Execute sql
    
    
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Select Case cboDestino.ListIndex
        Case 0
            Rep.Destination = crptToWindow
        Case 1
            Rep.Destination = crptToPrinter
        Case 2
            Rep.Destination = crptToFile
    End Select
    
    Rep.WindowTitle = "Listado de Cta-Cte Detalle x Deporte"
    Rep.ReportFileName = DRIVE & DirReport & "CtaCteDeportes.rpt"
    
    Rep.Action = 1
End Sub

Private Sub cmdCerrar_Click()
    Set frmCtaCteDeporte = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    If KeyAscii = 27 Then
        cmdCerrar_Click
    End If
End Sub

Private Sub Form_Load()
    Set Rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    
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
    
    cboDestino.ListIndex = 0
End Sub

Private Sub BuscarDeuda()
    'TITULARES
    sql = "SELECT S.SOC_NOMBRE, D.DEB_MES, D.DEB_ANO, S.SOC_CODIGO,"
    sql = sql & " D.DEB_ITEM, D.DEB_DETALLE, D.DEB_IMPORTE, D.DEB_FECHA, D.DEP_CODIGO"
    sql = sql & " FROM DEBITOS D, SOCIOS S, DEPORTE E"
    sql = sql & " WHERE "
    sql = sql & " S.SOC_CODIGO=D.SOC_CODIGO"
    sql = sql & " AND E.DEP_CODIGO=D.DEP_CODIGO"
    If mDep1 <> "" Then
        sql = sql & " AND D.DEP_CODIGO IN (" & mDep1 & ")"
    End If
    sql = sql & " AND D.DEB_SALDO > 0"
    sql = sql & " AND D.DEB_PAGADO IS NULL"
    
    'sql = sql & " UNION ALL"
    
'    'OPTATIVOS
'    sql = sql & " SELECT S1.SOC_NOMBRE, D1.DEB_MES, D1.DEB_ANO, S1.SOC_CODIGO,"
'    sql = sql & " D1.DEB_ITEM, D1.DEB_DETALLE, D1.DEB_IMPORTE, D1.DEB_FECHA"
'    sql = sql & " FROM DEBITOS D1, SOCIOS S1"
'    sql = sql & " WHERE"
'    sql = sql & " S1.SOC_CODIGO=D1.SOC_CODIGO"
'    sql = sql & " AND S1.SOC_TITULAR=" & XN(mCodigo)
'    sql = sql & " AND D1.DEB_SALDO > 0"
'    sql = sql & " AND D1.DEB_PAGADO IS NULL"
'    sql = sql & " ORDER BY D.DEB_FECHA, S.SOC_CODIGO, D.DEB_ITEM"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            sql = "INSERT INTO CTA_CTE_SOCIOS (SOC_CODIGO,SOC_NOMBRE,"
            sql = sql & " DEB_MES,DEB_ANO,DEB_DETALLE,DEB_IMPORTE,"
            sql = sql & " DEB_ITEM,DEB_FECHA,CTACTE_TITULAR,CTACTE_CODIGO)"
            sql = sql & " VALUES ("
            sql = sql & XN(Rec1!SOC_CODIGO) & ","
            sql = sql & XS(Rec1!SOC_NOMBRE) & ","
            sql = sql & XS(Format(Rec1!DEB_MES, "00")) & ","
            sql = sql & XS(Format(Rec1!DEB_ANO, "0000")) & ","
            sql = sql & XS(Rec1!DEB_DETALLE) & ","
            sql = sql & XN(Rec1!DEB_IMPORTE) & ","
            sql = sql & XN(Rec1!DEB_ITEM) & ","
            sql = sql & XDQ(Rec1!DEB_FECHA) & ","
            sql = sql & XS("") & ","
            sql = sql & XN(Rec1!DEP_CODIGO) & ")"
            DBConn.Execute sql
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Sub

