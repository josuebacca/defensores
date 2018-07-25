VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmComisiones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comisiones"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7605
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
   ScaleHeight     =   6165
   ScaleWidth      =   7605
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
      Left            =   105
      TabIndex        =   15
      Top             =   5370
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
         Picture         =   "frmComisiones.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   19
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
         Picture         =   "frmComisiones.frx":0102
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   18
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
         Picture         =   "frmComisiones.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   17
         Top             =   315
         Width           =   240
      End
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "frmComisiones.frx":0306
         Left            =   450
         List            =   "frmComisiones.frx":0313
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   270
         Width           =   1635
      End
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   5205
      Width           =   1500
   End
   Begin VB.TextBox txtTotalSocial 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4785
      Width           =   1500
   End
   Begin VB.TextBox txtTotalDeportiva 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   2340
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4785
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar por..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   60
      TabIndex        =   5
      Top             =   75
      Width           =   7455
      Begin VB.CommandButton CmdBuscAprox 
         Caption         =   "Buscar"
         Height          =   405
         Left            =   4815
         MaskColor       =   &H000000FF&
         TabIndex        =   2
         ToolTipText     =   "Buscar"
         Top             =   345
         UseMaskColor    =   -1  'True
         Width           =   2445
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   390
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   52363265
         CurrentDate     =   42925
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   3540
         TabIndex        =   1
         Top             =   390
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   52363265
         CurrentDate     =   42925
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecah Desde:"
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   450
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   2550
         TabIndex        =   6
         Top             =   450
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   4830
      TabIndex        =   3
      Top             =   5760
      Width           =   1300
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   6165
      TabIndex        =   4
      Top             =   5760
      Width           =   1300
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   3720
      Left            =   45
      TabIndex        =   8
      Top             =   1020
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   6562
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   280
      BackColorSel    =   16761024
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
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
   Begin Crystal.CrystalReport Rep 
      Left            =   2565
      Top             =   5445
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL GENERAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4080
      TabIndex        =   14
      Top             =   5250
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CUOTA SOCIAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4230
      TabIndex        =   12
      Top             =   4830
      Width           =   1710
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "CUOTA DEPORTIVA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   75
      TabIndex        =   10
      Top             =   4830
      Width           =   2175
   End
End
Attribute VB_Name = "frmComisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mComSocial As Double
Dim mComDeportiva As Double
Dim mSoc As Double
Dim mDep As Double
Dim I As Integer

Private Sub cboDestino_Click()
    picSalida(0).Visible = False
    picSalida(1).Visible = False
    picSalida(2).Visible = False
    picSalida(cboDestino.ListIndex).Visible = True
End Sub

Private Sub cmdAceptar_Click()
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
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

    Rep.WindowTitle = "Listado de Comisiones"
    Rep.ReportFileName = DRIVE & DirReport & "Comisiones.rpt"
    Rep.Action = 1
    
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
End Sub

Private Sub CmdBuscAprox_Click()
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
    cmdAceptar.Enabled = True
    mSoc = 0
    mDep = 0
    GrdModulos.Rows = 1
    
    sql = "DELETE FROM TMP_COMISION"
    DBConn.Execute sql
    
    sql = "SELECT REC_NUMERO, REC_FECHA, REC_IMPORTE,REC_COMISION,REC_NROTXT"
    sql = sql & " FROM RECIBO"
    sql = sql & " WHERE"
    sql = sql & " REC_ESTADO=1" 'SOLO RECIBOS DEFINITIVOS
    sql = sql & " AND REC_FECHA >=" & XDQ(FechaDesde.Value)
    sql = sql & " AND REC_FECHA <=" & XDQ(FechaHasta.Value)
    sql = sql & " AND REC_COMISION='S'"
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Do While Rec.EOF = False
            mSoc = 0
            mDep = 0
            sql = "SELECT DEB_DETALLE, DEB_IMPORTE, TIC_CODIGO, DEP_CODIGO"
            sql = sql & " FROM DEBITOS"
            sql = sql & " WHERE "
            sql = sql & " REC_NUMERO=" & XN(Rec!REC_NUMERO)
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                Do While Rec1.EOF = False
                    If Not IsNull(Rec1!TIC_CODIGO) Then
                        mSoc = mSoc + ((CDbl(Rec1!DEB_IMPORTE) * mComSocial) / 100)
                    ElseIf Not IsNull(Rec1!DEP_CODIGO) Then
                        mDep = mDep + (CDbl(mComDeportiva))
                    End If
                    Rec1.MoveNext
                Loop
            End If
            Rec1.Close
            
            GrdModulos.AddItem Rec!REC_FECHA & Chr(9) & Rec!REC_NROTXT & Chr(9) & Valido_Importe(Rec!REC_IMPORTE) & Chr(9) & _
                               Valido_Importe(CStr(mSoc)) & Chr(9) & Valido_Importe(CStr(mDep))
            
            sql = "INSERT INTO TMP_COMISION (FECHA,RECIBO,IMPORTE,COMSOC,COMDEP,NROREC)"
            sql = sql & " VALUES ("
            sql = sql & XDQ(Rec!REC_FECHA) & ","
            sql = sql & XS(Rec!REC_NROTXT) & ","
            sql = sql & XN(Rec!REC_IMPORTE) & ","
            sql = sql & XN(CStr(mSoc)) & ","
            sql = sql & XN(CStr(mDep)) & ","
            sql = sql & XN(Rec!REC_NUMERO) & ")"
            DBConn.Execute sql
            
            I = GrdModulos.Rows - 1
            If (Int(I / 2) * 2) = I Then
                CambiaColorAFilaDeGrilla GrdModulos, I, , &HE0E0E0
            Else
                CambiaColorAFilaDeGrilla GrdModulos, I, , vbWhite
            End If
            
            Rec.MoveNext
        Loop
    End If
    Rec.Close
    SUMAR_TOTALES
End Sub

Private Sub SUMAR_TOTALES()
    txtTotalDeportiva.Text = "0"
    txtTotalSocial.Text = "0"
    txtTotal.Text = "0"
    For I = 1 To GrdModulos.Rows - 1
        txtTotalDeportiva.Text = CDbl(txtTotalDeportiva.Text) + CDbl(Chk0(GrdModulos.TextMatrix(I, 4)))
        txtTotalSocial.Text = CDbl(txtTotalSocial.Text) + CDbl(Chk0(GrdModulos.TextMatrix(I, 3)))
    Next
    
    txtTotal.Text = CDbl(txtTotalDeportiva.Text) + CDbl(txtTotalSocial.Text)
    txtTotal.Text = Valido_Importe(txtTotal.Text)
    txtTotalDeportiva.Text = Valido_Importe(txtTotalDeportiva.Text)
    txtTotalSocial.Text = Valido_Importe(txtTotalSocial.Text)
End Sub

Private Sub cmdCerrar_Click()
    Set frmComisiones = Nothing
    Unload Me
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
    Set Rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    cboDestino.ListIndex = 0
    
    Me.Top = 0
    Me.Left = 0
    cmdAceptar.Enabled = False
    
    sql = "SELECT COM_CUTSOC,COM_CUTDEP"
    sql = sql & " FROM PARAMETROS"
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        mComSocial = Chk0(Rec!COM_CUTSOC)
        mComDeportiva = Chk0(Rec!COM_CUTDEP)
    End If
    Rec.Close
    
    GrdModulos.FormatString = "^Fecha|^Nro Recibo|>Importe|>Com. Cuota Soc.|>Com. Cuota Dep."
    GrdModulos.ColWidth(0) = 1300 'FECHA
    GrdModulos.ColWidth(1) = 1500 'NRO RECIBO
    GrdModulos.ColWidth(2) = 1300 'IMPORTE
    GrdModulos.ColWidth(3) = 1500 'COM CUOTA SOCIAL
    GrdModulos.ColWidth(4) = 1500 'COM CUOTA DEPORTIVA
    GrdModulos.Cols = 5
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
    GrdModulos.BorderStyle = flexBorderNone
    GrdModulos.row = 0
    For I = 0 To GrdModulos.Cols - 1
        GrdModulos.Col = I
        GrdModulos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdModulos.CellBackColor = &H808080    'GRIS OSCURO
        GrdModulos.CellFontBold = True
    Next
End Sub

