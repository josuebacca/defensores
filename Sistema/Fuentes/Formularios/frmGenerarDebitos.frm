VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGenerarDebitos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Débitos"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3810
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
   ScaleHeight     =   2325
   ScaleWidth      =   3810
   Begin MSComCtl2.DTPicker fecha1 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   52232193
      CurrentDate     =   42839
   End
   Begin VB.TextBox txtAno 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   2
      Top             =   870
      Width           =   975
   End
   Begin VB.TextBox txtMes 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   765
      MaxLength       =   2
      TabIndex        =   1
      Top             =   870
      Width           =   645
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   855
      TabIndex        =   3
      Top             =   1875
      Width           =   1300
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   2195
      TabIndex        =   4
      Top             =   1875
      Width           =   1300
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   90
      TabIndex        =   7
      Top             =   1395
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   1
      Min             =   1e-4
   End
   Begin VB.Label lblPor 
      AutoSize        =   -1  'True
      Caption         =   "100 %"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3075
      TabIndex        =   8
      Top             =   1440
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Periodo (mm/yyyy):"
      Height          =   195
      Left            =   765
      TabIndex        =   6
      Top             =   570
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   195
      Left            =   1710
      TabIndex        =   5
      Top             =   165
      Width           =   495
   End
End
Attribute VB_Name = "frmGenerarDebitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim Rec As ADODB.Recordset
Dim Rec1 As ADODB.Recordset
Dim Rec2 As ADODB.Recordset

'PARA EL PROGRES BAR
Dim Registro As Long
Dim Tamanio As Long
Dim I As Integer

Private Sub cmdAceptar_Click()
    If fecha1.Value = "" Then
        MsgBox "Falta Ingresar la Fecha", vbCritical, TIT_MSGBOX
        fecha1.SetFocus
        Exit Sub
    End If
    If txtMes.Text = "" Then
        MsgBox "Falta Ingresar el Mes del Débito", vbCritical, TIT_MSGBOX
        txtMes.SetFocus
        Exit Sub
    End If
    If txtAno.Text = "" Then
        MsgBox "Falta Ingresar el Año del Débito", vbCritical, TIT_MSGBOX
        txtAno.SetFocus
        Exit Sub
    End If
    
'    sql = "SELECT DEB_MES"
'    sql = sql & " FROM DEBITOS"
'    sql = sql & " WHERE DEB_MES = " & XN(txtMes.Text)
'    sql = sql & " AND DEB_ANO = " & XN(txtAno.Text)
'    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If Rec.EOF = False Then
'        MsgBox "El Periodo ingresado ya fue Generado", vbExclamation, TIT_MSGBOX
'        Rec.Close
'        txtMes.SetFocus
'        Exit Sub
'    End If
'    Rec.Close
    
    If MsgBox("¿Genera los Débitos?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    Registro = 0
    Tamanio = 0
    I = 1
    'BUSCO LOS TITULARES POR CUOTA SOCIAL
    sql = "SELECT SOC_CODIGO, T.TIC_DESCRI, T.TIC_CUOTA, T.TIC_CODIGO"
    sql = sql & " FROM SOCIOS S, TIPO_CUOTA T"
    sql = sql & " WHERE "
    sql = sql & " T.TIC_CODIGO=S.TIC_CODIGO"
    sql = sql & " AND TIS_CODIGO=1" 'SOLO TITULARES
    sql = sql & " AND ESS_CODIGO=1" 'ACTIVOS
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Tamanio = Rec.RecordCount
        Do While Rec.EOF = False
            I = 1
            sql = "SELECT DEB_MES"
            sql = sql & " FROM DEBITOS"
            sql = sql & " WHERE DEB_MES = " & XN(txtMes.Text)
            sql = sql & " AND DEB_ANO = " & XN(txtAno.Text)
            sql = sql & " AND SOC_CODIGO = " & XN(Rec!SOC_CODIGO)
            sql = sql & " AND TIC_CODIGO = " & XN(Rec!TIC_CODIGO)
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = True Then
                
                sql = "SELECT DEB_ITEM"
                sql = sql & " FROM DEBITOS"
                sql = sql & " WHERE DEB_MES = " & XN(txtMes.Text)
                sql = sql & " AND DEB_ANO = " & XN(txtAno.Text)
                sql = sql & " AND SOC_CODIGO = " & XN(Rec!SOC_CODIGO)
                sql = sql & " AND DEB_ITEM = " & I
                Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If Rec2.EOF = False Then
                    I = I + 1
                End If
                Rec2.Close
                
                sql = "INSERT INTO DEBITOS (DEB_MES,DEB_ANO,SOC_CODIGO,"
                sql = sql & " DEB_ITEM,DEB_DETALLE,DEB_IMPORTE,DEB_SALDO,DEB_FECHA,TIC_CODIGO)"
                sql = sql & " VALUES ("
                sql = sql & XN(txtMes.Text) & ","
                sql = sql & XN(txtAno.Text) & ","
                sql = sql & XN(Rec!SOC_CODIGO) & ","
                sql = sql & I & "," & XS("CUOTA " & Rec!TIC_DESCRI) & ","
                sql = sql & XN(Rec!TIC_CUOTA) & ","
                sql = sql & XN(Rec!TIC_CUOTA) & ","
                sql = sql & XDQ(fecha1.Value) & ","
                sql = sql & XN(Rec!TIC_CODIGO) & ")"
                DBConn.Execute sql
            End If
            Rec1.Close
            Rec.MoveNext
            'PARA LA BARRA DE PROGRESO
            Registro = Registro + 1
            ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
            lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
            lblPor.Refresh
        Loop
    End If
    Rec.Close
    
    Registro = 0
    Tamanio = 0
    'BUSCO EN LOS SOCIOS LOS DEPORTES QUE REALIZAN PARA COBRAR LA CUOTA DE C/DEPORTE
    sql = "SELECT S.SOC_CODIGO, D.DEP_DESCRI, D.DEP_CUOTA, D.DEP_CODIGO"
    sql = sql & " FROM SOCIOS S, SOCIOS_DEPORTES SD, DEPORTE D"
    sql = sql & " WHERE "
    sql = sql & " S.SOC_CODIGO=SD.SOC_CODIGO"
    sql = sql & " AND D.DEP_CODIGO=SD.DEP_CODIGO"
    sql = sql & " AND ESS_CODIGO=1" 'ACTIVOS
    sql = sql & " AND D.DEP_DEBITO IS NULL"
    sql = sql & " ORDER BY S.SOC_CODIGO, D.DEP_DESCRI"
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Tamanio = Rec.RecordCount
        I = 2
        Do While Rec.EOF = False
            sql = "SELECT DEB_MES"
            sql = sql & " FROM DEBITOS"
            sql = sql & " WHERE DEB_MES = " & XN(txtMes.Text)
            sql = sql & " AND DEB_ANO = " & XN(txtAno.Text)
            sql = sql & " AND SOC_CODIGO = " & XN(Rec!SOC_CODIGO)
            sql = sql & " AND DEP_CODIGO = " & XN(Rec!DEP_CODIGO)
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = True Then
                sql = "INSERT INTO DEBITOS (DEB_MES,DEB_ANO,SOC_CODIGO,"
                sql = sql & " DEB_ITEM,DEB_DETALLE,DEB_IMPORTE,DEB_SALDO,DEB_FECHA,DEP_CODIGO)"
                sql = sql & " VALUES ("
                sql = sql & XN(txtMes.Text) & ","
                sql = sql & XN(txtAno.Text) & ","
                sql = sql & XN(Rec!SOC_CODIGO) & ","
                sql = sql & I & "," & XS(Rec!DEP_DESCRI) & ","
                sql = sql & XN(Rec!DEP_CUOTA) & ","
                sql = sql & XN(Rec!DEP_CUOTA) & ","
                sql = sql & XDQ(fecha1.Value) & ","
                sql = sql & XN(Rec!DEP_CODIGO) & ")"
                DBConn.Execute sql
                I = I + 1
            End If
            Rec1.Close
            Rec.MoveNext
            'PARA LA BARRA DE PROGRESO
            Registro = Registro + 1
            ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
            lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
            lblPor.Refresh
        Loop
    End If
    Rec.Close
    
    MsgBox "Los Débitos se generaron con Exito", vbExclamation, TIT_MSGBOX
End Sub

Private Sub cmdCerrar_Click()
    Set frmGenerarDebitos = Nothing
    Unload Me
End Sub

Private Sub fecha1_LostFocus()
    If fecha1.Value = "" Then
        fecha1.Value = Date
    End If
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
    Set Rec2 = New ADODB.Recordset
    lblPor.Caption = "100 %"
End Sub

Private Sub txtAno_GotFocus()
    SelecTexto txtAno
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtAno_LostFocus()
    If txtAno.Text <> "" Then
        If Len(txtAno.Text) = 2 Then
            txtAno.Text = Mid(Year(Date), 1, 2) & txtAno.Text
        Else
            txtAno.Text = Format(txtAno.Text, "0000")
        End If
    End If
End Sub

Private Sub txtMes_GotFocus()
    SelecTexto txtMes
End Sub

Private Sub txtMes_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtMes_LostFocus()
    If txtMes.Text <> "" Then
        txtMes.Text = Format(txtMes.Text, "00")
    End If
End Sub
