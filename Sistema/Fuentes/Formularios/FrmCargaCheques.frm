VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCargaCheques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga de Cheques de Terceros"
   ClientHeight    =   4425
   ClientLeft      =   2535
   ClientTop       =   1005
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCargaCheques.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCuentaResponsable 
      Height          =   330
      Left            =   4560
      MaxLength       =   20
      TabIndex        =   8
      Top             =   2025
      Width           =   1860
   End
   Begin VB.TextBox TxtCheImport 
      Height          =   330
      Left            =   1380
      TabIndex        =   12
      Top             =   3075
      Width           =   1125
   End
   Begin VB.TextBox TxtCheObserv 
      Height          =   330
      Left            =   1380
      MaxLength       =   60
      TabIndex        =   13
      Top             =   3450
      Width           =   5040
   End
   Begin VB.TextBox TxtCheNombre 
      Height          =   330
      Left            =   1380
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1665
      Width           =   5040
   End
   Begin VB.TextBox TxtCheMotivo 
      Height          =   330
      Left            =   1380
      MaxLength       =   40
      TabIndex        =   9
      Top             =   2370
      Width           =   5040
   End
   Begin VB.Frame Frame2 
      Caption         =   "Banco"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   20
      Top             =   495
      Width           =   6300
      Begin VB.TextBox TxtCodInt 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5370
         TabIndex        =   22
         Top             =   660
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox TxtCODIGO 
         Height          =   285
         Left            =   4755
         MaxLength       =   6
         TabIndex        =   5
         Top             =   270
         Width           =   765
      End
      Begin VB.TextBox TxtLOCALIDAD 
         Height          =   285
         Left            =   2175
         MaxLength       =   3
         TabIndex        =   3
         Top             =   285
         Width           =   450
      End
      Begin VB.TextBox TxtBANCO 
         Height          =   285
         Left            =   780
         MaxLength       =   3
         TabIndex        =   2
         Top             =   285
         Width           =   450
      End
      Begin VB.TextBox TxtSUCURSAL 
         Height          =   285
         Left            =   3525
         MaxLength       =   3
         TabIndex        =   4
         Top             =   285
         Width           =   450
      End
      Begin VB.CommandButton CmdBanco 
         DisabledPicture =   "FrmCargaCheques.frx":08CA
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5655
         Picture         =   "FrmCargaCheques.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Agregar Banco"
         Top             =   270
         Width           =   375
      End
      Begin VB.TextBox TxtBanDescri 
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   210
         TabIndex        =   23
         Top             =   675
         Width           =   5820
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   4125
         TabIndex        =   27
         Top             =   315
         Width           =   555
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   2820
         TabIndex        =   26
         Top             =   315
         Width           =   660
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Banco:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   225
         TabIndex        =   25
         Top             =   315
         Width           =   495
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Localidad:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   1395
         TabIndex        =   24
         Top             =   315
         Width           =   720
      End
   End
   Begin VB.TextBox TxtCheNumero 
      Height          =   315
      Left            =   5010
      MaxLength       =   10
      TabIndex        =   1
      Top             =   150
      Width           =   1380
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5490
      TabIndex        =   17
      Top             =   3975
      Width           =   900
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   4575
      TabIndex        =   16
      Top             =   3975
      Width           =   900
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   3660
      TabIndex        =   15
      Top             =   3975
      Width           =   900
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   2745
      TabIndex        =   14
      Top             =   3975
      Width           =   900
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   120
      TabIndex        =   18
      Top             =   3870
      Width           =   6345
   End
   Begin MSMask.MaskEdBox txtCuit 
      Height          =   315
      Left            =   1380
      TabIndex        =   7
      Top             =   2025
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   13
      Mask            =   "##-########-#"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker TxtCheFecEmi 
      Height          =   315
      Left            =   1380
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   52494337
      CurrentDate     =   42925
   End
   Begin MSComCtl2.DTPicker TxtCheFecVto 
      Height          =   315
      Left            =   5160
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   52494337
      CurrentDate     =   42925
   End
   Begin MSComCtl2.DTPicker TxtCheFecEnt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   52494337
      CurrentDate     =   42925
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CUIT Librador:"
      Height          =   195
      Index           =   5
      Left            =   150
      TabIndex        =   37
      Top             =   2085
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nro Cuenta:"
      Height          =   195
      Index           =   8
      Left            =   3585
      TabIndex        =   36
      Top             =   2085
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Vto:"
      Height          =   195
      Index           =   3
      Left            =   4065
      TabIndex        =   35
      Top             =   2775
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha  Emisión:"
      Height          =   195
      Index           =   2
      Left            =   150
      TabIndex        =   34
      Top             =   2775
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Importe:"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   33
      Top             =   3105
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   195
      Index           =   0
      Left            =   795
      TabIndex        =   32
      Top             =   195
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones:"
      Height          =   195
      Index           =   6
      Left            =   150
      TabIndex        =   31
      Top             =   3480
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nro Cheque:"
      Height          =   195
      Index           =   7
      Left            =   3960
      TabIndex        =   30
      Top             =   210
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Responsable:"
      Height          =   195
      Index           =   9
      Left            =   150
      TabIndex        =   29
      Top             =   1710
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Concepto:"
      Height          =   195
      Index           =   10
      Left            =   150
      TabIndex        =   28
      Top             =   2400
      Width           =   750
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
      Left            =   180
      TabIndex        =   19
      Top             =   4020
      Width           =   660
   End
End
Attribute VB_Name = "FrmCargaCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mMeLlamo As String

Function Validar() As Boolean
   If Trim(TxtCheNumero.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Número de Cheque.", 16, TIT_MSGBOX
        TxtCheNumero.SetFocus
        Exit Function
        
   ElseIf Trim(TxtBanco.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Banco.", 16, TIT_MSGBOX
        TxtBanco.SetFocus
        Exit Function
        
   ElseIf Trim(TxtLOCALIDAD.Text) = "" Then
        Validar = False
        MsgBox "Ingrese la Localidad del Banco.", 16, TIT_MSGBOX
        TxtLOCALIDAD.SetFocus
        Exit Function
        
   ElseIf Trim(TxtSucursal.Text) = "" Then
        Validar = False
        MsgBox "Ingrese la Sucursal del Banco.", 16, TIT_MSGBOX
        TxtSucursal.SetFocus
        Exit Function
        
   ElseIf Trim(TxtCodigo.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Código del Banco.", 16, TIT_MSGBOX
        TxtCodigo.SetFocus
        Exit Function
        
   ElseIf Trim(Me.TxtCodInt.Text) = "" Then
        Validar = False
        MsgBox "Verifique el Código de Banco.", 16, TIT_MSGBOX
        TxtCodigo.SetFocus
        Exit Function
        
   ElseIf Trim(TxtCheMotivo.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Concepto del Cheque.", 16, TIT_MSGBOX
        TxtCheMotivo.SetFocus
        Exit Function
        
   ElseIf Trim(TxtCheFecEmi.Value) = "" Then
        Validar = False
        MsgBox "Ingrese la Fecha de Emisión.", 16, TIT_MSGBOX
        TxtCheFecEmi.SetFocus
        Exit Function
        
   ElseIf Trim(TxtCheFecVto.Value) = "" Then
        Validar = False
        MsgBox "Ingrese la Fecha de Vencimiento.", 16, TIT_MSGBOX
        TxtCheFecVto.SetFocus
        Exit Function
    
   ElseIf Me.TxtCheNombre.Text = "" Then
        Validar = False
        MsgBox "Debe ingresar la Persona Responsable.!", 16, TIT_MSGBOX
        TxtCheNombre.SetFocus
        Exit Function
        
   ElseIf Me.txtCuit.Text = "" Then
        Validar = False
        MsgBox "Debe ingresar el Nro de C.U.I.T. del Responsable.!", 16, TIT_MSGBOX
        txtCuit.SetFocus
        Exit Function
   
   ElseIf Me.txtCuentaResponsable.Text = "" Then
        Validar = False
        MsgBox "Debe ingresar el Nro de Cuenta del Responsable.!", 16, TIT_MSGBOX
        txtCuentaResponsable.SetFocus
        Exit Function
   ElseIf Trim(TxtCheImport.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Importe del Cheque.", 16, TIT_MSGBOX
        TxtCheImport.SetFocus
        Exit Function
   End If
   
   Validar = True
End Function

Private Sub CmdBanco_Click()
    Viene_Cheque = True
    ABMBancos.Show vbModal
    Viene_Cheque = False
End Sub

Private Sub cmdBorrar_Click()
    On Error GoTo CLAVOSE
    
    If Trim(TxtCheNumero.Text) <> "" And Trim(Me.TxtCodInt.Text) <> "" Then
        If MsgBox("Seguro desea eliminar el Cheque Nº: " & Trim(Me.TxtCheNumero.Text) & "? ", 36, TIT_MSGBOX) = vbYes Then
        
            Screen.MousePointer = vbHourglass
            lblEstado.Caption = "Borrando..."
            
            sql = "SELECT BOL_NUMERO "
            sql = sql & " FROM ChequeEstadoVigente "
            sql = sql & " WHERE CHE_NUMERO = " & XS(Me.TxtCheNumero.Text)
            sql = sql & " AND BAN_CODINT = " & XN(Me.TxtCodInt.Text)
            Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec.EOF = False Then
                If Not IsNull(Rec!BOL_NUMERO) Then
                   MsgBox "No se puede eliminar este Cheque porque fue depositado", vbExclamation, TIT_MSGBOX
                   Rec.Close
                   Screen.MousePointer = vbNormal
                   lblEstado.Caption = ""
                   Exit Sub
                 End If
            End If
            Rec.Close
    
            DBConn.BeginTrans
            DBConn.Execute "DELETE FROM CHEQUE_ESTADOS WHERE CHE_NUMERO = " & XS(Me.TxtCheNumero.Text) & " AND BAN_CODINT = " & XN(Me.TxtCodInt.Text)
                           
            DBConn.Execute "DELETE FROM CHEQUE WHERE CHE_NUMERO = " & XS(Me.TxtCheNumero.Text) & " AND BAN_CODINT = " & XN(Me.TxtCodInt.Text)
            
            Screen.MousePointer = vbNormal
            lblEstado.Caption = ""
            DBConn.CommitTrans
            CmdNuevo_Click
        End If
    End If
    Exit Sub
    
CLAVOSE:
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    If Rec.State = 1 Then Rec.Close
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdGrabar_Click()
    
    If Validar = True Then
    
    If mMeLlamo = "RECIBO" Then
        frmRecibo.grdPagos.AddItem "CHEQUE TERCERO" & Chr(9) & TxtCheImport.Text & Chr(9) & _
                                   "2" & Chr(9) & TxtCodInt.Text & Chr(9) & TxtCheNumero.Text & Chr(9) & _
                                   TxtCheFecEmi.Value & Chr(9) & TxtCheFecVto.Value & Chr(9) & TxtCheObserv.Text & Chr(9) & _
                                   TxtCheFecEnt.Value & Chr(9) & TxtCheNombre.Text & Chr(9) & TxtCheMotivo.Text & Chr(9) & _
                                   txtCuentaResponsable.Text & Chr(9) & txtCuit.Text
                                   
        cmdSalir_Click
        Exit Sub
        
    ElseIf mMeLlamo = "CAJA INGRESO" Then
        ABMIngresos.grdPagos.AddItem "CHEQUE TERCERO" & Chr(9) & TxtCheImport.Text & Chr(9) & _
                                   "2" & Chr(9) & TxtCodInt.Text & Chr(9) & TxtCheNumero.Text & Chr(9) & _
                                   TxtCheFecEmi.Value & Chr(9) & TxtCheFecVto.Value & Chr(9) & TxtCheObserv.Text & Chr(9) & _
                                   TxtCheFecEnt.Value & Chr(9) & TxtCheNombre.Text & Chr(9) & TxtCheMotivo.Text & Chr(9) & _
                                   txtCuentaResponsable.Text & Chr(9) & txtCuit.Text
                                   
        cmdSalir_Click
        Exit Sub
    End If
  
    On Error GoTo CLAVOSE
    
    DBConn.BeginTrans
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    Me.Refresh
    
    sql = "SELECT * FROM CHEQUE WHERE CHE_NUMERO = " & XS(TxtCheNumero.Text) & " AND BAN_CODINT = " & XN(TxtCodInt.Text)
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.RecordCount = 0 Then
         sql = "INSERT INTO CHEQUE(CHE_NUMERO,BAN_CODINT,CHE_NOMBRE,CHE_CUIT,CHE_NOMCTA,"
         sql = sql & " CHE_IMPORT,CHE_FECEMI,CHE_FECVTO,CHE_FECENT,CHE_MOTIVO,CHE_OBSERV)"
         sql = sql & " VALUES (" & XS(Me.TxtCheNumero.Text) & ","
         sql = sql & XN(Me.TxtCodInt.Text) & "," & XS(Me.TxtCheNombre.Text) & ","
         sql = sql & XS(Me.txtCuit.Text) & "," & XS(Me.txtCuentaResponsable.Text) & ","
         sql = sql & XN(Me.TxtCheImport.Text) & "," & XDQ(Me.TxtCheFecEmi.Value) & ","
         sql = sql & XDQ(Me.TxtCheFecVto.Value) & "," & XDQ(Me.TxtCheFecEnt.Value) & ","
         sql = sql & XS(Me.TxtCheMotivo.Text) & "," & XS(Me.TxtCheObserv.Text) & " )"
         DBConn.Execute sql
         
         'Insert en la Tabla de Estados de Cheques
        sql = "INSERT INTO CHEQUE_ESTADOS (CHE_NUMERO,BAN_CODINT,ECH_CODIGO,CES_FECHA,CES_DESCRI)"
        sql = sql & " VALUES ("
        sql = sql & XS(Me.TxtCheNumero.Text) & ","
        sql = sql & XN(Me.TxtCodInt.Text) & "," & XN(1) & ","
        sql = sql & XDQ(Date) & ",'CHEQUE EN CARTERA')"
        DBConn.Execute sql
    Else
         sql = "UPDATE CHEQUE SET CHE_NOMBRE = " & XS(Me.TxtCheNombre.Text)
         sql = sql & ",CHE_CUIT = " & XS(Me.txtCuit.Text)
         sql = sql & ",CHE_NOMCTA = " & XS(Me.txtCuentaResponsable.Text)
         sql = sql & ",CHE_IMPORT = " & XN(Me.TxtCheImport.Text)
         sql = sql & ",CHE_FECEMI =" & XDQ(Me.TxtCheFecEmi.Value)
         sql = sql & ",CHE_FECVTO =" & XDQ(Me.TxtCheFecVto.Value)
         sql = sql & ",CHE_FECENT = " & XDQ(Me.TxtCheFecEnt.Value)
         sql = sql & ",CHE_MOTIVO = " & XS(Me.TxtCheMotivo.Text)
         sql = sql & ",CHE_OBSERV = " & XS(Me.TxtCheObserv.Text)
         sql = sql & " WHERE CHE_NUMERO = " & XS(Me.TxtCheNumero.Text)
         sql = sql & " AND BAN_CODINT = " & XN(Me.TxtCodInt.Text)
         DBConn.Execute sql
    End If
    Rec.Close
        
    '************* PREGUNTAR POR SI DESEA IMPRIMIR ***************
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    DBConn.CommitTrans
    CmdNuevo_Click
 End If
 Exit Sub
      
CLAVOSE:
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    If Rec.State = 1 Then Rec.Close
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
End Sub

Private Sub CmdNuevo_Click()
    Me.TxtCheFecEnt.Value = ""
    Me.TxtCheNumero.Enabled = True
    Me.TxtBanco.Enabled = True
    Me.TxtLOCALIDAD.Enabled = True
    Me.TxtSucursal.Enabled = True
    Me.TxtCodigo.Enabled = True
    Me.TxtCheNombre.Enabled = True
    MtrObjetos = Array(TxtCheNumero, TxtBanco, TxtLOCALIDAD, TxtSucursal, TxtCodigo)
    Call CambiarColor(MtrObjetos, 5, &H80000005, "E")
    TxtCheNombre.ForeColor = &H80000008
    Me.TxtCheNumero.Text = ""
    Me.TxtBanco.Text = ""
    Me.TxtLOCALIDAD.Text = ""
    Me.TxtSucursal.Text = ""
    Me.TxtCodigo.Text = ""
    Me.TxtCodInt.Text = ""
    Me.TxtBanDescri.Text = ""
    Me.TxtCheNombre.Text = ""
    Me.TxtCheMotivo.Text = ""
    Me.TxtCheFecEmi.Value = ""
    Me.TxtCheFecVto.Value = ""
    Me.TxtCheImport.Text = ""
    Me.TxtCheObserv.Text = ""
    Me.txtCuit.Text = ""
    Me.txtCuentaResponsable.Text = ""
    Me.TxtCheFecEnt.SetFocus
    lblEstado.Caption = ""
End Sub

Private Sub cmdSalir_Click()
    Unload Me
    Set FrmCargaCheques = Nothing
End Sub

Private Sub Form_Activate()
    If mMeLlamo = "RECIBO" Then
        CmdNuevo.Enabled = False
        CmdBorrar.Enabled = False
        TxtCheFecEnt.Enabled = False
        TxtCheNumero.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    Set Rec = New ADODB.Recordset
    
    TxtCheFecEnt.Value = Date
    lblEstado.Caption = ""
End Sub

Private Sub TxtBANCO_GotFocus()
    SelecTexto TxtBanco
End Sub

Private Sub TxtBANCO_LostFocus()
    If Len(TxtBanco.Text) < 3 Then TxtBanco.Text = CompletarConCeros(TxtBanco.Text, 3)
End Sub

Private Sub TxtCheImport_GotFocus()
    SelecTexto TxtCheImport
End Sub

Private Sub TxtCheImport_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(TxtCheImport.Text, KeyAscii)
End Sub

Private Sub TxtCheImport_LostFocus()
   If Trim(TxtCheImport.Text) <> "" Then TxtCheImport.Text = Valido_Importe(TxtCheImport)
End Sub

Private Sub TxtCheMotivo_GotFocus()
    SelecTexto TxtCheMotivo
End Sub

Private Sub TxtCheNombre_GotFocus()
    SelecTexto TxtCheNombre
End Sub

Private Sub TxtCheNumero_GotFocus()
    SelecTexto TxtCheNumero
End Sub

Private Sub TxtCheNumero_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCheMotivo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCheNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCheNumero_LostFocus()
   If Len(TxtCheNumero.Text) < 10 Then TxtCheNumero.Text = CompletarConCeros(TxtCheNumero.Text, 10)
End Sub

Private Sub TxtCheObserv_GotFocus()
    SelecTexto TxtCheObserv
End Sub

Private Sub TxtCheObserv_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtBanco_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto TxtCodigo
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroTE(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    Dim MtrObjetos As Variant
    
    Set Rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
        
    'ChequeRegistrado = False
    
    If Len(TxtCodigo.Text) < 6 Then TxtCodigo.Text = CompletarConCeros(TxtCodigo.Text, 6)
     
    If Trim(Me.TxtCheNumero.Text) <> "" And _
       Trim(Me.TxtBanco.Text) <> "" And _
       Trim(Me.TxtLOCALIDAD.Text) <> "" And _
       Trim(Me.TxtSucursal.Text) <> "" And _
       Trim(Me.TxtCodigo.Text) <> "" Then
       
       'BUSCO EL CODIGO INTERNO
       sql = "SELECT BAN_CODINT, BAN_DESCRI FROM BANCO "
       sql = sql & " WHERE BAN_BANCO = " & XS(TxtBanco.Text)
       sql = sql & " AND BAN_LOCALIDAD = " & XS(Me.TxtLOCALIDAD.Text)
       sql = sql & " AND BAN_SUCURSAL = " & XS(Me.TxtSucursal.Text)
       sql = sql & " AND BAN_CODIGO = " & XS(TxtCodigo.Text)
       Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
       If Rec.RecordCount > 0 Then 'EXITE
          TxtCodInt.Text = Rec!BAN_CODINT
          TxtBanDescri.Text = Rec!BAN_DESCRI
          Rec.Close
       Else
          If Me.ActiveControl.Name <> "CmdSalir" And Me.ActiveControl.Name <> "CmdNuevo" Then
            MsgBox "Banco NO Registrado.", 16, TIT_MSGBOX
            Me.CmdBanco.SetFocus
          End If
          Rec.Close
          Exit Sub
       End If
       
       'CONSULTO SI EXISTE EL CHEQUE
        sql = "SELECT * FROM CHEQUE "
        sql = sql & " WHERE CHE_NUMERO = " & XS(TxtCheNumero.Text)
        sql = sql & " AND BAN_CODINT = " & XN(TxtCodInt.Text)
        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec.RecordCount > 0 Then 'EXITE
            Me.TxtCheFecEnt.Value = Rec!CHE_FECENT
            Me.TxtCheNumero.Text = Trim(Rec!CHE_NUMERO)
            
            'BUSCO LOS ATRIBUTOS DE BANCO
            sql = "SELECT BAN_BANCO,BAN_LOCALIDAD,BAN_SUCURSAL,BAN_CODIGO FROM BANCO " & _
                   "WHERE BAN_CODINT = " & XN(Me.TxtCodInt.Text)
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.RecordCount > 0 Then 'EXITE
                Me.TxtBanco.Text = Rec1!BAN_BANCO
                Me.TxtLOCALIDAD.Text = Rec1!BAN_LOCALIDAD
                Me.TxtSucursal.Text = Rec1!BAN_SUCURSAL
                Me.TxtCodigo.Text = Rec1!BAN_CODIGO
            End If
            Rec1.Close
            Me.TxtCheNombre.Text = ChkNull(Rec!CHE_NOMBRE)
            Me.txtCuit.Text = ChkNull(Rec!CHE_CUIT)
            Me.txtCuentaResponsable.Text = ChkNull(Rec!CHE_NOMCTA)
            Me.TxtCheMotivo.Text = Rec!CHE_MOTIVO
            Me.TxtCheFecEmi.Value = Rec!CHE_FECEMI
            Me.TxtCheFecVto.Value = Rec!CHE_FECVTO
            Me.TxtCheImport.Text = Valido_Importe(Rec!CHE_IMPORT)
            Me.TxtCheObserv.Text = ChkNull(Rec!CHE_OBSERV)
            
            TxtCheNumero.Enabled = False
            TxtBanco.Enabled = False
            TxtLOCALIDAD.Enabled = False
            TxtSucursal.Enabled = False
            TxtCodigo.Enabled = False
            
            MtrObjetos = Array(TxtCheNumero, TxtBanco, TxtLOCALIDAD, TxtSucursal, TxtCodigo)
            Call CambiarColor(MtrObjetos, 5, &H80000018, "D")
            
        Else
           TxtCheNombre.ForeColor = &H80000008
           Rec.Close
           Exit Sub
        End If
        If Rec.State = 1 Then Rec.Close
    End If
End Sub

Private Sub txtCuentaResponsable_GotFocus()
    SelecTexto txtCuentaResponsable
End Sub

Private Sub txtCuentaResponsable_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtCuit_GotFocus()
    txtCuit.SelStart = 0
    txtCuit.SelLength = Len(txtCuit) + 2
End Sub

Private Sub txtCuit_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCuit_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(Trim(txtCuit.ClipText)) = 12 Then
      txtCuit.SelStart = 12
  End If
End Sub

Private Sub txtCuit_LostFocus()
    If ActiveControl.Name = "CmdNuevo" Or ActiveControl.Name = "CmdSalir" Then Exit Sub
    If txtCuit.Text <> "" Then
        'rutina de validación de CUIT
        If Not ValidoCuit(txtCuit) Then
            txtCuit.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub TxtLOCALIDAD_GotFocus()
    SelecTexto TxtLOCALIDAD
End Sub

Private Sub Txtlocalidad_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtLOCALIDAD_LostFocus()
    If Len(TxtLOCALIDAD.Text) < 3 Then TxtLOCALIDAD.Text = CompletarConCeros(TxtLOCALIDAD.Text, 3)
End Sub

Private Sub txtSucursal_GotFocus()
    SelecTexto TxtSucursal
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtSucursal_LostFocus()
    If Len(TxtSucursal.Text) < 3 Then TxtSucursal.Text = CompletarConCeros(TxtSucursal.Text, 3)
End Sub
