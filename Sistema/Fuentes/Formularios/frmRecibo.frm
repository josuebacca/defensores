VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRecibo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Pagos"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
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
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAnular 
      Caption         =   "&Anular"
      Height          =   450
      Left            =   6030
      Picture         =   "frmRecibo.frx":0000
      TabIndex        =   52
      Top             =   5955
      Width           =   1110
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   450
      Left            =   4890
      Picture         =   "frmRecibo.frx":0442
      TabIndex        =   49
      Top             =   5955
      Width           =   1110
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   7170
      Picture         =   "frmRecibo.frx":110C
      TabIndex        =   19
      Top             =   5955
      Width           =   1110
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   8310
      Picture         =   "frmRecibo.frx":154E
      TabIndex        =   18
      Top             =   5955
      Width           =   1110
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   450
      Left            =   3750
      Picture         =   "frmRecibo.frx":2218
      TabIndex        =   17
      Top             =   5955
      Width           =   1110
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5835
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10292
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Datos"
      TabPicture(0)   =   "frmRecibo.frx":2EE2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label25"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkComision"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "GrdSocios"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FrameSocio"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FrameRecibo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtTotal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdFormaPago"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtRecargo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fraPagos"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmRecibo.frx":2EFE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "grdBuscar"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraPagos 
         Height          =   5175
         Left            =   1605
         TabIndex        =   22
         Top             =   585
         Width           =   4935
         Begin VB.CommandButton cmdAgregarPago 
            Caption         =   "Agregar"
            Height          =   315
            Left            =   3555
            TabIndex        =   25
            Top             =   1815
            Width           =   1245
         End
         Begin VB.CommandButton cmdCerrarPagos 
            Caption         =   "Cerrar"
            Height          =   375
            Left            =   3630
            TabIndex        =   31
            Top             =   4695
            Width           =   1095
         End
         Begin VB.Frame Frame3 
            Height          =   795
            Left            =   120
            TabIndex        =   28
            Top             =   570
            Width           =   4695
            Begin VB.TextBox txtTotalPagos 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   375
               Left            =   3120
               TabIndex        =   29
               Top             =   300
               Width           =   1515
            End
            Begin VB.Label Label35 
               Alignment       =   2  'Center
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "T O T A L"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   90
               TabIndex        =   30
               Top             =   300
               Width           =   3015
            End
         End
         Begin VB.CommandButton cmdBorroFila 
            Caption         =   "Borrar Fila"
            Height          =   375
            Left            =   90
            TabIndex        =   27
            Top             =   4695
            Width           =   1095
         End
         Begin VB.ComboBox cboFormaPago 
            Height          =   315
            ItemData        =   "frmRecibo.frx":2F1A
            Left            =   1470
            List            =   "frmRecibo.frx":2F1C
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1470
            Width           =   3330
         End
         Begin VB.CommandButton cmdAceptarPagos 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   2160
            TabIndex        =   26
            Top             =   4695
            Width           =   1425
         End
         Begin VB.TextBox txtImportePago 
            Height          =   315
            Left            =   1470
            TabIndex        =   24
            Top             =   1815
            Width           =   1245
         End
         Begin MSFlexGridLib.MSFlexGrid grdPagos 
            Height          =   2445
            Left            =   120
            TabIndex        =   32
            Top             =   2190
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   4313
            _Version        =   393216
            Rows            =   1
            Cols            =   13
            FixedCols       =   0
            RowHeightMin    =   280
            BackColorSel    =   16761024
            ForeColorSel    =   16777215
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            SelectionMode   =   1
            FormatString    =   $"frmRecibo.frx":2F1E
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Forma Pago"
            Height          =   330
            Left            =   120
            TabIndex        =   35
            Top             =   1470
            Width           =   1320
         End
         Begin VB.Label Label40 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            Caption         =   "Forma de Pago"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   45
            TabIndex        =   34
            Top             =   120
            Width           =   4845
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Importe:"
            Height          =   330
            Left            =   120
            TabIndex        =   33
            Top             =   1815
            Width           =   1320
         End
      End
      Begin VB.TextBox txtRecargo 
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
         Left            =   5550
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   5280
         Width           =   1290
      End
      Begin VB.Frame Frame4 
         Caption         =   "Buscar por..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   -74730
         TabIndex        =   37
         Top             =   480
         Width           =   8805
         Begin MSComCtl2.DTPicker FechaD 
            Height          =   315
            Left            =   1440
            TabIndex        =   39
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   52232193
            CurrentDate     =   42839
         End
         Begin VB.CommandButton cmdBuscarRecibos 
            Caption         =   "Buscar"
            Height          =   450
            Left            =   7170
            Picture         =   "frmRecibo.frx":2F24
            TabIndex        =   42
            Top             =   570
            Width           =   1395
         End
         Begin VB.CommandButton cmdBuscaSocioB 
            Height          =   315
            Left            =   6600
            Picture         =   "frmRecibo.frx":3BEE
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Buscar Socio"
            Top             =   330
            Width           =   330
         End
         Begin VB.TextBox txtNomSocB 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            TabIndex        =   41
            Top             =   330
            Width           =   4275
         End
         Begin VB.TextBox txtCodSocB 
            Alignment       =   2  'Center
            Height          =   330
            Left            =   1440
            TabIndex        =   38
            Top             =   330
            Width           =   855
         End
         Begin MSComCtl2.DTPicker FechaH 
            Height          =   315
            Left            =   5400
            TabIndex        =   40
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   52232193
            CurrentDate     =   42839
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4380
            TabIndex        =   47
            Top             =   735
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   330
            TabIndex        =   46
            Top             =   735
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Socio:"
            Height          =   210
            Left            =   330
            TabIndex        =   45
            Top             =   375
            Width           =   435
         End
      End
      Begin VB.CommandButton cmdFormaPago 
         Caption         =   "Forma Pago"
         Height          =   450
         Left            =   90
         TabIndex        =   36
         Top             =   5280
         Width           =   1965
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
         Left            =   7740
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   5280
         Width           =   1440
      End
      Begin VB.Frame FrameRecibo 
         Caption         =   "Recibo..."
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
         Left            =   6825
         TabIndex        =   12
         Top             =   465
         Width           =   2370
         Begin MSComCtl2.DTPicker fecha1 
            Height          =   375
            Left            =   720
            TabIndex        =   54
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   52232193
            CurrentDate     =   42839
         End
         Begin VB.TextBox txtNroRecibo 
            Height          =   375
            Left            =   1935
            TabIndex        =   14
            Top             =   1065
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label lblRecibo 
            Alignment       =   2  'Center
            Caption         =   "Numero"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   60
            TabIndex        =   15
            Top             =   345
            Width           =   2265
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   135
            TabIndex        =   13
            Top             =   870
            Width           =   495
         End
      End
      Begin VB.Frame FrameSocio 
         Caption         =   "Socio..."
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
         Left            =   105
         TabIndex        =   3
         Top             =   465
         Width           =   6720
         Begin VB.TextBox txtCodSoc 
            Alignment       =   2  'Center
            Height          =   330
            Left            =   915
            TabIndex        =   0
            Top             =   345
            Width           =   855
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
            Height          =   330
            Left            =   1785
            TabIndex        =   1
            Top             =   345
            Width           =   4275
         End
         Begin VB.CommandButton cmdBuscaSocio 
            Height          =   315
            Left            =   6105
            Picture         =   "frmRecibo.frx":3F78
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Buscar Socio"
            Top             =   345
            Width           =   330
         End
         Begin VB.TextBox txtTelefono 
            Enabled         =   0   'False
            Height          =   330
            Left            =   4665
            TabIndex        =   6
            Top             =   1035
            Width           =   1770
         End
         Begin VB.TextBox txtNroDoc 
            Enabled         =   0   'False
            Height          =   330
            Left            =   915
            TabIndex        =   5
            Top             =   1035
            Width           =   1770
         End
         Begin VB.TextBox txtDomici 
            Enabled         =   0   'False
            Height          =   330
            Left            =   915
            TabIndex        =   4
            Top             =   690
            Width           =   5520
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Socio:"
            Height          =   195
            Left            =   180
            TabIndex        =   11
            Top             =   375
            Width           =   435
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   3915
            TabIndex        =   10
            Top             =   1080
            Width           =   690
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Nro Doc:"
            Height          =   195
            Left            =   180
            TabIndex        =   9
            Top             =   1080
            Width           =   630
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   180
            TabIndex        =   8
            Top             =   735
            Width           =   660
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdSocios 
         Height          =   3225
         Left            =   90
         TabIndex        =   16
         Top             =   2025
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   5689
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         RowHeightMin    =   280
         BackColorSel    =   16761024
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
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
      Begin MSFlexGridLib.MSFlexGrid grdBuscar 
         Height          =   3825
         Left            =   -74730
         TabIndex        =   43
         Top             =   1830
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   6747
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         RowHeightMin    =   280
         BackColorSel    =   16761024
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
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
      Begin VB.CheckBox chkComision 
         Caption         =   "Recibo con Comisión"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   48
         Top             =   5355
         Width           =   2175
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Recargo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4500
         TabIndex        =   51
         Top             =   5325
         Width           =   1005
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6960
         TabIndex        =   21
         Top             =   5325
         Width           =   645
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   105
      Top             =   5955
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   735
      TabIndex        =   53
      Top             =   6030
      Width           =   825
   End
End
Attribute VB_Name = "frmRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim J As Integer
Dim mBuscar As Boolean
Dim mDiasRecargo As Integer
Dim Rec As ADODB.Recordset
Dim Rec1 As ADODB.Recordset
Dim Rec2 As ADODB.Recordset

Private Sub cboFormaPago_LostFocus()
    If Me.ActiveControl.Name = "grdPagos" Then
        Exit Sub
    End If
End Sub

Private Sub cmdAceptar_Click()
    If Validar() = False Then Exit Sub
    
    On Error GoTo CLAVOSE
    
    DBConn.BeginTrans
    Screen.MousePointer = vbHourglass
    
    sql = "INSERT INTO RECIBO (REC_NUMERO,REC_FECHA,SOC_CODIGO,REC_IMPORTE,REC_RECARGO,REC_NROTXT,REC_ESTADO,REC_COMISION)"
    sql = sql & " VALUES ("
    sql = sql & XN(txtNroRecibo.Text) & ","
    sql = sql & XDQ(fecha1.Value) & ","
    sql = sql & XN(txtCodSoc.Text) & ","
    sql = sql & XN(txtTotal.Text) & ","
    sql = sql & XN(txtRecargo.Text) & ","
    sql = sql & XS(Format(txtNroRecibo.Text, "00000000")) & ",1,"
    If chkComision.Value = Checked Then
        sql = sql & XS("S") & ")"
    Else
        sql = sql & "NULL)"
    End If
    DBConn.Execute sql
    
    For I = 1 To grdPagos.Rows - 1
        If grdPagos.TextMatrix(I, 2) = "1" Then 'EFECTIVO
            
            sql = "INSERT INTO RECIBO_PAGOS (REC_NUMERO,REC_ITEM,FPG_CODIGO,REC_PAGO)"
            sql = sql & " VALUES ("
            sql = sql & XN(txtNroRecibo.Text) & "," & I & ","
            sql = sql & XN(grdPagos.TextMatrix(I, 2)) & ","
            sql = sql & XN(grdPagos.TextMatrix(I, 1)) & ")"
            DBConn.Execute sql
            
        ElseIf grdPagos.TextMatrix(I, 2) = "2" Then 'CHEQUE TERCERO
            sql = "INSERT INTO RECIBO_PAGOS (REC_NUMERO,REC_ITEM,FPG_CODIGO,REC_PAGO,BAN_CODINT,CHE_NUMERO)"
            sql = sql & " VALUES ("
            sql = sql & XN(txtNroRecibo.Text) & "," & I & ","
            sql = sql & XN(grdPagos.TextMatrix(I, 2)) & ","
            sql = sql & XN(grdPagos.TextMatrix(I, 1)) & ","
            sql = sql & XN(grdPagos.TextMatrix(I, 3)) & ","
            sql = sql & XS(grdPagos.TextMatrix(I, 4)) & ")"
            DBConn.Execute sql
            
            'DOY DE ALTA EL CHEQUE
            sql = "SELECT * FROM CHEQUE WHERE CHE_NUMERO = " & XS(grdPagos.TextMatrix(I, 4)) & " AND BAN_CODINT = " & XN(grdPagos.TextMatrix(I, 3))
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.RecordCount = 0 Then
                 sql = "INSERT INTO CHEQUE(CHE_NUMERO,BAN_CODINT,CHE_NOMBRE,CHE_CUIT,CHE_NOMCTA,"
                 sql = sql & " CHE_IMPORT,CHE_FECEMI,CHE_FECVTO,CHE_FECENT,CHE_MOTIVO,CHE_OBSERV)"
                 sql = sql & " VALUES (" & XS(grdPagos.TextMatrix(I, 4)) & ","
                 sql = sql & XN(grdPagos.TextMatrix(I, 3)) & "," & XS(grdPagos.TextMatrix(I, 9)) & ","
                 sql = sql & XS(grdPagos.TextMatrix(I, 12)) & "," & XS(grdPagos.TextMatrix(I, 11)) & ","
                 sql = sql & XN(grdPagos.TextMatrix(I, 1)) & "," & XDQ(grdPagos.TextMatrix(I, 5)) & ","
                 sql = sql & XDQ(grdPagos.TextMatrix(I, 6)) & "," & XDQ(grdPagos.TextMatrix(I, 8)) & ","
                 sql = sql & XS(grdPagos.TextMatrix(I, 10)) & "," & XS(grdPagos.TextMatrix(I, 7)) & " )"
                 DBConn.Execute sql
                 
                 'Insert en la Tabla de Estados de Cheques
                sql = "INSERT INTO CHEQUE_ESTADOS (CHE_NUMERO,BAN_CODINT,ECH_CODIGO,CES_FECHA,CES_DESCRI)"
                sql = sql & " VALUES ("
                sql = sql & XS(grdPagos.TextMatrix(I, 4)) & ","
                sql = sql & XN(grdPagos.TextMatrix(I, 3)) & "," & XN(1) & ","
                sql = sql & XDQ(Date) & ",'CHEQUE EN CARTERA')"
                DBConn.Execute sql
            Else
                 sql = "UPDATE CHEQUE SET CHE_NOMBRE = " & XS(grdPagos.TextMatrix(I, 9))
                 sql = sql & ",CHE_CUIT = " & XS(grdPagos.TextMatrix(I, 12))
                 sql = sql & ",CHE_NOMCTA = " & XS(grdPagos.TextMatrix(I, 11))
                 sql = sql & ",CHE_IMPORT = " & XN(grdPagos.TextMatrix(I, 1))
                 sql = sql & ",CHE_FECEMI =" & XDQ(grdPagos.TextMatrix(I, 5))
                 sql = sql & ",CHE_FECVTO =" & XDQ(grdPagos.TextMatrix(I, 6))
                 sql = sql & ",CHE_FECENT = " & XDQ(grdPagos.TextMatrix(I, 8))
                 sql = sql & ",CHE_MOTIVO = " & XS(grdPagos.TextMatrix(I, 10))
                 sql = sql & ",CHE_OBSERV = " & XS(grdPagos.TextMatrix(I, 7))
                 sql = sql & " WHERE CHE_NUMERO = " & XS(grdPagos.TextMatrix(I, 4))
                 sql = sql & " AND BAN_CODINT = " & XN(grdPagos.TextMatrix(I, 3))
                 DBConn.Execute sql
            End If
            Rec1.Close
        End If
    Next I
    For J = 1 To GrdSocios.Rows - 1
        If GrdSocios.TextMatrix(J, 7) = "S" Then
            sql = "UPDATE DEBITOS"
            sql = sql & " SET REC_NUMERO=" & XN(txtNroRecibo.Text)
            sql = sql & " ,DEB_SALDO=DEB_SALDO-" & XN(GrdSocios.TextMatrix(J, 4))
            sql = sql & " ,DEB_PAGADO='S'"
            sql = sql & " WHERE "
            sql = sql & " DEB_MES=" & XN(GrdSocios.TextMatrix(J, 0))
            sql = sql & " AND DEB_ANO=" & XN(GrdSocios.TextMatrix(J, 1))
            sql = sql & " AND SOC_CODIGO=" & XN(GrdSocios.TextMatrix(J, 5))
            sql = sql & " AND DEB_ITEM=" & XN(GrdSocios.TextMatrix(J, 6))
            DBConn.Execute sql
        End If
    Next J
    
    DBConn.CommitTrans
    Screen.MousePointer = vbNormal
    cmdImprimir_Click
    
    CmdNuevo_Click
    Exit Sub
      
CLAVOSE:
    DBConn.RollbackTrans
    If Rec.State = 1 Then Rec.Close
    If Rec1.State = 1 Then Rec1.Close
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub cmdAceptarPagos_Click()
    fraPagos.Visible = False
End Sub

Private Sub cmdAgregarPago_Click()
    txtImportePago.Text = Valido_Importe(txtImportePago.Text)
    If cboFormaPago.Text = "" Then
        MsgBox "Debe Indicar la Forma de Pago", vbCritical, TIT_MSGBOX
        cboFormaPago.SetFocus
        Exit Sub
    End If
        
    Dim mTotalPagos As Double
    mTotalPagos = 0
    
    For I = 1 To grdPagos.Rows - 1
        mTotalPagos = mTotalPagos + CDbl(grdPagos.TextMatrix(I, 1))
    Next
    If mTotalPagos + CDbl(Chk0(txtImportePago.Text)) > CDbl(txtTotal.Text) Then
        MsgBox "El Importe Ingresado Exede el Monto!", vbInformation, TIT_MSGBOX
        txtImportePago.SetFocus
        Exit Sub
    End If
    If CDbl(Chk0(txtImportePago.Text)) > 0 Then
       If Trim(cboFormaPago.Text) = "CHEQUE TERCERO" Then
           FrmCargaCheques.mMeLlamo = "RECIBO"
           FrmCargaCheques.TxtCheImport.Text = txtImportePago.Text
           FrmCargaCheques.Show vbModal
       End If
       If Trim(UCase(Mid(cboFormaPago.Text, 1, 50))) = "EFECTIVO" Then
           grdPagos.AddItem ("")
           grdPagos.row = grdPagos.Rows - 1
           grdPagos.TextMatrix(grdPagos.row, 0) = Trim(Mid(cboFormaPago.Text, 1, 30))
           grdPagos.TextMatrix(grdPagos.row, 1) = txtImportePago.Text
           grdPagos.TextMatrix(grdPagos.row, 2) = cboFormaPago.ItemData(cboFormaPago.ListIndex)
       End If
    End If
    mTotalPagos = 0
    For I = 1 To grdPagos.Rows - 1
        mTotalPagos = CDbl(mTotalPagos) + CDbl(grdPagos.TextMatrix(I, 1))
    Next
    txtTotalPagos.Text = Format(CDbl(txtTotal.Text) - mTotalPagos, "0.00")
    txtImportePago.Text = Format(txtTotalPagos.Text, "0.00")
    If Val(txtTotalPagos.Text) = 0 Then
        cmdAceptarPagos.SetFocus
    Else
        cboFormaPago.ListIndex = 0
        cboFormaPago.SetFocus
    End If
End Sub

Private Sub CmdAnular_Click()
    sql = "SELECT MAX(CAJA_FECHA) AS FECHA FROM CAJA"
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        If CDate(fecha1.Value) <= CDate(Rec!Fecha) Then
            MsgBox "No se puede Anular el Recibo" & Chr(13) & _
                   "La Fecha del Recibo es Menor a la Feha de la Caja", vbExclamation, TIT_MSGBOX
                   
            Rec.Close
            Exit Sub
        End If
    End If
    Rec.Close
    
    If MsgBox("¿Seguro que Anula el Recibo?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        sql = "SELECT REC_ESTADO"
        sql = sql & " FROM RECIBO"
        sql = sql & " WHERE REC_NUMERO=" & XN(txtNroRecibo.Text)
        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec.EOF = False Then
            If Rec!REC_ESTADO = 1 Then
                For J = 1 To GrdSocios.Rows - 1
                    sql = "UPDATE DEBITOS"
                    sql = sql & " SET REC_NUMERO=NULL"
                    sql = sql & " ,DEB_SALDO=DEB_SALDO+" & XN(GrdSocios.TextMatrix(J, 4))
                    sql = sql & " ,DEB_PAGADO=NULL"
                    sql = sql & " WHERE "
                    sql = sql & " REC_NUMERO=" & XN(txtNroRecibo.Text)
                    sql = sql & " AND DEB_MES=" & XN(GrdSocios.TextMatrix(J, 0))
                    sql = sql & " AND DEB_ANO=" & XN(GrdSocios.TextMatrix(J, 1))
                    sql = sql & " AND SOC_CODIGO=" & XN(GrdSocios.TextMatrix(J, 5))
                    sql = sql & " AND DEB_ITEM=" & XN(GrdSocios.TextMatrix(J, 6))
                    DBConn.Execute sql
                Next
            End If
            
            sql = "UPDATE RECIBO"
            sql = sql & " SET REC_ESTADO=2"
            sql = sql & " WHERE "
            sql = sql & " REC_NUMERO=" & XN(txtNroRecibo.Text)
            DBConn.Execute sql
            
            MsgBox "Recibo Anulado", vbInformation, TIT_MSGBOX
        End If
        Rec.Close
    End If
    
    CmdNuevo_Click
End Sub

Private Sub cmdBorroFila_Click()
    If grdPagos.Rows <= 2 Then
        grdPagos.Rows = 1
    Else
        grdPagos.RemoveItem (grdPagos.row)
    End If
    Dim mTotalPagos As Double
    mTotalPagos = 0
    For I = 1 To grdPagos.Rows - 1
        mTotalPagos = CDbl(mTotalPagos) + CDbl(grdPagos.TextMatrix(I, 1))
    Next
    txtTotalPagos.Text = Format(CDbl(txtTotal.Text) - mTotalPagos, "0.00")
    cboFormaPago.SetFocus
End Sub

Private Sub cmdBuscarRecibos_Click()
    grdBuscar.Rows = 1
    grdBuscar.HighLight = flexHighlightNever
    
    sql = "SELECT S.SOC_NOMBRE, S.SOC_CODIGO, R.REC_FECHA, R.REC_ESTADO,"
    sql = sql & " R.REC_IMPORTE, R.REC_NUMERO, R.REC_COMISION, R.REC_RECARGO"
    sql = sql & " FROM RECIBO R, SOCIOS S"
    sql = sql & " WHERE S.SOC_CODIGO=R.SOC_CODIGO"
    If txtCodSocB.Text <> "" Then
        sql = sql & " AND S.SOC_CODIGO=" & XN(txtCodSocB.Text)
    End If
    If FechaD.Value <> "" Then
        sql = sql & " AND R.REC_FECHA>=" & XDQ(FechaD.Value)
    End If
    If FechaH.Value <> "" Then
        sql = sql & " AND R.REC_FECHA<=" & XDQ(FechaH.Value)
    End If
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        Do While Rec.EOF = False
            grdBuscar.AddItem Rec!REC_FECHA & Chr(9) & Format(Rec!REC_NUMERO, "00000000") & Chr(9) & _
                              Rec!SOC_NOMBRE & Chr(9) & Valido_Importe(Rec!REC_IMPORTE) & Chr(9) & _
                              Rec!SOC_CODIGO & Chr(9) & ChkNull(Rec!REC_COMISION) & Chr(9) & _
                              Valido_Importe(Chk0(Rec!REC_RECARGO)) & Chr(9) & Rec!REC_ESTADO
            Rec.MoveNext
        Loop
        grdBuscar.HighLight = flexHighlightAlways
        grdBuscar.row = 1
        grdBuscar.Col = 0
        grdBuscar.SetFocus
    Else
        grdBuscar.HighLight = flexHighlightNever
        MsgBox "No se encontraron Datos", vbExclamation, TIT_MSGBOX
        txtCodSocB.SetFocus
    End If
    Rec.Close
End Sub

Private Sub cmdBuscaSocio_Click()
    txtCodSoc.Text = ""
    BuscarSocios "txtCodSoc", "CODIGO"
    txtNomSoc.SetFocus
End Sub

Private Sub cmdBuscaSocioB_Click()
    txtCodSocB.Text = ""
    BuscarSocios "txtCodSocB", "CODIGO"
    txtNomSocB.SetFocus
End Sub

Private Sub cmdCerrarPagos_Click()
    fraPagos.Visible = False
End Sub

Private Sub cmdFormaPago_Click()
    If CDbl(Chk0(txtTotal.Text)) > 0 Then
        fraPagos.Top = 495
        fraPagos.Left = 2025
        fraPagos.Visible = True
        
        Dim mTotalPagos As Double
        mTotalPagos = 0
        For I = 1 To grdPagos.Rows - 1
            mTotalPagos = mTotalPagos + CDbl(grdPagos.TextMatrix(I, 1))
        Next
        txtTotalPagos.Text = Format(CDbl(txtTotal.Text) - mTotalPagos, "0.00")
        If mBuscar = False Then
            cboFormaPago.Enabled = True
            cboFormaPago.SetFocus
        End If
    End If
End Sub

Private Sub cmdImprimir_Click()
    If MsgBox("¿Imprime Recibo?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    Rep.SubreportToChange = ""
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    
    Rep.SelectionFormula = "{RECIBO.REC_NUMERO}=" & XN(CLng(txtNroRecibo.Text))
    
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.Destination = crptToWindow
    Rep.WindowTitle = "Recibo"
    Rep.ReportFileName = DRIVE & DirReport & "Recibo.rpt"
    
    'PARA EL SUBREPORTE
    Rep.SubreportToChange = "SubReporte_Recibo.rpt"
    Rep.SelectionFormula = ""
    Rep.SelectionFormula = "{RECIBO_PAGOS.REC_NUMERO}=" & XN(CLng(txtNroRecibo.Text))
    
    Rep.Action = 1
    
    Rep.SelectionFormula = ""
    Rep.SubreportToChange = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
End Sub

Private Sub CmdNuevo_Click()
    SSTab1.Tab = 0
    mBuscar = False
    cmdImprimir.Enabled = False
    fraPagos.Visible = False
    FrameRecibo.Enabled = True
    FrameSocio.Enabled = True
    GrdSocios.Enabled = True
    cmdAceptar.Enabled = True
    cboFormaPago.Enabled = True
    txtImportePago.Enabled = True
    cmdAgregarPago.Enabled = True
    cmdBorroFila.Enabled = True
    cmdAceptarPagos.Enabled = True
    cmdAnular.Enabled = False
    lblEstado.Caption = ""
    GrdSocios.Rows = 1
    grdPagos.Rows = 1
    grdBuscar.Rows = 1
    grdBuscar.HighLight = flexHighlightNever
    txtCodSocB.Text = ""
    FechaD.Value = Date
    FechaH.Value = Date
    chkComision.Value = Unchecked
    txtImportePago.Text = ""
    txtTotalPagos.Text = ""
    txtTotal.Text = ""
    txtRecargo.Text = ""
    txtCodSoc.Text = ""
    txtNroRecibo.Text = ""
    fecha1.Value = Date
    BuscarUltimoNumero
    txtCodSoc.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Set frmRecibo = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    If KeyAscii = 27 Then
        cmdSalir_Click
    End If
End Sub

Private Sub Form_Load()
    Set Rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    SSTab1.Tab = 0
    Centrar_pantalla Me
    mBuscar = False
    cmdImprimir.Enabled = False
    FrameRecibo.Enabled = False
    cmdAnular.Enabled = False
    lblEstado.Caption = ""
    
    fecha1.Value = Date
    
    'CARGO COMBO CON FORMA DE PAGO
    sql = "SELECT FPG_CODIGO, FPG_DESCRI FROM FORMA_PAGO WHERE FPG_CODIGO <> 3"
    sql = sql & " ORDER BY FPG_CODIGO"
    CargarControlItemdata cboFormaPago, sql
    
    GrdSocios.FormatString = "^Mes|^Año|<Socio|<Descripción|>Importe|Cod Socio|Cod ITEM|Paga|RECARGO"
    GrdSocios.ColWidth(0) = 800  'MES
    GrdSocios.ColWidth(1) = 800  'AÑO
    GrdSocios.ColWidth(2) = 3000 'SOCIO
    GrdSocios.ColWidth(3) = 3000 'DESCRIPCIÓN
    GrdSocios.ColWidth(4) = 1000 'IMPORTE
    GrdSocios.ColWidth(5) = 0    'COD SOCIO
    GrdSocios.ColWidth(6) = 0    'COD ITEM
    GrdSocios.ColWidth(7) = 0    'PAGA
    GrdSocios.ColWidth(8) = 0    'RECARGO
    GrdSocios.Cols = 9
    GrdSocios.Rows = 1
    GrdSocios.HighLight = flexHighlightNever
    GrdSocios.BorderStyle = flexBorderNone
    GrdSocios.row = 0
    For I = 0 To GrdSocios.Cols - 1
        GrdSocios.Col = I
        GrdSocios.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdSocios.CellBackColor = &H808080    'GRIS OSCURO
        GrdSocios.CellFontBold = True
    Next
    GrdSocios.MergeCells = 1
    GrdSocios.MergeCol(0) = True  'MES
    GrdSocios.MergeCol(1) = True  'AÑO
    GrdSocios.MergeCol(2) = True  'SOCIO
    GrdSocios.MergeCol(3) = False 'DESCRIPCION DEL CONCEPTO
    GrdSocios.MergeCol(4) = False 'IMPORTE
    GrdSocios.MergeCol(5) = False 'COD SOCIO
    GrdSocios.MergeCol(6) = False 'ITEM
    
    'grilla de busqueda
    grdBuscar.FormatString = "^Fecha|^Nro Recibo|<Socio|>Importe|Cod Socio|COMISION|RECARGO|ESTADO"
    grdBuscar.ColWidth(0) = 1300 'FECHA
    grdBuscar.ColWidth(1) = 1500 'NRO RECIBO
    grdBuscar.ColWidth(2) = 3500 'SOCIO
    grdBuscar.ColWidth(3) = 1200 'IMPORTE
    grdBuscar.ColWidth(4) = 0    'COD SOCIO
    grdBuscar.ColWidth(5) = 0    'COMISION
    grdBuscar.ColWidth(6) = 0    'RECARGO
    grdBuscar.ColWidth(7) = 0    'ESTADO
    grdBuscar.Cols = 8
    grdBuscar.Rows = 1
    grdBuscar.HighLight = flexHighlightNever
    grdBuscar.BorderStyle = flexBorderNone
    grdBuscar.row = 0
    For I = 0 To grdBuscar.Cols - 1
        grdBuscar.Col = I
        grdBuscar.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdBuscar.CellBackColor = &H808080    'GRIS OSCURO
        grdBuscar.CellFontBold = True
    Next
    
    sql = "SELECT DIAS_RECARGO FROM PARAMETROS"
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        mDiasRecargo = Chk0(Rec!DIAS_RECARGO)
    Else
        mDiasRecargo = 0
    End If
    Rec.Close
    
    ConfiguroGrillaPagos
    BuscarUltimoNumero
End Sub

Private Sub BuscarUltimoNumero()
    sql = "SELECT MAX(REC_NUMERO) AS NUMERO FROM RECIBO"
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec.EOF = False Then
        lblRecibo.Caption = "RECIBO   " & Format(CLng(Chk0(Rec!Numero)) + 1, "00000000")
        txtNroRecibo.Text = CLng(Chk0(Rec!Numero)) + 1
    Else
        lblRecibo.Caption = "RECIBO   " & Format(CLng(Chk0(Rec!Numero)), "00000001")
        txtNroRecibo.Text = "1"
    End If
    Rec.Close
End Sub


Private Sub ConfiguroGrillaPagos()
    grdPagos.FormatString = "^Forma Pago|^Importe|Cod.Forma Pago|BAN_CODINT|^Nro Cheque|CHE_FECEMI|" & _
                            "CHE_FECVTO|CHE_OBSERV|CHE_FECENT|CHE_NOMBRE|CHE_MOTIVO|CTA_NROCTA|CHE_CUIT"
    grdPagos.ColWidth(0) = 2000    'FORMA DE PAGO
    grdPagos.ColWidth(1) = 1000    'IMPORTE
    grdPagos.ColWidth(2) = 0       'COD FORMA PAGO
    grdPagos.ColWidth(3) = 0       'BAN_CODINT
    grdPagos.ColWidth(4) = 1500    'CHE_NUMERO
    grdPagos.ColWidth(5) = 0       'CHE_FECEMI
    grdPagos.ColWidth(6) = 0       'CHE_FECVTO
    grdPagos.ColWidth(7) = 0       'CHE_OBSERV
    grdPagos.ColWidth(8) = 0       'CHE_FECENT
    grdPagos.ColWidth(9) = 0       'CHE_NOMBRE
    grdPagos.ColWidth(10) = 0      'CHE_MOTIVO
    grdPagos.ColWidth(11) = 0      'CTA_NROCTA
    grdPagos.ColWidth(12) = 0      'CHE_CUIT
    grdPagos.Rows = 1
    'grdPagos.HighLight = flexHighlightNever
    grdPagos.BorderStyle = flexBorderNone
    grdPagos.row = 0
    For I = 0 To grdPagos.Cols - 1
        grdPagos.Col = I
        grdPagos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdPagos.CellBackColor = &H808080    'GRIS OSCURO
        grdPagos.CellFontBold = True
    Next
    fraPagos.Visible = False
End Sub

Private Sub grdBuscar_DblClick()
    If grdBuscar.Rows > 1 Then
        mBuscar = True
        GrdSocios.Rows = 1
        sql = "SELECT S.SOC_NOMBRE, D.DEB_MES, D.DEB_ANO, S.SOC_CODIGO,"
        sql = sql & " D.DEB_ITEM, D.DEB_DETALLE, D.DEB_IMPORTE, D.DEB_FECHA"
        sql = sql & " FROM DEBITOS D, SOCIOS S"
        sql = sql & " WHERE "
        sql = sql & " S.SOC_CODIGO=D.SOC_CODIGO"
        sql = sql & " AND D.REC_NUMERO=" & XN(grdBuscar.TextMatrix(grdBuscar.RowSel, 1))
        sql = sql & " ORDER BY D.DEB_FECHA, S.SOC_CODIGO, D.DEB_ITEM"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            Do While Rec1.EOF = False
                GrdSocios.AddItem Rec1!DEB_MES & Chr(9) & Rec1!DEB_ANO & Chr(9) & _
                              Rec1!SOC_NOMBRE & Chr(9) & Rec1!DEB_DETALLE & Chr(9) & _
                              Valido_Importe(Rec1!DEB_IMPORTE) & Chr(9) & _
                              Rec1!SOC_CODIGO & Chr(9) & Rec1!DEB_ITEM
                Rec1.MoveNext
            Loop
        End If
        Rec1.Close
        GrdSocios.MergeCells = 1
        GrdSocios.MergeCol(0) = True  'MES
        GrdSocios.MergeCol(1) = True  'AÑO
        GrdSocios.MergeCol(2) = True  'SOCIO
        GrdSocios.MergeCol(3) = False 'DESCRIPCION DEL CONCEPTO
        GrdSocios.MergeCol(4) = False 'IMPORTE
        GrdSocios.MergeCol(5) = False 'COD SOCIO
        GrdSocios.MergeCol(6) = False 'ITEM
    
        If grdBuscar.TextMatrix(grdBuscar.RowSel, 5) = "S" Then
            chkComision.Value = Checked
        Else
            chkComision.Value = Unchecked
        End If
        fecha1.Value = grdBuscar.TextMatrix(grdBuscar.RowSel, 0)
        lblRecibo.Caption = "RECIBO   " & Format(grdBuscar.TextMatrix(grdBuscar.RowSel, 1), "00000000")
        txtNroRecibo.Text = grdBuscar.TextMatrix(grdBuscar.RowSel, 1)
        txtCodSoc.Text = grdBuscar.TextMatrix(grdBuscar.RowSel, 4)
        txtCodSoc_LostFocus
        txtTotal.Text = grdBuscar.TextMatrix(grdBuscar.RowSel, 3)
        txtRecargo.Text = grdBuscar.TextMatrix(grdBuscar.RowSel, 6)
        FrameRecibo.Enabled = False
        FrameSocio.Enabled = False
        GrdSocios.Enabled = False
        cmdAceptar.Enabled = False
        
        'BUSCO LAS FORMAS DE PAGO
        grdPagos.Rows = 1
        sql = "SELECT F.FPG_DESCRI, R.REC_PAGO, R.CHE_NUMERO"
        sql = sql & " FROM FORMA_PAGO F, RECIBO_PAGOS R"
        sql = sql & " WHERE F.FPG_CODIGO=R.FPG_CODIGO"
        sql = sql & " AND R.REC_NUMERO=" & XN(grdBuscar.TextMatrix(grdBuscar.RowSel, 1))
        sql = sql & " ORDER BY REC_ITEM"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            Do While Rec1.EOF = False
                grdPagos.AddItem Rec1!FPG_DESCRI & Chr(9) & Valido_Importe(Rec1!REC_PAGO) _
                & Chr(9) & "" & Chr(9) & "" & Chr(9) & ChkNull(Rec1!CHE_NUMERO)
                Rec1.MoveNext
            Loop
        End If
        Rec1.Close
        cboFormaPago.Enabled = False
        txtImportePago.Enabled = False
        cmdAgregarPago.Enabled = False
        cmdBorroFila.Enabled = False
        cmdAceptarPagos.Enabled = False
        'cmdImprimir.Enabled = True
        If grdBuscar.TextMatrix(grdBuscar.RowSel, 7) = 1 Then
            cmdAnular.Enabled = True
            lblEstado.Caption = ""
            cmdImprimir.Enabled = True
        Else
            cmdAnular.Enabled = False
            lblEstado.Caption = "RECIBO ANULADO"
            cmdImprimir.Enabled = False
        End If
        SSTab1.Tab = 0
    End If
End Sub

Private Sub grdBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grdBuscar_DblClick
    End If
End Sub

Private Sub GrdSocios_Click()
    If GrdSocios.Rows > 1 Then
        If GrdSocios.MouseCol = 0 Then
            Dim mValor As String
            mValor = GrdSocios.TextMatrix(GrdSocios.RowSel, 0)
            For I = 1 To GrdSocios.Rows - 1
                If mValor = GrdSocios.TextMatrix(I, 0) Then
                    If GrdSocios.TextMatrix(I, 7) = "" Then
                        CambiaColorAFilaDeGrilla GrdSocios, I, vbRed
                        GrdSocios.TextMatrix(I, 7) = "S"
                    Else
                        CambiaColorAFilaDeGrilla GrdSocios, I, vbBlack
                        GrdSocios.TextMatrix(I, 7) = ""
                    End If
                End If
            Next
        ElseIf GrdSocios.MouseCol = 3 Then
        
            If GrdSocios.TextMatrix(GrdSocios.RowSel, 7) = "" Then
                CambiaColorAFilaDeGrilla GrdSocios, GrdSocios.RowSel, vbRed
                GrdSocios.TextMatrix(GrdSocios.RowSel, 7) = "S"
            Else
                CambiaColorAFilaDeGrilla GrdSocios, GrdSocios.RowSel, vbBlack
                GrdSocios.TextMatrix(GrdSocios.RowSel, 7) = ""
            End If
        End If
    End If
    SumarTotal
End Sub

Private Sub SumarTotal()
    txtTotal.Text = ""
    txtRecargo.Text = ""
    For I = 1 To GrdSocios.Rows - 1
        If GrdSocios.TextMatrix(I, 7) = "S" Then
            txtTotal.Text = CDbl(Chk0(txtTotal.Text)) + CDbl(GrdSocios.TextMatrix(I, 4))
            txtRecargo.Text = CDbl(Chk0(txtRecargo.Text)) + CDbl(Chk0(GrdSocios.TextMatrix(I, 8)))
        End If
    Next
    txtRecargo.Text = Valido_Importe(txtRecargo.Text)
    txtTotal.Text = Valido_Importe(CDbl(Chk0(txtTotal.Text)) + CDbl(txtRecargo.Text))
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
        txtCodSocB.SetFocus
    End If
End Sub

Private Sub txtCodSoc_Change()
    If txtCodSoc.Text = "" Then
        txtNomSoc.Text = ""
        txtDomici.Text = ""
        txtNroDoc.Text = ""
        txtTelefono.Text = ""
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
        sql = "SELECT SOC_CODIGO, SOC_NOMBRE, SOC_DOMICI, SOC_NRODOC, SOC_TELEFONO"
        sql = sql & " FROM SOCIOS"
        sql = sql & " WHERE "
        sql = sql & " SOC_CODIGO =" & XN(txtCodSoc.Text)
        sql = sql & " AND TIS_CODIGO=1"
        If Rec.State = 1 Then Rec.Close
        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec.EOF = False Then
            txtNomSoc.Text = ChkNull(Rec!SOC_NOMBRE)
            txtDomici.Text = ChkNull(Rec!SOC_DOMICI)
            txtNroDoc.Text = ChkNull(Rec!SOC_NRODOC)
            txtTelefono.Text = ChkNull(Rec!SOC_TELEFONO)
            
            If mBuscar = False Then
                Call BuscarDeuda(txtCodSoc.Text)
            End If
        Else
            MsgBox "El Código no existe", vbInformation
            txtNomSoc.Text = ""
            txtCodSoc.Text = ""
            txtCodSoc.SetFocus
        End If
        If Rec.State = 1 Then Rec.Close
    End If
End Sub

Private Sub txtCodSocB_Change()
    If txtCodSocB.Text = "" Then
        txtNomSocB.Text = ""
    End If
End Sub

Private Sub txtCodSocB_GotFocus()
    SelecTexto txtCodSocB
End Sub

Private Sub txtCodSocB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        txtCodSocB.Text = ""
        BuscarSocios "txtCodSocB", "CODIGO"
    End If
End Sub

Private Sub txtCodSocB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodSocB_LostFocus()
    If txtCodSocB.Text <> "" Then
        sql = "SELECT SOC_CODIGO, SOC_NOMBRE"
        sql = sql & " FROM SOCIOS"
        sql = sql & " WHERE "
        sql = sql & " SOC_CODIGO =" & XN(txtCodSocB.Text)
        If Rec.State = 1 Then Rec.Close
        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec.EOF = False Then
            txtNomSocB.Text = ChkNull(Rec!SOC_NOMBRE)
        Else
            MsgBox "El Código no existe", vbInformation
            txtNomSocB.Text = ""
            txtCodSocB.Text = ""
            txtCodSocB.SetFocus
        End If
        If Rec.State = 1 Then Rec.Close
    End If
End Sub

Private Sub txtImportePago_GotFocus()
    txtImportePago.Text = txtTotalPagos.Text
    SelecTexto txtImportePago
End Sub

Private Sub txtImportePago_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImportePago, KeyAscii)
End Sub

Private Sub txtImportePago_LostFocus()
    If txtImportePago.Text <> "" Then
        txtImportePago.Text = Valido_Importe(txtImportePago.Text)
    End If
End Sub

Private Sub txtNomSoc_Change()
    If txtNomSoc.Text = "" Then
        txtCodSoc.Text = ""
        txtDomici.Text = ""
        txtNroDoc.Text = ""
        txtTelefono.Text = ""
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

Private Sub txtNomSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtNomSoc_LostFocus()
    If txtCodSoc.Text = "" And txtNomSoc.Text <> "" Then
        sql = "SELECT SOC_CODIGO,SOC_NOMBRE, SOC_DOMICI, SOC_NRODOC, SOC_TELEFONO"
        sql = sql & " FROM SOCIOS"
        sql = sql & " WHERE "
        sql = sql & " SOC_NOMBRE LIKE '" & XN(Trim(txtNomSoc.Text)) & "%'"
        sql = sql & " AND TIS_CODIGO=1"
        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec.EOF = False Then
            If Rec.RecordCount > 1 Then
                BuscarSocios "txtCodSoc", "CADENA", Trim(txtNomSoc.Text)
                If Rec.State = 1 Then Rec.Close
                txtNomSoc.SetFocus
            Else
                txtCodSoc.Text = Rec!SOC_CODIGO
                txtNomSoc.Text = Rec!SOC_NOMBRE
                txtDomici.Text = ChkNull(Rec!SOC_DOMICI)
                txtNroDoc.Text = ChkNull(Rec!SOC_NRODOC)
                txtTelefono.Text = ChkNull(Rec!SOC_TELEFONO)
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
            cSQL = cSQL & " AND TIS_CODIGO=1"
        Else
            cSQL = cSQL & " WHERE TIS_CODIGO=1"
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
            ElseIf Txt = "txtCodSocB" Then
                txtCodSocB.Text = .ResultFields(3)
                txtCodSocB_LostFocus
            End If
        End If
    End With
    
    Set B = Nothing
End Sub

Private Sub BuscarDeuda(Codigo As String)
    GrdSocios.Rows = 1
    'TITULARES
    sql = "SELECT S.SOC_NOMBRE, D.DEB_MES, D.DEB_ANO, S.SOC_CODIGO,"
    sql = sql & " D.DEB_ITEM, D.DEB_DETALLE, D.DEB_IMPORTE, D.DEB_FECHA"
    sql = sql & " FROM DEBITOS D, SOCIOS S"
    sql = sql & " WHERE "
    sql = sql & " S.SOC_CODIGO=D.SOC_CODIGO"
    sql = sql & " AND S.SOC_CODIGO=" & XN(Codigo)
    sql = sql & " AND D.DEB_SALDO > 0"
    sql = sql & " AND D.DEB_PAGADO IS NULL"
    
    sql = sql & " UNION ALL"
    
    'OPTATIVOS
    sql = sql & " SELECT S1.SOC_NOMBRE, D1.DEB_MES, D1.DEB_ANO, S1.SOC_CODIGO,"
    sql = sql & " D1.DEB_ITEM, D1.DEB_DETALLE, D1.DEB_IMPORTE, D1.DEB_FECHA"
    sql = sql & " FROM DEBITOS D1, SOCIOS S1"
    sql = sql & " WHERE"
    sql = sql & " S1.SOC_CODIGO=D1.SOC_CODIGO"
    sql = sql & " AND S1.SOC_TITULAR=" & XN(Codigo)
    sql = sql & " AND D1.DEB_SALDO > 0"
    sql = sql & " AND D1.DEB_PAGADO IS NULL"
    sql = sql & " ORDER BY D.DEB_FECHA, S.SOC_CODIGO, D.DEB_ITEM"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            GrdSocios.AddItem Rec1!DEB_MES & Chr(9) & Rec1!DEB_ANO & Chr(9) & _
                              Rec1!SOC_NOMBRE & Chr(9) & Rec1!DEB_DETALLE & Chr(9) & _
                              Valido_Importe(Rec1!DEB_IMPORTE) & Chr(9) & _
                              Rec1!SOC_CODIGO & Chr(9) & Rec1!DEB_ITEM & Chr(9) & "" & Chr(9) & _
                              BuscoRecargo(Rec1!DEB_MES, Rec1!DEB_ANO, Rec1!SOC_CODIGO, Rec1!DEB_ITEM)
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
     
    GrdSocios.MergeCells = 1
    GrdSocios.MergeCol(0) = True  'MES
    GrdSocios.MergeCol(1) = True  'AÑO
    GrdSocios.MergeCol(2) = True  'SOCIO
    GrdSocios.MergeCol(3) = False 'DESCRIPCION DEL CONCEPTO
    GrdSocios.MergeCol(4) = False 'IMPORTE
    GrdSocios.MergeCol(5) = False 'COD SOCIO
    GrdSocios.MergeCol(6) = False 'ITEM
End Sub

Private Function BuscoRecargo(mes As String, ano As String, soc As String, Item As String) As Double
    If mDiasRecargo > 0 Then
        Dim mFechaRec As Date
        Dim mFechita As String
        mFechita = "01/" & Month(fecha1.Value) & "/" & Year(fecha1.Value)
        mFechaRec = CDate(mFechita) + mDiasRecargo
        
        Select Case Weekday(mFechaRec)
            Case 1 'domingo
                mFechaRec = mFechaRec + 1
            Case 7 'sabado
                mFechaRec = mFechaRec + 2
        End Select
        
        sql = "SELECT E.DEP_RECARGO"
        sql = sql & " FROM DEBITOS D, DEPORTE E"
        sql = sql & " WHERE D.DEP_CODIGO=E.DEP_CODIGO"
        sql = sql & " AND E.DEP_RECARGO > 0"
        sql = sql & " AND D.DEB_MES=" & XN(mes)
        sql = sql & " AND D.DEB_ANO=" & XN(ano)
        sql = sql & " AND D.SOC_CODIGO=" & XN(soc)
        sql = sql & " AND D.DEB_ITEM=" & XN(Item)
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec2.EOF = False Then
            If CDbl(Chk0(Rec2!DEP_RECARGO)) > 0 Then
                If CInt(ano) = Year(fecha1.Value) Then 'And Month(fecha1.value) <= CInt(MES) And CDate(fecha1.value) > mFechaRec Then
                    If CInt(mes) < Month(fecha1.Value) Then
                        BuscoRecargo = Valido_Importe(Chk0(Rec2!DEP_RECARGO))
                        
                    ElseIf (CInt(mes) = Month(fecha1.Value)) And (CDate(fecha1.Value) >= mFechaRec) Then
                        BuscoRecargo = Valido_Importe(Chk0(Rec2!DEP_RECARGO))
                        
                    Else
                        BuscoRecargo = 0
                        
                    End If
                    
                ElseIf CInt(ano) < Year(fecha1.Value) Then
                    
                    BuscoRecargo = Valido_Importe(Chk0(Rec2!DEP_RECARGO))
                Else
                    
                    BuscoRecargo = 0
                End If
            Else
                BuscoRecargo = 0
            End If
        Else
            BuscoRecargo = 0
        End If
        Rec2.Close
    Else
        BuscoRecargo = 0
    End If
End Function

Private Function Validar() As Boolean
    Validar = False
    
    If txtCodSoc.Text = "" Then
        MsgBox "Falta Ingresar el Socios", vbCritical, TIT_MSGBOX
        txtCodSoc.SetFocus
        Exit Function
    End If
    If GrdSocios.Rows = 1 Then
        MsgBox "No Hay conceptos para Abonar", vbCritical, TIT_MSGBOX
        txtCodSoc.SetFocus
        Exit Function
    End If
    Dim mBandera As Boolean
    mBandera = False
    For I = 1 To GrdSocios.Rows - 1
        If GrdSocios.TextMatrix(I, 7) = "S" Then
            mBandera = True
            Exit For
        End If
    Next
    If mBandera = False Then
        MsgBox "No Hay conceptos para Abonar", vbCritical, TIT_MSGBOX
        txtCodSoc.SetFocus
        Exit Function
    End If
    If grdPagos.Rows = 1 Then
        MsgBox "No a ingresado la Forma de Pago", vbCritical, TIT_MSGBOX
        cmdFormaPago.SetFocus
        Exit Function
    End If
    
    Validar = True
End Function

Private Sub txtNomSocB_Change()
    If txtNomSocB.Text = "" Then
        txtCodSocB.Text = ""
    End If
End Sub

Private Sub txtNomSocB_GotFocus()
    SelecTexto txtNomSocB
End Sub

Private Sub txtNomSocB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarSocios "txtcodCliB", "CODIGO"
    End If
End Sub

Private Sub txtNomSocB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtNomSocB_LostFocus()
    If txtCodSocB.Text = "" And txtNomSocB.Text <> "" Then
        sql = "SELECT SOC_CODIGO,SOC_NOMBRE"
        sql = sql & " FROM SOCIOS"
        sql = sql & " WHERE "
        sql = sql & " SOC_NOMBRE LIKE '" & XN(Trim(txtNomSocB.Text)) & "%'"
        sql = sql & " AND TIS_CODIGO=1"
        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec.EOF = False Then
            If Rec.RecordCount > 1 Then
                BuscarSocios "txtCodSocB", "CADENA", Trim(txtNomSocB.Text)
                If Rec.State = 1 Then Rec.Close
                txtNomSocB.SetFocus
            Else
                txtCodSocB.Text = Rec!SOC_CODIGO
                txtNomSocB.Text = Rec!SOC_NOMBRE
            End If
        Else
            MsgBox "El Socio no existe", vbExclamation, TIT_MSGBOX
            txtCodSocB.Text = ""
            txtNomSocB.SetFocus
        End If
        If Rec.State = 1 Then Rec.Close
    End If
End Sub
