VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ABMIngresos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga Ingresos de Caja"
   ClientHeight    =   5895
   ClientLeft      =   1620
   ClientTop       =   1950
   ClientWidth     =   9720
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
   ScaleHeight     =   5895
   ScaleWidth      =   9720
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   435
      Left            =   5955
      Picture         =   "ABMIngresos.frx":0000
      TabIndex        =   38
      Top             =   5415
      Width           =   915
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   6900
      TabIndex        =   8
      Top             =   5415
      Width           =   915
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   5010
      TabIndex        =   7
      Top             =   5415
      Width           =   915
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   8745
      TabIndex        =   10
      Top             =   5415
      Width           =   915
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      Height          =   435
      Left            =   7830
      TabIndex        =   9
      Top             =   5415
      Width           =   915
   End
   Begin TabDlg.SSTab TabTB 
      Height          =   5355
      Left            =   45
      TabIndex        =   11
      Top             =   30
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   9446
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
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
      TabPicture(0)   =   "ABMIngresos.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraPagos"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "B&uscar"
      TabPicture(1)   =   "ABMIngresos.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fraPagos 
         Height          =   2520
         Left            =   75
         TabIndex        =   28
         Top             =   2175
         Width           =   9495
         Begin VB.TextBox txtImportePago 
            Height          =   315
            Left            =   1395
            TabIndex        =   5
            Top             =   1695
            Width           =   1245
         End
         Begin VB.ComboBox cboFormaPago 
            Height          =   315
            ItemData        =   "ABMIngresos.frx":0D02
            Left            =   1395
            List            =   "ABMIngresos.frx":0D04
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1350
            Width           =   3330
         End
         Begin VB.CommandButton cmdBorroFila 
            Caption         =   "Borrar Fila"
            Height          =   315
            Left            =   3480
            TabIndex        =   31
            Top             =   2025
            Width           =   1245
         End
         Begin VB.Frame Frame6 
            Height          =   795
            Left            =   45
            TabIndex        =   29
            Top             =   525
            Width           =   4695
            Begin VB.TextBox txtTotalPagos 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3120
               TabIndex        =   3
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
         Begin VB.CommandButton cmdAgregarPago 
            Caption         =   "Agregar"
            Height          =   315
            Left            =   3480
            TabIndex        =   6
            Top             =   1695
            Width           =   1245
         End
         Begin MSFlexGridLib.MSFlexGrid grdPagos 
            Height          =   1815
            Left            =   4740
            TabIndex        =   32
            Top             =   585
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   3201
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
            FormatString    =   $"ABMIngresos.frx":0D06
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Importe:"
            Height          =   315
            Left            =   45
            TabIndex        =   35
            Top             =   1695
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
            Width           =   9330
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Forma Pago"
            Height          =   330
            Left            =   45
            TabIndex        =   33
            Top             =   1350
            Width           =   1320
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Total General"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   6420
         TabIndex        =   24
         Top             =   4680
         Width           =   3150
         Begin VB.TextBox txtTotalGeneral 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF0000&
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
            Height          =   360
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   180
            Width           =   1305
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Total Cheques"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   3240
         TabIndex        =   22
         Top             =   4680
         Width           =   3150
         Begin VB.TextBox txtTotalCheques 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF0000&
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
            Height          =   360
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   180
            Width           =   1305
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Total Efectivo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   75
         TabIndex        =   20
         Top             =   4680
         Width           =   3150
         Begin VB.TextBox txtTotalMoneda 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF0000&
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
            Height          =   360
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   180
            Width           =   1305
         End
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
         Left            =   -74835
         TabIndex        =   14
         Top             =   525
         Width           =   9315
         Begin VB.CommandButton CmdBuscAprox 
            Caption         =   "Buscar"
            Height          =   405
            Left            =   6615
            MaskColor       =   &H000000FF&
            TabIndex        =   26
            ToolTipText     =   "Buscar"
            Top             =   345
            UseMaskColor    =   -1  'True
            Width           =   2205
         End
         Begin MSComCtl2.DTPicker mFechaD 
            Height          =   315
            Left            =   1800
            TabIndex        =   39
            Top             =   405
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   20971521
            CurrentDate     =   42841
         End
         Begin MSComCtl2.DTPicker mFechaH 
            Height          =   315
            Left            =   4800
            TabIndex        =   40
            Top             =   405
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   20971521
            CurrentDate     =   42841
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   3780
            TabIndex        =   18
            Top             =   465
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecah Desde:"
            Height          =   195
            Left            =   705
            TabIndex        =   17
            Top             =   465
            Width           =   990
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Datos del Ingreso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         Left            =   75
         TabIndex        =   12
         Top             =   345
         Width           =   9495
         Begin VB.ComboBox CboTipoIngreso 
            Height          =   315
            Left            =   1215
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   720
            Width           =   5160
         End
         Begin VB.TextBox TxtCodigoE 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1215
            TabIndex        =   0
            Top             =   375
            Width           =   1305
         End
         Begin VB.TextBox TxtDescrip 
            Height          =   570
            Left            =   1215
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   1065
            Width           =   8040
         End
         Begin MSComCtl2.DTPicker txtcing_fecha 
            Height          =   315
            Left            =   4920
            TabIndex        =   41
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   20971521
            CurrentDate     =   42841
         End
         Begin VB.Label lblEstadoEgreso 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   6705
            TabIndex        =   37
            Top             =   615
            Width           =   825
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Ingreso:"
            Height          =   195
            Left            =   210
            TabIndex        =   36
            Top             =   765
            Width           =   960
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Index           =   0
            Left            =   4290
            TabIndex        =   16
            Top             =   435
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Ingreso:"
            Height          =   195
            Index           =   3
            Left            =   210
            TabIndex        =   15
            Top             =   435
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Index           =   2
            Left            =   210
            TabIndex        =   13
            Top             =   1095
            Width           =   870
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   3705
         Left            =   -74850
         TabIndex        =   27
         Top             =   1530
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   6535
         _Version        =   393216
         Cols            =   4
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
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   2070
      Top             =   5445
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      TabIndex        =   19
      Top             =   5505
      Width           =   660
   End
End
Attribute VB_Name = "ABMIngresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim J As Integer
Dim mTotalPagos As Double
Dim mTotalPagosCH As Double
Dim mTotalPagosEFT As Double

Private Sub BuscoDatos()
    sql = "SELECT * FROM CAJA_INGRESO"
    sql = sql & " WHERE CIGR_NUMERO = " & XN(TxtCodigoE.Text)
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then ' si existe
        txtcing_fecha.Value = ChkNull(Rec1!CIGR_FECHA)
        BuscaCodigoProxItemData Rec1!TIG_CODIGO, CboTipoIngreso
        TxtDescrip.Text = ChkNull(Rec1!CIGR_DESCRI)
        TxtDescrip.SetFocus
        BuscoEstado Rec1!EST_CODIGO, lblEstadoEgreso
        If Rec1!EST_CODIGO = 2 Then
            cmdBorrar.Enabled = False
            CmdGrabar.Enabled = False
            cmdImprimir.Enabled = False
        Else
            cmdBorrar.Enabled = True
            CmdGrabar.Enabled = False
            cmdImprimir.Enabled = True
        End If
        Rec1.Close
        
        sql = "SELECT D.*, F.FPG_DESCRI FROM DETALLE_CAJA_INGRESO D, FORMA_PAGO F"
        sql = sql & " WHERE CIGR_NUMERO = " & XN(TxtCodigoE.Text)
        sql = sql & " AND F.FPG_CODIGO=D.FPG_CODIGO"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        Do While Rec1.EOF = False
            grdPagos.AddItem Rec1!FPG_DESCRI & Chr(9) & Valido_Importe(Rec1!CIGR_IMPORTE) & Chr(9) & "" & Chr(9) & "" & Chr(9) & ChkNull(Rec1!CHE_NUMERO)
            Rec1.MoveNext
        Loop
        Rec1.Close
        
        mTotalPagos = 0
        mTotalPagosCH = 0
        mTotalPagosEFT = 0
        For I = 1 To grdPagos.Rows - 1
            mTotalPagos = CDbl(mTotalPagos) + CDbl(grdPagos.TextMatrix(I, 1))
            Select Case grdPagos.TextMatrix(I, 0)
                Case "CHEQUE TERCERO"
                    mTotalPagosCH = CDbl(mTotalPagosCH) + CDbl(grdPagos.TextMatrix(I, 1))
                Case "EFECTIVO"
                    mTotalPagosEFT = CDbl(mTotalPagosEFT) + CDbl(grdPagos.TextMatrix(I, 1))
            End Select
        Next
        txtTotalMoneda.Text = Valido_Importe(CStr(mTotalPagosEFT))
        txtTotalCheques = Valido_Importe(CStr(mTotalPagosCH))
        txtTotalGeneral = Valido_Importe(CStr(mTotalPagos))
    Else
        MsgBox "Ingreso Inexistente", vbCritical
        TxtCodigoE.Text = ""
        TxtCodigoE.SetFocus
        Rec1.Close
        Exit Sub
    End If
    If Rec1.State = 1 Then Rec1.Close
End Sub

Private Sub cboFormaPago_LostFocus()
    If Me.ActiveControl.Name = "grdPagos" Then
        Exit Sub
    End If
End Sub

Private Sub cmdAgregarPago_Click()
    txtImportePago.Text = Valido_Importe(txtImportePago.Text)
    If cboFormaPago.Text = "" Then
        MsgBox "Debe Indicar la Forma de Pago", vbCritical, TIT_MSGBOX
        cboFormaPago.SetFocus
        Exit Sub
    End If
    mTotalPagos = 0
    mTotalPagosCH = 0
    mTotalPagosEFT = 0
    
    For I = 1 To grdPagos.Rows - 1
        mTotalPagos = mTotalPagos + CDbl(grdPagos.TextMatrix(I, 1))
    Next
    If mTotalPagos + CDbl(Chk0(txtImportePago.Text)) > CDbl(txtTotalPagos.Text) Then
        MsgBox "El Importe Ingresado Exede el Monto!", vbInformation, TIT_MSGBOX
        txtImportePago.SetFocus
        Exit Sub
    End If
    If CDbl(Chk0(txtImportePago.Text)) > 0 Then
       If Trim(cboFormaPago.Text) = "CHEQUE TERCERO" Then
           FrmCargaCheques.mMeLlamo = "CAJA INGRESO"
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
    mTotalPagosCH = 0
    mTotalPagosEFT = 0
    For I = 1 To grdPagos.Rows - 1
        mTotalPagos = CDbl(mTotalPagos) + CDbl(grdPagos.TextMatrix(I, 1))
        Select Case grdPagos.TextMatrix(I, 0)
            Case "CHEQUE TERCERO"
                mTotalPagosCH = CDbl(mTotalPagosCH) + CDbl(grdPagos.TextMatrix(I, 1))
            Case "EFECTIVO"
                mTotalPagosEFT = CDbl(mTotalPagosEFT) + CDbl(grdPagos.TextMatrix(I, 1))
        End Select
    Next
    txtTotalMoneda.Text = Valido_Importe(CStr(mTotalPagosEFT))
    txtTotalCheques = Valido_Importe(CStr(mTotalPagosCH))
    txtTotalGeneral = Valido_Importe(CStr(mTotalPagos))
    
    'txtTotalPagos.Text = Format(CDbl(txtTotalGeneral.Text) - mTotalPagos, "0.00")
    
    txtImportePago.Text = Format(CDbl(txtTotalPagos.Text) - mTotalPagos, "0.00")
'    If Val(txtTotalPagos.Text) = 0 Then
'        cmdAceptarPagos.SetFocus
'    Else
    cboFormaPago.ListIndex = 0
    cboFormaPago.SetFocus
'    End If
End Sub

Private Sub cmdBorrar_Click()
    On Error GoTo CLAVOSE
    If Trim(TxtCodigoE.Text) <> "" Then
        If MsgBox("Seguro desea eliminar el Ingreso '" & Trim(TxtDescrip.Text) & "' ?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            Screen.MousePointer = vbHourglass
            lblEstado.Caption = "Eliminando ..."
            DBConn.BeginTrans
            'DBConn.Execute "DELETE FROM DETALLE_CAJA_INGRESO WHERE CIGR_NUMERO = " & XN(TxtCodigoE.Text)
            'DBConn.Execute "DELETE FROM CAJA_INGRESO WHERE CIGR_NUMERO = " & XN(TxtCodigoE.Text)
            sql = "UPDATE CAJA_INGRESO"
            sql = sql & " SET EST_CODIGO=2"
            sql = sql & " WHERE CIGR_NUMERO = " & XN(TxtCodigoE.Text)
            DBConn.Execute sql
            
            DBConn.CommitTrans
            If TxtDescrip.Enabled Then TxtDescrip.SetFocus
            
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            CmdNuevo_Click
        End If
    End If
    Exit Sub
    
CLAVOSE:
    If Rec.State = 1 Then Rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
End Sub

Private Sub cmdBorroFila_Click()
    If grdPagos.Rows <= 2 Then
        grdPagos.Rows = 1
    Else
        grdPagos.RemoveItem (grdPagos.row)
    End If
    
    mTotalPagos = 0
    mTotalPagosCH = 0
    mTotalPagosEFT = 0
    For I = 1 To grdPagos.Rows - 1
        mTotalPagos = CDbl(mTotalPagos) + CDbl(grdPagos.TextMatrix(I, 1))
        Select Case grdPagos.TextMatrix(I, 0)
            Case "CHEQUE TERCERO"
                mTotalPagosCH = CDbl(mTotalPagosCH) + CDbl(grdPagos.TextMatrix(I, 1))
            Case "EFECTIVO"
                mTotalPagosEFT = CDbl(mTotalPagosEFT) + CDbl(grdPagos.TextMatrix(I, 1))
        End Select
    Next
    txtTotalMoneda.Text = Valido_Importe(CStr(mTotalPagosEFT))
    txtTotalCheques = Valido_Importe(CStr(mTotalPagosCH))
    txtTotalGeneral = Valido_Importe(CStr(mTotalPagos))
    
    For I = 1 To grdPagos.Rows - 1
        mTotalPagos = CDbl(mTotalPagos) + CDbl(grdPagos.TextMatrix(I, 1))
        Select Case grdPagos.TextMatrix(I, 0)
            Case "CHEQUE TERCERO"
                mTotalPagosCH = CDbl(mTotalPagosCH) + CDbl(grdPagos.TextMatrix(I, 1))
            Case "EFECTIVO"
                mTotalPagosEFT = CDbl(mTotalPagosEFT) + CDbl(grdPagos.TextMatrix(I, 1))
        End Select
    Next
    txtTotalMoneda.Text = Valido_Importe(CStr(mTotalPagosEFT))
    txtTotalCheques = Valido_Importe(CStr(mTotalPagosCH))
    txtTotalGeneral = Valido_Importe(CStr(mTotalPagos))
    
    'txtTotalPagos.Text = Format(CDbl(txtTotalGeneral.Text) - mTotalPagos, "0.00")
    txtImportePago.Text = Format(CDbl(txtTotalPagos.Text) - mTotalPagos, "0.00")
    cboFormaPago.SetFocus
End Sub

Private Sub CmdBuscAprox_Click()
    Set Rec = New ADODB.Recordset
    GrdModulos.HighLight = flexHighlightNever
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    Me.Refresh
    
    sql = "SELECT I.*, T.TIG_DESCRI"
    sql = sql & " FROM CAJA_INGRESO I, TIPO_INGRESO T"
    sql = sql & " WHERE "
    sql = sql & " T.TIG_CODIGO=I.TIG_CODIGO"
    If mfechaD.Value <> "" Then
        sql = sql & " AND CIGR_FECHA >= " & XDQ(mfechaD.Value)
    End If
    If mfechaH.Value <> "" Then
        sql = sql & " AND CIGR_FECHA <= " & XDQ(mfechaH.Value)
    End If
    sql = sql & " ORDER BY CIGR_NUMERO"
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If Rec.EOF = False Then
        Rec.MoveFirst
        'Número|Descripción|^Fecha|>Importe|Tipo de Ingreso|CODIGO Tipo de Ingreso
        Do While Not Rec.EOF
            GrdModulos.AddItem Rec!CIGR_NUMERO & Chr(9) & Trim(Rec!TIG_DESCRI) & Chr(9) & _
                        Rec!CIGR_FECHA & Chr(9) & Valido_Importe(Rec!CIGR_IMPORTE)
            Rec.MoveNext
        Loop
        GrdModulos.HighLight = flexHighlightAlways
        If GrdModulos.Enabled Then GrdModulos.SetFocus
        lblEstado.Caption = ""
    Else
        lblEstado.Caption = ""
        MsgBox "No se encontraron items con este Criterio", vbExclamation, TIT_MSGBOX
        If mfechaD.Enabled Then mfechaD.SetFocus
    End If
    lblEstado.Caption = ""
    Rec.Close
    Screen.MousePointer = vbNormal
End Sub

Private Function ValidoIngreso() As Boolean
    If Trim(TxtDescrip.Text) = "" Then
        MsgBox "No ha ingresado la descripción", vbExclamation, TIT_MSGBOX
        TxtDescrip.SetFocus
        ValidoIngreso = False
        Exit Function
    End If
    If txtcing_fecha.Value = "" Then
        MsgBox "No ha ingresado la Fecha del Ingreso", vbExclamation, TIT_MSGBOX
        txtcing_fecha.SetFocus
        ValidoIngreso = False
        Exit Function
    End If
    If CboTipoIngreso.ListIndex = -1 Then
        MsgBox "No ha Ingresado Tipo Ingreso", vbExclamation, TIT_MSGBOX
        CboTipoIngreso.SetFocus
        ValidoIngreso = False
        Exit Function
    End If
    If grdPagos.Rows = 1 Then
        MsgBox "No ha Ingresado Forma de Pago", vbExclamation, TIT_MSGBOX
        cboFormaPago.SetFocus
        ValidoIngreso = False
        Exit Function
    End If
    If CDbl(txtTotalPagos.Text) <> CDbl(txtTotalGeneral.Text) Then
        MsgBox "El Total de Pagos no coincide con el Total General", vbExclamation, TIT_MSGBOX
        cboFormaPago.SetFocus
        ValidoIngreso = False
        Exit Function
    End If
    ValidoIngreso = True
End Function

Private Sub CmdGrabar_Click()
    On Error GoTo CLAVOSE
    
    If ValidoIngreso = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando ..."
    
    Set Rec = New ADODB.Recordset
    If TxtCodigoE.Text = "" Then
        TxtCodigoE.Text = "1"
        sql = "SELECT MAX(CIGR_NUMERO) as MAXIMO FROM CAJA_INGRESO"
        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not IsNull(Rec.Fields!Maximo) Then TxtCodigoE.Text = Val(Trim(Rec.Fields!Maximo)) + 1
        Rec.Close
    End If
    DBConn.BeginTrans
    
    sql = "SELECT * FROM CAJA_INGRESO WHERE CIGR_NUMERO = " & XN(TxtCodigoE.Text)
    Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If Rec.EOF = True Then
        sql = "INSERT INTO CAJA_INGRESO"
        sql = sql & " (CIGR_NUMERO, CIGR_DESCRI, CIGR_FECHA,"
        sql = sql & " CIGR_IMPORTE, TIG_CODIGO,EST_CODIGO)"
        sql = sql & " VALUES ("
        sql = sql & XN(TxtCodigoE.Text) & ","
        sql = sql & XS(TxtDescrip.Text) & ","
        sql = sql & XDQ(txtcing_fecha.Value) & ","
        sql = sql & XN(txtTotalGeneral.Text) & ","
        sql = sql & CboTipoIngreso.ItemData(CboTipoIngreso.ListIndex) & ",3)"
        DBConn.Execute sql
        
        J = 1
        For I = 1 To grdPagos.Rows - 1
            If grdPagos.TextMatrix(I, 2) = "1" Then 'EFECTIVO
                sql = "INSERT INTO DETALLE_CAJA_INGRESO"
                sql = sql & " (CIGR_NUMERO, CIGR_NROITEM, FPG_CODIGO, CIGR_IMPORTE)"
                sql = sql & " VALUES ("
                sql = sql & XN(TxtCodigoE.Text) & ","
                sql = sql & J & ","
                sql = sql & XN(grdPagos.TextMatrix(I, 2)) & "," 'FORMA PAGO
                sql = sql & XN(grdPagos.TextMatrix(I, 1)) & ")"
                DBConn.Execute sql

                
            ElseIf grdPagos.TextMatrix(I, 2) = "2" Then 'CHEQUE TERCERO
                sql = "INSERT INTO DETALLE_CAJA_INGRESO"
                sql = sql & " (CIGR_NUMERO, CIGR_NROITEM, FPG_CODIGO, CIGR_IMPORTE, "
                sql = sql & " BAN_CODINT,CHE_NUMERO, CHE_IMPORTE)"
                sql = sql & " VALUES ("
                sql = sql & XN(TxtCodigoE.Text) & ","
                sql = sql & J & ","
                sql = sql & XN(grdPagos.TextMatrix(I, 2)) & "," 'FORMA PAGO
                sql = sql & XN(grdPagos.TextMatrix(I, 1)) & ","
                sql = sql & XN(grdPagos.TextMatrix(I, 3)) & ","
                sql = sql & XS(grdPagos.TextMatrix(I, 4)) & ","
                sql = sql & XN(grdPagos.TextMatrix(I, 1)) & ")"
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
            J = J + 1
        Next
    Else
        MsgBox "El Ingreso ya fue registrado", vbExclamation, TIT_MSGBOX
    End If
    Rec.Close
    
    DBConn.CommitTrans
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    cmdImprimir_Click
    CmdNuevo_Click
    Exit Sub
    
CLAVOSE:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
End Sub

Private Sub cmdImprimir_Click()
    If MsgBox("¿Imprime Ingreso?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    Rep.SubreportToChange = ""
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    
    Rep.SelectionFormula = "{CAJA_INGRESO.CIGR_NUMERO}=" & XN(CLng(TxtCodigoE.Text))
    
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.Destination = crptToWindow
    Rep.WindowTitle = "Caja Ingreso"
    Rep.ReportFileName = DRIVE & DirReport & "CajaIngreso.rpt"
    
    'PARA EL SUBREPORTE
    Rep.SubreportToChange = "SubReporte_CajaIngreso.rpt"
    Rep.SelectionFormula = ""
    Rep.SelectionFormula = "{DETALLE_CAJA_INGRESO.CIGR_NUMERO}=" & XN(CLng(TxtCodigoE.Text))
    
    Rep.Action = 1
    
    Rep.SelectionFormula = ""
    Rep.SubreportToChange = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
End Sub

Private Sub CmdNuevo_Click()
    TabTB.Tab = 0
    LIMPIAR
    cmdImprimir.Enabled = False
    GrdModulos.HighLight = flexHighlightNever
    GrdModulos.Rows = 1
    If TxtCodigoE.Enabled And Me.Visible Then TxtCodigoE.SetFocus
End Sub

Private Sub LIMPIAR()
    txtTotalCheques.Text = "0,00"
    txtTotalMoneda.Text = "0,00"
    cboFormaPago.ListIndex = -1
    grdPagos.Rows = 1
    CboTipoIngreso.ListIndex = -1
    txtTotalPagos.Text = ""
    txtImportePago.Text = ""
    TxtCodigoE.Text = ""
    TxtDescrip.Text = ""
    lblEstado.Caption = ""
    txtcing_fecha.Value = ""
    txtTotalGeneral.Text = "0,00"
    cmdBorrar.Enabled = True
    CmdGrabar.Enabled = True
    txtcing_fecha.SetFocus
    BuscoEstado 1, lblEstadoEgreso
End Sub

Private Sub cmdSalir_Click()
    Unload Me
    Set ABMIngresos = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'si preciona f1 voy a la busqueda
    If KeyCode = vbKeyF1 Then TabTB.Tab = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'si presiono ESCAPE salgo del form
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
    'If KeyAscii = vbKeyReturn And _
        Me.ActiveControl.Name <> "TxtDescriB" And _
        Me.ActiveControl.Name <> "GrdContactos" Then  'avanza de campo
    If KeyAscii = vbKeyReturn Then
            MySendKeys Chr(9)
            KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Set Rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    cmdImprimir.Enabled = False
    
    fraPagos.Visible = True
    BuscoEstado 1, lblEstadoEgreso
    
    'CARGO COMBO CON FORMA DE PAGO
    sql = "SELECT FPG_CODIGO, FPG_DESCRI FROM FORMA_PAGO WHERE FPG_CODIGO <> 3"
    sql = sql & " ORDER BY FPG_CODIGO"
    CargarControlItemdata cboFormaPago, sql
    
    'CARGO COMBO CON TIPO INGRESO
    sql = "SELECT TIG_CODIGO, TIG_DESCRI FROM TIPO_INGRESO"
    sql = sql & " ORDER BY TIG_DESCRI"
    CargarControlItemdata CboTipoIngreso, sql
    
    
    lblEstado.Caption = ""
    CmdGrabar.Enabled = True
    CmdNuevo.Enabled = True
    cmdBorrar.Enabled = False
    
    GrdModulos.FormatString = "^Número|Tipo Ingreso|^Fecha|>Importe"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 5500
    GrdModulos.ColWidth(2) = 1300
    GrdModulos.ColWidth(3) = 1100
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
    
    TabTB.Tab = 0
    txtTotalCheques.Text = "0,00"
    txtTotalMoneda.Text = "0,00"
    txtTotalGeneral.Text = "0,00"
    
    ConfiguroGrillaPagos
    
    Screen.MousePointer = vbNormal
    Me.Left = 0
    Me.Top = 0
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
End Sub

Private Sub GrdModulos_dblClick()
    If GrdModulos.Rows > 1 Then
        'paso el item seleccionado al tab 'DATOS'
        LIMPIAR
        TxtCodigoE.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
        TxtCodigoE_LostFocus
        TabTB.Tab = 0
    End If
End Sub

Private Sub GrdModulos_GotFocus()
    GrdModulos.Col = 0
    GrdModulos.ColSel = 1
    GrdModulos.HighLight = flexHighlightAlways
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cmdBorrar_Click
    If KeyCode = vbKeyReturn Then GrdModulos_dblClick
End Sub

Private Sub GrdModulos_LostFocus()
    GrdModulos.HighLight = flexHighlightNever
End Sub

Private Sub mFechaD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        MySendKeys Chr(9)
    End If
End Sub

Private Sub mFechaH_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        MySendKeys Chr(9)
    End If
End Sub

Private Sub tabTB_Click(PreviousTab As Integer)
    'Si cambio de 'Pestaña' en el tab
    'pongo el foco en el primer campo de la misma
    If TabTB.Tab = 0 And Me.Visible Then TxtDescrip.SetFocus
    If TabTB.Tab = 1 Then
        If mfechaD.Enabled Then mfechaD.SetFocus
    End If
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
      KeyAscii = CarNumeroTE(KeyAscii)
End Sub

Private Sub txtcing_fecha_LostFocus()
    If txtcing_fecha.Value = "" Then txtcing_fecha.Value = Date
End Sub

Private Sub TxtCodigoE_GotFocus()
    SelecTexto TxtCodigoE
End Sub

Private Sub TxtCodigoE_LostFocus()
    If Trim(TxtCodigoE.Text) <> "" Then ' si no viene vacio
        BuscoDatos
    Else
        CmdGrabar.Enabled = True
        CmdNuevo.Enabled = True
        cmdBorrar.Enabled = False
    End If
End Sub

Private Sub TxtDescrip_GotFocus()
    SelecTexto TxtDescrip
End Sub

Private Sub TxtDescrip_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCodigoE_Change()
    If Trim(TxtCodigoE.Text) = "" And cmdBorrar.Enabled Then
        cmdBorrar.Enabled = False
    ElseIf Trim(TxtCodigoE.Text) <> "" Then
        cmdBorrar.Enabled = True
    End If
End Sub

Private Sub TxtDescrip_Change()
    If Trim(TxtDescrip.Text) = "" And CmdGrabar.Enabled Then
        CmdGrabar.Enabled = False
    Else
        CmdGrabar.Enabled = True
    End If
End Sub

Private Sub txtImportePago_GotFocus()
    txtImportePago.Text = Format(CDbl(txtTotalPagos.Text) - CDbl(txtTotalGeneral.Text), "0.00")
    SelecTexto txtImportePago
End Sub

Private Sub txtImportePago_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImportePago, KeyAscii)
End Sub

Private Sub txtImportePago_LostFocus()
    If txtImportePago.Text <> "" Then
        txtImportePago.Text = Valido_Importe(txtImportePago)
    End If
End Sub

Private Sub txtTotalPagos_GotFocus()
    SelecTexto txtTotalPagos
End Sub

Private Sub txtTotalPagos_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtTotalPagos, KeyAscii)
End Sub

Private Sub txtTotalPagos_LostFocus()
    If txtTotalPagos.Text <> "" Then
        txtTotalPagos.Text = Valido_Importe(txtTotalPagos)
    End If
End Sub
