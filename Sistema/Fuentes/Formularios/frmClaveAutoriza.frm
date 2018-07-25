VERSION 5.00
Begin VB.Form frmClaveAutoriza 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Validación"
   ClientHeight    =   3525
   ClientLeft      =   2985
   ClientTop       =   2850
   ClientWidth     =   5835
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
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   345
      TabIndex        =   2
      Top             =   990
      Width           =   5070
      Begin VB.TextBox TxtUsuario 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1980
         MaxLength       =   20
         TabIndex        =   0
         Top             =   660
         Width           =   2085
      End
      Begin VB.TextBox TxtClave 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1980
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1170
         Width           =   2085
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   135
         Picture         =   "frmClaveAutoriza.frx":0000
         Top             =   810
         Width           =   480
      End
      Begin VB.Label LblUsuario 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         Height          =   180
         Left            =   1005
         TabIndex        =   8
         Top             =   720
         Width           =   885
      End
      Begin VB.Label LblTitulo2 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese el nombre de Usuario y contraseña"
         Height          =   225
         Left            =   915
         TabIndex        =   7
         Top             =   240
         Width           =   3105
      End
      Begin VB.Label LblContrasena 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña:"
         Height          =   195
         Index           =   1
         Left            =   1005
         TabIndex        =   3
         Top             =   1200
         Width           =   900
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   105
      Picture         =   "frmClaveAutoriza.frx":030A
      Top             =   150
      Width           =   480
   End
   Begin VB.Label LblTitulo1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Debe ingresar su clave para poder continuar "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   810
      TabIndex        =   6
      Top             =   585
      Width           =   3870
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "<ENTER>  Aceptar    -   <ESC> Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1635
      TabIndex        =   5
      Top             =   3060
      Width           =   3570
   End
   Begin VB.Label LblClave 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "El matriculado debe mas de dos cuotas ! "
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
      Left            =   825
      TabIndex        =   4
      Top             =   225
      Width           =   3960
   End
End
Attribute VB_Name = "frmClaveAutoriza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r As ADODB.Recordset
Dim r1 As ADODB.Recordset
Dim clave_okk As Boolean
Dim sql As String
Dim CUANTAS_VECES As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
    
        If Me.ActiveControl.Name <> "TxtUsuario" Then

            sql = "SELECT * FROM USUARIO WHERE " & _
            "USU_NOMBRE = '" & Trim(TxtUsuario) & "' AND " & _
            "USU_CLAVE = '" & Trim(TxtClave) & "'"
            r.Open sql, DBConn, adOpenStatic, adLockOptimistic
            clave_okk = True
            If r.RecordCount = 0 Then
                Me.TxtUsuario.Text = ""
                Me.TxtClave.Text = ""
                MsgBox "Clave incorrecta !" & Chr(13) & _
                "El nombre de usuario o la clave no han sido cargadas correctamente.", vbCritical, "Mensaje:"
                TxtUsuario.SetFocus
                r.Close
                CUANTAS_VECES = CUANTAS_VECES + 1
                If CUANTAS_VECES = 3 Then
                    MsgBox "El sistema se cerrará.", vbExclamation, TIT_MSGBOX
                    End
                End If
                Exit Sub
            Else
                sql = "SELECT * FROM USUARIO WHERE " & _
                "USU_NOMBRE = '" & Trim(TxtUsuario) & "' AND " & _
                "USU_CLAVE = '" & Trim(TxtClave) & "' AND " & _
                "USU_NIVEL = 'ALTO'"
                r1.Open sql, DBConn, adOpenStatic, adLockOptimistic
                clave_okk = True
                If r1.RecordCount = 0 Then
                    Me.TxtUsuario.Text = ""
                    Me.TxtClave.Text = ""
                    MsgBox "Autorización Denegada !" & Chr(13) & _
                    "El usuario No Posee Autorización para Realizar esta Acción.", vbCritical, "Mensaje:"
                    TxtUsuario.SetFocus
                    r.Close
                    r1.Close
                    CUANTAS_VECES = CUANTAS_VECES + 1
                    If CUANTAS_VECES = 3 Then
                        MsgBox "El sistema se cerrará.", vbExclamation, TIT_MSGBOX
                        End
                    End If
                    Exit Sub
                End If
            End If
            r.Close
            r1.Close
            Menu.USUARIO_LOCAL = Trim(TxtUsuario)
            Unload Me
            Set frmClaveAutoriza = Nothing
        Else
            TxtClave.SetFocus
        End If
    End If
    
    If KeyAscii = vbKeyEscape Then
        Menu.USUARIO_LOCAL = ""
        Unload Me
        Set frmClaveAutoriza = Nothing
    End If
End Sub

Private Sub Form_Load()
    'CLAVE_OK = False
    clave_okk = False
    CUANTAS_VECES = 0
    Set r = New ADODB.Recordset
    Set r1 = New ADODB.Recordset
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Esto se agrego porque cuando salian de formulario haciendo click en al (X) que cierra el formulario
    'lo mismo emitia los ordenes para aquellos bebe que tienen mas de 6 meses
    If clave_okk = False Then
        If r.State = 1 Then
            r.Close
        End If
        Me.TxtUsuario.Text = ""
        Me.TxtClave.Text = ""
        Menu.USUARIO_LOCAL = Trim(TxtUsuario)
        Unload Me
        Set frmClaveAutoriza = Nothing
    End If
End Sub

Private Sub TxtClave_GotFocus()
    SelecTexto TxtClave
End Sub

Private Sub TxtUsuario_GotFocus()
    SelecTexto TxtUsuario
End Sub

Private Sub TxtUsuario_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub
