Attribute VB_Name = "modDllListaBaseABM"
Option Explicit

'definición de forms y opbjetos a utilizar en ABMs

'forms
Public vFormClPr As frmCListaBaseABM
Public vFormLocalidad As frmCListaBaseABM
Public vFormTipoIngreso As frmCListaBaseABM
Public vFormTipoEgreso As frmCListaBaseABM
Public vFormSexo As frmCListaBaseABM
Public vFormParentesco As frmCListaBaseABM
Public vFormEstadoCivil As frmCListaBaseABM
Public vFormEstadoJuicio As frmCListaBaseABM
Public vFormTipoCuota As frmCListaBaseABM
Public vFormDeporte As frmCListaBaseABM
Public vFormSocios As frmCListaBaseABM

Public vFormEstadoDocumento As frmCListaBaseABM
Public vFormFormaPago As frmCListaBaseABM

Public vFormEmpleados As frmCListaBaseABM

Public vFormBancos As frmCListaBaseABM
Public vFormTMONEDA As frmCListaBaseABM
Public vFormEstadosCheques As frmCListaBaseABM
Public vFormTCuenta As frmCListaBaseABM
Public vFormCuentaBancaria As frmCListaBaseABM
Public vFormGastoBancario As frmCListaBaseABM
Public vFormDebCreBancario As frmCListaBaseABM

'objetos
Public vABMClPr As CListaBaseABM
Public vABMLocalidad As CListaBaseABM
Public vABMTipoIngreso As CListaBaseABM
Public vABMTipoEgreso As CListaBaseABM
Public vABMSexo As CListaBaseABM
Public vABMParentesco As CListaBaseABM
Public vABMEstadoCivil As CListaBaseABM
Public vABMEmpleados As CListaBaseABM
Public vABMTipoCuota As CListaBaseABM
Public vABMDeporte As CListaBaseABM

Public vABMEstadoDocumento As CListaBaseABM
Public vABMFormaPago As CListaBaseABM
Public vABMSocios As CListaBaseABM
Public vABMBancos As CListaBaseABM
Public vABMTMONEDA As CListaBaseABM
Public vABMEstadosCheques As CListaBaseABM
Public vABMTCuenta As CListaBaseABM
Public vABMCuentaBancaria As CListaBaseABM
Public vABMGastoBancario As CListaBaseABM
Public vABMDebCreBancario As CListaBaseABM

'variable para mantener el objeto base de ABM activo
Public auxDllActiva As CListaBaseABM
'Public auxDllActivaCta As CListaBaseABMCta

