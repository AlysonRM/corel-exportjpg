VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fExportPNG 
   Caption         =   "Formulário Exporta PNG"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4050
   OleObjectBlob   =   "fExportPNG.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fExportPNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btExport_Click()
exportPNG tLayer, tPlayer, tNumber, tSize
Unload Me
End Sub
