VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Calculate()
    ' Este evento se dispara cada vez que las fórmulas FILTRAR cambian o el archivo se abre
    
    On Error Resume Next
    Application.EnableEvents = False
    
    ' Cambia "15:50" por el rango de filas exacto donde viven tus tablas
    Me.Rows("15:50").AutoFit
    
    Application.EnableEvents = True
    On Error GoTo 0
End Sub
