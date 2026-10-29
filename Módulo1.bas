Attribute VB_Name = "Módulo1"
Sub GenerarPDFsMasivos()
    Dim wsDash As Worksheet
    Dim wsConfig As Worksheet
    Dim rngClientes As Range
    Dim celdaCliente As Range
    Dim ultimaFila As Long
    Dim rutaPrincipal As String
    Dim rutaCarpetaMes As String
    Dim nombreArchivo As String
    Dim mesReporte As String
    
    ' 1. Definir hojas
    Set wsDash = ThisWorkbook.Sheets("DASHBOARD")
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    
    ' 2. Leer el mes que está actualmente en L3
    mesReporte = wsDash.Range("L3").Value
    
    ' 3. RUTINA DE CREACIÓN DE CARPETA
    ' Definir la ruta donde está el Excel y agregarle el nombre del mes
    rutaPrincipal = ThisWorkbook.Path & "\"
    rutaCarpetaMes = rutaPrincipal & mesReporte & "\"
    
    ' Evaluar si la carpeta del mes NO existe. Si no existe, ordenarle a Windows que la cree.
    If Dir(rutaCarpetaMes, vbDirectory) = "" Then
        MkDir rutaCarpetaMes
    End If
    
    ' 4. Encontrar dinámicamente la última fila con datos en la columna A de CONFIG
    ultimaFila = wsConfig.Cells(wsConfig.Rows.Count, "A").End(xlUp).Row
    Set rngClientes = wsConfig.Range("A2:A" & ultimaFila)
    
    ' 5. Ciclo de generación
    For Each celdaCliente In rngClientes
        If celdaCliente.Value <> "" Then
            
            wsDash.Range("L2").Value = celdaCliente.Value
            Calculate
            DoEvents
            
            ' Armar el nombre del archivo apuntando a la NUEVA ruta (rutaCarpetaMes)
            nombreArchivo = rutaCarpetaMes & "Reporte_" & celdaCliente.Value & "_" & mesReporte & ".pdf"
            
            ' Exportar PDF
            wsDash.ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=nombreArchivo, _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=False
                
        End If
    Next celdaCliente
    
    MsgBox "¡Proceso finalizado! Los reportes se guardaron ordenados dentro de la carpeta: " & mesReporte, vbInformation, "STYBA Logística"
End Sub

