Attribute VB_Name = "modDocumentos"
Option Explicit

' =============================================================================
' MODULO: modDocumentos
' Centraliza la lista de documentos obligatorios (hoja CONFIG, tabla
' "TablaDocumentos") y la creacion de filas en DATA.
' Reemplaza los dos Array(...) que estaban duplicados en Hoja1 (DATA) y Hoja4
' (CLIENTES).
'
' Requisitos en el libro:
'   - Hoja "CONFIG" con la tabla "TablaDocumentos" (columnas: Documento | Activo)
'   - Hoja "DATA" con la tabla "Tabla1"
'   - Hoja "CLIENTES" con la tabla "TablaClientes"
' =============================================================================

' Devuelve los documentos vigentes de la lista maestra.
' Un documento cuenta como activo salvo que la columna Activo diga "No".
Public Function ObtenerDocumentos() As Collection
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim docs As New Collection
    Dim nombre As String

    On Error Resume Next
    Set tbl = ThisWorkbook.Sheets("CONFIG").ListObjects("TablaDocumentos")
    On Error GoTo 0

    If tbl Is Nothing Then
        Set ObtenerDocumentos = docs
        Exit Function
    End If
    If tbl.DataBodyRange Is Nothing Then
        Set ObtenerDocumentos = docs
        Exit Function
    End If

    For Each fila In tbl.ListRows
        nombre = Trim(CStr(fila.Range(1, 1).Value))
        If nombre <> "" And UCase(Trim(CStr(fila.Range(1, 2).Value))) <> "NO" Then
            docs.Add nombre
        End If
    Next fila

    Set ObtenerDocumentos = docs
End Function

' Indica si una clave ya existe en una Collection (sin distinguir mayusculas).
Private Function YaTiene(ByVal col As Collection, ByVal clave As String) As Boolean
    Dim tmp As Variant
    On Error Resume Next
    tmp = col(UCase(clave))
    YaTiene = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' -----------------------------------------------------------------------------
' Agrega a DATA SOLO los documentos que le faltan al cliente.
'   - Cliente nuevo            -> se crean todos los documentos.
'   - Cliente existente        -> se agregan unicamente los que no tiene
'                                 (esto cubre el anti-duplicados).
' Parametro opcional celdaID: la celda de ID que el usuario escribio a mano en
' DATA; el primer documento faltante se escribe en esa misma fila.
' Devuelve cuantas filas se crearon.
' -----------------------------------------------------------------------------
Public Function SincronizarCliente(ByVal clientID As String, _
                                   ByVal clientName As String, _
                                   Optional ByVal celdaID As Range = Nothing) As Long
    Dim tblData As ListObject
    Dim docs As Collection
    Dim existentes As New Collection
    Dim fila As ListRow
    Dim d As Variant
    Dim newRow As ListRow
    Dim agregados As Long
    Dim usoCelda As Boolean
    Dim evPrevio As Boolean
    Dim pantallaPrevia As Boolean
    Dim clave As String

    Set tblData = ThisWorkbook.Sheets("DATA").ListObjects("Tabla1")
    Set docs = ObtenerDocumentos()
    If docs.Count = 0 Then Exit Function

    evPrevio = Application.EnableEvents
    pantallaPrevia = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Salida

    ' 1) Que documentos ya tiene este cliente
    If Not tblData.DataBodyRange Is Nothing Then
        For Each fila In tblData.ListRows
            If Trim(CStr(fila.Range(1, 1).Value)) = clientID Then
                clave = UCase(Trim(CStr(fila.Range(1, 3).Value)))
                If clave <> "" Then
                    On Error Resume Next
                    existentes.Add True, clave
                    On Error GoTo Salida
                End If
            End If
        Next fila
    End If

    ' 2) Agregar los que faltan
    For Each d In docs
        If Not YaTiene(existentes, CStr(d)) Then
            If Not celdaID Is Nothing And Not usoCelda Then
                celdaID.Offset(0, 2).Value = d
                celdaID.Offset(0, 3).Value = "Pendiente"
                If clientName <> "" Then celdaID.Offset(0, 1).Value = clientName
                usoCelda = True
            Else
                Set newRow = tblData.ListRows.Add
                newRow.Range(1, 1).Value = clientID
                newRow.Range(1, 2).Value = clientName
                newRow.Range(1, 3).Value = d
                newRow.Range(1, 4).Value = "Pendiente"
            End If
            agregados = agregados + 1
        End If
    Next d

    ' 3) Si el usuario escribio el ID a mano pero el cliente ya estaba completo,
    '    la fila quedaria con solo el ID: se elimina.
    If Not celdaID Is Nothing And Not usoCelda Then
        tblData.ListRows(celdaID.Row - tblData.HeaderRowRange.Row).Delete
    End If

Salida:
    Application.EnableEvents = evPrevio
    Application.ScreenUpdating = pantallaPrevia
    SincronizarCliente = agregados
End Function

' -----------------------------------------------------------------------------
' Recorre todos los clientes ACTIVOS y les agrega los documentos nuevos de la
' lista maestra. Se usa cuando se agrega un tipo de documento y ya existen
' expedientes abiertos.
' -----------------------------------------------------------------------------
Public Sub SincronizarTodos()
    Dim tblC As ListObject
    Dim fila As ListRow
    Dim id As String, nombre As String
    Dim n As Long, total As Long, clientes As Long

    On Error Resume Next
    Set tblC = ThisWorkbook.Sheets("CLIENTES").ListObjects("TablaClientes")
    On Error GoTo 0
    If tblC Is Nothing Then Exit Sub
    If tblC.DataBodyRange Is Nothing Then Exit Sub

    For Each fila In tblC.ListRows
        id = Trim(CStr(fila.Range(1, 1).Value))
        nombre = Trim(CStr(fila.Range(1, 2).Value))
        If id <> "" And UCase(Trim(CStr(fila.Range(1, 3).Value))) = "ACTIVO" Then
            n = SincronizarCliente(id, nombre)
            If n > 0 Then
                clientes = clientes + 1
                total = total + n
            End If
        End If
    Next fila

    ' La tabla dinamica (vigencia) no se actualiza sola; se refresca aqui.
    On Error Resume Next
    ThisWorkbook.Sheets("Tabla dinamica").PivotTables(1).RefreshTable
    On Error GoTo 0

    If total = 0 Then
        MsgBox "Todos los clientes activos ya tienen la lista completa de documentos.", _
               vbInformation, "Styba Automatizacion"
    Else
        MsgBox "Se agregaron " & total & " documento(s) a " & clientes & _
               " cliente(s) activo(s).", vbInformation, "Styba Automatizacion"
    End If
End Sub
