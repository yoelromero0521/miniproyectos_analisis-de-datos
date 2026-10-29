VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim tbl As ListObject
    Dim tblClientes As ListObject
    Dim clientID As String
    Dim clientName As String
    Dim isBelowTable As Boolean
    Dim formulaRange As Range
    Dim rng1 As Range, rng2 As Range, rng6 As Range, rng7 As Range, rng8 As Range
    
    On Error Resume Next
    Set tbl = Me.ListObjects("Tabla1")
    Set tblClientes = ThisWorkbook.Sheets("CLIENTES").ListObjects(1)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Sub
    If Target.Cells.Count > 1 Or IsEmpty(Target) Then Exit Sub
    
    ' =========================================================================
    ' EL GUARDIA VIRTUAL (Protege columnas 6, 7 y 8 sin colapsar por tablas vacías)
    ' =========================================================================
    On Error Resume Next
    Set rng6 = tbl.ListColumns(6).DataBodyRange
    Set rng7 = tbl.ListColumns(7).DataBodyRange
    Set rng8 = tbl.ListColumns(8).DataBodyRange
    On Error GoTo 0
    
    If Not rng6 Is Nothing And Not rng7 Is Nothing And Not rng8 Is Nothing Then
        Set formulaRange = Union(rng6, rng7, rng8)
        If Not Intersect(Target, formulaRange) Is Nothing Then
            Application.EnableEvents = False
            
            ' Evita el Error 1004 cuando el historial de deshacer de Excel está vacío
            On Error Resume Next
            Application.Undo
            On Error GoTo 0
            
            MsgBox "¡Acción denegada! Estas columnas contienen fórmulas automáticas y no deben modificarse.", vbCritical, "Protección Styba"
            Application.EnableEvents = True
            Exit Sub
        End If
    End If
    
    ' Determinar zonas de captura seguras
    If Target.Column = tbl.ListColumns(1).Range.Column And Target.Row = tbl.Range.Row + tbl.Range.Rows.Count Then isBelowTable = True
    
    On Error Resume Next
    Set rng1 = tbl.ListColumns(1).DataBodyRange
    Set rng2 = tbl.ListColumns(2).DataBodyRange
    On Error GoTo 0
    
    ' =========================================================================
    ' ACCIÓN 1: Al escribir el ID del Cliente Manualmente en DATA
    ' =========================================================================
    If isBelowTable Or (Not rng1 Is Nothing And Not Intersect(Target, rng1) Is Nothing) Then
        If Target.Offset(0, 2).Value <> "" Then Exit Sub
        
        clientID = Trim(Target.Value)
        If clientID = "" Then Exit Sub
        
        If isBelowTable Then
            Application.EnableEvents = False
            tbl.Resize tbl.Range.Resize(tbl.Range.Rows.Count + 1, tbl.Range.Columns.Count)
            Application.EnableEvents = True
        End If
        
        clientName = ""
        If Not tblClientes Is Nothing Then
            On Error Resume Next
            clientName = Application.WorksheetFunction.VLookup(clientID, tblClientes.DataBodyRange, 2, False)
            On Error GoTo 0
        End If
        
        ' Llamada al nuevo módulo dinámico en lugar del Array quemado
        SincronizarCliente clientID, clientName, Target
        Exit Sub
    End If

    ' =========================================================================
    ' ACCIÓN 2: Respaldo para el Nombre
    ' =========================================================================
    If Not rng2 Is Nothing And Not Intersect(Target, rng2) Is Nothing Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        
        clientName = Target.Value
        clientID = Target.Offset(0, -1).Value
        
        ' Bucle dinámico que reemplaza al antiguo "For i = 1 To 13"
        Dim fila As ListRow
        For Each fila In tbl.ListRows
            If Trim(CStr(fila.Range(1, 1).Value)) = Trim(CStr(clientID)) Then
                fila.Range(1, 2).Value = clientName
            End If
        Next fila
        
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Exit Sub
    End If
End Sub

' =========================================================================
' NUEVO EVENTO: Desplegar calendario
' =========================================================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim tbl As ListObject
    Dim rng5 As Range
    
    On Error Resume Next
    Set tbl = Me.ListObjects("Tabla1")
    On Error GoTo 0
    If tbl Is Nothing Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub
    
    On Error Resume Next
    Set rng5 = tbl.ListColumns(5).DataBodyRange
    On Error GoTo 0
    
    If Not rng5 Is Nothing Then
        If Not Intersect(Target, rng5) Is Nothing Then frmCalendario.Show
    End If
End Sub
