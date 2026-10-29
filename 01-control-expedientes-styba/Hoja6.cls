VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim tbl As ListObject
    Dim changedDocName As String
    
    On Error Resume Next
    Set tbl = Me.ListObjects("TablaDocumentos")
    On Error GoTo 0
    
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub
    
    ' ==============================================================
    ' UNIFICACIÓN: Solo reaccionar si el usuario modifica la Columna 2 (Activo)
    ' ==============================================================
    If Intersect(Target, tbl.ListColumns(2).DataBodyRange) Is Nothing Then Exit Sub
    If Trim(CStr(Target.Value)) = "" Then Exit Sub

    ' Si el usuario selecciona o escribe "Sí" (con o sin tilde)
    If UCase(Trim(CStr(Target.Value))) = "SÍ" Or UCase(Trim(CStr(Target.Value))) = "SI" Then
        
        ' Extraer el nombre del documento que está en la celda de la izquierda
        changedDocName = Trim(CStr(Target.Offset(0, -1).Value))
        
        ' Candado: No hacer nada si pusieron "Sí" pero la celda del documento está vacía
        If changedDocName = "" Then Exit Sub
        
        ' Pregunta unificada para documentos nuevos o reactivados
        If MsgBox("El documento """ & changedDocName & """ está configurado como Activo." & vbCrLf & vbCrLf & _
                  "¿Deseas sincronizar la base de datos para inyectarlo a los clientes activos que aún no lo tengan?", _
                  vbYesNo + vbQuestion, "Styba Automatización - Sincronización") = vbYes Then
            
            Application.ScreenUpdating = False
            SincronizarTodos
            OrdenarDatos ' Mantiene el orden automático que programamos
            Application.ScreenUpdating = True
            
            MsgBox "Sincronización completada. La tabla DATA ha sido actualizada y ordenada.", vbInformation, "Éxito"
        End If
    End If
End Sub

' ==============================================================
' SUBRUTINA: Ordenar alfabéticamente la hoja DATA por Nombre Cliente
' ==============================================================
Private Sub OrdenarDatos()
    Dim tblData As ListObject
    
    On Error Resume Next
    Set tblData = ThisWorkbook.Sheets("DATA").ListObjects("Tabla1")
    On Error GoTo 0
    
    If Not tblData Is Nothing Then
        With tblData.Sort
            .SortFields.Clear
            ' Ordena basándose en la columna 2 (Nombre Cliente) de la Tabla1
            .SortFields.Add Key:=tblData.ListColumns(2).Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Header = xlYes
            .Apply
        End With
    End If
End Sub

