VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim tblClientes As ListObject
    Dim clientID As String
    Dim clientName As String
    Dim n As Long

    On Error Resume Next
    Set tblClientes = Me.ListObjects("TablaClientes")
    On Error GoTo 0
    
    If tblClientes Is Nothing Then Exit Sub
    If tblClientes.DataBodyRange Is Nothing Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub

    ' =================================================================
    ' DISPARADOR ÚNICO: Solo reaccionar cuando se modifique el Estatus
    ' =================================================================
    If Intersect(Target, tblClientes.ListColumns(3).DataBodyRange) Is Nothing Then Exit Sub
    If Trim(Target.Value) = "" Then Exit Sub

    ' PARTE 1: INYECCIÓN DE DOCUMENTOS Y ORDENAMIENTO EN LA HOJA DATA
    If Trim(UCase(Target.Value)) = "ACTIVO" Then
        clientID = Trim(Target.Offset(0, -2).Value)
        clientName = Trim(Target.Offset(0, -1).Value)

        ' Candado: No hacer nada si el ID o Nombre están vacíos
        If clientID = "" Or clientName = "" Then
            Application.EnableEvents = False
            Target.Value = ""
            Application.EnableEvents = True
            MsgBox "Error: Debes ingresar el ID y el Nombre del cliente ANTES de cambiar su estatus a Activo.", _
                   vbCritical, "Faltan Datos"
            Exit Sub
        End If

        ' Inyección a la hoja DATA
        n = SincronizarCliente(clientID, clientName)

        If n > 0 Then
            MsgBox "Cliente activado. Se generaron " & n & " documentos para " & clientName & _
                   " en la hoja DATA.", vbInformation, "Styba Automatización"
                   
            ' [NUEVO] Ordenar automáticamente la hoja DATA por Nombre (Columna 2)
            Dim tblData As ListObject
            On Error Resume Next
            Set tblData = ThisWorkbook.Sheets("DATA").ListObjects("Tabla1")
            On Error GoTo 0
            
            If Not tblData Is Nothing Then
                Application.EnableEvents = False
                With tblData.Sort
                    .SortFields.Clear
                    .SortFields.Add Key:=tblData.ListColumns(2).Range, _
                                    SortOn:=xlSortOnValues, Order:=xlAscending, _
                                    DataOption:=xlSortNormal
                    .Header = xlYes
                    .Apply
                End With
                Application.EnableEvents = True
            End If
        End If
    End If

    ' =================================================================
    ' PARTE 2: ORDENAMIENTO ALFABÉTICO POR NOMBRE EN LA HOJA CLIENTES
    ' =================================================================
    ' Una vez sincronizado, ordenamos toda la tabla junta
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    With tblClientes.Sort
        .SortFields.Clear
        ' Ordenamos usando la Columna 2 (Nombre Cliente)
        .SortFields.Add Key:=tblClientes.ListColumns(2).Range, _
                        SortOn:=xlSortOnValues, Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .Header = xlYes
        .Apply
    End With

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

