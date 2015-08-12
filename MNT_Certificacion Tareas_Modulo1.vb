Public iniTime, endTime As Variant

Sub TareasRegion()
    Dim Dict1, Dict2, Dict3, WSF As Object
    Dim rng As Range
    Set WSF = WorksheetFunction

    ' desactivo el refresco de pantalla
    Application.ScreenUpdating = False
    ' empieza el contador
    iniTime = Now()

    Sheets("Territorios").Activate
    ' borra el contenido de las hojas de los territorios
    Sheets("BALEARES").Range("8:1000").ClearContents
    Sheets("ARAGÓN").Range("8:1000").ClearContents
    Sheets("CANARIAS").Range("8:1000").ClearContents
    Sheets("SUR").Range("8:1000").ClearContents
    Sheets("CATALUNYA").Range("8:1000").ClearContents
    Sheets("Resumen").Range("2:1000").ClearContents
    Sheets("Territorios").Range("A3:F7").ClearContents ' borramos los datos resumen de la anterior lectura
    Sheets("Territorios").Range("CB10:CC" & Range("CB65536").End(xlUp).Row).ClearContents ' borramos las columnas del sumatorio de baremos y Cotizaciones de cada tarea
    
    'ordena la tabla importada por el Territorio, muestra el autofiltro y da formato a las celdas de Resumen
    OrdenaDatos1
    
    ' sustituye los baremos del JIRA por los de ENDESA
    Sheets("Equivalencias").Select
    Range("G2:G60").Select
    Selection.Copy
    Sheets("TERRITORIOS").Select
    Range("U9").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    ' inserta las cotizaciones de cada baremo
    Sheets("Equivalencias").Select
    Range("H2:H60").Select
    Selection.Copy
    Sheets("TERRITORIOS").Select
    Range("U8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True

    ' generamos las formulas de suma de baremos que falten
    For f = 10 To Range("A65536").End(xlUp).Row
        'Cells(f, 80).FormulaR1C1 = "=SUM(RC[-59]:RC[-1])"
        Cells(f, 80) = WSF.Sum(Range(Cells(f, 21), Cells(f, 79)))
    Next
    
    ' creamos los objetos a usar
    Set Dict1 = CreateObject("scripting.dictionary")
    Set Dict2 = CreateObject("scripting.dictionary")
    Set WSF = Application.WorksheetFunction
    ' ultima fila con valores de la tabla
    endRow = Range("A65536").End(xlUp).Row
    
        
    ' crea una tabla con los territorios y tareas leidos de Jira
    ' tabla para contar los territorios y las tareas que tiene
    With Dict1
        .comparemode = 1
        For Each rng In Sheets("Territorios").Range("E10:E" & endRow)
             If Dict1.Exists(rng.Text) Then
                 Dict1.Item(rng.Value) = Dict1.Item(rng.Value) + 1
             Else
                 Dict1.Add rng.Text, 1
             End If
        Next rng
    End With
    ' tabla para contar los territorios y los baremos que tiene
    With Dict2
        For Each rng In Sheets("Territorios").Range("E10:E" & endRow)
             If Dict2.Exists(rng.Text) Then
                 Dict2.Item(rng.Value) = Dict2.Item(rng.Value) + Cells(rng.Row, 80)
             Else
                 Dict2.Add rng.Text, Cells(rng.Row, 80)
             End If
        Next rng
    End With
    
    ' muestra el resumen de los datos de la tabla
    Sheets("Territorios").Range("A3").Resize(Dict1.Count, 1).Value = WSF.Transpose(Dict1.Keys)
    Sheets("Territorios").Range("B3").Resize(Dict1.Count, 1).Value = WSF.Transpose(Dict1.Items)
    Sheets("Territorios").Range("C3").Resize(Dict2.Count, 1).Value = WSF.Transpose(Dict2.Items)
    'Sheets("Territorios").Range("B8") = WSF.Sum(Range("B3:B7"))
    'Sheets("Territorios").Range("C8") = WSF.Sum(Range("U10:CA" & endRow))
    
    ' borra los dict
    Dict1.RemoveAll
    Dict2.RemoveAll
    ' activo el refresco de pantalla
    Application.ScreenUpdating = True

    ' llama a la rutina de Certificacion
    Certifica
End Sub


Sub Certifica()
    Dim Dict3, WSF As Object
    Dim rng As Range
    Dim hoy, FFin As Date
    
    Application.ScreenUpdating = False
    Sheets("Territorios").Activate
    
    Set WSF = WorksheetFunction    
    Set Dict3 = CreateObject("scripting.dictionary")
    Set FSO = CreateObject("Scripting.FileSystemObject")
   
   ' generamos un fichero de Log de Lectura
    fileLog = "JIRA_Certificacion_Tareas_LOG_" & Format(Now(), "yyyymmdd_hhmm")
	Set a = FSO.CreateTextFile (fileLog)
    a.WriteLine "Log Lectura Tareas Pte. Entregar en GOM " & Now()
    a.WriteLine ""
    
	' defino variables para lineas
    iRowIni = Sheets("Territorios").Range("B:B").Find("Código tarea GOM", LookAt:=xlWhole, SearchOrder:=xlRows, SearchDirection:=xlDown).Row + 1
    iRowEnd = Sheets("Territorios").Range("B65536").End(xlUp).Row
    iEndesa = Sheets("Territorios").Range("A:A").Find("Clave", LookAt:=xlWhole, SearchOrder:=xlRows, SearchDirection:=xlDown).Row
    
    ' recorre la hoja de Territorios y va copiando los datos de cada tarea en la hoja del territorio
    For n = iRowIni To iRowEnd
        ' completa el codigo de la tarea con los ceros
        nceros = 8 - Len(Sheets("Territorios").Cells(n, 2))
        idTar = Sheets("Territorios").Cells(n, 2).Text
        For C = 1 To nceros
            idTar = "0" & idTar
        Next C
        ' modifica el formato de la Tarea GOM al que tiene los ceros
        Sheets("Territorios").Cells(n, 2) = idTar
                    
        ' comprueba la fecha limite de la Tarea (si no tiene le asigna la del dia en curso)
        hoy = CDate(Format(Now(), "dd-mmm-yy"))
        If Sheets("Territorios").Cells(n, 15) = vbNullString Then
            FFin = hoy
            Sheets("Territorios").Cells(n, 15) = " "
            Else
            'FFin = CDate(Sheets("Territorios").Cells(n, 14))
            Set f = Sheets("Descargos").ListObjects("ConsultaSQL").ListColumns("Tarea").DataBodyRange.Find(idTar, LookAt:=xlWhole)
            If Not f Is Nothing Then
                iRow = Range(f.Address).Row
                FFin = CDate(Sheets("Descargos").Cells(iRow, 9))
                Sheets("Territorios").Cells(n, 15) = FFin
            End If
        End If
        
        ' si la Resolucion de Tarea esta vacio lo rellena
        If Sheets("Territorios").Cells(n, 17) = vbNullString Then Sheets("Territorios").Cells(n, 17) = "Resuelto SIN Incidencias"
        
        ' añade un hipervinculo a la Tarea del JIRA
        pathJira = "https://herramientasdp.sadiel.net/HGP/browse/"
        With Sheets("Territorios")
            .Hyperlinks.Add Anchor:=.Cells(n, 1), _
            Address:=pathJira & .Cells(n, 1).Text, _
            TextToDisplay:=.Cells(n, 1).Text
        End With

        ' Comprueba si la tarea tiene baremos para valorar
        If Cells(n, 80) = 0 Then
            Sheets("Territorios").Range(Cells(n, 1), Cells(n, 81)).Style = "Incorrecto"
            a.writeLine "Linea " & n & " del fichero. La tarea " & Cells(n, 1).Text & " de " & Cells(n, 5).Text & " no tiene Baremos."
			Else
            ' comprueba que la tarea es SIN INCIDENCIA y que la Fecha Fin Descargo es menor que el dia en curso
            If Sheets("Territorios").Cells(n, 17) <> "Resuelto CON Incidencias" And FFin <= hoy Then
                ' para la tarea vigente recorre las columnas de los baremos y añade los que tengan datos
                Region = Sheets("Territorios").Cells(n, 5)
                f = Sheets(Region).Range("A5536").End(xlUp).Row + 1 'ultima fila en blanco en la hoja del territorio
                For m = 21 To 79
                    If Sheets("Territorios").Cells(n, m) <> vbNullString Then
                        Sheets(Region).Cells(f, 1) = idTar 'codigo GOM
                        Sheets(Region).Cells(f, 2) = "CI"
                        Sheets(Region).Cells(f, 3) = "GDD"
                        Sheets(Region).Cells(f, 4) = Sheets("Territorios").Cells(iEndesa, m) 'id Baremo
                        Sheets(Region).Cells(f, 5) = Sheets("Territorios").Cells(n, m) 'unidades del Baremo
                        Sheets(Region).Cells(f, 6) = Sheets("Territorios").Cells(n, 20) 'comentario JIRA
                        Sheets("Territorios").Cells(n, 81) = Sheets("Territorios").Cells(n, 81) + (Cells(n, m).Value * Cells(8, m).Value) / 100
                        f = f + 1
                    End If
                Next m
                ' cuenta a origen de las cotizaciones
                impCert = impCert + Sheets("Territorios").Cells(n, 81)
                ' marca la tarea como leida correctamente
                Sheets("Territorios").Cells(n, 1).Style = "Buena"
                Else
                Sheets("Territorios").Range(Cells(n, 1), Cells(n, 81)).Style = "Neutral"
				a.WriteLine "Linea " & n & " del fichero. La tarea " & Cells(n, 1).Text & " de " & Cells(n, 5).Text & " esta CON Incidencia o Fuera de Fecha."
			End If
        End If
        ' reinicia la variable nBar
        nBar = Null
    Next n
	
    'cierra el fichero de LOG
	a.writeline "Fin de la lectura"
	a.close
	Shell "notepad.exe " & fileLog, vbNormalFocus
	
    ' muestra la valoracion de
    Sheets("Territorios").Range("C1") = "Valoracion: " & Format(WSF.Sum(Range("CC10:CC" & iRowEnd)), "0.00 €")
    Sheets("Territorios").Range("C1") = "Valoracion: " & Format(impCert, "0.00 €")
    
    ' Recuento y verifico datos
    nReg = Sheets("Territorios").Range("A3").End(xlDown).Row
    For Reg = 3 To nReg
        Region = Sheets("Territorios").Cells(Reg, 1).Text
        f = Sheets(Region).Range("A65536").End(xlUp).Row
        With Dict3
            .comparemode = 1
            For Each rng In Sheets(Region).Range("A8:A" & f)
            If rng <> vbNullString Then
                If Dict3.Exists(rng.Text) Then
                    Dict3.Item(rng.Value) = Dict3.Item(rng.Value) + 1
                    Else
                    Dict3.Add rng.Text, 1
                End If
            End If
            Next rng
        End With
        Sheets("Territorios").Range("D" & Reg) = Dict3.Count
        Sheets("Territorios").Range("E" & Reg) = WSF.Sum(Sheets(Region).Range("E8:E" & f))
        Sheets("Territorios").Range("F" & Reg) = WSF.SumIf(Range(Cells(10, 5), Cells(iRowEnd, 5)), Region, Range(Cells(10, 81), Cells(iRowEnd, 81)))
        ' añado informacion resumida de los elementos exportados de cada territorio a la hoja Resumen
        COL = (Reg - 2) * 2
        Sheets("Resumen").Cells(1, COL) = Region
        Sheets("Resumen").Cells(2, COL).Resize(Dict3.Count, 1).Value = WSF.Transpose(Dict3.Keys)
        Sheets("Resumen").Cells(2, COL + 1).Resize(Dict3.Count, 1).Value = WSF.Transpose(Dict3.Items)
        Dict3.RemoveAll
    Next Reg
    ' totales despues de generar los ficheros
    'Sheets("Territorios").Range("D8") = WSF.Sum(Range("D3:D7"))
    'Sheets("Territorios").Range("E8") = WSF.Sum(Range("E3:E7"))
    OrdenaDatos2 ' ordena la tabla por la Resolucion de la tarea y el Territorio
    endTime = Now()
    
    LapTime = Format((endTime - iniTime), "hh:mm:ss")
    Opc = MsgBox("El proceso de preparacion de las hojas a tardado " & LapTime & vbCrLf & _
                "Desea Generar los ficheros para Certificar Tareas?", vbYesNo, "Generar Ficheros")
    If Opc = 6 Then
        ' exportar las Hojas
        ruta = "\\Espsev10ca01\mcartobase\00000_Mantenimiento_2015\LANZADOR\CERTIFICACIONES\"
        nReg = Sheets("Territorios").Range("A3").End(xlDown).Row
        For Reg = 3 To nReg
            Region = Sheets("Territorios").Cells(Reg, 1).Text
            nTareas = Sheets("Territorios").Cells(Reg, 4)
            Sheets(Region).Activate
            file = "JIRA_Certificacion_Tareas_" & Format(Now(), "yyyymmdd_hhmm") & "_" & nTareas & "_" & Region
            Sheets(Region).Copy
            Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs Filename:=ruta & file, FileFormat:=xlExcel8
            ActiveWorkbook.Close
            Application.DisplayAlerts = True
        Next Reg
    ' mensaje de finalizacion del proceso
    endTime = Now()
    LapTime = Format((endTime - iniTime), "hh:mm:ss")
    MsgBox "El proceso de exportar las hojas a tardado " & LapTime & vbCrLf & _
            "Se han generado los ficheros de forma satisfacoria", vbInformation, " Generar Ficheros"
    End If
        
    Sheets("Territorios").Activate
    Application.ScreenUpdating = True
    
End Sub
