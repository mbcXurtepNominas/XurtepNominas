Imports ClosedXML.Excel
Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Word 'control de office
Imports Microsoft.Office.Interop
Imports System.Data
Public Class frmImportarEmpleadosAlta
    Dim sheetIndex As Integer = -1
    Dim SQL As String
    Dim contacolumna As Integer

    Private Sub frmImportarEmpleadosAlta_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        ''MostrarEmpresasC()
        tsbImportar_Click(sender, e)
    End Sub

    'Private Sub MostrarEmpresasC()
    '    SQL = "select (nombre + ' ' + ruta) AS nombre, iIdEmpresaC from empresaC ORDER BY nombre"
    '    nCargaCBO(cbEmpresasC, SQL, "nombre", "iIdEmpresaC")
    '    cbEmpresasC.SelectedIndex = 0

    'End Sub


    Private Sub tsbNuevo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles tsbNuevo.Click
        tsbNuevo.Enabled = False
        tsbImportar.Enabled = True
        tsbImportar_Click(sender, e)
    End Sub

    Private Sub tsbImportar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles tsbImportar.Click
        Dim dialogo As New OpenFileDialog
        lblRuta.Text = ""
        With dialogo
            .Title = "Búsqueda de archivos de saldos."
            .Filter = "Hoja de cálculo de excel (xlsx)|*.xlsx;"
            .CheckFileExists = True
            If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
                lblRuta.Text = .FileName
            End If
        End With
        tsbProcesar.Enabled = lblRuta.Text.Length > 0
        If tsbProcesar.Enabled Then
            tsbProcesar_Click(sender, e)
        End If
    End Sub

    Private Sub tsbProcesar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles tsbProcesar.Click
        lsvLista.Items.Clear()
        lsvLista.Columns.Clear()
        lsvLista.Clear()

        pnlCatalogo.Enabled = False
        tsbGuardar.Enabled = False
        tsbCancelar.Enabled = False
        lsvLista.Visible = False
        tsbImportar.Enabled = False
        Me.cmdCerrar.Enabled = False
        Me.Cursor = Cursors.WaitCursor
        Me.Enabled = False
        'Application.DoEvents()

        Try
            If File.Exists(lblRuta.Text) Then
                Dim Archivo As String = lblRuta.Text
                Dim Hoja As String


                Dim book As New ClosedXML.Excel.XLWorkbook(Archivo)
                If book.Worksheets.Count >= 1 Then
                    sheetIndex = 1
                    If book.Worksheets.Count > 1 Then
                        Dim Forma As New frmHojasNomina
                        Dim Hojas As String = ""
                        For i As Integer = 0 To book.Worksheets.Count - 1
                            Hojas &= book.Worksheets(i).Name & IIf(i < (book.Worksheets.Count - 1), "|", "")
                        Next
                        Forma.Hojas = Hojas
                        If Forma.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                            sheetIndex = Forma.selectedIndex + 1
                        Else
                            Exit Sub
                        End If
                    End If
                    Hoja = book.Worksheet(sheetIndex).Name
                    Dim sheet As IXLWorksheet = book.Worksheet(sheetIndex)

                    Dim colIni As Integer = sheet.FirstColumnUsed().ColumnNumber()
                    Dim colFin As Integer = sheet.LastColumnUsed().ColumnNumber()
                    Dim Columna As String
                    Dim numerocolumna As Integer = 1


                    lsvLista.Columns.Add("#")
                    For c As Integer = colIni To colFin

                        lsvLista.Columns.Add(sheet.Cell(1, c).Value.ToString)
                        'lsvLista.Columns.Add(numerocolumna)
                        'numerocolumna = numerocolumna + 1

                    Next

                    'lsvLista.Columns.Add("conciliacion")
                    'lsvLista.Columns.Add("color")

                    lsvLista.Columns(1).Width = 90

                    lsvLista.Columns(2).Width = 200
                    lsvLista.Columns(3).Width = 100
                    lsvLista.Columns(4).Width = 200
                    lsvLista.Columns(5).Width = 50
                    lsvLista.Columns(6).Width = 200
                    lsvLista.Columns(7).Width = 150
                    lsvLista.Columns(7).TextAlign = 1
                    lsvLista.Columns(8).Width = 150
                    lsvLista.Columns(8).TextAlign = 1
                    lsvLista.Columns(9).Width = 150
                    lsvLista.Columns(9).TextAlign = 1
                    lsvLista.Columns(10).Width = 100
                    lsvLista.Columns(11).Width = 400





                    Dim Filas As Long = sheet.RowsUsed().Count()
                    For f As Integer = 2 To Filas
                        Dim item As ListViewItem = lsvLista.Items.Add((f - 1).ToString())
                        For c As Integer = colIni To colFin
                            Try

                                Dim Valor As String = ""
                                If (sheet.Cell(f, c).ValueCached Is Nothing) Then
                                    Valor = sheet.Cell(f, c).Value.ToString()
                                Else
                                    Valor = sheet.Cell(f, c).ValueCached.ToString()
                                End If
                                Valor = Valor.Trim()
                                item.SubItems.Add(Valor)


                                If f = 6 And c >= 12 Then

                                    'If Valor <> "" AndAlso InStr(Valor, "-") > 0 Then
                                    '    Dim sValores() As String = Valor.Split("-")
                                    '    Select Case sValores(0).ToUpper()
                                    '        Case "P"
                                    '            item.SubItems(item.SubItems.Count - 1).Tag = "1" 'Percepción
                                    '        Case "D"
                                    '            item.SubItems(item.SubItems.Count - 1).Tag = "2" 'Deducción
                                    '        Case "I"
                                    '            item.SubItems(item.SubItems.Count - 1).Tag = "3" 'Incapacidad
                                    '    End Select
                                    '    Valor = sValores(1).Trim()
                                    'End If
                                    item.SubItems(item.SubItems.Count - 1).Text = Valor
                                End If



                            Catch ex As Exception

                            End Try

                        Next
                    Next

                    book.Dispose()
                    book = Nothing
                    GC.Collect()
                    'If lsvNominaFile.Items.Count >= 9 Then
                    '    If chkTipo.Checked Then
                    '        ProcesarNomina()
                    '    Else
                    '        ProcesarNomina1()
                    '    End If

                    'End If
                    pnlCatalogo.Enabled = True
                    If lsvLista.Items.Count = 0 Then
                        MessageBox.Show("El catálogo no puso ser importado o no contiene registros." & vbCrLf & "¿Por favor verifique?", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Else
                        MessageBox.Show("Se han encontrado " & FormatNumber(lsvLista.Items.Count, 0) & " registros en el archivo.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                        tsbGuardar.Enabled = True
                        tsbCancelar.Enabled = True
                        lblRuta.Text = FormatNumber(lsvLista.Items.Count, 0) & " registros en el archivo."
                        Me.Enabled = True
                        Me.cmdCerrar.Enabled = True
                        Me.Cursor = Cursors.Default
                        tsbImportar.Enabled = True
                        lsvLista.Visible = True
                    End If




                ElseIf book.Worksheets.Count = 0 Then
                    MessageBox.Show("El archivo no contiene hojas.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Else
                MessageBox.Show("El archivo ya no se encuentra en la ruta indicada.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Function getColumnName(ByVal index As Single) As String
        Dim numletter As Single = 26
        Dim sGrupo As Single = index / numletter
        Dim Modulo As Single = sGrupo - Math.Truncate(sGrupo)

        If Modulo = 0 Then Modulo = 1
        Dim Grupo As Integer = sGrupo - Modulo

        Dim Indice As Integer = index - (Grupo * numletter)
        Dim ColumnName As String = ""

        If Grupo > 0 Then
            ColumnName = Chr(64 + Grupo)
        End If
        ColumnName &= Chr(64 + Indice)
        Return ColumnName

    End Function

    Private Sub tsbGuardar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles tsbGuardar.Click
        Dim SQL As String, nombresistema As String = ""
        Dim bandera As Boolean

        Dim x As Integer

        Try
            If lsvLista.CheckedItems.Count > 0 Then
                Dim mensaje As String


                pnlProgreso.Visible = True
                pnlCatalogo.Enabled = False
                'Application.DoEvents()


                Dim IdEmpleado As Long
                Dim i As Integer = 0

                Dim t As Integer = 0
                Dim conta As Integer = 0



                pgbProgreso.Minimum = 0
                pgbProgreso.Value = 0
                pgbProgreso.Maximum = lsvLista.CheckedItems.Count


                'Dim fila As New DataRow
                SQL = "Select * from usuarios where idUsuario = " & idUsuario
                Dim rwFilas As DataRow() = nConsulta(SQL)

                If rwFilas Is Nothing = False Then
                    Dim Fila As DataRow = rwFilas(0)
                    nombresistema = Fila.Item("nombre")
                End If

                Dim empleadofull As ListViewItem
                Dim mensa As String
                '' mensa = "Datos incompletos en el empleado: "

                For Each empleado As ListViewItem In lsvLista.CheckedItems


                    For x = 0 To empleado.SubItems.Count - 1

                        If empleado.SubItems(x).Text = "" Then
                            mensa = " Datos incompletos en el empleado: Empleado: " & empleado.Text & " Columna:" & x.ToString() & " "


                            '' MessageBox.Show(mensa, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            bandera = False
                            x = empleado.SubItems.Count - 1

                        Else

                            empleadofull = empleado

                            '' MessageBox.Show("Pasa" & empleado.SubItems(x).Text, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                            bandera = True

                        End If
                    Next x

                    If bandera <> False Then

                        Dim b As String = Trim(empleadofull.SubItems(24).Text)
                        Dim idbanco As Integer
                        If b <> "" Then
                            Dim banco As DataRow() = nConsulta("select * from bancos where clave =" & b)
                            If banco Is Nothing Then
                                idbanco = 1
                                mensa = "Revise el tipo de banco"
                                bandera = False
                            Else
                                idbanco = banco(0).Item("iIdBanco")
                            End If

                        Else
                            b = 0
                        End If
                        Dim p As String = Trim(empleadofull.SubItems(15).Text) ''idPuesto
                        Dim cPuesto As String
                        If p <> "" Then
                            Dim puesto As DataRow() = nConsulta("SELECT * FROM Puestos where iIdPuesto =" & p)
                            If puesto Is Nothing Then
                                cPuesto = ""
                                mensa = "Revise el tipo de banco"
                                bandera = False
                            Else
                                cPuesto = puesto(0).Item("cNombre")
                            End If

                        Else
                            p = 0
                        End If

                        Dim factor As Integer
                        Select Case Trim(empleadofull.SubItems(32).Text)
                            Case "VSM"
                                factor = 0
                                ' The following is the only Case clause that evaluates to True.
                            Case "PORCENTAJE"
                                factor = 1
                            Case "CUOTA FIJA"
                                factor = 2
                            Case Else
                                factor = 0
                        End Select

                        Dim number As Integer
                        Select Case Trim(empleadofull.SubItems(33).Text)
                            Case "QUINCENAL"
                                number = 4
                                ' The following is the only Case clause that evaluates to True.
                            Case "MENSUAL"
                                number = 5
                            Case "SEMANAL"
                                number = 2
                            Case Else
                                number = 10
                        End Select



                        Dim dFechaNac, dFechaCap, dFechaPlanta As String ''--, dFechaPatrona, dFechaTerminoContrato, dFechaSindicato, dFechaAntiguedad As String

                        dFechaNac = Trim(empleadofull.SubItems(13).Text) ''Format(Trim(empleadofull.SubItems(18).Text), "yyyy/dd/MM")
                        dFechaCap = (Trim(empleadofull.SubItems(14).Text))
                        dFechaPlanta = Trim(empleadofull.SubItems(40).Text)
                        'dFechaPatrona = (Trim(empleadofull.SubItems(14).Text))
                        'dFechaTerminoContrato = ((Trim(empleadofull.SubItems(44).Text))) ''No asignado
                        'dFechaSindicato = (Trim(empleadofull.SubItems(14).Text))
                        'dFechaAntiguedad = Trim(empleadofull.SubItems(14).Text)


                        SQL = "EXEC setempleadosCInsertar 0,'" & Trim(empleadofull.SubItems(1).Text) & "','" & Trim(empleadofull.SubItems(2).Text)
                        SQL &= "','" & Trim(empleadofull.SubItems(3).Text)
                        SQL &= "','" & Trim(empleadofull.SubItems(4).Text) & "','" & Trim(empleadofull.SubItems(3).Text) & " " & Trim(empleadofull.SubItems(4).Text) & " " & Trim(empleadofull.SubItems(2).Text)
                        SQL &= "','" & Trim(empleadofull.SubItems(5).Text) & "','" & Trim(empleadofull.SubItems(6).Text) & "','" & Trim(empleadofull.SubItems(7).Text)
                        SQL &= "','" & Trim(empleadofull.SubItems(8).Text)
                        SQL &= "','" & Trim(empleadofull.SubItems(9).Text) & "'," & Trim(empleadofull.SubItems(10).Text) & ",'" & Trim(empleadofull.SubItems(11).Text)
                        SQL &= "'," & IIf(Trim(empleadofull.SubItems(12).Text) = "FEMENINO", 0, 1) & ",'" & dFechaNac & "','" & dFechaCap
                        SQL &= "','" & cPuesto & "','" & Trim(empleadofull.SubItems(16).Text)
                        SQL &= "'," & IIf(Trim(empleadofull.SubItems(17).Text) = "", 0, Trim(empleadofull.SubItems(17).Text)) & "," & IIf(Trim(empleadofull.SubItems(18).Text) = "", 0, Trim(empleadofull.SubItems(18).Text))
                        SQL &= ",'" & Trim(empleadofull.SubItems(19).Text) & "','" & Trim(empleadofull.SubItems(20).Text) & "','','','" & Trim(empleadofull.SubItems(21).Text) & "','" & Trim(empleadofull.SubItems(22).Text)
                        SQL &= "',1," & IIf((empleadofull.SubItems(23).Text) = "", 0, (empleadofull.SubItems(23).Text)) & ",0" & ",-1" & "," & 1 & "," & idbanco
                        SQL &= ",'" & Trim(empleadofull.SubItems(25).Text) & "',1,'" & Trim(empleadofull.SubItems(26).Text)
                        SQL &= "','" & Trim(empleadofull.SubItems(27).Text) & "'," & Trim(empleadofull.SubItems(28).Text) & ",'" & Trim(empleadofull.SubItems(29).Text)
                        SQL &= "','" & dFechaCap & "','" & dFechaCap & "','" & dFechaCap
                        SQL &= "'," & 0 & ",'" & Trim(empleadofull.SubItems(30).Text) & "','" & " "
                        SQL &= "'," & 1 & ",'" & Trim(empleadofull.SubItems(31).Text) & "','" & Trim(empleadofull.SubItems(32).Text) ''factor
                        SQL &= "'," & 0 & ",'" & Trim(empleadofull.SubItems(33).Text) & "','" & Trim(empleadofull.SubItems(34).Text)
                        SQL &= "','" & Trim(empleadofull.SubItems(35).Text) & "','" & Trim(empleadofull.SubItems(36).Text) & "','" & Trim(empleadofull.SubItems(37).Text) & "'," & -1 ''estatus 
                        SQL &= "," & Trim(empleadofull.SubItems(15).Text) & "," & Trim(empleadofull.SubItems(38).Text)
                        SQL &= "," & IIf(Trim(empleadofull.SubItems(39).Text) = "SOLTERO", 0, 1)
                        SQL &= "," & 1
                        SQL &= ",'" & " "
                        SQL &= "','" & "" & "'"
                        SQL &= "," & 0 & ",'" & dFechaPlanta & "','" & Trim(empleadofull.SubItems(41).Text) & "','" & Trim(empleadofull.SubItems(42).Text) & "'"
                        SQL &= ",'" & Trim(empleadofull.SubItems(43).Text) & "','" & " " & "'"

                        If nExecute(SQL) = False Then
                            MessageBox.Show("Error en el registro con los siguiente datos:   Empleado:  " & Trim(empleado.SubItems(3).Text), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                            Exit Sub
                        End If
                        pgbProgreso.Value += 1
                        ''Application.DoEvents()
                        t = t + 1
                    Else
                        MessageBox.Show(mensa, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        tsbCancelar_Click(sender, e)
                        pnlProgreso.Visible = False
                    End If




                Next

                If bandera <> False Then
                    tsbCancelar_Click(sender, e)
                    pnlProgreso.Visible = False

                    MessageBox.Show(t.ToString() & "  Proceso terminado", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    pnlProgreso.Visible = False
                    MessageBox.Show("No se guardo ninguna dato, revise y vuelva a intentarlo ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If


            Else

                MessageBox.Show("Por favor seleccione al menos una registro para importar.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
            pnlCatalogo.Enabled = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub chkAll_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkAll.CheckedChanged
        For Each item As ListViewItem In lsvLista.Items
            item.Checked = chkAll.Checked
        Next
        chkAll.Text = IIf(chkAll.Checked, "Desmarcar todos", "Marcar todos")
    End Sub

    Private Sub cmdCerrar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdCerrar.Click
        Me.Close()
    End Sub

    Private Sub tsbCancelar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles tsbCancelar.Click
        pnlCatalogo.Enabled = False
        lsvLista.Items.Clear()
        chkAll.Checked = False
        lblRuta.Text = ""
        tsbImportar.Enabled = False
        tsbProcesar.Enabled = False
        tsbGuardar.Enabled = False
        tsbCancelar.Enabled = False
        tsbNuevo.Enabled = True
    End Sub

    Private Sub frmImportarEmpladosAlta_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

    End Sub


    Private Sub abiriEmpresasC()
        'Declaramos la variable nombre
        Dim nombre As String
        'Entrada de datos mediante un inputbox
        nombre = InputBox("Ingrese Nombre de empresa ",
                         "Registro de Datos Personales",
                         "Nombre", 100, 0)
        MessageBox.Show("Bienvenido Usuario: " + nombre,
                        "Registro de Datos Personales",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)
    End Sub

    Private Sub tsbContrato_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim MSWord As New Word.Application
        Dim Documento As Word.Document
        Dim Ruta As String, strPWD As String
        Dim SQL As String
        Try

            Ruta = System.Windows.Forms.Application.StartupPath & "\Archivos\test.docx"

            FileCopy(Ruta, "C:\Temp\TMM.docx")
            Documento = MSWord.Documents.Open("C:\Temp\TMM.docx")

            'SQL = "select iIdEmpleadoAlta,cCodigoEmpleado,empleadosAlta.cNombre,cApellidoP,cApellidoM,cRFC,cCURP,"
            'SQL &= "cIMSS,cDescanso,cCalleNumero,cCiudadP,cCP,iSexo,iEstadoCivil, dFechaNac,puestos.cNombre as cPuesto,fSueldoBase,"
            'SQL &= "cNacionalidad,empleadosAlta.cFuncionesPuesto, fSueldoOrd, iOrigen,empresa.iIdEmpresa ,empresa.calle +' '+ empresa.numero AS cDireccionP, empresa.localidad as cCiudadP, empresa.cp as cCPP, iCategoria, cJornada, cHorario,"
            'SQL &= "cHoras, cDescanso, cFechaPago, cFormaPago, cLugarPago, cLugarFirmaContrato,empleadosAlta.cLugarPrestacion, dFechaPatrona,"
            'SQL &= "empresa.nombrefiscal, empresa.RFC AS cRFCP, empresa.cRepresentanteP, empresa.cObjetoSocialP,  Cat_SindicatosAlta.cNombre AS cNombreSindicato"
            'SQL &= " from ((empleadosAlta"
            'SQL &= " inner join empresa on fkiIdEmpresa= iIdEmpresa)"
            'SQL &= " inner join puestos on fkiIdPuesto= iIdPuesto)"
            'SQL &= " inner join (clientes inner join Cat_SindicatosAlta on fkiIdSindicato= iIdSindicato) on fkiIdCliente=iIdCliente"
            'SQL &= " where iIdEmpleadoAlta = " & gIdEmpleado
            SQL = "SELECT * FROM (empleadosC INNER JOIN familiar on iIdEmpleadoC=fkiIdEmpleadoC) WHERE iIdEmpleado="
            Dim rwEmpleado As DataRow() = nConsulta(SQL)

            If rwEmpleado Is Nothing = False Then
                Dim fEmpleado As DataRow = rwEmpleado(0)


                Documento.Bookmarks.Item("cNombreLargo").Range.Text = fEmpleado.Item("cNombre") & " " & fEmpleado.Item("cApellidoP") & " " & fEmpleado.Item("cApellidoM")
                Documento.Bookmarks.Item("cNombreLargo2").Range.Text = fEmpleado.Item("cNombre") & " " & fEmpleado.Item("cApellidoP") & " " & fEmpleado.Item("cApellidoM")
                Documento.Bookmarks.Item("cNombreFiscal").Range.Text = fEmpleado.Item("nombrefiscal")
                Documento.Bookmarks.Item("cFecha").Range.Text = DateTime.Now.ToString("dd/MM/yyyy")
                Documento.Bookmarks.Item("cFecha2").Range.Text = DateTime.Now.ToString("dd/MM/yyyy")
                Documento.Bookmarks.Item("cLugarFirma").Range.Text = fEmpleado.Item("cLugarFirmaContrato")
                Documento.Bookmarks.Item("cCURP").Range.Text = fEmpleado.Item("cCURP")
                Documento.Bookmarks.Item("cRFC").Range.Text = fEmpleado.Item("cRFC")
                Documento.Bookmarks.Item("cRFCP").Range.Text = fEmpleado.Item("cRFCP")

                Documento.Bookmarks.Item("cDireccionP").Range.Text = fEmpleado.Item("cDireccionP") & ", " & fEmpleado.Item("cCiudadP") & ", " & fEmpleado.Item("cCPP")
                Documento.Bookmarks.Item("cDireccionP2").Range.Text = fEmpleado.Item("cDireccionP") & ", " & fEmpleado.Item("cCiudadP") & ", " & fEmpleado.Item("cCPP")
                Documento.Bookmarks.Item("cDireccion").Range.Text = fEmpleado.Item("cCalleNumero") & ", " & fEmpleado.Item("cCiudadP") & ", " & fEmpleado.Item("cCP")
                If IsDBNull(fEmpleado.Item("cRepresentanteP")) = False Then
                    Documento.Bookmarks.Item("cRepresentanteP").Range.Text = fEmpleado.Item("cRepresentanteP")
                    Documento.Bookmarks.Item("cRepresentanteP2").Range.Text = fEmpleado.Item("cRepresentanteP")
                Else
                    MessageBox.Show("Falta agregar Representante Patrona", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                End If
                Documento.Save()
                MSWord.Visible = True
            End If

        Catch ex As Exception
            Documento.Close()
            MessageBox.Show(ex.ToString(), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Try
    End Sub
End Class