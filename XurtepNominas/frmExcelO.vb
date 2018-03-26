Imports ClosedXML.Excel
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Net.Mime.MediaTypeNames
Imports Microsoft.Office.Interop

Public Class frmExcelO
    Dim sheetIndex As Integer = -1
    Dim SQL As String
    Dim contacolumna As Integer

    Private Sub frmExcel_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        'MostrarEmpresasC()
        Dim moment As Date = Date.Now()


        cboMes.SelectedIndex = moment.Month - 1
        cboTipoR.SelectedIndex = 1


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
            .Filter = "Hoja de cálculo de excel (xlsx)|*.xlsm;"
            .CheckFileExists = True
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
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
        tsbGuardar2.Enabled = False
        cmdNN.Enabled = False

        tsbCancelar.Enabled = False
        lsvLista.Visible = False
        tsbImportar.Enabled = False
        Me.cmdCerrar.Enabled = False
        Me.Cursor = Cursors.WaitCursor
        Me.Enabled = False
        System.Windows.Forms.Application.DoEvents()

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
                        If Forma.ShowDialog = Windows.Forms.DialogResult.OK Then
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


                    lsvLista.Columns(1).Width = 100

                    lsvLista.Columns(2).Width = 250
                    lsvLista.Columns(3).Width = 100
                    lsvLista.Columns(4).Width = 200
                    lsvLista.Columns(5).Width = 100
                    lsvLista.Columns(6).Width = 200
                    lsvLista.Columns(7).Width = 100
                    ''lsvLista.Columns(7).TextAlign = 1
                    lsvLista.Columns(8).Width = 70
                    '' lsvLista.Columns(8).TextAlign = 1
                    lsvLista.Columns(9).Width = 150
                    ''lsvLista.Columns(9).TextAlign = 1
                    lsvLista.Columns(10).Width = 150
                    lsvLista.Columns(11).Width = 90
                    lsvLista.Columns(12).Width = 91
                    lsvLista.Columns(13).Width = 92
                    lsvLista.Columns(14).Width = 93
                    lsvLista.Columns(15).Width = 94
                    lsvLista.Columns(16).Width = 95
                    lsvLista.Columns(17).Width = 96
                    lsvLista.Columns(18).Width = 97
                    lsvLista.Columns(19).Width = 98
                    lsvLista.Columns(20).Width = 99
                    lsvLista.Columns(21).Width = 100
                    lsvLista.Columns(22).Width = 101
                    lsvLista.Columns(23).Width = 102
                    lsvLista.Columns(24).Width = 103
                    lsvLista.Columns(25).Width = 104
                    lsvLista.Columns(26).Width = 105
                    lsvLista.Columns(27).Width = 106
                    lsvLista.Columns(28).Width = 107
                    lsvLista.Columns(29).Width = 108
                    lsvLista.Columns(30).Width = 109
                    lsvLista.Columns(31).Width = 110
                    lsvLista.Columns(32).Width = 111
                    lsvLista.Columns(33).Width = 112
                    lsvLista.Columns(34).Width = 113
                    lsvLista.Columns(35).Width = 114
                    lsvLista.Columns(36).Width = 115
                    lsvLista.Columns(37).Width = 116
                    lsvLista.Columns(38).Width = 117
                    lsvLista.Columns(39).Width = 118
                    lsvLista.Columns(40).Width = 119


                    Dim Filas As Long = sheet.RowsUsed().Count() + 1
                    For f As Integer = 6 To Filas
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

                                '' Existe(Valor)

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

                                'If f = 2 And c = 4 Then
                                Dim fecha As Date = sheet.Cell(2, 4).Value.ToString()

                                Dim fec As Date = fecha
                                cboMes.SelectedIndex = fec.Month - 1
                                'End If




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
                        MessageBox.Show("El catálogo no pudo ser importado o no contiene registros." & vbCrLf & "¿Por favor verifique?", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Else
                        MessageBox.Show("Se han encontrado " & FormatNumber(lsvLista.Items.Count, 0) & " registros en el archivo.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                        tsbGuardar.Enabled = True
                        tsbGuardar2.Enabled = True
                        cmdNN.Enabled = True

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

    Private Function Existe(ByVal valor As String) As Boolean

        For Each item As ListViewItem In lsvLista.Items

            If item.Text = valor _
                OrElse item.SubItems(1).Text = valor _
                OrElse item.SubItems(1).Text = valor Then
                item.SubItems(1).BackColor = Color.AliceBlue
                Return True

            End If

        Next

        Return False

    End Function

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
        Try


            Dim filaExcel As Integer = 2
            Dim dialogo As New SaveFileDialog()

            If lsvLista.CheckedItems.Count > 0 Then

                Dim libro As New ClosedXML.Excel.XLWorkbook
                Dim hoja As IXLWorksheet = libro.Worksheets.Add("Generales")
                Dim hoja2 As IXLWorksheet = libro.Worksheets.Add("Percepciones")
                Dim hoja3 As IXLWorksheet = libro.Worksheets.Add("Deducciones")
                Dim hoja4 As IXLWorksheet = libro.Worksheets.Add("Otros Pagos")

                hoja.Column("A").Width = 20
                hoja.Column("B").Width = 15
                hoja.Column("C").Width = 15
                hoja.Column("D").Width = 12
                hoja.Column("E").Width = 12
                hoja.Column("F").Width = 25
                hoja.Column("G").Width = 15
                hoja.Column("H").Width = 15
                hoja.Column("I").Width = 25
                hoja.Column("J").Width = 15
                hoja.Column("K").Width = 15
                hoja.Column("L").Width = 15
                hoja.Column("M").Width = 15
                hoja.Column("N").Width = 50
                hoja.Column("O").Width = 12

                hoja.Range(1, 1, 1, 31).Style.Font.FontSize = 10
                hoja.Range(1, 1, 1, 31).Style.Font.SetBold(True)
                hoja.Range(1, 1, 1, 31).Style.Alignment.WrapText = True
                hoja.Range(1, 1, 1, 31).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                hoja.Range(1, 1, 1, 31).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                hoja.Range(1, 1, 1, 31).Style.Fill.BackgroundColor = XLColor.FromHtml("#BDBDBD")
                hoja.Range(1, 1, 1, 31).Style.Font.FontColor = XLColor.FromHtml("#000000")


                hoja.Cell(1, 1).Value = "No. Empleado"
                hoja.Cell(1, 2).Value = "RFC"
                hoja.Cell(1, 3).Value = "Nombre"
                hoja.Cell(1, 4).Value = "CURP"
                hoja.Cell(1, 5).Value = "SSA"
                hoja.Cell(1, 6).Value = "Cuenta Bancaria"
                hoja.Cell(1, 7).Value = "SBC"
                hoja.Cell(1, 8).Value = "SDI"
                hoja.Cell(1, 9).Value = "Reg. Patronal"
                hoja.Cell(1, 10).Value = "Ent. Federativa"
                hoja.Cell(1, 11).Value = "Días Pagados"
                hoja.Cell(1, 12).Value = "FechaInicioRelLaboral"
                hoja.Cell(1, 13).Value = "TipoContrato" ''Numero
                hoja.Cell(1, 14).Value = "TipoContrato"
                hoja.Cell(1, 15).Value = "Sindicalizado"
                hoja.Cell(1, 16).Value = "TipoJornada"
                hoja.Cell(1, 17).Value = "TipoJornada"
                hoja.Cell(1, 18).Value = "Tipo Regimen" ''Numero
                hoja.Cell(1, 19).Value = "Tipo Regimen"
                hoja.Cell(1, 20).Value = "Departamento"
                hoja.Cell(1, 21).Value = "Puesto"
                hoja.Cell(1, 22).Value = "Riesgo Puesto" ''Numero
                hoja.Cell(1, 23).Value = "Riesgo Puesto"
                hoja.Cell(1, 24).Value = "Periodicidad Pago" ''Numero
                hoja.Cell(1, 25).Value = "Periodicidad Pago"
                hoja.Cell(1, 26).Value = "Banco" ''Numero
                hoja.Cell(1, 27).Value = "Banco"
                hoja.Cell(1, 28).Value = "Subcontratacion"
                hoja.Cell(1, 29).Value = "Tipo Recibo"
                hoja.Cell(1, 30).Value = "Mes Pago"
                hoja.Cell(1, 31).Value = "Buque"

                ''Percepciones
                hoja2.Range(3, 1, 3, 23).Style.Font.FontSize = 10
                hoja2.Range(3, 1, 3, 23).Style.Font.SetBold(True)
                hoja2.Range(3, 1, 3, 23).Style.Alignment.WrapText = True
                hoja2.Range(3, 1, 3, 23).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                hoja2.Range(3, 1, 3, 23).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                hoja2.Range(3, 1, 3, 23).Style.Fill.BackgroundColor = XLColor.FromHtml("#BDBDBD")
                hoja2.Range(3, 1, 3, 23).Style.Font.FontColor = XLColor.FromHtml("#000000")
                hoja2.Range(2, 3, 2, 23).Style.Border.BottomBorderColor = XLColor.FromHtml("#000000")

                hoja2.Cell(3, 1).Value = "	RFC	"
                hoja2.Cell(3, 2).Value = "	Nombre	"
                hoja2.Cell(1, 3).Value = "37/001"
                hoja2.Range(2, 3, 2, 4).Merge(True)
                hoja2.Cell(2, 3).Value = "VACACIONES PROPORCIONALES"
                hoja2.Cell(3, 3).Value = "	Gravado	"
                hoja2.Cell(3, 4).Value = "	Exento	"
                hoja2.Cell(1, 5).Value = "36/001"
                hoja2.Range(2, 5, 2, 6).Merge(True)
                hoja2.Cell(2, 5).Value = "DESC. SEM OBLIGATORIO"
                hoja2.Cell(3, 5).Value = "	Gravado	"
                hoja2.Cell(3, 6).Value = "	Exento	"
                hoja2.Cell(1, 7).Value = "35/001"
                hoja2.Range(2, 7, 2, 8).Merge(True)
                hoja2.Cell(2, 7).Value = "TIEMPO EXTRA OCASIONAL"
                hoja2.Cell(3, 7).Value = "	Gravado	"
                hoja2.Cell(3, 8).Value = "   Exento    "
                hoja2.Cell(1, 9).Value = "34/001"
                hoja2.Range(2, 9, 2, 10).Merge(True)
                hoja2.Cell(2, 9).Value = "TIEMPO EXTRA FIJO"
                hoja2.Cell(3, 9).Value = "   Gravado   "
                hoja2.Cell(3, 10).Value = "  Exento    "
                hoja2.Cell(1, 11).Value = "33/001"
                hoja2.Range(2, 11, 2, 12).Merge(True)
                hoja2.Cell(2, 11).Value = "SUELDO BASE"
                hoja2.Cell(3, 11).Value = "  Gravado   "
                hoja2.Cell(3, 12).Value = "  Exento    "
                hoja2.Cell(1, 13).Value = "38/001"
                hoja2.Range(2, 13, 2, 14).Merge(True)
                hoja2.Cell(2, 13).Value = "AGUINALDO"
                hoja2.Cell(3, 13).Value = "  Gravado   "
                hoja2.Cell(3, 14).Value = "  Exento    "
                hoja2.Cell(1, 15).Value = "39/001"
                hoja2.Range(2, 15, 2, 16).Merge(True)
                hoja2.Cell(2, 15).Value = "PRIMA VACACIONAL"
                hoja2.Cell(3, 15).Value = "  Gravado   "
                hoja2.Cell(3, 16).Value = "  Exento    "
                hoja2.Cell(1, 17).Value = "40/001"
                hoja2.Range(2, 17, 2, 22).Merge(True)
                hoja2.Cell(2, 17).Value = "PRIMA DE ANTIGÜEDAD"
                hoja2.Cell(3, 17).Value = "  Gravado   "
                hoja2.Cell(3, 18).Value = "  Exento    "
                hoja2.Cell(3, 19).Value = "  Total Pagado   "
                hoja2.Cell(3, 20).Value = "  Años Servicio  "
                hoja2.Cell(3, 21).Value = "  Ult sueldo mensual ord   "
                hoja2.Cell(3, 22).Value = "  Ing Acumulable "
                hoja2.Cell(3, 23).Value = "  Ing No Acumulable   "

                ''Deducciones
                hoja3.Range(3, 1, 3, 9).Style.Font.FontSize = 10
                hoja3.Range(3, 1, 3, 9).Style.Font.SetBold(True)
                hoja3.Range(3, 1, 3, 9).Style.Alignment.WrapText = True
                hoja3.Range(3, 1, 3, 9).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                hoja3.Range(3, 1, 3, 9).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center)

                hoja3.Range(3, 1, 3, 9).Style.Fill.BackgroundColor = XLColor.FromHtml("#BDBDBD")
                hoja3.Range(3, 1, 3, 9).Style.Font.FontColor = XLColor.FromHtml("#000000")

                hoja3.Cell(3, 1).Value = "   RFC  "
                hoja3.Cell(3, 2).Value = "   Nombre    "
                hoja3.Cell(1, 3).Value = "42/001"
                hoja3.Cell(2, 3).Value = "IMSS"
                hoja3.Cell(3, 3).Value = "   Importe   "
                hoja3.Cell(1, 4).Value = "41/002"
                hoja3.Cell(2, 4).Value = "ISR"
                hoja3.Cell(3, 4).Value = "   Importe   "
                hoja3.Cell(1, 5).Value = "40/006"
                hoja3.Range(2, 5, 2, 7).Merge(True)
                hoja3.Cell(2, 5).Value = "INCAPACIDAD"
                hoja3.Cell(3, 5).Value = "   Dias Incapacidad"
                hoja3.Cell(3, 6).Value = "   Tipo "
                hoja3.Cell(3, 7).Value = "   Importe   "
                hoja3.Cell(1, 8).Value = "43/007"
                hoja3.Cell(2, 8).Value = "PENSIÓN ALIMENTICIA"
                hoja3.Cell(3, 8).Value = "   Importe   "
                hoja3.Cell(1, 8).Value = "44/010"
                hoja3.Cell(2, 8).Value = "INFONAVIT"
                hoja3.Cell(3, 9).Value = "   Importe   "

                ''Otros Pagos
                hoja4.Range(3, 1, 3, 4).Style.Font.FontSize = 10
                hoja4.Range(3, 1, 3, 4).Style.Font.SetBold(True)
                hoja4.Range(3, 1, 3, 4).Style.Alignment.WrapText = True
                hoja4.Range(3, 1, 3, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                hoja4.Range(3, 1, 3, 4).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center)


                hoja4.Range(3, 1, 3, 4).Style.Fill.BackgroundColor = XLColor.FromHtml("#BDBDBD")
                hoja4.Range(3, 1, 3, 4).Style.Font.FontColor = XLColor.FromHtml("#000000")
                hoja4.Cell(3, 1).Value = "   RFC  "
                hoja4.Cell(3, 2).Value = "   Nombre    "
                hoja4.Cell(1, 3).Value = "60/002"
                hoja4.Range(2, 3, 2, 4).Merge(True)
                hoja4.Cell(2, 3).Value = "SUBSIDIO"
                hoja4.Cell(3, 3).Value = "   Importe   "
                hoja4.Cell(3, 4).Value = "   Subsidio Causado"

                '' filaExcel = 6
                For Each dato As ListViewItem In lsvLista.CheckedItems
                    hoja.Range(2, 1, filaExcel, 1).Style.NumberFormat.Format = "@" '' "0000"
                    hoja.Range(2, 5, filaExcel, 5).Style.NumberFormat.Format = "@"
                    hoja.Range(2, 6, filaExcel, 6).Style.NumberFormat.Format = "@" ''"000000000000000000"
                    hoja.Range(2, 26, filaExcel, 26).Style.NumberFormat.Format = "@" ''"000"

                    ''hoja.Range(2, 6, filaExcel, 6).Style.NumberFormat.Format = "##################"
                    ''hoja.Range(2, 6, filaExcel, 6).Style.NumberFormat.NumberFormatId = 7

                    ''Generales
                    hoja.Cell(filaExcel, 1).Value = dato.SubItems(1).Text
                    hoja.Cell(filaExcel, 2).Value = dato.SubItems(4).Text
                    hoja.Cell(filaExcel, 3).Value = dato.SubItems(2).Text
                    hoja.Cell(filaExcel, 4).Value = dato.SubItems(5).Text
                    hoja.Cell(filaExcel, 5).Value = dato.SubItems(6).Text
                    hoja.Cell(filaExcel, 6).Value = dato.SubItems(42).Text
                    hoja.Cell(filaExcel, 7).Value = dato.SubItems(14).Text
                    hoja.Cell(filaExcel, 8).Value = dato.SubItems(13).Text
                    hoja.Cell(filaExcel, 9).Value = "A1131077105" ''dato.SubItems(8).Text 
                    hoja.Cell(filaExcel, 10).Value = "CAM" ''dato.SubItems(9).Text  
                    hoja.Cell(filaExcel, 11).Value = dato.SubItems(15).Text
                    hoja.Cell(filaExcel, 12).Value = dato.SubItems(43).Text
                    hoja.Cell(filaExcel, 13).Value = "3" ''dato.SubItems(12).Text 
                    hoja.Cell(filaExcel, 14).Value = ""  ''dato.SubItems(14).Text
                    hoja.Cell(filaExcel, 15).Value = ""  ''dato.SubItems(15).Text
                    hoja.Cell(filaExcel, 16).Value = "1"  ''dato.SubItems(16).Text
                    hoja.Cell(filaExcel, 17).Value = ""  ''dato.SubItems(17).Text
                    hoja.Cell(filaExcel, 18).Value = "2"  ''dato.SubItems(18).Text
                    hoja.Cell(filaExcel, 19).Value = ""  ''dato.SubItems(19).Text
                    hoja.Cell(filaExcel, 20).Value = ""
                    hoja.Cell(filaExcel, 21).Value = dato.SubItems(9).Text  '' dato.SubItems(21).Text
                    hoja.Cell(filaExcel, 22).Value = "4"  ''dato.SubItems(22).Text
                    hoja.Cell(filaExcel, 23).Value = ""  ''dato.SubItems(23).Text
                    hoja.Cell(filaExcel, 24).Value = "5"  ''dato.SubItems(24).Text
                    hoja.Cell(filaExcel, 25).Value = ""
                    hoja.Cell(filaExcel, 26).Value = dato.SubItems(41).Text  ''dato.SubItems(26).Text
                    hoja.Cell(filaExcel, 27).Value = ""  ''dato.SubItems(27).Text
                    hoja.Cell(filaExcel, 28).Value = "" ''dato.SubItems(28).Text
                    hoja.Cell(filaExcel, 29).Value = cboTipoR.SelectedItem.ToString() ''"NA" '' dato.SubItems(29).Text MES DE PAGO
                    hoja.Cell(filaExcel, 30).Value = cboMes.SelectedIndex + 1
                    hoja.Cell(filaExcel, 31).Value = dato.SubItems(10).Text
                    filaExcel = filaExcel + 1
                Next

                filaExcel = 4
                For Each dato As ListViewItem In lsvLista.CheckedItems

                    hoja2.Cell(filaExcel, 1).Value = dato.SubItems(4).Text
                    hoja2.Cell(filaExcel, 2).Value = dato.SubItems(2).Text
                    hoja2.Cell(filaExcel, 3).Value = dato.SubItems(23).Text
                    hoja2.Cell(filaExcel, 4).Value = ""
                    hoja2.Cell(filaExcel, 5).Value = dato.SubItems(22).Text
                    hoja2.Cell(filaExcel, 6).Value = ""
                    hoja2.Cell(filaExcel, 7).Value = dato.SubItems(21).Text
                    hoja2.Cell(filaExcel, 8).Value = ""
                    hoja2.Cell(filaExcel, 9).Value = dato.SubItems(19).Text
                    hoja2.Cell(filaExcel, 10).Value = dato.SubItems(20).Text
                    hoja2.Cell(filaExcel, 11).Value = dato.SubItems(18).Text
                    hoja2.Cell(filaExcel, 12).Value = ""
                    hoja2.Cell(filaExcel, 13).Value = dato.SubItems(24).Text
                    hoja2.Cell(filaExcel, 14).Value = dato.SubItems(25).Text
                    hoja2.Cell(filaExcel, 15).Value = dato.SubItems(27).Text
                    hoja2.Cell(filaExcel, 16).Value = dato.SubItems(28).Text
                    hoja2.Cell(filaExcel, 17).Value = ""
                    hoja2.Cell(filaExcel, 18).Value = ""
                    hoja2.Cell(filaExcel, 19).Value = ""
                    hoja2.Cell(filaExcel, 20).Value = ""
                    hoja2.Cell(filaExcel, 21).Value = ""
                    hoja2.Cell(filaExcel, 22).Value = ""
                    hoja2.Cell(filaExcel, 23).Value = ""

                    ''Deducciones
                    hoja3.Cell(filaExcel, 1).Value = dato.SubItems(4).Text
                    hoja3.Cell(filaExcel, 2).Value = dato.SubItems(2).Text
                    hoja3.Cell(filaExcel, 3).Value = dato.SubItems(34).Text
                    hoja3.Cell(filaExcel, 4).Value = dato.SubItems(33).Text
                    hoja3.Cell(filaExcel, 5).Value = ""
                    hoja3.Cell(filaExcel, 6).Value = ""
                    hoja3.Cell(filaExcel, 7).Value = dato.SubItems(32).Text
                    hoja3.Cell(filaExcel, 8).Value = dato.SubItems(36).Text
                    hoja3.Cell(filaExcel, 9).Value = dato.SubItems(35).Text


                    ''Otros Pagos
                    hoja4.Columns("A").Width = 20
                    hoja4.Columns("B").Width = 20
                    hoja4.Cell(filaExcel, 1).Value = dato.SubItems(4).Text
                    hoja4.Cell(filaExcel, 2).Value = dato.SubItems(2).Text
                    hoja4.Cell(filaExcel, 3).Value = dato.SubItems(37).Text
                    hoja4.Cell(filaExcel, 4).Value = dato.SubItems(48).Text

                    filaExcel = filaExcel + 1

                Next
                Dim moment As Date = Date.Now()
                Dim month As Integer = moment.Month
                Dim year As Integer = moment.Year
                dialogo.DefaultExt = "*.xlsx"
                dialogo.FileName = "Isla-Arca " & Format(moment.Date, "yyyy dd MMMM") & " " & cboTipoR.SelectedItem.ToString()
                dialogo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
                dialogo.ShowDialog()
                libro.SaveAs(dialogo.FileName)
                libro = Nothing

                MessageBox.Show("Archivo generado correctamente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            Else

                MessageBox.Show("Por favor seleccione al menos una registro para importar.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If

        Catch ex As Exception
            '' MessageBox.Show(ex.ToString(), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

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
        tsbGuardar2.Enabled = False
        cmdNN.Enabled = False
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

    Private Sub tsbGuardar2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbGuardar2.Click

        cargarND()


    End Sub



    Public Sub cargarND()

        Try

            Dim filaExcel As Integer = 2
            Dim dialogo As New SaveFileDialog()

            If lsvLista.CheckedItems.Count > 0 Then

                Dim libro As New ClosedXML.Excel.XLWorkbook
                Dim hoja As IXLWorksheet = libro.Worksheets.Add("Generales")
                Dim hoja2 As IXLWorksheet = libro.Worksheets.Add("Percepciones")
                Dim hoja3 As IXLWorksheet = libro.Worksheets.Add("Deducciones")
                Dim hoja4 As IXLWorksheet = libro.Worksheets.Add("Otros Pagos")

                hoja.Column("A").Width = 20
                hoja.Column("B").Width = 15
                hoja.Column("C").Width = 15
                hoja.Column("D").Width = 12
                hoja.Column("E").Width = 12
                hoja.Column("F").Width = 25
                hoja.Column("G").Width = 15
                hoja.Column("H").Width = 15
                hoja.Column("I").Width = 25
                hoja.Column("J").Width = 15
                hoja.Column("K").Width = 15
                hoja.Column("L").Width = 15
                hoja.Column("M").Width = 15
                hoja.Column("N").Width = 50
                hoja.Column("O").Width = 12

                hoja.Range(1, 1, 1, 30).Style.Font.FontSize = 10
                hoja.Range(1, 1, 1, 30).Style.Font.SetBold(True)
                hoja.Range(1, 1, 1, 30).Style.Alignment.WrapText = True
                hoja.Range(1, 1, 1, 30).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                hoja.Range(1, 1, 1, 30).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                hoja.Range(1, 1, 1, 30).Style.Fill.BackgroundColor = XLColor.FromHtml("#BDBDBD")
                hoja.Range(1, 1, 1, 30).Style.Font.FontColor = XLColor.FromHtml("#000000")


                hoja.Cell(1, 1).Value = "No. Empleado"
                hoja.Cell(1, 2).Value = "RFC"
                hoja.Cell(1, 3).Value = "Nombre"
                hoja.Cell(1, 4).Value = "CURP"
                hoja.Cell(1, 5).Value = "SSA"
                hoja.Cell(1, 6).Value = "Cuenta Bancaria"
                hoja.Cell(1, 7).Value = "SBC"
                hoja.Cell(1, 8).Value = "SDI"
                hoja.Cell(1, 9).Value = "Reg. Patronal"
                hoja.Cell(1, 10).Value = "Ent. Federativa"
                hoja.Cell(1, 11).Value = "Días Pagados"
                hoja.Cell(1, 12).Value = "FechaInicioRelLaboral"
                hoja.Cell(1, 13).Value = "TipoContrato" ''Numero
                hoja.Cell(1, 14).Value = "TipoContrato"
                hoja.Cell(1, 15).Value = "Sindicalizado"
                hoja.Cell(1, 16).Value = "TipoJornada"
                hoja.Cell(1, 17).Value = "TipoJornada"
                hoja.Cell(1, 18).Value = "Tipo Regimen" ''Numero
                hoja.Cell(1, 19).Value = "Tipo Regimen"
                hoja.Cell(1, 20).Value = "Departamento"
                hoja.Cell(1, 21).Value = "Puesto"
                hoja.Cell(1, 22).Value = "Riesgo Puesto" ''Numero
                hoja.Cell(1, 23).Value = "Riesgo Puesto"
                hoja.Cell(1, 24).Value = "Periodicidad Pago" ''Numero
                hoja.Cell(1, 25).Value = "Periodicidad Pago"
                hoja.Cell(1, 26).Value = "Banco" ''Numero
                hoja.Cell(1, 27).Value = "Banco"
                hoja.Cell(1, 28).Value = "Subcontratacion"
                hoja.Cell(1, 29).Value = "Tipo Recibo"
                hoja.Cell(1, 30).Value = "Mes Pago"
                hoja.Cell(1, 31).Value = "Buque"


                ''Percepciones
                hoja2.Range(3, 1, 3, 23).Style.Font.FontSize = 10
                hoja2.Range(3, 1, 3, 23).Style.Font.SetBold(True)
                hoja2.Range(3, 1, 3, 23).Style.Alignment.WrapText = True
                hoja2.Range(3, 1, 3, 23).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                hoja2.Range(3, 1, 3, 23).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                hoja2.Range(3, 1, 3, 23).Style.Fill.BackgroundColor = XLColor.FromHtml("#BDBDBD")
                hoja2.Range(3, 1, 3, 23).Style.Font.FontColor = XLColor.FromHtml("#000000")
                hoja2.Range(2, 3, 2, 23).Style.Border.BottomBorderColor = XLColor.FromHtml("#000000")

                hoja2.Cell(3, 1).Value = "	RFC	"
                hoja2.Cell(3, 2).Value = "	Nombre	"
                hoja2.Cell(1, 3).Value = "37/001"
                hoja2.Range(2, 3, 2, 4).Merge(True)
                hoja2.Cell(2, 3).Value = "VACACIONES PROPORCIONALES"
                hoja2.Cell(3, 3).Value = "	Gravado	"
                hoja2.Cell(3, 4).Value = "	Exento	"
                hoja2.Cell(1, 5).Value = "36/001"
                hoja2.Range(2, 5, 2, 6).Merge(True)
                hoja2.Cell(2, 5).Value = "DESC. SEM OBLIGATORIO"
                hoja2.Cell(3, 5).Value = "	Gravado	"
                hoja2.Cell(3, 6).Value = "	Exento	"
                hoja2.Cell(1, 7).Value = "35/001"
                hoja2.Range(2, 7, 2, 8).Merge(True)
                hoja2.Cell(2, 7).Value = "TIEMPO EXTRA OCASIONAL"
                hoja2.Cell(3, 7).Value = "	Gravado	"
                hoja2.Cell(3, 8).Value = "   Exento    "
                hoja2.Cell(1, 9).Value = "34/001"
                hoja2.Range(2, 9, 2, 10).Merge(True)
                hoja2.Cell(2, 9).Value = "TIEMPO EXTRA FIJO"
                hoja2.Cell(3, 9).Value = "   Gravado   "
                hoja2.Cell(3, 10).Value = "  Exento    "
                hoja2.Cell(1, 11).Value = "33/001"
                hoja2.Range(2, 11, 2, 12).Merge(True)
                hoja2.Cell(2, 11).Value = "SUELDO BASE"
                hoja2.Cell(3, 11).Value = "  Gravado   "
                hoja2.Cell(3, 12).Value = "  Exento    "
                hoja2.Cell(1, 13).Value = "38/001"
                hoja2.Range(2, 13, 2, 14).Merge(True)
                hoja2.Cell(2, 13).Value = "AGUINALDO"
                hoja2.Cell(3, 13).Value = "  Gravado   "
                hoja2.Cell(3, 14).Value = "  Exento    "
                hoja2.Cell(1, 15).Value = "39/001"
                hoja2.Range(2, 15, 2, 16).Merge(True)
                hoja2.Cell(2, 15).Value = "PRIMA VACACIONAL"
                hoja2.Cell(3, 15).Value = "  Gravado   "
                hoja2.Cell(3, 16).Value = "  Exento    "
                hoja2.Cell(1, 17).Value = "40/001"
                hoja2.Range(2, 17, 2, 22).Merge(True)
                hoja2.Cell(2, 17).Value = "PRIMA DE ANTIGÜEDAD"
                hoja2.Cell(3, 17).Value = "  Gravado   "
                hoja2.Cell(3, 18).Value = "  Exento    "
                hoja2.Cell(3, 19).Value = "  Total Pagado   "
                hoja2.Cell(3, 20).Value = "  Años Servicio  "
                hoja2.Cell(3, 21).Value = "  Ult sueldo mensual ord   "
                hoja2.Cell(3, 22).Value = "  Ing Acumulable "
                hoja2.Cell(3, 23).Value = "  Ing No Acumulable   "

                ''Deducciones
                hoja3.Range(3, 1, 3, 9).Style.Font.FontSize = 10
                hoja3.Range(3, 1, 3, 9).Style.Font.SetBold(True)
                hoja3.Range(3, 1, 3, 9).Style.Alignment.WrapText = True
                hoja3.Range(3, 1, 3, 9).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                hoja3.Range(3, 1, 3, 9).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center)

                hoja3.Range(3, 1, 3, 9).Style.Fill.BackgroundColor = XLColor.FromHtml("#BDBDBD")
                hoja3.Range(3, 1, 3, 9).Style.Font.FontColor = XLColor.FromHtml("#000000")

                hoja3.Cell(3, 1).Value = "   RFC  "
                hoja3.Cell(3, 2).Value = "   Nombre    "
                hoja3.Cell(1, 3).Value = "42/001"
                hoja3.Cell(2, 3).Value = "IMSS"
                hoja3.Cell(3, 3).Value = "   Importe   "
                hoja3.Cell(1, 4).Value = "41/002"
                hoja3.Cell(2, 4).Value = "ISR"
                hoja3.Cell(3, 4).Value = "   Importe   "
                hoja3.Cell(1, 5).Value = "40/006"
                hoja3.Range(2, 5, 2, 7).Merge(True)
                hoja3.Cell(2, 5).Value = "INCAPACIDAD"
                hoja3.Cell(3, 5).Value = "   Dias Incapacidad"
                hoja3.Cell(3, 6).Value = "   Tipo "
                hoja3.Cell(3, 7).Value = "   Importe   "
                hoja3.Cell(1, 8).Value = "43/007"
                hoja3.Cell(2, 8).Value = "PENSIÓN ALIMENTICIA"
                hoja3.Cell(3, 8).Value = "   Importe   "
                hoja3.Cell(1, 8).Value = "44/010"
                hoja3.Cell(2, 8).Value = "INFONAVIT"
                hoja3.Cell(3, 9).Value = "   Importe   "

                ''Otros Pagos
                hoja4.Range(3, 1, 3, 4).Style.Font.FontSize = 10
                hoja4.Range(3, 1, 3, 4).Style.Font.SetBold(True)
                hoja4.Range(3, 1, 3, 4).Style.Alignment.WrapText = True
                hoja4.Range(3, 1, 3, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                hoja4.Range(3, 1, 3, 4).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center)


                hoja4.Range(3, 1, 3, 4).Style.Fill.BackgroundColor = XLColor.FromHtml("#BDBDBD")
                hoja4.Range(3, 1, 3, 4).Style.Font.FontColor = XLColor.FromHtml("#000000")
                hoja4.Cell(3, 1).Value = "   RFC  "
                hoja4.Cell(3, 2).Value = "   Nombre    "
                hoja4.Cell(3, 3).Value = "   Importe   "
                hoja4.Cell(3, 4).Value = "   Subsidio Causado"

                For Each dato As ListViewItem In lsvLista.CheckedItems
                    hoja.Range(2, 1, filaExcel, 1).Style.NumberFormat.Format = "@" '' "0000"
                    hoja.Range(2, 5, filaExcel, 5).Style.NumberFormat.Format = "@"
                    hoja.Range(2, 6, filaExcel, 6).Style.NumberFormat.Format = "@" ''"000000000000000000"
                    hoja.Range(2, 26, filaExcel, 26).Style.NumberFormat.Format = "@" ''"000"



                    ''hoja.Range(2, 6, filaExcel, 6).Style.NumberFormat.NumberFormatId = 7
                    ''Generales
                    hoja.Cell(filaExcel, 1).Value = dato.SubItems(1).Text
                    hoja.Cell(filaExcel, 2).Value = dato.SubItems(4).Text
                    hoja.Cell(filaExcel, 3).Value = dato.SubItems(2).Text
                    hoja.Cell(filaExcel, 4).Value = dato.SubItems(5).Text
                    hoja.Cell(filaExcel, 5).Value = dato.SubItems(6).Text
                    hoja.Cell(filaExcel, 6).Value = dato.SubItems(42).Text
                    hoja.Cell(filaExcel, 7).Value = dato.SubItems(14).Text
                    hoja.Cell(filaExcel, 8).Value = dato.SubItems(13).Text
                    hoja.Cell(filaExcel, 9).Value = "A1131077105" ''dato.SubItems(8).Text 
                    hoja.Cell(filaExcel, 10).Value = "CAM" ''dato.SubItems(9).Text  
                    hoja.Cell(filaExcel, 11).Value = dato.SubItems(15).Text
                    hoja.Cell(filaExcel, 12).Value = dato.SubItems(43).Text
                    hoja.Cell(filaExcel, 13).Value = "3" ''dato.SubItems(12).Text 
                    hoja.Cell(filaExcel, 14).Value = ""  ''dato.SubItems(14).Text
                    hoja.Cell(filaExcel, 15).Value = ""  ''dato.SubItems(15).Text
                    hoja.Cell(filaExcel, 16).Value = "1"  ''dato.SubItems(16).Text
                    hoja.Cell(filaExcel, 17).Value = ""  ''dato.SubItems(17).Text
                    hoja.Cell(filaExcel, 18).Value = "2"  ''dato.SubItems(18).Text
                    hoja.Cell(filaExcel, 19).Value = ""  ''dato.SubItems(19).Text
                    hoja.Cell(filaExcel, 20).Value = ""
                    hoja.Cell(filaExcel, 21).Value = dato.SubItems(9).Text  '' dato.SubItems(21).Text
                    hoja.Cell(filaExcel, 22).Value = "4"  ''dato.SubItems(22).Text
                    hoja.Cell(filaExcel, 23).Value = ""  ''dato.SubItems(23).Text
                    hoja.Cell(filaExcel, 24).Value = "5"  ''dato.SubItems(24).Text
                    hoja.Cell(filaExcel, 25).Value = ""
                    hoja.Cell(filaExcel, 26).Value = dato.SubItems(41).Text  ''dato.SubItems(26).Text
                    hoja.Cell(filaExcel, 27).Value = ""  ''dato.SubItems(27).Text
                    hoja.Cell(filaExcel, 28).Value = "" ''dato.SubItems(28).Text
                    hoja.Cell(filaExcel, 29).Value = cboTipoR.SelectedItem.ToString() ''"NA" 
                    hoja.Cell(filaExcel, 30).Value = cboMes.SelectedIndex + 1
                    hoja.Cell(filaExcel, 31).Value = dato.SubItems(10).Text

                    filaExcel = filaExcel + 1
                Next

                filaExcel = 4
                For Each dato As ListViewItem In lsvLista.CheckedItems

                    hoja2.Cell(filaExcel, 1).Value = dato.SubItems(4).Text
                    hoja2.Cell(filaExcel, 2).Value = dato.SubItems(2).Text
                    hoja2.Cell(filaExcel, 3).Value = dato.SubItems(23).Text
                    hoja2.Cell(filaExcel, 4).Value = ""
                    hoja2.Cell(filaExcel, 5).Value = dato.SubItems(22).Text
                    hoja2.Cell(filaExcel, 6).Value = ""
                    hoja2.Cell(filaExcel, 7).Value = dato.SubItems(21).Text
                    hoja2.Cell(filaExcel, 8).Value = ""
                    hoja2.Cell(filaExcel, 9).Value = dato.SubItems(19).Text
                    hoja2.Cell(filaExcel, 10).Value = dato.SubItems(20).Text
                    hoja2.Cell(filaExcel, 11).Value = dato.SubItems(18).Text
                    hoja2.Cell(filaExcel, 12).Value = ""
                    hoja2.Cell(filaExcel, 13).Value = dato.SubItems(24).Text
                    hoja2.Cell(filaExcel, 14).Value = dato.SubItems(25).Text
                    hoja2.Cell(filaExcel, 15).Value = dato.SubItems(27).Text
                    hoja2.Cell(filaExcel, 16).Value = dato.SubItems(28).Text
                    hoja2.Cell(filaExcel, 17).Value = ""
                    hoja2.Cell(filaExcel, 18).Value = ""
                    hoja2.Cell(filaExcel, 19).Value = ""
                    hoja2.Cell(filaExcel, 20).Value = ""
                    hoja2.Cell(filaExcel, 21).Value = ""
                    hoja2.Cell(filaExcel, 22).Value = ""
                    hoja2.Cell(filaExcel, 23).Value = ""

                    ''Deducciones
                    hoja3.Cell(filaExcel, 1).Value = dato.SubItems(4).Text
                    hoja3.Cell(filaExcel, 2).Value = dato.SubItems(2).Text
                    hoja3.Cell(filaExcel, 3).Value = dato.SubItems(34).Text
                    hoja3.Cell(filaExcel, 4).Value = dato.SubItems(33).Text
                    hoja3.Cell(filaExcel, 5).Value = ""
                    hoja3.Cell(filaExcel, 6).Value = ""
                    hoja3.Cell(filaExcel, 7).Value = dato.SubItems(32).Text
                    hoja3.Cell(filaExcel, 8).Value = dato.SubItems(36).Text
                    hoja3.Cell(filaExcel, 9).Value = dato.SubItems(35).Text


                    ''Otros Pagos
                    hoja4.Columns("A").Width = 20
                    hoja4.Columns("B").Width = 20
                    hoja4.Cell(filaExcel, 1).Value = dato.SubItems(4).Text
                    hoja4.Cell(filaExcel, 2).Value = dato.SubItems(2).Text
                    hoja4.Cell(filaExcel, 3).Value = dato.SubItems(37).Text
                    hoja4.Cell(filaExcel, 4).Value = dato.SubItems(48).Text

                    filaExcel = filaExcel + 1

                Next
                Dim moment As Date = Date.Now()
                Dim month As Integer = moment.Month.ToString()
                Dim year As Integer = moment.Year
                dialogo.DefaultExt = "*.xlsx"
                ''  dialogo.FileName = "Isla-Arca " & Format(moment.Date, "yyyy dd MMMM") & " ND"
                dialogo.FileName = "Isla-Arca " & Format(moment.Date, "yyyy dd MMMM") & cboTipoR.SelectedItem.ToString()
                dialogo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
                dialogo.ShowDialog()
                libro.SaveAs(dialogo.FileName)
                libro = Nothing

                MessageBox.Show("Archivo generado correctamente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            Else

                MessageBox.Show("Por favor seleccione al menos una registro para importar.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString(), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Try
    End Sub

    Private Sub cmdVerificar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdVerificar.Click

        recorrerLista()
        ''recorrerLista()

    End Sub


    Public Sub recorrerLista()

        Dim filas, filas2 As Integer
        Dim contador As Integer = 0

        For filas = 1 To lsvLista.Items.Count - 1
            For filas2 = 1 + filas To lsvLista.Items.Count - 1
                ''MsgBox(lsvLista.Items.Item(filas).SubItems(1).Text)

                If lsvLista.Items(filas).SubItems(1).Text = lsvLista.Items(filas2).SubItems(1).Text Then
                    lsvLista.Items(filas2).BackColor = Color.GreenYellow

                    contador = contador + 1
                End If

                If filas2 = lsvLista.Items.Count Then
                    Exit For
                End If

            Next
            If filas = lsvLista.Items.Count Then

                Exit Sub

            End If
        Next
        MsgBox(contador.ToString & " Datos repetidos")


    End Sub
    Public Sub recorrerLista2()
        Try


            Dim lsvDate As New ListView
            Dim lsvDate2 As ListViewItem '' = lsvDate.SelectedItems(0)

            If lsvLista.CheckedItems.Count > 0 Then
                For Each dato As ListViewItem In lsvLista.CheckedItems
                    lsvDate.Items.Add(dato.SubItems(1).Text)
                    '' MsgBox(dato.SubItems(1).Text)
                Next
            End If


            Dim filas, filas2 As Integer
            Dim contador As Integer = 0

            For filas = 1 To lsvDate.Items.Count - 1
                For filas2 = 1 + filas To lsvDate.Items.Count - 1
                    ''MsgBox(lsvDate.Items(filas2).Text)

                    If lsvDate.Items(filas).Text = lsvDate.Items(filas2).Text Then
                        lsvLista.Items(filas2).BackColor = Color.GreenYellow
                        lsvLista.Items(filas).BackColor = Color.Yellow
                        lsvDate2 = lsvDate.Items.Add(filas2)
                        '' lsvLista.Items.Add(filas2)
                        contador = contador + 1
                    End If

                    If filas2 = lsvDate.Items.Count Then
                        Exit For
                    End If

                Next
                If filas = lsvDate.Items.Count Then

                    Exit Sub

                End If
            Next

            MsgBox(contador.ToString & " Datos repetidos")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    'Private Sub cmdBuscar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBuscar.Click

    '    Dim item As ListViewItem = lsvLista.FindItemWithText(txtBuscar.Text)

    '    Dim i As Integer = lsvLista.Items.IndexOf(item).ToString() + 1


    '    If lsvLista IsNot Nothing Then
    '        lsvLista.Items(i - 1).BackColor = Color.Coral
    '    Else
    '        MessageBox.Show("No tiene datos la Lista")
    '    End If


    'End Sub



    Private Sub cmdNN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNN.Click

        Try

            Dim filaExcel As Integer = 2
            Dim dialogo As New SaveFileDialog()

            If lsvLista.CheckedItems.Count > 0 Then

                'Abrimos el machote
                Dim ruta As String
                ruta = My.Application.Info.DirectoryPath() & "\Archivos\Machote _Marinos.xlsx"

                Dim book As New ClosedXML.Excel.XLWorkbook(ruta)


                Dim libro As New ClosedXML.Excel.XLWorkbook


                book.Worksheet(1).CopyTo(libro, "Generales")
                book.Worksheet(2).CopyTo(libro, "Percepciones")
                book.Worksheet(3).CopyTo(libro, "Deducciones")
                book.Worksheet(4).CopyTo(libro, "Otros Pagos")


                Dim hoja As IXLWorksheet = libro.Worksheets(0)
                Dim hoja2 As IXLWorksheet = libro.Worksheets(1)
                Dim hoja3 As IXLWorksheet = libro.Worksheets(2)
                Dim hoja4 As IXLWorksheet = libro.Worksheets(3)

                'hoja.Column("A").Width = 20
                'hoja.Column("B").Width = 15
                'hoja.Column("C").Width = 15
                'hoja.Column("D").Width = 12
                'hoja.Column("E").Width = 12
                'hoja.Column("F").Width = 25
                'hoja.Column("G").Width = 15
                'hoja.Column("H").Width = 15
                'hoja.Column("I").Width = 25
                'hoja.Column("J").Width = 15
                'hoja.Column("K").Width = 15
                'hoja.Column("L").Width = 15
                'hoja.Column("M").Width = 15
                'hoja.Column("N").Width = 50
                'hoja.Column("O").Width = 12

                'hoja.Range(1, 1, 1, 31).Style.Font.FontSize = 10
                'hoja.Range(1, 1, 1, 31).Style.Font.SetBold(True)
                'hoja.Range(1, 1, 1, 31).Style.Alignment.WrapText = True
                'hoja.Range(1, 1, 1, 31).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                'hoja.Range(1, 1, 1, 31).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                'hoja.Range(1, 1, 1, 31).Style.Fill.BackgroundColor = XLColor.FromHtml("#BDBDBD")
                'hoja.Range(1, 1, 1, 31).Style.Font.FontColor = XLColor.FromHtml("#000000")


                'hoja.Cell(1, 1).Value = "No. Empleado"
                'hoja.Cell(1, 2).Value = "RFC"
                'hoja.Cell(1, 3).Value = "Nombre"
                'hoja.Cell(1, 4).Value = "CURP"
                'hoja.Cell(1, 5).Value = "SSA"
                'hoja.Cell(1, 6).Value = "Cuenta Bancaria"
                'hoja.Cell(1, 7).Value = "SBC"
                'hoja.Cell(1, 8).Value = "SDI"
                'hoja.Cell(1, 9).Value = "Reg. Patronal"
                'hoja.Cell(1, 10).Value = "Ent. Federativa"
                'hoja.Cell(1, 11).Value = "Días Pagados"
                'hoja.Cell(1, 12).Value = "FechaInicioRelLaboral"
                'hoja.Cell(1, 13).Value = "TipoContrato" ''Numero
                'hoja.Cell(1, 14).Value = "TipoContrato"
                'hoja.Cell(1, 15).Value = "Sindicalizado"
                'hoja.Cell(1, 16).Value = "TipoJornada"
                'hoja.Cell(1, 17).Value = "TipoJornada"
                'hoja.Cell(1, 18).Value = "Tipo Regimen" ''Numero
                'hoja.Cell(1, 19).Value = "Tipo Regimen"
                'hoja.Cell(1, 20).Value = "Departamento"
                'hoja.Cell(1, 21).Value = "Puesto"
                'hoja.Cell(1, 22).Value = "Riesgo Puesto" ''Numero
                'hoja.Cell(1, 23).Value = "Riesgo Puesto"
                'hoja.Cell(1, 24).Value = "Periodicidad Pago" ''Numero
                'hoja.Cell(1, 25).Value = "Periodicidad Pago"
                'hoja.Cell(1, 26).Value = "Banco" ''Numero
                'hoja.Cell(1, 27).Value = "Banco"
                'hoja.Cell(1, 28).Value = "Subcontratacion"
                'hoja.Cell(1, 29).Value = "Tipo Recibo"
                'hoja.Cell(1, 30).Value = "Mes Pago"
                'hoja.Cell(1, 31).Value = "Buque"


                ' ''Percepciones
                'hoja2.Range(3, 1, 3, 16).Style.Font.FontSize = 10
                'hoja2.Range(3, 1, 3, 16).Style.Font.SetBold(True)
                'hoja2.Range(3, 1, 3, 16).Style.Alignment.WrapText = True
                'hoja2.Range(3, 1, 3, 16).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                'hoja2.Range(3, 1, 3, 16).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                'hoja2.Range(3, 1, 3, 16).Style.Fill.BackgroundColor = XLColor.FromHtml("#BDBDBD")
                'hoja2.Range(3, 1, 3, 16).Style.Font.FontColor = XLColor.FromHtml("#000000")
                'hoja2.Range(2, 3, 2, 16).Style.Border.BottomBorderColor = XLColor.FromHtml("#000000")

                'hoja2.Cell(3, 1).Value = "	RFC	"
                'hoja2.Cell(3, 2).Value = "	Nombre	"
                'hoja2.Cell(1, 3).Value = "15/001"
                'hoja2.Range(2, 3, 2, 4).Merge(True)
                'hoja2.Cell(2, 3).Value = "SUELDO BASE"
                'hoja2.Cell(3, 3).Value = "  Gravado   "
                'hoja2.Cell(3, 4).Value = "  Exento    "

                'hoja2.Cell(1, 5).Value = "16/001"
                'hoja2.Range(2, 5, 2, 6).Merge(True)
                'hoja2.Cell(2, 5).Value = "TIEMPO EXTRA FIJO"
                'hoja2.Cell(3, 5).Value = "   Gravado   "
                'hoja2.Cell(3, 6).Value = "  Exento    "

                'hoja2.Cell(1, 7).Value = "17/001"
                'hoja2.Range(2, 7, 2, 8).Merge(True)
                'hoja2.Cell(2, 7).Value = "TIEMPO EXTRA OCASIONAL"
                'hoja2.Cell(3, 7).Value = "	Gravado	"
                'hoja2.Cell(3, 8).Value = "   Exento    "

                'hoja2.Cell(1, 9).Value = "18/001"
                'hoja2.Range(2, 9, 2, 10).Merge(True)
                'hoja2.Cell(2, 9).Value = "DESC. SEM OBLIGATORIO"
                'hoja2.Cell(3, 9).Value = "	Gravado	"
                'hoja2.Cell(3, 10).Value = "	Exento	"

                'hoja2.Cell(1, 11).Value = "19/001"
                'hoja2.Range(2, 11, 2, 12).Merge(True)
                'hoja2.Cell(2, 11).Value = "VACACIONES PROPORCIONALES"
                'hoja2.Cell(3, 11).Value = "	Gravado	"
                'hoja2.Cell(3, 12).Value = "	Exento	"
                'hoja2.Cell(1, 13).Value = "20/002"
                'hoja2.Range(2, 13, 2, 14).Merge(True)
                'hoja2.Cell(2, 13).Value = "AGUINALDO"
                'hoja2.Cell(3, 13).Value = "  Gravado   "
                'hoja2.Cell(3, 14).Value = "  Exento    "
                'hoja2.Cell(1, 15).Value = "21/021"
                'hoja2.Range(2, 15, 2, 16).Merge(True)
                'hoja2.Cell(2, 15).Value = "PRIMA VACACIONAL"
                'hoja2.Cell(3, 15).Value = "  Gravado   "
                'hoja2.Cell(3, 16).Value = "  Exento    "


                ' ''Deducciones
                'hoja3.Range(3, 1, 3, 12).Style.Font.FontSize = 10
                'hoja3.Range(3, 1, 3, 12).Style.Font.SetBold(True)
                'hoja3.Range(3, 1, 3, 12).Style.Alignment.WrapText = True
                'hoja3.Range(3, 1, 3, 12).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                'hoja3.Range(3, 1, 3, 12).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center)

                'hoja3.Range(3, 1, 3, 12).Style.Fill.BackgroundColor = XLColor.FromHtml("#BDBDBD")
                'hoja3.Range(3, 1, 3, 12).Style.Font.FontColor = XLColor.FromHtml("#000000")

                'hoja3.Cell(3, 1).Value = "   RFC  "
                'hoja3.Cell(3, 2).Value = "   Nombre    "
                'hoja3.Cell(1, 3).Value = "24/001"
                'hoja3.Cell(2, 3).Value = "IMSS"
                'hoja3.Cell(3, 3).Value = "   Importe   "
                'hoja3.Cell(1, 4).Value = "23/002"
                'hoja3.Cell(2, 4).Value = "ISR"
                'hoja3.Cell(3, 4).Value = "   Importe   "

                'hoja3.Cell(1, 5).Value = "46/004"
                'hoja3.Cell(2, 5).Value = "PRESTAMO"
                'hoja3.Cell(3, 5).Value = "Importe"

                'hoja3.Cell(1, 6).Value = "22/006"
                'hoja3.Range(2, 6, 2, 8).Merge(True)
                'hoja3.Cell(2, 6).Value = "INCAPACIDAD"
                'hoja3.Cell(3, 6).Value = "   Dias Incapacidad"
                'hoja3.Cell(3, 7).Value = "   Tipo "
                'hoja3.Cell(3, 8).Value = "   Importe   "

                'hoja3.Cell(1, 9).Value = "25/007"
                'hoja3.Cell(2, 9).Value = "PENSIÓN ALIMENTICIA"
                'hoja3.Cell(3, 9).Value = "   Importe   "

                'hoja3.Cell(1, 10).Value = "61/010"
                'hoja3.Cell(2, 10).Value = "INFONAVIT MES ANTERIOR"
                'hoja3.Cell(3, 10).Value = "   Importe   "

                'hoja3.Cell(1, 11).Value = "26/010"
                'hoja3.Cell(2, 11).Value = "INFONAVIT"
                'hoja3.Cell(3, 11).Value = "   Importe   "

                'hoja3.Cell(1, 12).Value = "58/011"
                'hoja3.Cell(2, 12).Value = "FONACOT"
                'hoja3.Cell(3, 12).Value = "Importe"

                ''Otros Pagos
                'hoja4.Range(3, 1, 3, 4).Style.Font.FontSize = 10
                'hoja4.Range(3, 1, 3, 4).Style.Font.SetBold(True)
                'hoja4.Range(3, 1, 3, 4).Style.Alignment.WrapText = True
                'hoja4.Range(3, 1, 3, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                'hoja4.Range(3, 1, 3, 4).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center)


                'hoja4.Range(3, 1, 3, 4).Style.Fill.BackgroundColor = XLColor.FromHtml("#BDBDBD")
                'hoja4.Range(3, 1, 3, 4).Style.Font.FontColor = XLColor.FromHtml("#000000")
                'hoja4.Cell(3, 1).Value = "   RFC  "
                'hoja4.Cell(3, 2).Value = "   Nombre    "
                'hoja4.Cell(3, 3).Value = "   Importe   "
                'hoja4.Cell(3, 4).Value = "   Subsidio Causado"

                For Each dato As ListViewItem In lsvLista.CheckedItems
                    hoja.Range(2, 1, filaExcel, 1).Style.NumberFormat.Format = "@" '' "0000"
                    hoja.Range(2, 5, filaExcel, 5).Style.NumberFormat.Format = "@"
                    hoja.Range(2, 6, filaExcel, 6).Style.NumberFormat.Format = "@" '' "000000000000000000"
                    hoja.Range(2, 26, filaExcel, 26).Style.NumberFormat.Format = "@" ''"000"



                    ''hoja.Range(2, 6, filaExcel, 6).Style.NumberFormat.NumberFormatId = 7
                    ''Generales
                    hoja.Cell(filaExcel, 1).Value = dato.SubItems(1).Text
                    hoja.Cell(filaExcel, 2).Value = dato.SubItems(4).Text
                    hoja.Cell(filaExcel, 3).Value = dato.SubItems(2).Text
                    hoja.Cell(filaExcel, 4).Value = dato.SubItems(5).Text
                    hoja.Cell(filaExcel, 5).Value = dato.SubItems(6).Text
                    hoja.Cell(filaExcel, 6).Value = dato.SubItems(44).Text
                    hoja.Cell(filaExcel, 7).Value = dato.SubItems(14).Text
                    hoja.Cell(filaExcel, 8).Value = dato.SubItems(13).Text
                    hoja.Cell(filaExcel, 9).Value = "G0666980109" ''dato.SubItems(8).Text 
                    hoja.Cell(filaExcel, 10).Value = "VER" ''dato.SubItems(9).Text  
                    hoja.Cell(filaExcel, 11).Value = dato.SubItems(15).Text
                    hoja.Cell(filaExcel, 12).Value = dato.SubItems(45).Text
                    hoja.Cell(filaExcel, 13).Value = "3" ''dato.SubItems(12).Text 
                    hoja.Cell(filaExcel, 14).Value = ""  ''dato.SubItems(14).Text
                    hoja.Cell(filaExcel, 15).Value = ""  ''dato.SubItems(15).Text
                    hoja.Cell(filaExcel, 16).Value = "1"  ''dato.SubItems(16).Text
                    hoja.Cell(filaExcel, 17).Value = ""  ''dato.SubItems(17).Text
                    hoja.Cell(filaExcel, 18).Value = "2"  ''dato.SubItems(18).Text
                    hoja.Cell(filaExcel, 19).Value = ""  ''dato.SubItems(19).Text
                    hoja.Cell(filaExcel, 20).Value = ""
                    hoja.Cell(filaExcel, 21).Value = dato.SubItems(9).Text  '' dato.SubItems(21).Text
                    hoja.Cell(filaExcel, 22).Value = "4"  ''dato.SubItems(22).Text
                    hoja.Cell(filaExcel, 23).Value = ""  ''dato.SubItems(23).Text
                    hoja.Cell(filaExcel, 24).Value = "5"  ''dato.SubItems(24).Text
                    hoja.Cell(filaExcel, 25).Value = ""
                    hoja.Cell(filaExcel, 26).Value = dato.SubItems(43).Text  ''dato.SubItems(26).Text
                    hoja.Cell(filaExcel, 27).Value = ""  ''dato.SubItems(27).Text
                    hoja.Cell(filaExcel, 28).Value = "" ''dato.SubItems(28).Text
                    hoja.Cell(filaExcel, 29).Value = cboTipoR.SelectedItem.ToString() '' dato.SubItems(29).Text MES DE PAGO
                    hoja.Cell(filaExcel, 30).Value = cboMes.SelectedIndex + 1
                    hoja.Cell(filaExcel, 31).Value = dato.SubItems(10).Text

                    filaExcel = filaExcel + 1
                Next

                filaExcel = 4
                For Each dato As ListViewItem In lsvLista.CheckedItems
                    ''Percepciones
                    hoja2.Cell(filaExcel, 1).Value = dato.SubItems(4).Text
                    hoja2.Cell(filaExcel, 2).Value = dato.SubItems(2).Text
                    hoja2.Cell(filaExcel, 3).Value = dato.SubItems(18).Text
                    hoja2.Cell(filaExcel, 4).Value = ""
                    hoja2.Cell(filaExcel, 5).Value = dato.SubItems(19).Text
                    hoja2.Cell(filaExcel, 6).Value = dato.SubItems(20).Text
                    hoja2.Cell(filaExcel, 7).Value = dato.SubItems(21).Text
                    hoja2.Cell(filaExcel, 8).Value = ""
                    hoja2.Cell(filaExcel, 9).Value = dato.SubItems(22).Text
                    hoja2.Cell(filaExcel, 10).Value = ""
                    hoja2.Cell(filaExcel, 11).Value = dato.SubItems(23).Text
                    hoja2.Cell(filaExcel, 12).Value = ""
                    hoja2.Cell(filaExcel, 13).Value = dato.SubItems(24).Text
                    hoja2.Cell(filaExcel, 14).Value = dato.SubItems(25).Text
                    hoja2.Cell(filaExcel, 15).Value = dato.SubItems(27).Text
                    hoja2.Cell(filaExcel, 16).Value = dato.SubItems(28).Text
                    'hoja2.Cell(filaExcel, 17).Value = ""
                    'hoja2.Cell(filaExcel, 18).Value = ""
                    'hoja2.Cell(filaExcel, 19).Value = ""
                    'hoja2.Cell(filaExcel, 20).Value = ""
                    'hoja2.Cell(filaExcel, 21).Value = ""
                    'hoja2.Cell(filaExcel, 22).Value = ""
                    'hoja2.Cell(filaExcel, 23).Value = ""

                    ''Deducciones
                    hoja3.Cell(filaExcel, 1).Value = dato.SubItems(4).Text
                    hoja3.Cell(filaExcel, 2).Value = dato.SubItems(2).Text
                    hoja3.Cell(filaExcel, 3).Value = dato.SubItems(34).Text
                    hoja3.Cell(filaExcel, 4).Value = dato.SubItems(33).Text
                    hoja3.Cell(filaExcel, 5).Value = dato.SubItems(38).Text
                    hoja3.Cell(filaExcel, 6).Value = ""
                    hoja3.Cell(filaExcel, 7).Value = ""
                    hoja3.Cell(filaExcel, 8).Value = dato.SubItems(32).Text
                    hoja3.Cell(filaExcel, 9).Value = dato.SubItems(37).Text
                    hoja3.Cell(filaExcel, 10).Value = dato.SubItems(36).Text
                    hoja3.Cell(filaExcel, 11).Value = dato.SubItems(35).Text
                    hoja3.Cell(filaExcel, 12).Value = dato.SubItems(39).Text


                    ''Otros Pagos
                    'hoja4.Columns("A").Width = 20
                    'hoja4.Columns("B").Width = 20
                    'hoja4.Cell(filaExcel, 1).Value = dato.SubItems(4).Text
                    'hoja4.Cell(filaExcel, 2).Value = dato.SubItems(2).Text
                    'hoja4.Cell(filaExcel, 3).Value = dato.SubItems(37).Text
                    'hoja4.Cell(filaExcel, 4).Value = dato.SubItems(48).Text

                    filaExcel = filaExcel + 1

                Next
                Dim moment As Date = Date.Now()
                Dim month As Integer = moment.Month.ToString()
                Dim year As Integer = moment.Year
                dialogo.DefaultExt = "*.xlsx"
                'dialogo.FileName = Format(moment.Date, "yyyy dd MMMM") & " "
                dialogo.FileName = Format(Date.Now, "MMMM yyyy") & " " & cboTipoR.SelectedItem.ToString()
                dialogo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
                dialogo.ShowDialog()
                libro.SaveAs(dialogo.FileName)
                libro = Nothing

                MessageBox.Show("Archivo generado correctamente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            Else

                MessageBox.Show("Por favor seleccione al menos una registro para importar.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString(), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Try

    End Sub
End Class