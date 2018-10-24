Imports ClosedXML.Excel

Public Class frmnominasproceso
    Private m_currentControl As Control = Nothing
    Public gIdEmpresa As String
    Public gIdTipoPeriodo As String
    Public gNombrePeriodo As String
    Dim Ruta As String
    Dim nombre As String
    Dim cargado As Boolean = False
    Dim diasperiodo As Integer
    Dim aniocostosocial As Integer
    Dim dgvCombo As DataGridViewComboBoxEditingControl
    Dim campoordenamiento As String
    Dim TipoNomina As Boolean
    Dim IDCalculoInfonavit As Integer
    Dim FechaInicioPeriodoGlobal As Date

    Private Sub dvgCombo_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        Try
            '
            ' se recupera el valor del combo
            ' a modo de ejemplo se escribe en consola el valor seleccionado
            '



            Dim combo As ComboBox = TryCast(sender, ComboBox)

            If dgvCombo IsNot Nothing Then
                Dim sql As String
                'Console.WriteLine(combo.SelectedValue)
                'MessageBox.Show(combo.Text)
                '
                ' se accede a la fila actual, para trabajr con otor de sus campos
                ' en este caso se marca el check si se cambia la seleccion
                '
                Dim row As DataGridViewRow = dtgDatos.CurrentRow

                'Dim cell As DataGridViewCheckBoxCell = TryCast(row.Cells("Seleccionado"), DataGridViewCheckBoxCell)
                'cell.Value = True

                'Poner los datos necesarios para poner el nuevo sueldo diario y el integrado


                sql = "Select salariod,sbc,salariodTopado,sbcTopado from costosocial "
                sql &= " where fkiIdPuesto = " & combo.SelectedValue & " and anio=" & aniocostosocial

                Dim rwDatosSalario As DataRow() = nConsulta(sql)

                If rwDatosSalario Is Nothing = False Then
                    If row.Cells(10).Value >= 55 Then
                        row.Cells(16).Value = rwDatosSalario(0)("salariodTopado")
                        row.Cells(17).Value = rwDatosSalario(0)("sbcTopado")
                    Else
                        row.Cells(16).Value = rwDatosSalario(0)("salariod")
                        row.Cells(17).Value = rwDatosSalario(0)("sbc")
                    End If

                Else
                    MessageBox.Show("No se encontraron datos")
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try




    End Sub

    Private Sub tsbIEmpleados_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbIEmpleados.Click
        Try
            Dim Forma As New frmEmpleados
            Forma.gIdEmpresa = gIdEmpresa
            Forma.gIdPeriodo = cboperiodo.SelectedValue
            Forma.gIdTipoPuesto = 1
            Forma.ShowDialog()
        Catch ex As Exception

        End Try
    End Sub

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub frmnominasproceso_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim sql As String
            cargarperiodos()
            Me.dtgDatos.ContextMenuStrip = Me.cMenu
            cboserie.SelectedIndex = 0
            cboTipoNomina.SelectedIndex = 0
            sql = "select * from periodos where iIdPeriodo= " & cboperiodo.SelectedValue
            Dim rwPeriodo As DataRow() = nConsulta(sql)
            If rwPeriodo Is Nothing = False Then

                aniocostosocial = Date.Parse(rwPeriodo(0)("dFechaInicio").ToString).Year

            End If

            campoordenamiento = "Nomina.Buque,cNombreLargo"
            TipoNomina = False


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cargarbancosasociados()
        Dim sql As String
        Try
            sql = "select * from bancos inner join ( select distinct(fkiidBanco) from DatosBanco where fkiIdEmpresa=" & gIdEmpresa & ") bancos2 on bancos.iIdBanco=bancos2.fkiidBanco order by cBanco"
            nCargaCBO(cbobancos, sql, "cBanco", "iIdBanco")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub cargarperiodos()
        'Verificar si se tienen permisos
        Dim sql As String
        Try
            sql = "Select (CONVERT(nvarchar(12),dFechaInicio,103) + ' - ' + CONVERT(nvarchar(12),dFechaFin,103)) as dFechaInicio,iIdPeriodo  from periodos order by iEjercicio,iNumeroPeriodo"
            nCargaCBO(cboperiodo, sql, "dFechainicio", "iIdPeriodo")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub cmdverdatos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdverdatos.Click
        Try
            'If cargado Then



            '    dtgDatos.DataSource = Nothing
            '    llenargrid()
            'Else
            '    cargado = True
            '    llenargrid()
            'End If
            If dtgDatos.RowCount > 0 Then
                Dim resultado As Integer = MessageBox.Show("ya se tienen empleados cargados en la lista, si continua estos se borraran,¿Desea continuar?", "Pregunta", MessageBoxButtons.YesNo)
                If resultado = DialogResult.Yes Then

                    dtgDatos.Columns.Clear()
                    llenargrid()

                End If
            Else
                dtgDatos.Columns.Clear()
                llenargrid()

            End If




        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub llenargrid()
        'Cargar grid
        Try
            Dim sql As String
            Dim sql2 As String
            Dim infonavit As Double
            Dim prestamo As Double
            Dim incidencia As Double
            Dim bCalcular As Boolean
            Dim PrimaSA As Double
            Dim cadenabanco As String
            dtgDatos.Columns.Clear()
            dtgDatos.DataSource = Nothing


            dtgDatos.DefaultCellStyle.Font = New Font("Calibri", 8)
            dtgDatos.ColumnHeadersDefaultCellStyle.Font = New Font("Calibri", 9)
            Dim chk As New DataGridViewCheckBoxColumn()
            dtgDatos.Columns.Add(chk)
            chk.HeaderText = ""
            chk.Name = "chk"
            'dtgDatos.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

            'dtgDatos.Columns("chk").SortMode = DataGridViewColumnSortMode.NotSortable

            'dtgDatos.Columns.Add("idempleado", "idempleado")
            'dtgDatos.Columns(0).Width = 30
            'dtgDatos.Columns(0).ReadOnly = True
            ''dtgDatos.Columns(0).DataPropertyName("idempleado")

            'dtgDatos.Columns.Add("departamento", "Departamento")
            'dtgDatos.Columns(1).Width = 100
            'dtgDatos.Columns(1).ReadOnly = True
            'dtgDatos.Columns.Add("nombre", "Trabajador")
            'dtgDatos.Columns(2).Width = 250
            'dtgDatos.Columns(2).ReadOnly = True
            'dtgDatos.Columns.Add("sueldo", "Sueldo Ordinario")
            'dtgDatos.Columns(3).Width = 75
            'dtgDatos.Columns.Add("neto", "Neto")
            'dtgDatos.Columns(4).Width = 75
            'dtgDatos.Columns.Add("infonavit", "Infonavit")
            'dtgDatos.Columns(5).Width = 75
            'dtgDatos.Columns.Add("descuento", "Descuento")
            'dtgDatos.Columns(6).Width = 75
            'dtgDatos.Columns.Add("prestamo", "Prestamo")
            'dtgDatos.Columns(7).Width = 75
            'dtgDatos.Columns.Add("sindicato", "Sindicato")
            'dtgDatos.Columns(8).Width = 75
            'dtgDatos.Columns.Add("neto", "Sueldo Neto")
            'dtgDatos.Columns(9).Width = 75
            'dtgDatos.Columns.Add("imss", "Retención IMSS")
            'dtgDatos.Columns(10).Width = 75
            'dtgDatos.Columns.Add("subsidiado", "Retenciones")
            'dtgDatos.Columns(11).Width = 75
            'dtgDatos.Columns.Add("costosocial", "Costo Social")
            'dtgDatos.Columns(12).Width = 75
            'dtgDatos.Columns.Add("comision", "Comisión")
            'dtgDatos.Columns(13).Width = 75
            'dtgDatos.Columns.Add("subtotal", "Subtotal")
            'dtgDatos.Columns(14).Width = 75
            'dtgDatos.Columns.Add("iva", "IVA")
            'dtgDatos.Columns(15).Width = 75
            'dtgDatos.Columns.Add("total", "Total")
            'dtgDatos.Columns(16).Width = 75


            Dim dsPeriodo As New DataSet
            dsPeriodo.Tables.Add("Tabla")
            dsPeriodo.Tables("Tabla").Columns.Add("Consecutivo")
            dsPeriodo.Tables("Tabla").Columns.Add("Id_empleado")
            dsPeriodo.Tables("Tabla").Columns.Add("CodigoEmpleado")
            dsPeriodo.Tables("Tabla").Columns.Add("Nombre")
            dsPeriodo.Tables("Tabla").Columns.Add("Status")
            dsPeriodo.Tables("Tabla").Columns.Add("RFC")
            dsPeriodo.Tables("Tabla").Columns.Add("CURP")
            dsPeriodo.Tables("Tabla").Columns.Add("Num_IMSS")
            dsPeriodo.Tables("Tabla").Columns.Add("Fecha_Nac")
            dsPeriodo.Tables("Tabla").Columns.Add("Edad")
            dsPeriodo.Tables("Tabla").Columns.Add("Puesto")
            dsPeriodo.Tables("Tabla").Columns.Add("Buque")
            dsPeriodo.Tables("Tabla").Columns.Add("Tipo_Infonavit")
            dsPeriodo.Tables("Tabla").Columns.Add("Valor_Infonavit")
            dsPeriodo.Tables("Tabla").Columns.Add("Sueldo_Base")
            dsPeriodo.Tables("Tabla").Columns.Add("Salario_Diario")
            dsPeriodo.Tables("Tabla").Columns.Add("Salario_Cotización")
            dsPeriodo.Tables("Tabla").Columns.Add("Dias_Trabajados")
            dsPeriodo.Tables("Tabla").Columns.Add("Tipo_Incapacidad")
            dsPeriodo.Tables("Tabla").Columns.Add("Número_días")
            dsPeriodo.Tables("Tabla").Columns.Add("Sueldo_Bruto")
            dsPeriodo.Tables("Tabla").Columns.Add("Aguinaldo_gravado")
            dsPeriodo.Tables("Tabla").Columns.Add("Aguinaldo_exento")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_Aguinaldo")
            dsPeriodo.Tables("Tabla").Columns.Add("Prima_vac_gravado")
            dsPeriodo.Tables("Tabla").Columns.Add("Prima_vac_exento")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_Prima_vac")
            dsPeriodo.Tables("Tabla").Columns.Add("Vacaciones_proporcionales")
            dsPeriodo.Tables("Tabla").Columns.Add("Bono_Puntualidad")
            dsPeriodo.Tables("Tabla").Columns.Add("Bono_Asistencia")
            dsPeriodo.Tables("Tabla").Columns.Add("Fomento_Deporte")
            dsPeriodo.Tables("Tabla").Columns.Add("Bono_Proceso")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_percepciones")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_percepciones_p/isr")
            dsPeriodo.Tables("Tabla").Columns.Add("Incapacidad")
            dsPeriodo.Tables("Tabla").Columns.Add("ISR")
            dsPeriodo.Tables("Tabla").Columns.Add("IMSS")
            dsPeriodo.Tables("Tabla").Columns.Add("Infonavit")
            dsPeriodo.Tables("Tabla").Columns.Add("Infonavit_bim_anterior")
            dsPeriodo.Tables("Tabla").Columns.Add("Ajuste_infonavit")
            dsPeriodo.Tables("Tabla").Columns.Add("Pension_Alimenticia")
            dsPeriodo.Tables("Tabla").Columns.Add("Prestamo")
            dsPeriodo.Tables("Tabla").Columns.Add("Fonacot")
            dsPeriodo.Tables("Tabla").Columns.Add("Subsidio_Generado")
            dsPeriodo.Tables("Tabla").Columns.Add("Subsidio_Aplicado")
            dsPeriodo.Tables("Tabla").Columns.Add("Operadora")
            dsPeriodo.Tables("Tabla").Columns.Add("Prestamo_Personal_A")
            dsPeriodo.Tables("Tabla").Columns.Add("Adeudo_Infonavit_A")
            dsPeriodo.Tables("Tabla").Columns.Add("Diferencia_Infonavit_A")
            dsPeriodo.Tables("Tabla").Columns.Add("Asimilados")
            dsPeriodo.Tables("Tabla").Columns.Add("Retenciones_Operadora")
            dsPeriodo.Tables("Tabla").Columns.Add("%_Comisión")
            dsPeriodo.Tables("Tabla").Columns.Add("Comisión_Operadora")
            dsPeriodo.Tables("Tabla").Columns.Add("Comisión_Asimilados")
            dsPeriodo.Tables("Tabla").Columns.Add("IMSS_CS")
            dsPeriodo.Tables("Tabla").Columns.Add("RCV_CS")
            dsPeriodo.Tables("Tabla").Columns.Add("Infonavit_CS")
            dsPeriodo.Tables("Tabla").Columns.Add("ISN_CS")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_Costo_Social")
            dsPeriodo.Tables("Tabla").Columns.Add("Subtotal")
            dsPeriodo.Tables("Tabla").Columns.Add("IVA")
            dsPeriodo.Tables("Tabla").Columns.Add("TOTAL_DEPOSITO")



            'verificamos que no sea una nomina ya guardada como final
            sql = "select * from NominaProceso inner join EmpleadosC on fkiIdEmpleadoC=iIdEmpleadoC"
            sql &= " where Nomina.fkiIdEmpresa = 1 And fkiIdPeriodo = " & cboperiodo.SelectedValue
            sql &= " and Nomina.iEstatus=1 and iEstatusEmpleado=" & cboserie.SelectedIndex
            sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex
            sql &= " order by " & campoordenamiento 'cNombreLargo"
            'sql = "EXEC getNominaXEmpresaXPeriodo " & gIdEmpresa & "," & cboperiodo.SelectedValue & ",1"

            bCalcular = True
            Dim rwNominaGuardada As DataRow() = nConsulta(sql)

            'If rwNominaGuardadaFinal Is Nothing = False Then
            If rwNominaGuardada Is Nothing = False Then
                'Cargamos los datos de guardados como final
                For x As Integer = 0 To rwNominaGuardada.Count - 1

                    Dim fila As DataRow = dsPeriodo.Tables("Tabla").NewRow

                    fila.Item("Consecutivo") = (x + 1).ToString
                    fila.Item("Id_empleado") = rwNominaGuardada(x)("fkiIdEmpleadoC").ToString





                    fila.Item("CodigoEmpleado") = rwNominaGuardada(x)("cCodigoEmpleado").ToString
                    fila.Item("Nombre") = rwNominaGuardada(x)("cNombreLargo").ToString.ToUpper()
                    fila.Item("Status") = IIf(rwNominaGuardada(x)("iOrigen").ToString = "1", "INTERINO", "PLANTA")
                    fila.Item("RFC") = rwNominaGuardada(x)("cRFC").ToString
                    fila.Item("CURP") = rwNominaGuardada(x)("cCURP").ToString
                    fila.Item("Num_IMSS") = rwNominaGuardada(x)("cIMSS").ToString

                    fila.Item("Fecha_Nac") = Date.Parse(rwNominaGuardada(x)("dFechaNac").ToString).ToShortDateString()
                    'Dim tiempo As TimeSpan = Date.Now - Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString)

                    fila.Item("Edad") = CalcularEdad(Date.Parse(rwNominaGuardada(x)("dFechaNac").ToString).Day, Date.Parse(rwNominaGuardada(x)("dFechaNac").ToString).Month, Date.Parse(rwNominaGuardada(x)("dFechaNac").ToString).Year)
                    fila.Item("Puesto") = rwNominaGuardada(x)("Puesto").ToString
                    fila.Item("Buque") = rwNominaGuardada(x)("Buque").ToString

                    fila.Item("Tipo_Infonavit") = rwNominaGuardada(x)("TipoInfonavit").ToString
                    fila.Item("Valor_Infonavit") = rwNominaGuardada(x)("fValor").ToString
                    '
                    fila.Item("Sueldo_Base") = rwNominaGuardada(x)("fSalarioBase").ToString
                    fila.Item("Salario_Diario") = rwNominaGuardada(x)("fSalarioDiario").ToString
                    fila.Item("Salario_Cotización") = rwNominaGuardada(x)("fSalarioBC").ToString


                    fila.Item("Dias_Trabajados") = rwNominaGuardada(x)("iDiasTrabajados").ToString
                    fila.Item("Tipo_Incapacidad") = rwNominaGuardada(x)("TipoIncapacidad").ToString
                    fila.Item("Número_días") = rwNominaGuardada(x)("iNumeroDias").ToString
                    fila.Item("Sueldo_Bruto") = rwNominaGuardada(x)("fSueldoBruto").ToString
                    fila.Item("Aguinaldo_gravado") = rwNominaGuardada(x)("fAguinaldoGravado").ToString
                    fila.Item("Aguinaldo_exento") = rwNominaGuardada(x)("fAguinaldoExento").ToString
                    fila.Item("Total_Aguinaldo") = Math.Round(Double.Parse(rwNominaGuardada(x)("fAguinaldoGravado").ToString) + Double.Parse(rwNominaGuardada(x)("fAguinaldoExento").ToString), 2)
                    fila.Item("Prima_vac_gravado") = rwNominaGuardada(x)("fPrimaVacacionalGravado").ToString
                    fila.Item("Prima_vac_exento") = rwNominaGuardada(x)("fPrimaVacacionalExento").ToString
                    fila.Item("Total_Prima_vac") = Math.Round(Double.Parse(rwNominaGuardada(x)("fPrimaVacacionalGravado").ToString) + Double.Parse(rwNominaGuardada(x)("fPrimaVacacionalExento").ToString), 2)
                    fila.Item("Vacaciones_proporcionales") = rwNominaGuardada(x)("fVacacionesProporcionales").ToString
                    fila.Item("Bono_Puntualidad") = rwNominaGuardada(x)("fBonoPuntualidad").ToString
                    fila.Item("Bono_Asistencia") = rwNominaGuardada(x)("fBonoAsistencia").ToString
                    fila.Item("Fomento_Deporte") = rwNominaGuardada(x)("fFomentoDeporte").ToString
                    fila.Item("Bono_Proceso") = rwNominaGuardada(x)("fBonoProceso").ToString
                    fila.Item("Total_percepciones") = rwNominaGuardada(x)("fTotalPercepciones").ToString
                    fila.Item("Total_percepciones_p/isr") = rwNominaGuardada(x)("fTotalPercepcionesISR").ToString
                    fila.Item("Incapacidad") = rwNominaGuardada(x)("fIncapacidad").ToString
                    fila.Item("ISR") = rwNominaGuardada(x)("fIsr").ToString
                    fila.Item("IMSS") = rwNominaGuardada(x)("fImss").ToString
                    fila.Item("Infonavit") = rwNominaGuardada(x)("fInfonavit").ToString
                    fila.Item("Infonavit_bim_anterior") = rwNominaGuardada(x)("fInfonavitBanterior").ToString
                    fila.Item("Ajuste_infonavit") = rwNominaGuardada(x)("fAjusteInfonavit").ToString
                    fila.Item("Pension_Alimenticia") = rwNominaGuardada(x)("fPensionAlimenticia").ToString
                    fila.Item("Prestamo") = rwNominaGuardada(x)("fPrestamo").ToString
                    fila.Item("Fonacot") = rwNominaGuardada(x)("fFonacot").ToString
                    fila.Item("Subsidio_Generado") = rwNominaGuardada(x)("fSubsidioGenerado").ToString
                    fila.Item("Subsidio_Aplicado") = rwNominaGuardada(x)("fSubsidioAplicado").ToString
                    fila.Item("Operadora") = rwNominaGuardada(x)("fOperadora").ToString
                    fila.Item("Prestamo_Personal_A") = rwNominaGuardada(x)("fPrestamoPerA").ToString
                    fila.Item("Adeudo_Infonavit_A") = rwNominaGuardada(x)("fAdeudoInfonavitA").ToString
                    fila.Item("Diferencia_Infonavit_A") = rwNominaGuardada(x)("fDiferenciaInfonavitA").ToString
                    fila.Item("Asimilados") = rwNominaGuardada(x)("fAsimilados").ToString
                    fila.Item("Retenciones_Operadora") = rwNominaGuardada(x)("fRetencionOperadora").ToString
                    fila.Item("%_Comisión") = rwNominaGuardada(x)("fPorComision").ToString
                    fila.Item("Comisión_Operadora") = rwNominaGuardada(x)("fComisionOperadora").ToString
                    fila.Item("Comisión_Asimilados") = rwNominaGuardada(x)("fComisionAsimilados").ToString
                    fila.Item("IMSS_CS") = rwNominaGuardada(x)("fImssCS").ToString
                    fila.Item("RCV_CS") = rwNominaGuardada(x)("fRcvCS").ToString
                    fila.Item("Infonavit_CS") = rwNominaGuardada(x)("fInfonavitCS").ToString
                    fila.Item("ISN_CS") = rwNominaGuardada(x)("fInsCS").ToString
                    fila.Item("Total_Costo_Social") = rwNominaGuardada(x)("fTotalCostoSocial").ToString
                    fila.Item("Subtotal") = rwNominaGuardada(x)("fSubtotal").ToString
                    fila.Item("IVA") = rwNominaGuardada(x)("fIVA").ToString
                    fila.Item("TOTAL_DEPOSITO") = rwNominaGuardada(x)("fTotalDeposito").ToString


                    dsPeriodo.Tables("Tabla").Rows.Add(fila)
                Next

                dtgDatos.DataSource = dsPeriodo.Tables("Tabla")

                dtgDatos.Columns(0).Width = 30
                dtgDatos.Columns(0).ReadOnly = True
                dtgDatos.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                'consecutivo
                dtgDatos.Columns(1).Width = 60
                dtgDatos.Columns(1).ReadOnly = True
                dtgDatos.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'idempleado
                dtgDatos.Columns(2).Width = 100
                dtgDatos.Columns(2).ReadOnly = True
                dtgDatos.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'codigo empleado
                dtgDatos.Columns(3).Width = 100
                dtgDatos.Columns(3).ReadOnly = True
                dtgDatos.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'Nombre
                dtgDatos.Columns(4).Width = 250
                dtgDatos.Columns(4).ReadOnly = True
                'Estatus
                dtgDatos.Columns(5).Width = 100
                dtgDatos.Columns(5).ReadOnly = True
                'RFC
                dtgDatos.Columns(6).Width = 100
                dtgDatos.Columns(6).ReadOnly = True
                'dtgDatos.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                'CURP
                dtgDatos.Columns(7).Width = 150
                dtgDatos.Columns(7).ReadOnly = True
                'IMSS 

                dtgDatos.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(8).ReadOnly = True
                'Fecha_Nac
                dtgDatos.Columns(9).Width = 150
                dtgDatos.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(9).ReadOnly = True

                'Edad
                dtgDatos.Columns(10).ReadOnly = True
                dtgDatos.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                'Puesto
                dtgDatos.Columns(11).ReadOnly = True
                dtgDatos.Columns(11).Width = 200
                dtgDatos.Columns.Remove("Puesto")

                Dim combo As New DataGridViewComboBoxColumn

                sql = "select * from puestos where iTipo=1 order by cNombre"

                'Dim rwPuestos As DataRow() = nConsulta(sql)
                'If rwPuestos Is Nothing = False Then
                '    combo.Items.Add("uno")
                '    combo.Items.Add("dos")
                '    combo.Items.Add("tres")
                'End If

                nCargaCBO(combo, sql, "cNombre", "iIdPuesto")

                combo.HeaderText = "Puesto"

                combo.Width = 150
                dtgDatos.Columns.Insert(11, combo)
                'DirectCast(dtgDatos.Columns(11), DataGridViewComboBoxColumn).Sorted = True
                'Dim combo2 As New DataGridViewComboBoxCell
                'combo2 = CType(Me.dtgDatos.Rows(2).Cells(11), DataGridViewComboBoxCell)
                'combo2.Value = combo.Items(11)



                'dtgDatos.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                'Buque
                'dtgDatos.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(12).ReadOnly = True
                dtgDatos.Columns(12).Width = 150
                dtgDatos.Columns.Remove("Buque")

                Dim combo2 As New DataGridViewComboBoxColumn

                sql = "select * from departamentos where iEstatus=1 order by cNombre"

                'Dim rwPuestos As DataRow() = nConsulta(sql)
                'If rwPuestos Is Nothing = False Then
                '    combo.Items.Add("uno")
                '    combo.Items.Add("dos")
                '    combo.Items.Add("tres")
                'End If

                nCargaCBO(combo2, sql, "cNombre", "iIdDepartamento")

                combo2.HeaderText = "Buque"
                combo2.Width = 150
                dtgDatos.Columns.Insert(12, combo2)

                'Tipo_Infonavit
                dtgDatos.Columns(13).ReadOnly = True
                dtgDatos.Columns(13).Width = 150
                'dtgDatos.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight



                'Valor_Infonavit
                dtgDatos.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(14).ReadOnly = True
                dtgDatos.Columns(14).Width = 150
                'Sueldo_Base
                dtgDatos.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(15).ReadOnly = True
                dtgDatos.Columns(15).Width = 150
                'Salario_Diario
                dtgDatos.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(16).ReadOnly = True
                dtgDatos.Columns(16).Width = 150
                'Salario_Cotización
                dtgDatos.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(17).ReadOnly = True
                dtgDatos.Columns(17).Width = 150
                'Dias_Trabajados
                dtgDatos.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(18).Width = 150
                'Tipo_Incapacidad
                dtgDatos.Columns(19).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(19).ReadOnly = True
                dtgDatos.Columns(19).Width = 150
                'Número_días
                dtgDatos.Columns(20).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(20).ReadOnly = True
                dtgDatos.Columns(20).Width = 150
                'Sueldo_Bruto
                dtgDatos.Columns(21).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(21).ReadOnly = True
                dtgDatos.Columns(21).Width = 150
                'Tiempo_Extra_Fijo_Gravado
                dtgDatos.Columns(22).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(22).ReadOnly = True
                dtgDatos.Columns(22).Width = 150

                'Tiempo_Extra_Fijo_Exento
                dtgDatos.Columns(23).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(23).ReadOnly = True
                dtgDatos.Columns(23).Width = 150

                'Tiempo_Extra_Ocasional
                dtgDatos.Columns(24).Width = 150
                dtgDatos.Columns(24).ReadOnly = True
                dtgDatos.Columns(24).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'Desc_Sem_Obligatorio
                dtgDatos.Columns(25).Width = 150
                dtgDatos.Columns(25).ReadOnly = True
                dtgDatos.Columns(25).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'Vacaciones_proporcionales
                dtgDatos.Columns(26).Width = 150
                dtgDatos.Columns(26).ReadOnly = True
                dtgDatos.Columns(26).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'Aguinaldo_gravado
                dtgDatos.Columns(27).Width = 150
                dtgDatos.Columns(27).ReadOnly = True
                dtgDatos.Columns(27).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'Aguinaldo_exento
                dtgDatos.Columns(28).Width = 150
                dtgDatos.Columns(28).ReadOnly = True
                dtgDatos.Columns(28).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'Total_Aguinaldo
                dtgDatos.Columns(29).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(29).Width = 150
                dtgDatos.Columns(29).ReadOnly = True

                'Prima_vac_gravado
                dtgDatos.Columns(30).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(30).ReadOnly = True
                dtgDatos.Columns(30).Width = 150
                'Prima_vac_exento 
                dtgDatos.Columns(31).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(31).ReadOnly = True
                dtgDatos.Columns(31).Width = 150

                'Total_Prima_vac
                dtgDatos.Columns(32).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(32).ReadOnly = True
                dtgDatos.Columns(32).Width = 150


                'Total_percepciones
                dtgDatos.Columns(33).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(33).ReadOnly = True
                dtgDatos.Columns(33).Width = 150
                'Total_percepciones_p/isr
                dtgDatos.Columns(34).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(34).ReadOnly = True
                dtgDatos.Columns(34).Width = 150

                'Incapacidad
                dtgDatos.Columns(35).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(35).ReadOnly = True
                dtgDatos.Columns(35).Width = 150

                'ISR
                dtgDatos.Columns(36).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(36).ReadOnly = True
                dtgDatos.Columns(36).Width = 150


                'IMSS
                dtgDatos.Columns(37).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(37).ReadOnly = True
                dtgDatos.Columns(37).Width = 150

                'Infonavit
                dtgDatos.Columns(38).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(38).ReadOnly = True
                dtgDatos.Columns(38).Width = 150
                'Infonavit_bim_anterior
                dtgDatos.Columns(39).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(39).ReadOnly = True
                dtgDatos.Columns(39).Width = 150
                'Ajuste_infonavit
                dtgDatos.Columns(40).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(40).ReadOnly = True
                dtgDatos.Columns(40).Width = 150
                'Pension_Alimenticia
                dtgDatos.Columns(41).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(40).ReadOnly = True
                dtgDatos.Columns(41).Width = 150
                'Prestamo
                dtgDatos.Columns(42).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(42).ReadOnly = True
                dtgDatos.Columns(42).Width = 150
                'Fonacot
                dtgDatos.Columns(43).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(43).ReadOnly = True
                dtgDatos.Columns(43).Width = 150
                'Subsidio_Generado
                dtgDatos.Columns(44).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(44).ReadOnly = True
                dtgDatos.Columns(44).Width = 150
                'Subsidio_Aplicado
                dtgDatos.Columns(45).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(45).ReadOnly = True
                dtgDatos.Columns(45).Width = 150
                'Operadora
                dtgDatos.Columns(46).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(46).ReadOnly = True
                dtgDatos.Columns(46).Width = 150

                'Prestamo Personal Asimilado
                dtgDatos.Columns(47).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(48).ReadOnly = True
                dtgDatos.Columns(47).Width = 150

                'Adeudo_Infonavit_Asimilado
                dtgDatos.Columns(48).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(49).ReadOnly = True
                dtgDatos.Columns(48).Width = 150

                'Difencia infonavit Asimilado
                dtgDatos.Columns(49).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(50).ReadOnly = True
                dtgDatos.Columns(49).Width = 150

                'Complemento Asimilado
                dtgDatos.Columns(50).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(50).ReadOnly = True
                dtgDatos.Columns(50).Width = 150

                'Retenciones_Operadora
                dtgDatos.Columns(51).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(51).ReadOnly = True
                dtgDatos.Columns(51).Width = 150

                '% Comision
                dtgDatos.Columns(52).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(52).ReadOnly = True
                dtgDatos.Columns(52).Width = 150

                'Comision_Operadora
                dtgDatos.Columns(53).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(53).ReadOnly = True
                dtgDatos.Columns(53).Width = 150

                'Comision asimilados
                dtgDatos.Columns(54).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(54).ReadOnly = True
                dtgDatos.Columns(54).Width = 150

                'IMSS_CS
                dtgDatos.Columns(55).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(55).ReadOnly = True
                dtgDatos.Columns(55).Width = 150

                'RCV_CS
                dtgDatos.Columns(56).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(56).ReadOnly = True
                dtgDatos.Columns(56).Width = 150

                'Infonavit_CS
                dtgDatos.Columns(57).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(57).ReadOnly = True
                dtgDatos.Columns(57).Width = 150

                'ISN_CS
                dtgDatos.Columns(58).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(58).ReadOnly = True
                dtgDatos.Columns(58).Width = 150

                'Total Costo Social
                dtgDatos.Columns(59).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(59).ReadOnly = True
                dtgDatos.Columns(59).Width = 150

                'Subtotal
                dtgDatos.Columns(60).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(60).ReadOnly = True
                dtgDatos.Columns(60).Width = 150

                'IVA
                dtgDatos.Columns(61).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(61).ReadOnly = True
                dtgDatos.Columns(61).Width = 150

                'TOTAL DEPOSITO
                dtgDatos.Columns(62).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(62).ReadOnly = True
                dtgDatos.Columns(62).Width = 150

                'calcular()

                'Cambiamos index del combo en el grid

                'For x As Integer = 0 To dtgDatos.Rows.Count - 1

                '    sql = "select * from nomina where fkiIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                '    sql &= " and fkiIdPeriodo=" & cboperiodo.SelectedValue
                '    sql &= " and iEstatusEmpleado=" & cboserie.SelectedIndex
                '    sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex
                '    Dim rwFila As DataRow() = nConsulta(sql)



                '    CType(Me.dtgDatos.Rows(x).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("Puesto").ToString()
                '    CType(Me.dtgDatos.Rows(x).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("Buque").ToString()
                'Next


                'verificar costo social

                Dim contador, Posicion1, Posicion2, Posicion3, Posicion4, Posicion5 As Integer


                For x As Integer = 0 To dtgDatos.Rows.Count - 1
                    contador = 0


                    For y As Integer = 0 To dtgDatos.Rows.Count - 1
                        If dtgDatos.Rows(x).Cells(2).Value = dtgDatos.Rows(y).Cells(2).Value Then
                            contador = contador + 1
                            If contador = 1 Then
                                Posicion1 = y
                            End If
                            If contador = 2 Then
                                Posicion2 = y
                            End If
                            If contador = 3 Then
                                Posicion3 = y
                            End If
                            If contador = 4 Then
                                Posicion4 = y
                            End If
                            If contador = 5 Then
                                Posicion5 = y
                            End If
                        End If



                    Next
                    sql = "select * from Nomina inner join EmpleadosC on fkiIdEmpleadoC=iIdEmpleadoC  where fkiIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                    'sql = "select * from nomina inner join EmpleadosC on nomin where fkiIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                    sql &= " and fkiIdPeriodo=" & cboperiodo.SelectedValue
                    sql &= " and iEstatusEmpleado=" & cboserie.SelectedIndex
                    sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex
                    sql &= " order by " & campoordenamiento

                    Dim rwFila As DataRow() = nConsulta(sql)

                    If rwFila.Length = 1 Then
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("Buque").ToString()

                    End If

                    If rwFila.Length = 2 Then
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion2).Cells(11), DataGridViewComboBoxCell).Value = rwFila(1)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion2).Cells(12), DataGridViewComboBoxCell).Value = rwFila(1)("Buque").ToString()

                    End If
                    If rwFila.Length = 3 Then
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion2).Cells(11), DataGridViewComboBoxCell).Value = rwFila(1)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion2).Cells(12), DataGridViewComboBoxCell).Value = rwFila(1)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion3).Cells(11), DataGridViewComboBoxCell).Value = rwFila(2)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion3).Cells(12), DataGridViewComboBoxCell).Value = rwFila(2)("Buque").ToString()
                    End If
                    If rwFila.Length = 4 Then
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion2).Cells(11), DataGridViewComboBoxCell).Value = rwFila(1)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion2).Cells(12), DataGridViewComboBoxCell).Value = rwFila(1)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion3).Cells(11), DataGridViewComboBoxCell).Value = rwFila(2)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion3).Cells(12), DataGridViewComboBoxCell).Value = rwFila(2)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion4).Cells(11), DataGridViewComboBoxCell).Value = rwFila(3)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion4).Cells(12), DataGridViewComboBoxCell).Value = rwFila(3)("Buque").ToString()
                    End If
                    If rwFila.Length = 5 Then
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion2).Cells(11), DataGridViewComboBoxCell).Value = rwFila(1)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion2).Cells(12), DataGridViewComboBoxCell).Value = rwFila(1)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion3).Cells(11), DataGridViewComboBoxCell).Value = rwFila(2)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion3).Cells(12), DataGridViewComboBoxCell).Value = rwFila(2)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion4).Cells(11), DataGridViewComboBoxCell).Value = rwFila(3)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion4).Cells(12), DataGridViewComboBoxCell).Value = rwFila(3)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion5).Cells(11), DataGridViewComboBoxCell).Value = rwFila(4)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion5).Cells(12), DataGridViewComboBoxCell).Value = rwFila(4)("Buque").ToString()
                    End If
                Next



                'Cambiamos el index del combro de departamentos

                'For x As Integer = 0 To dtgDatos.Rows.Count - 1

                '    sql = "select * from empleadosC where iIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                '    Dim rwFila As DataRow() = nConsulta(sql)




                'Next

                MessageBox.Show("Datos cargados", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)


            Else

                If cboTipoNomina.SelectedIndex = 0 Then
                    If cboserie.SelectedIndex = 0 Then
                        'Buscamos los datos de sindicato solamente
                        sql = "select  * from empleadosC where fkiIdClienteInter=-1"
                        'sql = "select iIdEmpleadoC,NumCuenta, (cApellidoP + ' ' + cApellidoM + ' ' + cNombre) as nombre, fkiIdEmpresa,fSueldoOrd,fCosto from empleadosC"
                        'sql &= " where empleadosC.iOrigen=2 and empleadosC.iEstatus=1"
                        'sql &= " and empleadosC.fkiIdEmpresa =" & gIdEmpresa
                        sql &= " order by cFuncionesPuesto,cNombreLargo"

                    ElseIf cboserie.SelectedIndex > 0 Or cboserie.SelectedIndex - 1 Then
                        sql = "select * from Nomina inner join EmpleadosC on fkiIdEmpleadoC=iIdEmpleadoC"
                        sql &= " where Nomina.fkiIdEmpresa = 1 And fkiIdPeriodo = " & cboperiodo.SelectedValue
                        sql &= " and Nomina.iEstatus=1 and iEstatusEmpleado=20"
                        sql &= " order by cNombreLargo"

                    End If


                    Dim rwDatosEmpleados As DataRow() = nConsulta(sql)
                    If rwDatosEmpleados Is Nothing = False Then
                        For x As Integer = 0 To rwDatosEmpleados.Length - 1


                            Dim fila As DataRow = dsPeriodo.Tables("Tabla").NewRow

                            fila.Item("Consecutivo") = (x + 1).ToString
                            fila.Item("Id_empleado") = rwDatosEmpleados(x)("iIdEmpleadoC").ToString
                            fila.Item("CodigoEmpleado") = rwDatosEmpleados(x)("cCodigoEmpleado").ToString
                            fila.Item("Nombre") = rwDatosEmpleados(x)("cNombreLargo").ToString.ToUpper()
                            fila.Item("Status") = IIf(rwDatosEmpleados(x)("iOrigen").ToString = "1", "INTERINO", "PLANTA")
                            fila.Item("RFC") = rwDatosEmpleados(x)("cRFC").ToString
                            fila.Item("CURP") = rwDatosEmpleados(x)("cCURP").ToString
                            fila.Item("Num_IMSS") = rwDatosEmpleados(x)("cIMSS").ToString

                            fila.Item("Fecha_Nac") = Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString).ToShortDateString()
                            'Dim tiempo As TimeSpan = Date.Now - Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString)
                            fila.Item("Edad") = CalcularEdad(Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString).Day, Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString).Month, Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString).Year)
                            fila.Item("Puesto") = rwDatosEmpleados(x)("cPuesto").ToString
                            fila.Item("Buque") = "ECO III"

                            fila.Item("Tipo_Infonavit") = rwDatosEmpleados(x)("cTipoFactor").ToString
                            fila.Item("Valor_Infonavit") = rwDatosEmpleados(x)("fFactor").ToString
                            fila.Item("Sueldo_Base") = "0.00"
                            fila.Item("Salario_Diario") = rwDatosEmpleados(x)("fSueldoBase").ToString
                            fila.Item("Salario_Cotización") = rwDatosEmpleados(x)("fSueldoIntegrado").ToString
                            fila.Item("Dias_Trabajados") = "30"
                            fila.Item("Tipo_Incapacidad") = TipoIncapacidad(rwDatosEmpleados(x)("iIdEmpleadoC").ToString, cboperiodo.SelectedValue)
                            fila.Item("Número_días") = NumDiasIncapacidad(rwDatosEmpleados(x)("iIdEmpleadoC").ToString, cboperiodo.SelectedValue)
                            fila.Item("Sueldo_Bruto") = ""
                            fila.Item("Tiempo_Extra_Fijo_Gravado") = ""
                            fila.Item("Tiempo_Extra_Fijo_Exento") = ""
                            fila.Item("Tiempo_Extra_Ocasional") = ""
                            fila.Item("Desc_Sem_Obligatorio") = ""
                            fila.Item("Vacaciones_proporcionales") = ""
                            fila.Item("Aguinaldo_gravado") = ""
                            fila.Item("Aguinaldo_exento") = ""
                            fila.Item("Total_Aguinaldo") = ""
                            fila.Item("Prima_vac_gravado") = ""
                            fila.Item("Prima_vac_exento") = ""

                            fila.Item("Total_Prima_vac") = ""
                            fila.Item("Total_percepciones") = ""
                            fila.Item("Total_percepciones_p/isr") = ""
                            fila.Item("Incapacidad") = ""
                            fila.Item("ISR") = ""
                            fila.Item("IMSS") = ""
                            fila.Item("Infonavit") = ""
                            fila.Item("Infonavit_bim_anterior") = ""
                            fila.Item("Ajuste_infonavit") = ""
                            fila.Item("Pension_Alimenticia") = ""
                            fila.Item("Prestamo") = ""
                            fila.Item("Fonacot") = ""
                            fila.Item("Subsidio_Generado") = ""
                            fila.Item("Subsidio_Aplicado") = ""
                            fila.Item("Operadora") = ""
                            fila.Item("Prestamo_Personal_A") = ""
                            fila.Item("Adeudo_Infonavit_A") = ""
                            fila.Item("Diferencia_Infonavit_A") = ""
                            fila.Item("Asimilados") = ""
                            fila.Item("Retenciones_Operadora") = ""
                            fila.Item("%_Comisión") = ""
                            fila.Item("Comisión_Operadora") = ""
                            fila.Item("Comisión_Asimilados") = ""
                            fila.Item("IMSS_CS") = ""
                            fila.Item("RCV_CS") = ""
                            fila.Item("Infonavit_CS") = ""
                            fila.Item("ISN_CS") = ""
                            fila.Item("Total_Costo_Social") = ""
                            fila.Item("Subtotal") = ""
                            fila.Item("IVA") = ""
                            fila.Item("TOTAL_DEPOSITO") = ""


                            dsPeriodo.Tables("Tabla").Rows.Add(fila)




                        Next




                        dtgDatos.DataSource = dsPeriodo.Tables("Tabla")

                        dtgDatos.Columns(0).Width = 30
                        dtgDatos.Columns(0).ReadOnly = True
                        dtgDatos.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                        'consecutivo
                        dtgDatos.Columns(1).Width = 60
                        dtgDatos.Columns(1).ReadOnly = True
                        dtgDatos.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'idempleado
                        dtgDatos.Columns(2).Width = 100
                        dtgDatos.Columns(2).ReadOnly = True
                        dtgDatos.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'codigo empleado
                        dtgDatos.Columns(3).Width = 100
                        dtgDatos.Columns(3).ReadOnly = True
                        dtgDatos.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Nombre
                        dtgDatos.Columns(4).Width = 250
                        dtgDatos.Columns(4).ReadOnly = True
                        'Estatus
                        dtgDatos.Columns(5).Width = 100
                        dtgDatos.Columns(5).ReadOnly = True
                        'RFC
                        dtgDatos.Columns(6).Width = 100
                        dtgDatos.Columns(6).ReadOnly = True
                        'dtgDatos.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                        'CURP
                        dtgDatos.Columns(7).Width = 150
                        dtgDatos.Columns(7).ReadOnly = True
                        'IMSS 

                        dtgDatos.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(8).ReadOnly = True
                        'Fecha_Nac
                        dtgDatos.Columns(9).Width = 150
                        dtgDatos.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(9).ReadOnly = True

                        'Edad
                        dtgDatos.Columns(10).ReadOnly = True
                        dtgDatos.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                        'Puesto
                        dtgDatos.Columns(11).ReadOnly = True
                        dtgDatos.Columns(11).Width = 200
                        dtgDatos.Columns.Remove("Puesto")

                        Dim combo As New DataGridViewComboBoxColumn

                        sql = "select * from puestos where iTipo=1 order by cNombre"

                        'Dim rwPuestos As DataRow() = nConsulta(sql)
                        'If rwPuestos Is Nothing = False Then
                        '    combo.Items.Add("uno")
                        '    combo.Items.Add("dos")
                        '    combo.Items.Add("tres")
                        'End If

                        nCargaCBO(combo, sql, "cNombre", "iIdPuesto")

                        combo.HeaderText = "Puesto"

                        combo.Width = 150
                        dtgDatos.Columns.Insert(11, combo)
                        'DirectCast(dtgDatos.Columns(11), DataGridViewComboBoxColumn).Sorted = True
                        'Dim combo2 As New DataGridViewComboBoxCell
                        'combo2 = CType(Me.dtgDatos.Rows(2).Cells(11), DataGridViewComboBoxCell)
                        'combo2.Value = combo.Items(11)



                        'dtgDatos.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                        'Buque
                        'dtgDatos.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(12).ReadOnly = True
                        dtgDatos.Columns(12).Width = 150
                        dtgDatos.Columns.Remove("Buque")

                        Dim combo2 As New DataGridViewComboBoxColumn

                        sql = "select * from departamentos where iEstatus=1 order by cNombre"

                        'Dim rwPuestos As DataRow() = nConsulta(sql)
                        'If rwPuestos Is Nothing = False Then
                        '    combo.Items.Add("uno")
                        '    combo.Items.Add("dos")
                        '    combo.Items.Add("tres")
                        'End If

                        nCargaCBO(combo2, sql, "cNombre", "iIdDepartamento")

                        combo2.HeaderText = "Buque"
                        combo2.Width = 150
                        dtgDatos.Columns.Insert(12, combo2)

                        'Tipo_Infonavit
                        dtgDatos.Columns(13).ReadOnly = True
                        dtgDatos.Columns(13).Width = 150
                        'dtgDatos.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight



                        'Valor_Infonavit
                        dtgDatos.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(14).ReadOnly = True
                        dtgDatos.Columns(14).Width = 150
                        'Sueldo_Base
                        dtgDatos.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(15).ReadOnly = True
                        dtgDatos.Columns(15).Width = 150
                        'Salario_Diario
                        dtgDatos.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(16).ReadOnly = True
                        dtgDatos.Columns(16).Width = 150
                        'Salario_Cotización
                        dtgDatos.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(17).ReadOnly = True
                        dtgDatos.Columns(17).Width = 150
                        'Dias_Trabajados
                        dtgDatos.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(18).Width = 150
                        'Tipo_Incapacidad
                        dtgDatos.Columns(19).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(19).ReadOnly = True
                        dtgDatos.Columns(19).Width = 150
                        'Número_días
                        dtgDatos.Columns(20).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(20).ReadOnly = True
                        dtgDatos.Columns(20).Width = 150
                        'Sueldo_Bruto
                        dtgDatos.Columns(21).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(21).ReadOnly = True
                        dtgDatos.Columns(21).Width = 150
                        'Tiempo_Extra_Fijo_Gravado
                        dtgDatos.Columns(22).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(22).ReadOnly = True
                        dtgDatos.Columns(22).Width = 150

                        'Tiempo_Extra_Fijo_Exento
                        dtgDatos.Columns(23).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(23).ReadOnly = True
                        dtgDatos.Columns(23).Width = 150

                        'Tiempo_Extra_Ocasional
                        dtgDatos.Columns(24).Width = 150
                        dtgDatos.Columns(24).ReadOnly = True
                        dtgDatos.Columns(24).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Desc_Sem_Obligatorio
                        dtgDatos.Columns(25).Width = 150
                        dtgDatos.Columns(25).ReadOnly = True
                        dtgDatos.Columns(25).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Vacaciones_proporcionales
                        dtgDatos.Columns(26).Width = 150
                        dtgDatos.Columns(26).ReadOnly = True
                        dtgDatos.Columns(26).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Aguinaldo_gravado
                        dtgDatos.Columns(27).Width = 150
                        dtgDatos.Columns(27).ReadOnly = True
                        dtgDatos.Columns(27).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Aguinaldo_exento
                        dtgDatos.Columns(28).Width = 150
                        dtgDatos.Columns(28).ReadOnly = True
                        dtgDatos.Columns(28).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Total_Aguinaldo
                        dtgDatos.Columns(29).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(29).Width = 150
                        dtgDatos.Columns(29).ReadOnly = True

                        'Prima_vac_gravado
                        dtgDatos.Columns(30).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(30).ReadOnly = True
                        dtgDatos.Columns(30).Width = 150
                        'Prima_vac_exento 
                        dtgDatos.Columns(31).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(31).ReadOnly = True
                        dtgDatos.Columns(31).Width = 150

                        'Total_Prima_vac
                        dtgDatos.Columns(32).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(32).ReadOnly = True
                        dtgDatos.Columns(32).Width = 150


                        'Total_percepciones
                        dtgDatos.Columns(33).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(33).ReadOnly = True
                        dtgDatos.Columns(33).Width = 150
                        'Total_percepciones_p/isr
                        dtgDatos.Columns(34).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(34).ReadOnly = True
                        dtgDatos.Columns(34).Width = 150

                        'Incapacidad
                        dtgDatos.Columns(35).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(35).ReadOnly = True
                        dtgDatos.Columns(35).Width = 150

                        'ISR
                        dtgDatos.Columns(36).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(36).ReadOnly = True
                        dtgDatos.Columns(36).Width = 150


                        'IMSS
                        dtgDatos.Columns(37).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(37).ReadOnly = True
                        dtgDatos.Columns(37).Width = 150

                        'Infonavit
                        dtgDatos.Columns(38).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(38).ReadOnly = True
                        dtgDatos.Columns(38).Width = 150
                        'Infonavit_bim_anterior
                        dtgDatos.Columns(39).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(39).ReadOnly = True
                        dtgDatos.Columns(39).Width = 150
                        'Ajuste_infonavit
                        dtgDatos.Columns(40).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(40).ReadOnly = True
                        dtgDatos.Columns(40).Width = 150
                        'Pension_Alimenticia
                        dtgDatos.Columns(41).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(40).ReadOnly = True
                        dtgDatos.Columns(41).Width = 150
                        'Prestamo
                        dtgDatos.Columns(42).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(42).ReadOnly = True
                        dtgDatos.Columns(42).Width = 150
                        'Fonacot
                        dtgDatos.Columns(43).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(43).ReadOnly = True
                        dtgDatos.Columns(43).Width = 150
                        'Subsidio_Generado
                        dtgDatos.Columns(44).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(44).ReadOnly = True
                        dtgDatos.Columns(44).Width = 150
                        'Subsidio_Aplicado
                        dtgDatos.Columns(45).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(45).ReadOnly = True
                        dtgDatos.Columns(45).Width = 150
                        'Operadora
                        dtgDatos.Columns(46).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(46).ReadOnly = True
                        dtgDatos.Columns(46).Width = 150

                        'Prestamo Personal Asimilado
                        dtgDatos.Columns(47).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(48).ReadOnly = True
                        dtgDatos.Columns(47).Width = 150

                        'Adeudo_Infonavit_Asimilado
                        dtgDatos.Columns(48).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(49).ReadOnly = True
                        dtgDatos.Columns(48).Width = 150

                        'Difencia infonavit Asimilado
                        dtgDatos.Columns(49).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(50).ReadOnly = True
                        dtgDatos.Columns(49).Width = 150

                        'Complemento Asimilado
                        dtgDatos.Columns(50).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(50).ReadOnly = True
                        dtgDatos.Columns(50).Width = 150

                        'Retenciones_Operadora
                        dtgDatos.Columns(51).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(51).ReadOnly = True
                        dtgDatos.Columns(51).Width = 150

                        '% Comision
                        dtgDatos.Columns(52).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(52).ReadOnly = True
                        dtgDatos.Columns(52).Width = 150

                        'Comision_Operadora
                        dtgDatos.Columns(53).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(53).ReadOnly = True
                        dtgDatos.Columns(53).Width = 150

                        'Comision asimilados
                        dtgDatos.Columns(54).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(54).ReadOnly = True
                        dtgDatos.Columns(54).Width = 150

                        'IMSS_CS
                        dtgDatos.Columns(55).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(55).ReadOnly = True
                        dtgDatos.Columns(55).Width = 150

                        'RCV_CS
                        dtgDatos.Columns(56).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(56).ReadOnly = True
                        dtgDatos.Columns(56).Width = 150

                        'Infonavit_CS
                        dtgDatos.Columns(57).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(57).ReadOnly = True
                        dtgDatos.Columns(57).Width = 150

                        'ISN_CS
                        dtgDatos.Columns(58).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(58).ReadOnly = True
                        dtgDatos.Columns(58).Width = 150

                        'Total Costo Social
                        dtgDatos.Columns(59).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(59).ReadOnly = True
                        dtgDatos.Columns(59).Width = 150

                        'Subtotal
                        dtgDatos.Columns(60).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(60).ReadOnly = True
                        dtgDatos.Columns(60).Width = 150

                        'IVA
                        dtgDatos.Columns(61).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(61).ReadOnly = True
                        dtgDatos.Columns(61).Width = 150

                        'TOTAL DEPOSITO
                        dtgDatos.Columns(62).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(62).ReadOnly = True
                        dtgDatos.Columns(62).Width = 150
                        'calcular()

                        'Cambiamos index del combo en el grid

                        For x As Integer = 0 To dtgDatos.Rows.Count - 1

                            sql = "select * from empleadosC where iIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                            Dim rwFila As DataRow() = nConsulta(sql)



                            CType(Me.dtgDatos.Rows(x).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("cPuesto").ToString()
                            CType(Me.dtgDatos.Rows(x).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("cFuncionesPuesto").ToString()
                        Next


                        'Cambiamos el index del combro de departamentos

                        'For x As Integer = 0 To dtgDatos.Rows.Count - 1

                        '    sql = "select * from empleadosC where iIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                        '    Dim rwFila As DataRow() = nConsulta(sql)




                        'Next


                        MessageBox.Show("Datos cargados", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("No hay datos en este período", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If




                    'No hay datos en este período
                Else
                    MessageBox.Show("Para la nomina Descanso, solo se mostraran datos guardados, no se podrá calcular de 0", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If




            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Function TipoIncapacidad(ByVal idempleado As String, ByVal periodo As Integer) As String
        Dim sql As String
        Dim cadena As String = "Ninguno"

        Try
            sql = "select * from periodos where iIdPeriodo= " & periodo
            Dim rwPeriodo As DataRow() = nConsulta(sql)

            If rwPeriodo Is Nothing = False Then

                sql = "select * from incapacidad where iIdIncapacidad= "
                sql &= " (select Max(iIdIncapacidad) from incapacidad where iEstatus=1 and fkiIdEmpleado=" & idempleado & ") "
                Dim rwIncapacidad As DataRow() = nConsulta(sql)

                If rwIncapacidad Is Nothing = False Then
                    Dim FechaBuscar As Date = Date.Parse(rwIncapacidad(0)("FechaInicio"))
                    Dim FechaInicial As Date = Date.Parse(rwPeriodo(0)("dFechaInicio"))
                    Dim FechaFinal As Date = Date.Parse(rwPeriodo(0)("dFechaFin"))
                    'Dim FechaAntiguedad As Date = Date.Parse(rwDatosBanco(0)("dFechaAntiguedad"))

                    If FechaBuscar.CompareTo(FechaInicial) >= 0 And FechaBuscar.CompareTo(FechaFinal) <= 0 Then
                        'Estamos dentro del rango inicial
                        Return Identificadorincapacidad(rwIncapacidad(0)("RamoRiesgo"))

                    ElseIf FechaBuscar.CompareTo(FechaInicial) <= 0 Then
                        FechaBuscar = Date.Parse(rwIncapacidad(0)("fechafin"))
                        If FechaBuscar.CompareTo(FechaFinal) <= 0 Then
                            Return Identificadorincapacidad(rwIncapacidad(0)("RamoRiesgo"))
                        End If

                    End If

                Else
                    cadena = "Ninguno"
                    Return cadena
                End If


            Else
                Return "Ninguno"

            End If
            Return "Ninguno"
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Function

    Private Function NumDiasIncapacidad(ByVal idempleado As String, ByVal periodo As Integer) As String
        Dim sql As String
        Dim cadena As String

        Try
            sql = "select * from periodos where iIdPeriodo= " & periodo
            Dim rwPeriodo As DataRow() = nConsulta(sql)

            If rwPeriodo Is Nothing = False Then

                sql = "select * from incapacidad where iIdIncapacidad= "
                sql &= " (select Max(iIdIncapacidad) from incapacidad where iEstatus=1 and fkiIdEmpleado=" & idempleado & ") "
                Dim rwIncapacidad As DataRow() = nConsulta(sql)

                If rwIncapacidad Is Nothing = False Then
                    Dim FechaBuscar As Date = Date.Parse(rwIncapacidad(0)("FechaInicio"))
                    Dim FechaInicial As Date = Date.Parse(rwPeriodo(0)("dFechaInicio"))
                    Dim FechaFinal As Date = Date.Parse(rwPeriodo(0)("dFechaFin"))
                    'Dim FechaAntiguedad As Date = Date.Parse(rwDatosBanco(0)("dFechaAntiguedad"))

                    If FechaBuscar.CompareTo(FechaInicial) >= 0 And FechaBuscar.CompareTo(FechaFinal) <= 0 Then
                        'Estamos dentro del rango inicial
                        FechaBuscar = Date.Parse(rwIncapacidad(0)("fechafin"))
                        If FechaBuscar.CompareTo(FechaFinal) <= 0 Then
                            'Restamos entre final incapacidad menos la inicial incapacidad
                            Return (DateDiff(DateInterval.Day, Date.Parse(rwIncapacidad(0)("FechaInicio")), Date.Parse(rwIncapacidad(0)("fechafin"))) + 1).ToString
                        Else
                            'restamos final del periodo menos inicial incapacidad
                            Return (DateDiff(DateInterval.Day, Date.Parse(rwIncapacidad(0)("FechaInicio")), Date.Parse(rwPeriodo(0)("dFechaFin"))) + 1).ToString


                        End If

                    ElseIf FechaBuscar.CompareTo(FechaInicial) <= 0 Then
                        FechaBuscar = Date.Parse(rwIncapacidad(0)("fechafin"))
                        If FechaBuscar.CompareTo(FechaFinal) <= 0 Then
                            'Restamos fecha final incapacidad menos la fechainicial  periodo
                            Return (DateDiff(DateInterval.Day, Date.Parse(rwPeriodo(0)("dFechaInicio")), Date.Parse(rwIncapacidad(0)("fechafin"))) + 1).ToString
                        Else
                            'todos los dias del periodo tiene incapaciddad
                            Return (DateDiff(DateInterval.Day, Date.Parse(rwPeriodo(0)("dFechaInicio")), Date.Parse(rwPeriodo(0)("dFechaFin"))) + 1).ToString
                        End If

                    End If
                Else
                    cadena = "0"
                    Return cadena
                End If


            Else
                Return "0"

            End If
            Return "0"
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Function

    Private Function Identificadorincapacidad(ByVal identificador As String) As String
        Try
            Dim TipoIncidencia As String = ""

            If identificador = "0" Then
                TipoIncidencia = "Riesgo de trabajo"
            ElseIf identificador = "1" Then
                TipoIncidencia = "Enfermedad general"
            ElseIf identificador = "2" Then
                TipoIncidencia = "Maternidad"

            End If

            Return TipoIncidencia
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Function

    Private Function CalcularEdad(ByVal DiaNacimiento As Integer, ByVal MesNacimiento As Integer, ByVal AñoNacimiento As Integer)
        ' SE DEFINEN LAS FECHAS ACTUALES
        Dim AñoActual As Integer = Year(Now)
        Dim MesActual As Integer = Month(Now)
        Dim DiaActual As Integer = Now.Day
        Dim Cumplidos As Boolean = False
        ' SE COMPRUEBA CUANDO FUE EL ULTIMOS CUMPLEAÑOS
        ' FORMULA:
        '   Años cumplidos = (Año del ultimo cumpleaños - Año de nacimiento)
        If (MesNacimiento <= MesActual) Then
            If (DiaNacimiento <= DiaActual) Then
                If (DiaNacimiento = DiaActual And MesNacimiento = MesActual) Then
                    'MsgBox("Feliz Cumpleaños!")
                End If
                ' MsgBox("Ya cumplio")
                Cumplidos = True
            End If
        End If

        If (Cumplidos = False) Then
            AñoActual = (AñoActual - 1)
            'MsgBox("Ultimo cumpleaños: " & AñoActual)
        End If
        ' Se realiza la resta de años para definir los años cumplidos
        Dim EdadAños As Integer = (AñoActual - AñoNacimiento)
        ' DEFINICION DE LOS MESES LUEGO DEL ULTIMO CUMPLEAÑOS
        Dim EdadMes As Integer
        If Not (AñoActual = Now.Year) Then
            EdadMes = (12 - MesNacimiento)
            EdadMes = EdadMes + Now.Month
        Else
            EdadMes = Math.Abs(Now.Month - MesNacimiento)
        End If
        'SACAMOS LA CANTIDAD DE DIAS EXACTOS
        Dim EdadDia As Integer = (DiaActual - DiaNacimiento)

        'RETORNAMOS LOS VALORES EN UNA CADENA STRING
        Return (EdadAños)


    End Function

    Private Sub cmdexcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdexcel.Click
        Try

            Dim filaExcel As Integer = 0
            Dim filatmp As Integer = 0
            Dim dialogo As New SaveFileDialog()
            Dim periodo As String

            pnlProgreso.Visible = True
            pnlCatalogo.Enabled = False
            Application.DoEvents()

            pgbProgreso.Minimum = 0
            pgbProgreso.Value = 0
            pgbProgreso.Maximum = dtgDatos.Rows.Count


            If dtgDatos.Rows.Count > 0 Then
                Dim ruta As String
                ruta = My.Application.Info.DirectoryPath() & "\Archivos\nominasprocesos.xlsx"

                Dim book As New ClosedXML.Excel.XLWorkbook(ruta)
                Dim libro As New ClosedXML.Excel.XLWorkbook

                book.Worksheet(1).CopyTo(libro, "PLANTA PROCESO OK")
                book.Worksheet(2).CopyTo(libro, "XURTEP ABORDO")
                book.Worksheet(3).CopyTo(libro, "XURTEP DESCANSO")
                book.Worksheet(4).CopyTo(libro, "ASIMILADOS DESCANSO")
                book.Worksheets(5).CopyTo(libro, "RESUMEN")


                Dim hoja As IXLWorksheet = libro.Worksheets(0)
                Dim hoja2 As IXLWorksheet = libro.Worksheets(1)
                Dim hoja3 As IXLWorksheet = libro.Worksheets(2)
                Dim hoja4 As IXLWorksheet = libro.Worksheets(3)
                Dim hoja5 As IXLWorksheet = libro.Worksheets(4)

                filaExcel = 11
                Dim nombrebuque As String
                Dim inicio As Integer = 0
                Dim contadorexcelbuqueinicial As Integer = 0
                Dim contadorexcelbuquefinal As Integer = 0
                Dim total As Integer = dtgDatos.Rows.Count - 1
                'Dim filatmp As Integer = 13 - 4
                'Dim filatmp2 As Integer = filaExcel
                Dim fecha As String


                If cboTipoNomina.SelectedIndex = 1 Then
                    llenargridD("0")

                End If

                '<<<<<<<<<<<<<<<<<<<<<<<<<<PLANTA PROCESO>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                recorrerFilasColumnas(hoja, 11, dtgDatos.Rows.Count + 10, 40, "clear")

                Dim rwPeriodo0 As DataRow() = nConsulta("Select * from periodos where iIdPeriodo=" & cboperiodo.SelectedValue)
                If rwPeriodo0 Is Nothing = False Then
                    Dim Fechafin As Date = rwPeriodo0(0).Item("dFechaFin")
                    periodo = "1 " & MonthString(rwPeriodo0(0).Item("iMes")).ToUpper & " AL " & Fechafin.Day & " " & MonthString(rwPeriodo0(0).Item("iMes")).ToUpper & " " & rwPeriodo0(0).Item("iEjercicio")
                    'periodo = MonthString(rwPeriodo0(0).Item("iMes")).ToUpper & " DE " & (rwPeriodo0(0).Item("iEjercicio"))
                    fecha = MonthString(rwPeriodo0(0).Item("iMes")).ToUpper
                    hoja.Cell(8, 1).Style.Font.SetBold(True)
                    hoja.Cell(8, 1).Style.NumberFormat.Format = "@"
                    hoja.Cell(8, 1).Value = periodo
                    hoja.Cell(8, 1).Style.Font.FontSize = 12

                End If

                For x As Integer = 0 To dtgDatos.Rows.Count - 1

                    hoja.Cell(filaExcel + x, 1).Value = dtgDatos.Rows(x).Cells(3).Value ' NO TRABAJADOR
                    hoja.Cell(filaExcel + x, 2).Value = dtgDatos.Rows(x).Cells(10).Value ' EDAD 
                    hoja.Cell(filaExcel + x, 3).Value = dtgDatos.Rows(x).Cells(4).Value ' TRABAJADOR 
                    hoja.Cell(filaExcel + x, 4).Value = dtgDatos.Rows(x).Cells(5).Value 'STATUS/Tipo de nomina
                    hoja.Cell(filaExcel + x, 5).Value = "ECO III"
                    hoja.Cell(filaExcel + x, 6).Value = dtgDatos.Rows(x).Cells(11).FormattedValue 'PUESTO 
                    hoja.Cell(filaExcel + x, 7).Value = dtgDatos.Rows(x).Cells(18).Value ' DIAS ABORDO
                    hoja.Cell(filaExcel + x, 8).Value = dtgDatos.Rows(x).Cells(18).Value ' DIAS DESCANSO
                    hoja.Cell(filaExcel + x, 9).FormulaA1 = "=K" & filaExcel + x & "/2"  'NOMINA ABORDO
                    hoja.Cell(filaExcel + x, 10).FormulaA1 = "=K" & filaExcel + x & "/2" 'NOMINA DESCANSO
                    hoja.Cell(filaExcel + x, 11).Value = dtgDatos.Rows(x).Cells(15).Value ' SUELDO ORDINARIO
                    hoja.Cell(filaExcel + x, 12).FormulaA1 = "='XURTEP ABORDO'!AE" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AE" & filaExcel + x + 1 & "+'XURTEP ABORDO'!AF" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AF" & filaExcel + x + 1 ' INFONAVIT
                    hoja.Cell(filaExcel + x, 13).FormulaA1 = "='XURTEP ABORDO'!AI" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AI" & filaExcel + x + 1 'FONACOT
                    hoja.Cell(filaExcel + x, 14).FormulaA1 = "='XURTEP ABORDO'!AG" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AG" & filaExcel + x + 1 ' PENSION ALIMENTICIA
                    hoja.Cell(filaExcel + x, 15).FormulaA1 = "='XURTEP ABORDO'!AH" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AH" & filaExcel + x + 1 'Anticipo SA
                    hoja.Cell(filaExcel + x, 16).FormulaA1 = "0.0"
                    'hoja.Cell(filaExcel + x, 17).FormulaA1 = "0.0"
                    hoja.Cell(filaExcel + x, 17).FormulaA1 = "=K" & filaExcel + x & "-L" & filaExcel + x & "-M" & filaExcel + x & "-N" & filaExcel + x & "-O" & filaExcel + x & "-P" & filaExcel + x & "-Q" & filaExcel + x  'SUELDO ORDINARIO
                    hoja.Cell(filaExcel + x, 18).FormulaA1 = "" '
                    hoja.Cell(filaExcel + x, 19).FormulaA1 = "='XURTEP ABORDO'!AKK" & filaExcel + x & "+'XURTEP DESCANSO'!AKK" & filaExcel + x  ' XURTEP
                    hoja.Cell(filaExcel + x, 20).FormulaA1 = "=R" & filaExcel + x & "-T" & filaExcel + x 'COMPLEMENTO (ASIM NETOS)
                    hoja.Cell(filaExcel + x, 21).FormulaA1 = "='XURTEP ABORDO'!AB" & filaExcel + x + 1 & "+'XURTEP ABORDO'!AC" & filaExcel + x + 1 & "+'XURTEP ABORDO'!AD" & filaExcel + x + 1 & "+'XURTEP ABORDO'!AE" & filaExcel + x + 1 & "+'XURTEP ABORDO'!AF" & filaExcel + x + 1 & "+'XURTEP ABORDO'!AG" & filaExcel + x + 1 & "+'XURTEP ABORDO'!AH" & filaExcel + x + 1 & "+'XURTEP ABORDO'!AI" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AB" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AC" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AD" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AE" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AF" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AG" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AH" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AI" & filaExcel + x + 1
                    hoja.Cell(filaExcel + x, 22).FormulaA1 = ""
                    hoja.Cell(filaExcel + x, 23).FormulaA1 = "2%" '%COMISION
                    hoja.Cell(filaExcel + x, 24).FormulaA1 = "=+((T" & filaExcel + x & "+V" & filaExcel + x & ")*X" & filaExcel + x & ")" 'COMISION XURTEP
                    hoja.Cell(filaExcel + x, 25).FormulaA1 = "=+(P" & filaExcel + x & "+Q" & filaExcel + x & "+U" & filaExcel + x & ")*X" & filaExcel + x  ' COMPLEMENTO COMISION
                    hoja.Cell(filaExcel + x, 26).FormulaA1 = dtgDatos.Rows(x).Cells(55).Value 'IMSS
                    hoja.Cell(filaExcel + x, 27).FormulaA1 = dtgDatos.Rows(x).Cells(56).Value 'RCV
                    hoja.Cell(filaExcel + x, 28).FormulaA1 = dtgDatos.Rows(x).Cells(57).Value 'INFONAVIT
                    hoja.Cell(filaExcel + x, 29).FormulaA1 = dtgDatos.Rows(x).Cells(58).Value
                    hoja.Cell(filaExcel + x, 30).FormulaA1 = "=AA" & filaExcel + x & "+AB" & filaExcel + x & "+AC" & filaExcel + x & "+AD" & filaExcel + x
                    hoja.Cell(filaExcel + x, 31).FormulaA1 = dtgDatos.Rows(x).Cells(59).Value 'COSTO SOCIAL 
                    hoja.Cell(filaExcel + x, 32).FormulaA1 = "=+P" & filaExcel + x & "+Q" & filaExcel + x & "+T" & filaExcel + x & "+U" & filaExcel + x & "+V" & filaExcel + x & "+Y" & filaExcel + x & "+AF" & filaExcel + x & "+Z" & filaExcel + x
                    hoja.Cell(filaExcel + x, 33).FormulaA1 = "=+AG" & filaExcel + x & "*0.16" 'IVA 16%
                    hoja.Cell(filaExcel + x, 34).FormulaA1 = ("=+AG" & filaExcel + x & "+AH" & filaExcel + x)

                    hoja.Cell(filaExcel + x, 36).FormulaA1 = dtgDatos.Rows(x).Cells(59).Value
                    hoja.Cell(filaExcel + x, 37).FormulaA1 = "30"
                Next x
                filaExcel = filaExcel + 2
                contadorexcelbuquefinal = filaExcel + total - 1

                hoja.Cell(filaExcel + total, 9).FormulaA1 = "=SUM(I11:I" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 10).FormulaA1 = "=SUM(J11:J" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 11).FormulaA1 = "=SUM(K11:K" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 12).FormulaA1 = "=SUM(L11:L" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 13).FormulaA1 = "=SUM(M11 :M" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 14).FormulaA1 = "=SUM(N11 :N" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 15).FormulaA1 = "=SUM(O11 :O" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 16).FormulaA1 = "=SUM(P11:P" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 17).FormulaA1 = "=SUM(Q11:Q" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 18).FormulaA1 = "=SUM(R11:R" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 19).FormulaA1 = "=SUM(S11:S" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 20).FormulaA1 = "=SUM(T11:T" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 21).FormulaA1 = "=SUM(U11:U" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 22).FormulaA1 = "=SUM(V11:V" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 23).FormulaA1 = "=SUM(W11:W" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 24).FormulaA1 = "=SUM(X11:X" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 25).FormulaA1 = "=SUM(Y11:Y" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 26).FormulaA1 = "=SUM(Z11:Z" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 27).FormulaA1 = "=SUM(AA11:AA" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 28).FormulaA1 = "=SUM(AB11:AB" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 29).FormulaA1 = "=SUM(AC11:AC" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 30).FormulaA1 = "=SUM(AD11:AD" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 31).FormulaA1 = "=SUM(AE11:AE" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 32).FormulaA1 = "=SUM(AF11:AF" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 33).FormulaA1 = "=SUM(AG11:AG" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 34).FormulaA1 = "=SUM(AH11:AH" & contadorexcelbuquefinal & ")"
                hoja.Cell(filaExcel + total, 35).FormulaA1 = "=SUM(AI11:AI" & contadorexcelbuquefinal & ")"

                hoja.Range(filaExcel + total, 9, filaExcel + total, 35).Style.Fill.BackgroundColor = XLColor.PowderBlue
                hoja.Range(filaExcel + total, 9, filaExcel + total, 35).Style.Font.SetBold(True)


                Dim cuenta, banco, clabe As String

                '<<<<<<<<<<<<<<ASIMILADOS >>>>>>>>>>>>
                filaExcel = 2
                filatmp = 11
                recorrerFilasColumnas(hoja4, 2, dtgDatos.Rows.Count + 30, 13, "clear")

                Dim app, apm, nom As String
                For x As Integer = 0 To dtgDatos.Rows.Count - 1

                    Dim empleado As DataRow() = nConsulta("Select * from empleadosC where cCodigoEmpleado=" & dtgDatos.Rows(x).Cells(3).Value)
                    If empleado Is Nothing = False Then
                        cuenta = empleado(0).Item("NumCuenta")
                        clabe = empleado(0).Item("Clabe")
                        app = empleado(0).Item("cApellidoP")
                        apm = empleado(0).Item("cApellidoM")
                        nom = empleado(0).Item("cNombre")
                        Dim bank As DataRow() = nConsulta("select * from bancos where iIdBanco =" & empleado(0).Item("fkiIdBanco"))
                        If bank Is Nothing = False Then
                            banco = bank(0).Item("cBANCO")
                        End If
                    End If

                    hoja4.Cell(filaExcel, 5).Style.NumberFormat.Format = "@"
                    hoja4.Cell(filaExcel, 8).Style.NumberFormat.Format = "@"
                    hoja4.Cell(filaExcel, 9).Style.NumberFormat.Format = "@"
                    hoja4.Cell(filaExcel, 11).Style.NumberFormat.Format = "@"
                    hoja4.Cell(filaExcel, 12).Style.NumberFormat.Format = "@"

                    hoja4.Cell(filaExcel, 1).Value = app ' Apellido Paterno
                    hoja4.Cell(filaExcel, 2).Value = apm ' Apellido Materno
                    hoja4.Cell(filaExcel, 3).Value = nom ' Nombre
                    hoja4.Cell(filaExcel, 4).FormulaA1 = "=MARINOS!U11" 'Asimilado
                    hoja4.Cell(filaExcel, 5).Value = dtgDatos.Rows(x).Cells(8).Value '# Afiliacion IMSS
                    hoja4.Cell(filaExcel, 6).Value = dtgDatos.Rows(x).Cells(18).Value 'Dias Trabajandos
                    hoja4.Cell(filaExcel, 7).Value = banco
                    hoja4.Cell(filaExcel, 8).Value = cuenta ' IIf(cuenta = 0, "", cuenta)
                    hoja4.Cell(filaExcel, 9).Value = clabe
                    hoja4.Cell(filaExcel, 10).Value = "SIN TARJT" ' Tarjeta
                    hoja4.Cell(filaExcel, 11).Value = dtgDatos.Rows(x).Cells(6).Value ' RFC
                    hoja4.Cell(filaExcel, 12).Value = dtgDatos.Rows(x).Cells(7).Value ' CURP

                    filaExcel = filaExcel + 1
                Next x


                '<<<<<<<<<<<<<<<RESUMEN>>>>>>>>>>>>>>>>>>

                filaExcel = 5
                hoja5.Cell(4, 3).Style.Font.SetBold(True)
                hoja5.Cell(4, 3).Style.NumberFormat.Format = "@"
                ' hoja5.Cell(4, 3).Value = periodo

                For x As Integer = 0 To dtgDatos.Rows.Count - 1
                    hoja5.Cell(filaExcel, 2).Style.NumberFormat.Format = "@"
                    hoja5.Range(filaExcel, 4, filaExcel, 9).Style.NumberFormat.Format = "@"
                    hoja5.Range(filaExcel, 10, filaExcel, 11).Style.NumberFormat.NumberFormatId = 4

                    hoja5.Range(filaExcel, 2, filaExcel, 9).Style.Font.SetBold(False)
                    'hoja5.Range(filaExcel, 8, filaExcel, 9).Style.NumberFormat.NumberFormatId = 4
                    hoja5.Range(filaExcel, 2, filaExcel, 9).Style.Font.SetFontColor(XLColor.Black)
                    hoja5.Range(filaExcel, 2, filaExcel, 9).Style.Font.SetFontName("Arial")
                    hoja5.Range(filaExcel, 2, filaExcel, 9).Style.Font.SetFontSize(8)
                    hoja5.Range(filaExcel, 2, filaExcel, 9).Style.Font.SetBold(False)
                    hoja5.Range(filaExcel, 2, filaExcel, 9).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.General)

                    Dim empleado As DataRow() = nConsulta("Select * from empleadosC where cCodigoEmpleado=" & dtgDatos.Rows(x).Cells(3).Value)
                    If empleado Is Nothing = False Then
                        cuenta = empleado(0).Item("NumCuenta")
                        clabe = empleado(0).Item("Clabe")
                        Dim bank As DataRow() = nConsulta("select * from bancos where iIdBanco =" & empleado(0).Item("fkiIdBanco"))
                        If bank Is Nothing = False Then
                            banco = bank(0).Item("cBANCO")
                        End If
                    End If



                    hoja5.Cell(filaExcel, 2).Value = dtgDatos.Rows(x).Cells(3).Value 'Codigo
                    hoja5.Cell(filaExcel, 3).Value = ""
                    hoja5.Cell(filaExcel, 4).Value = dtgDatos.Rows(x).Cells(4).Value ' Trabajador
                    hoja5.Cell(filaExcel, 5).Value = "ECO III" ' BUQUE
                    hoja5.Cell(filaExcel, 6).Value = dtgDatos.Rows(x).Cells(6).Value 'rfc 
                    hoja5.Cell(filaExcel, 7).Value = banco
                    hoja5.Cell(filaExcel, 8).Value = cuenta ' IIf(cuenta = 0, "", cuenta)
                    hoja5.Cell(filaExcel, 9).Value = clabe
                    hoja5.Cell(filaExcel, 10).Value = dtgDatos.Rows(x).Cells(46).Value ' XURTEP
                    hoja5.Cell(filaExcel, 11).Value = dtgDatos.Rows(x).Cells(50).Value ' ASIMILADOS


                    filaExcel = filaExcel + 1

                    pgbProgreso.Value += 1
                    Application.DoEvents()

                Next x


                'Formulas
                hoja5.Range(filaExcel + 2, 10, filaExcel + 4, 11).Style.Font.SetBold(True)
                hoja5.Cell(filaExcel + 2, 10).FormulaA1 = "=SUM(J5:J" & filaExcel & ")"
                hoja5.Cell(filaExcel + 2, 11).FormulaA1 = "=SUM(K5:K" & filaExcel & ")"



                '<<<<<<<<<<<<<<<<<Xurtep Abordo>>>>>>>>>>>>>>>>>>>>>>>>

                'Limpiar encabezado y relleno
                recorrerFilasColumnas(hoja2, 1, 10, 50, "clear", 13)
                recorrerFilasColumnas(hoja2, 12, dtgDatos.Rows.Count + 30, 50, "clear", 1)


                'Validamos en que nomina esta

                Dim rwPeriodo As DataRow() = nConsulta("Select (CONVERT(nvarchar(12),dFechaInicio,103) + ' al ' + CONVERT(nvarchar(12),dFechaFin,103)) as dFechaInicio from periodos where iIdPeriodo=" & cboperiodo.SelectedValue)
                If rwPeriodo Is Nothing = False Then
                    hoja2.Cell(7, 2).Value = "Periodo Mensual del " & rwPeriodo(0).Item("dFechaInicio")
                    hoja3.Cell(7, 2).Value = "Periodo Mensual del " & rwPeriodo(0).Item("dFechaInicio")

                End If


                ''XURTEP ABORDO
                filaExcel = 12
                For x As Integer = 0 To dtgDatos.Rows.Count - 1
                    'Style
                    hoja2.Cell(filaExcel, 1).Style.NumberFormat.Format = "@"
                    hoja2.Cell(filaExcel, 7).Style.NumberFormat.Format = "@"

                    hoja2.Range(filaExcel, 1, filaExcel, 45).Unmerge()
                    hoja2.Range(filaExcel, 1, filaExcel, 45).Style.Font.SetFontColor(XLColor.Black)
                    hoja2.Range(filaExcel, 11, filaExcel, 12).Style.NumberFormat.NumberFormatId = 4
                    hoja2.Range(filaExcel, 14, filaExcel, 45).Style.NumberFormat.NumberFormatId = 4

                    hoja2.Range(filaExcel, 1, filaExcel, 45).Style.Font.SetFontName("Arial")
                    hoja2.Range(filaExcel, 1, filaExcel, 45).Style.Font.SetFontSize(8)
                    hoja2.Range(filaExcel, 1, filaExcel, 45).Style.Font.SetBold(False)

                    'hoja2.Range(filaExcel, 1, filaExcel, 11).Style.NumberFormat.Format = "@"
                    'hoja2.Cell(filaExcel, 15).Style.NumberFormat.Format = "@"
                    hoja2.Range(filaExcel, 1, filaExcel, 45).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.General)
                    'Datos
                    hoja2.Cell(filaExcel, 1).Value = dtgDatos.Rows(x).Cells(3).Value 'N° Trabajador
                    hoja2.Cell(filaExcel, 2).Value = dtgDatos.Rows(x).Cells(4).Value ' Nombre
                    hoja2.Cell(filaExcel, 3).Value = dtgDatos.Rows(x).Cells(5).Value 'Status
                    hoja2.Cell(filaExcel, 4).Value = dtgDatos.Rows(x).Cells(12).FormattedValue 'buque 
                    hoja2.Cell(filaExcel, 5).Value = dtgDatos.Rows(x).Cells(6).Value 'rfc 
                    hoja2.Cell(filaExcel, 6).Value = dtgDatos.Rows(x).Cells(7).Value 'curp
                    hoja2.Cell(filaExcel, 7).Value = dtgDatos.Rows(x).Cells(8).Value 'imss 
                    hoja2.Cell(filaExcel, 8).Value = dtgDatos.Rows(x).Cells(9).Value 'fecha nac 
                    hoja2.Cell(filaExcel, 9).Value = dtgDatos.Rows(x).Cells(10).Value 'edad
                    hoja2.Cell(filaExcel, 10).Value = dtgDatos.Rows(x).Cells(11).FormattedValue 'puesto
                    hoja2.Cell(filaExcel, 11).Value = dtgDatos.Rows(x).Cells(16).Value 'Salario Diario
                    hoja2.Cell(filaExcel, 12).Value = dtgDatos.Rows(x).Cells(17).Value 'SDI  
                    hoja2.Cell(filaExcel, 13).Value = dtgDatos.Rows(x).Cells(18).Value ' Dias Trabajados 
                    hoja2.Cell(filaExcel, 14).Value = dtgDatos.Rows(x).Cells(21).Value 'Sueldo base
                    hoja2.Cell(filaExcel, 15).Value = dtgDatos.Rows(x).Cells(22).Value ' Tiempo Extra Fijo Gravado 
                    hoja2.Cell(filaExcel, 16).Value = dtgDatos.Rows(x).Cells(23).Value 'Tiempo Extra Fijo Exento
                    hoja2.Cell(filaExcel, 17).Value = dtgDatos.Rows(x).Cells(24).Value ' Tiempo extra ocasional  
                    hoja2.Cell(filaExcel, 18).Value = dtgDatos.Rows(x).Cells(25).Value ' Desc. Sem Oblig.
                    hoja2.Cell(filaExcel, 19).Value = dtgDatos.Rows(x).Cells(26).Value ' VAC. PROPOR 
                    hoja2.Cell(filaExcel, 20).Value = dtgDatos.Rows(x).Cells(27).Value ' AGINALDO GRA 
                    hoja2.Cell(filaExcel, 21).Value = dtgDatos.Rows(x).Cells(28).Value ' AGUINALDO EXENTO 
                    hoja2.Cell(filaExcel, 22).Value = dtgDatos.Rows(x).Cells(29).Value ' TOTAL AGUINALDO 
                    hoja2.Cell(filaExcel, 23).Value = dtgDatos.Rows(x).Cells(30).Value ' P. VAC. GRAVADO 
                    hoja2.Cell(filaExcel, 24).Value = dtgDatos.Rows(x).Cells(31).Value ' P. VAC. EXENTO 
                    hoja2.Cell(filaExcel, 25).Value = dtgDatos.Rows(x).Cells(32).Value ' TOTAL P. VAC 
                    hoja2.Cell(filaExcel, 26).Value = dtgDatos.Rows(x).Cells(33).Value ' TOTAL PERCEPCIONES
                    hoja2.Cell(filaExcel, 27).Value = dtgDatos.Rows(x).Cells(34).Value ' TOTAL PERCEPC P/ISR
                    hoja2.Cell(filaExcel, 28).Value = dtgDatos.Rows(x).Cells(35).Value ' INCAPACIDAD
                    hoja2.Cell(filaExcel, 29).Value = dtgDatos.Rows(x).Cells(36).Value ' ISR
                    hoja2.Cell(filaExcel, 30).Value = dtgDatos.Rows(x).Cells(37).Value ' IMSS
                    hoja2.Cell(filaExcel, 31).Value = dtgDatos.Rows(x).Cells(38).Value ' INFONAVIT
                    hoja2.Cell(filaExcel, 32).Value = dtgDatos.Rows(x).Cells(39).Value ' INFONAVIT
                    hoja2.Cell(filaExcel, 33).Value = dtgDatos.Rows(x).Cells(41).Value ' PENSION ALIMENTICIA
                    hoja2.Cell(filaExcel, 34).Value = dtgDatos.Rows(x).Cells(42).Value ' PRESTAMOS/ANTICIPO NOMINA?
                    hoja2.Cell(filaExcel, 35).Value = dtgDatos.Rows(x).Cells(43).Value ' FONACOT
                    hoja2.Cell(filaExcel, 36).FormulaA1 = "=AB" & filaExcel & "+AC" & filaExcel & "+AD" & filaExcel & "+AE" & filaExcel & "+AF" & filaExcel & "+AG" & filaExcel & "+AH" & filaExcel & "+AI" & filaExcel
                    hoja2.Cell(filaExcel, 37).Value = dtgDatos.Rows(x).Cells(46).Value ' NETO A PAGAR


                    'hoja2.Cell(filaExcel, 40).Value = dtgDatos.Rows(x).Cells(55).Value
                    'hoja2.Cell(filaExcel, 41).Value = dtgDatos.Rows(x).Cells(56).Value
                    'hoja2.Cell(filaExcel, 42).Value = dtgDatos.Rows(x).Cells(57).Value
                    'hoja2.Cell(filaExcel, 43).Value = dtgDatos.Rows(x).Cells(58).Value
                    'hoja2.Cell(filaExcel, 44).FormulaA1 = "=SUM(AN" & filaExcel & ":AQ" & filaExcel & ")"
                    'hoja2.Cell(filaExcel, 45).Value = dtgDatos.Rows(x).Cells(59).Value

                    filaExcel = filaExcel + 1


                Next x

                'STYLE
                hoja2.Range(filaExcel + 2, 13, filaExcel + 4, 39).Style.Font.SetFontColor(XLColor.Black)
                hoja2.Range(filaExcel + 2, 13, filaExcel + 4, 39).Style.NumberFormat.NumberFormatId = 4
                hoja2.Range(filaExcel + 2, 13, filaExcel + 4, 39).Style.Font.SetBold(True)

                'Xurtep Abordo       

                hoja2.Cell(filaExcel + 2, 13).FormulaA1 = "=SUM(M12:M" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 14).FormulaA1 = "=SUM(N12:N" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 15).FormulaA1 = "=SUM(O12:O" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 16).FormulaA1 = "=SUM(P12:P" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 17).FormulaA1 = "=SUM(Q12:Q" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 18).FormulaA1 = "=SUM(R12:R" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 19).FormulaA1 = "=SUM(S12:S" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 20).FormulaA1 = "=SUM(T12:T" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 21).FormulaA1 = "=SUM(U12:U" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 22).FormulaA1 = "=SUM(V12:V" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 23).FormulaA1 = "=SUM(W12:W" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 24).FormulaA1 = "=SUM(X12:X" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 25).FormulaA1 = "=SUM(Y12:Y" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 26).FormulaA1 = "=SUM(Z12:Z" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 27).FormulaA1 = "=SUM(AA12:AA" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 28).FormulaA1 = "=SUM(AB12:AB" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 29).FormulaA1 = "=SUM(AC12:AC" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 30).FormulaA1 = "=SUM(AD12:AD" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 31).FormulaA1 = "=SUM(AE12:AE" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 32).FormulaA1 = "=SUM(AF12:AF" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 33).FormulaA1 = "=SUM(AG12:AG" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 34).FormulaA1 = "=SUM(AH12:AH" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 35).FormulaA1 = "=SUM(AI12:AI" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 36).FormulaA1 = "=SUM(AJ12:AJ" & filaExcel & ")"
                hoja2.Cell(filaExcel + 2, 37).FormulaA1 = "=SUM(AK12:AK" & filaExcel & ")"
                'hoja2.Cell(filaExcel + 2, 38).FormulaA1 = "=SUM(AL12:AL" & filaExcel & ")"
                'hoja2.Cell(filaExcel + 2, 39).FormulaA1 = "=SUM(AM12:AM" & filaExcel & ")"


                '<<<<<<<<<<<<<<<xurtep Descanso>>>>>>>>>>>>>>>>>>

                'Limpiar encabezado y relleno
                recorrerFilasColumnas(hoja3, 1, 10, 50, "clear", 13)
                recorrerFilasColumnas(hoja3, 12, dtgDatos.Rows.Count + 30, 50, "clear", 1)

                llenargridD("1")

                ''XURTEP Descanso
                filaExcel = 12
                For x As Integer = 0 To dtgDatos.Rows.Count - 1

                    'Style
                    hoja3.Cell(filaExcel, 1).Style.NumberFormat.Format = "@"
                    hoja3.Cell(filaExcel, 1).Style.NumberFormat.Format = "@"
                    hoja3.Range(filaExcel, 1, filaExcel, 45).Unmerge()
                    hoja3.Range(filaExcel, 1, filaExcel, 45).Style.Font.SetFontColor(XLColor.Black)
                    hoja3.Range(filaExcel, 11, filaExcel, 12).Style.NumberFormat.NumberFormatId = 4
                    hoja3.Range(filaExcel, 14, filaExcel, 45).Style.NumberFormat.NumberFormatId = 4

                    hoja3.Range(filaExcel, 1, filaExcel, 45).Style.Font.SetFontName("Arial")
                    hoja3.Range(filaExcel, 1, filaExcel, 45).Style.Font.SetFontSize(8)
                    hoja3.Range(filaExcel, 1, filaExcel, 45).Style.Font.SetBold(False)


                    hoja3.Range(filaExcel, 1, filaExcel, 11).Style.NumberFormat.Format = "@"
                    hoja3.Cell(filaExcel, 15).Style.NumberFormat.Format = "@"
                    hoja3.Range(filaExcel, 1, filaExcel, 45).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.General)
                    'Datos
                    hoja3.Cell(filaExcel, 1).Value = dtgDatos.Rows(x).Cells(3).Value 'N° Trabajador
                    hoja3.Cell(filaExcel, 2).Value = dtgDatos.Rows(x).Cells(4).Value ' Nombre
                    hoja3.Cell(filaExcel, 3).Value = dtgDatos.Rows(x).Cells(5).Value 'Status
                    hoja3.Cell(filaExcel, 4).Value = dtgDatos.Rows(x).Cells(12).FormattedValue 'buque 
                    hoja3.Cell(filaExcel, 5).Value = dtgDatos.Rows(x).Cells(6).Value 'rfc 
                    hoja3.Cell(filaExcel, 6).Value = dtgDatos.Rows(x).Cells(7).Value 'curp
                    hoja3.Cell(filaExcel, 7).Value = dtgDatos.Rows(x).Cells(8).Value 'imss 
                    hoja3.Cell(filaExcel, 8).Value = dtgDatos.Rows(x).Cells(9).Value 'fecha nac 
                    hoja3.Cell(filaExcel, 9).Value = dtgDatos.Rows(x).Cells(10).Value 'edad
                    hoja3.Cell(filaExcel, 10).Value = dtgDatos.Rows(x).Cells(11).FormattedValue 'puesto
                    hoja3.Cell(filaExcel, 11).Value = dtgDatos.Rows(x).Cells(16).Value 'Salario Diario
                    hoja3.Cell(filaExcel, 12).Value = dtgDatos.Rows(x).Cells(17).Value 'SDI  
                    hoja3.Cell(filaExcel, 13).Value = dtgDatos.Rows(x).Cells(18).Value ' Dias Trabajados 
                    hoja3.Cell(filaExcel, 14).Value = dtgDatos.Rows(x).Cells(21).Value 'Sueldo base
                    hoja3.Cell(filaExcel, 15).Value = dtgDatos.Rows(x).Cells(22).Value ' Tiempo Extra Fijo Gravado 
                    hoja3.Cell(filaExcel, 16).Value = dtgDatos.Rows(x).Cells(23).Value 'Tiempo Extra Fijo Exento
                    hoja3.Cell(filaExcel, 17).Value = dtgDatos.Rows(x).Cells(24).Value ' Tiempo extra ocasional  
                    hoja3.Cell(filaExcel, 18).Value = dtgDatos.Rows(x).Cells(25).Value ' Desc. Sem Oblig.
                    hoja3.Cell(filaExcel, 19).Value = dtgDatos.Rows(x).Cells(26).Value ' VAC. PROPOR 
                    hoja3.Cell(filaExcel, 20).Value = dtgDatos.Rows(x).Cells(27).Value ' AGINALDO GRA 
                    hoja3.Cell(filaExcel, 21).Value = dtgDatos.Rows(x).Cells(28).Value ' AGUINALDO EXENTO 
                    hoja3.Cell(filaExcel, 22).Value = dtgDatos.Rows(x).Cells(29).Value ' TOTAL AGUINALDO 
                    hoja3.Cell(filaExcel, 23).Value = dtgDatos.Rows(x).Cells(30).Value ' P. VAC. GRAVADO 
                    hoja3.Cell(filaExcel, 24).Value = dtgDatos.Rows(x).Cells(31).Value ' P. VAC. EXENTO 
                    hoja3.Cell(filaExcel, 25).Value = dtgDatos.Rows(x).Cells(32).Value ' TOTAL P. VAC 
                    hoja3.Cell(filaExcel, 26).Value = dtgDatos.Rows(x).Cells(33).Value ' TOTAL PERCEPCIONES
                    hoja3.Cell(filaExcel, 27).Value = dtgDatos.Rows(x).Cells(34).Value ' TOTAL PERCEPC P/ISR
                    hoja3.Cell(filaExcel, 28).Value = dtgDatos.Rows(x).Cells(35).Value ' INCAPACIDAD
                    hoja3.Cell(filaExcel, 29).Value = dtgDatos.Rows(x).Cells(36).Value ' ISR
                    hoja3.Cell(filaExcel, 30).Value = dtgDatos.Rows(x).Cells(37).Value ' IMSS
                    hoja3.Cell(filaExcel, 31).Value = dtgDatos.Rows(x).Cells(38).Value ' INFONAVIT
                    hoja3.Cell(filaExcel, 32).Value = dtgDatos.Rows(x).Cells(39).Value ' INFONAVIT
                    hoja3.Cell(filaExcel, 33).Value = dtgDatos.Rows(x).Cells(41).Value ' PENSION ALIMENTICIA
                    hoja3.Cell(filaExcel, 34).Value = dtgDatos.Rows(x).Cells(42).Value ' PRESTAMOS/ANTICIPO NOMINA?
                    hoja3.Cell(filaExcel, 35).Value = dtgDatos.Rows(x).Cells(43).Value ' FONACOT
                    hoja3.Cell(filaExcel, 36).FormulaA1 = "=AB" & filaExcel & "+AC" & filaExcel & "+AD" & filaExcel & "+AE" & filaExcel & "+AF" & filaExcel & "+AG" & filaExcel & "+AH" & filaExcel & "+AI" & filaExcel
                    hoja3.Cell(filaExcel, 37).Value = dtgDatos.Rows(x).Cells(46).Value ' NETO A PAGAR


                    filaExcel = filaExcel + 1


                Next x

                'STYLE
                hoja3.Range(filaExcel + 4, 18, filaExcel + 4, 39).Style.Font.SetFontColor(XLColor.Black)
                hoja3.Range(filaExcel + 4, 18, filaExcel + 4, 39).Style.NumberFormat.NumberFormatId = 4
                hoja3.Range(filaExcel + 4, 18, filaExcel + 4, 39).Style.Font.SetBold(True)

                'Xurtep Descanso
                hoja3.Cell(filaExcel + 2, 13).FormulaA1 = "=SUM(M12:M" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 14).FormulaA1 = "=SUM(N12:N" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 15).FormulaA1 = "=SUM(O12:O" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 16).FormulaA1 = "=SUM(P12:P" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 17).FormulaA1 = "=SUM(Q12:Q" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 18).FormulaA1 = "=SUM(R12:R" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 19).FormulaA1 = "=SUM(S12:S" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 20).FormulaA1 = "=SUM(T12:T" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 21).FormulaA1 = "=SUM(U12:U" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 22).FormulaA1 = "=SUM(V12:V" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 23).FormulaA1 = "=SUM(W12:W" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 24).FormulaA1 = "=SUM(X12:X" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 25).FormulaA1 = "=SUM(Y12:Y" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 26).FormulaA1 = "=SUM(Z12:Z" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 27).FormulaA1 = "=SUM(AA12:AA" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 28).FormulaA1 = "=SUM(AB12:AB" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 29).FormulaA1 = "=SUM(AC12:AC" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 30).FormulaA1 = "=SUM(AD12:AD" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 31).FormulaA1 = "=SUM(AE12:AE" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 32).FormulaA1 = "=SUM(AF12:AF" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 33).FormulaA1 = "=SUM(AG12:AG" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 34).FormulaA1 = "=SUM(AH12:AH" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 35).FormulaA1 = "=SUM(AI12:AI" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 36).FormulaA1 = "=SUM(AJ12:AJ" & filaExcel & ")"
                hoja3.Cell(filaExcel + 2, 37).FormulaA1 = "=SUM(AK12:AK" & filaExcel & ")"
                'hoja3.Cell(filaExcel + 2, 38).FormulaA1 = "=SUM(AL12:AL" & filaExcel & ")"
                'hoja3.Cell(filaExcel + 2, 39).FormulaA1 = "=SUM(AM12:AM" & filaExcel & ")"


                'Titulo
                Dim moment As Date = Date.Now()
                Dim month As Integer = moment.Month
                Dim year As Integer = moment.Year

                pnlProgreso.Visible = False
                pnlCatalogo.Enabled = True

                dialogo.FileName = "MARINOS " + fecha + " " + year.ToString + " OK"
                dialogo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
                ''  dialogo.ShowDialog()

                If dialogo.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    ' OK button pressed
                    libro.SaveAs(dialogo.FileName)
                    libro = Nothing
                    MessageBox.Show("Archivo generado correctamente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                    If cboTipoNomina.SelectedIndex = "0" Then
                        llenargridD("0")
                    Else
                        llenargridD("1")
                    End If
                Else
                    MessageBox.Show("No se guardo el archivo", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Try
    End Sub


    Public Sub recorrerFilasColumnas(ByRef hoja As IXLWorksheet, ByRef filainicio As Integer, ByRef filafinal As Integer, ByRef colTotal As Integer, ByRef tipo As String, Optional ByVal inicioCol As Integer = 1)

        For f As Integer = filainicio To filafinal
            For c As Integer = IIf(inicioCol = Nothing, 1, inicioCol) To colTotal

                Select Case tipo
                    Case "bold"
                        hoja.Cell(f, c).Style.Font.SetFontColor(XLColor.Black)
                    Case "bold false"
                        hoja.Cell(f, c).Style.Font.SetBold(False)
                    Case "clear"
                        hoja.Cell(f, c).Clear()
                    Case "sin relleno"
                        hoja.Cell(f, c).Style.Fill.BackgroundColor = XLColor.NoColor
                    Case "text black"
                        hoja.Cell(f, c).Style.Font.SetFontColor(XLColor.Black)
                End Select
            Next
        Next

    End Sub

    Private Sub llenargridD(ByRef tiponom As String)
        'Cargar grid
        Try
            Dim sql As String
            Dim sql2 As String
            Dim infonavit As Double
            Dim prestamo As Double
            Dim incidencia As Double
            Dim bCalcular As Boolean
            Dim PrimaSA As Double
            Dim cadenabanco As String
            dtgDatos.Columns.Clear()
            dtgDatos.DataSource = Nothing


            dtgDatos.DefaultCellStyle.Font = New Font("Calibri", 8)
            dtgDatos.ColumnHeadersDefaultCellStyle.Font = New Font("Calibri", 9)
            Dim chk As New DataGridViewCheckBoxColumn()
            dtgDatos.Columns.Add(chk)
            chk.HeaderText = ""
            chk.Name = "chk"
            'dtgDatos.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

            'dtgDatos.Columns("chk").SortMode = DataGridViewColumnSortMode.NotSortable

            'dtgDatos.Columns.Add("idempleado", "idempleado")
            'dtgDatos.Columns(0).Width = 30
            'dtgDatos.Columns(0).ReadOnly = True
            ''dtgDatos.Columns(0).DataPropertyName("idempleado")

            'dtgDatos.Columns.Add("departamento", "Departamento")
            'dtgDatos.Columns(1).Width = 100
            'dtgDatos.Columns(1).ReadOnly = True
            'dtgDatos.Columns.Add("nombre", "Trabajador")
            'dtgDatos.Columns(2).Width = 250
            'dtgDatos.Columns(2).ReadOnly = True
            'dtgDatos.Columns.Add("sueldo", "Sueldo Ordinario")
            'dtgDatos.Columns(3).Width = 75
            'dtgDatos.Columns.Add("neto", "Neto")
            'dtgDatos.Columns(4).Width = 75
            'dtgDatos.Columns.Add("infonavit", "Infonavit")
            'dtgDatos.Columns(5).Width = 75
            'dtgDatos.Columns.Add("descuento", "Descuento")
            'dtgDatos.Columns(6).Width = 75
            'dtgDatos.Columns.Add("prestamo", "Prestamo")
            'dtgDatos.Columns(7).Width = 75
            'dtgDatos.Columns.Add("sindicato", "Sindicato")
            'dtgDatos.Columns(8).Width = 75
            'dtgDatos.Columns.Add("neto", "Sueldo Neto")
            'dtgDatos.Columns(9).Width = 75
            'dtgDatos.Columns.Add("imss", "Retención IMSS")
            'dtgDatos.Columns(10).Width = 75
            'dtgDatos.Columns.Add("subsidiado", "Retenciones")
            'dtgDatos.Columns(11).Width = 75
            'dtgDatos.Columns.Add("costosocial", "Costo Social")
            'dtgDatos.Columns(12).Width = 75
            'dtgDatos.Columns.Add("comision", "Comisión")
            'dtgDatos.Columns(13).Width = 75
            'dtgDatos.Columns.Add("subtotal", "Subtotal")
            'dtgDatos.Columns(14).Width = 75
            'dtgDatos.Columns.Add("iva", "IVA")
            'dtgDatos.Columns(15).Width = 75
            'dtgDatos.Columns.Add("total", "Total")
            'dtgDatos.Columns(16).Width = 75


            Dim dsPeriodo As New DataSet
            dsPeriodo.Tables.Add("Tabla")
            dsPeriodo.Tables("Tabla").Columns.Add("Consecutivo")
            dsPeriodo.Tables("Tabla").Columns.Add("Id_empleado")
            dsPeriodo.Tables("Tabla").Columns.Add("CodigoEmpleado")
            dsPeriodo.Tables("Tabla").Columns.Add("Nombre")
            dsPeriodo.Tables("Tabla").Columns.Add("Status")
            dsPeriodo.Tables("Tabla").Columns.Add("RFC")
            dsPeriodo.Tables("Tabla").Columns.Add("CURP")
            dsPeriodo.Tables("Tabla").Columns.Add("Num_IMSS")
            dsPeriodo.Tables("Tabla").Columns.Add("Fecha_Nac")
            dsPeriodo.Tables("Tabla").Columns.Add("Edad")
            dsPeriodo.Tables("Tabla").Columns.Add("Puesto")
            dsPeriodo.Tables("Tabla").Columns.Add("Buque")
            dsPeriodo.Tables("Tabla").Columns.Add("Tipo_Infonavit")
            dsPeriodo.Tables("Tabla").Columns.Add("Valor_Infonavit")
            dsPeriodo.Tables("Tabla").Columns.Add("Sueldo_Base")
            dsPeriodo.Tables("Tabla").Columns.Add("Salario_Diario")
            dsPeriodo.Tables("Tabla").Columns.Add("Salario_Cotización")
            dsPeriodo.Tables("Tabla").Columns.Add("Dias_Trabajados")
            dsPeriodo.Tables("Tabla").Columns.Add("Tipo_Incapacidad")
            dsPeriodo.Tables("Tabla").Columns.Add("Número_días")
            dsPeriodo.Tables("Tabla").Columns.Add("Sueldo_Bruto")
            dsPeriodo.Tables("Tabla").Columns.Add("Aguinaldo_gravado")
            dsPeriodo.Tables("Tabla").Columns.Add("Aguinaldo_exento")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_Aguinaldo")
            dsPeriodo.Tables("Tabla").Columns.Add("Prima_vac_gravado")
            dsPeriodo.Tables("Tabla").Columns.Add("Prima_vac_exento")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_Prima_vac")
            dsPeriodo.Tables("Tabla").Columns.Add("Vacaciones_proporcionales")
            dsPeriodo.Tables("Tabla").Columns.Add("Bono_Puntualidad")
            dsPeriodo.Tables("Tabla").Columns.Add("Bono_Asistencia")
            dsPeriodo.Tables("Tabla").Columns.Add("Fomento_Deporte")
            dsPeriodo.Tables("Tabla").Columns.Add("Bono_Proceso")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_percepciones")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_percepciones_p/isr")
            dsPeriodo.Tables("Tabla").Columns.Add("Incapacidad")
            dsPeriodo.Tables("Tabla").Columns.Add("ISR")
            dsPeriodo.Tables("Tabla").Columns.Add("IMSS")
            dsPeriodo.Tables("Tabla").Columns.Add("Infonavit")
            dsPeriodo.Tables("Tabla").Columns.Add("Infonavit_bim_anterior")
            dsPeriodo.Tables("Tabla").Columns.Add("Ajuste_infonavit")
            dsPeriodo.Tables("Tabla").Columns.Add("Pension_Alimenticia")
            dsPeriodo.Tables("Tabla").Columns.Add("Prestamo")
            dsPeriodo.Tables("Tabla").Columns.Add("Fonacot")
            dsPeriodo.Tables("Tabla").Columns.Add("Subsidio_Generado")
            dsPeriodo.Tables("Tabla").Columns.Add("Subsidio_Aplicado")
            dsPeriodo.Tables("Tabla").Columns.Add("Operadora")
            dsPeriodo.Tables("Tabla").Columns.Add("Prestamo_Personal_A")
            dsPeriodo.Tables("Tabla").Columns.Add("Adeudo_Infonavit_A")
            dsPeriodo.Tables("Tabla").Columns.Add("Diferencia_Infonavit_A")
            dsPeriodo.Tables("Tabla").Columns.Add("Asimilados")
            dsPeriodo.Tables("Tabla").Columns.Add("Retenciones_Operadora")
            dsPeriodo.Tables("Tabla").Columns.Add("%_Comisión")
            dsPeriodo.Tables("Tabla").Columns.Add("Comisión_Operadora")
            dsPeriodo.Tables("Tabla").Columns.Add("Comisión_Asimilados")
            dsPeriodo.Tables("Tabla").Columns.Add("IMSS_CS")
            dsPeriodo.Tables("Tabla").Columns.Add("RCV_CS")
            dsPeriodo.Tables("Tabla").Columns.Add("Infonavit_CS")
            dsPeriodo.Tables("Tabla").Columns.Add("ISN_CS")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_Costo_Social")
            dsPeriodo.Tables("Tabla").Columns.Add("Subtotal")
            dsPeriodo.Tables("Tabla").Columns.Add("IVA")
            dsPeriodo.Tables("Tabla").Columns.Add("TOTAL_DEPOSITO")



            'verificamos que no sea una nomina ya guardada como final
            sql = "select * from NominaProceso inner join EmpleadosC on fkiIdEmpleadoC=iIdEmpleadoC"
            sql &= " where Nomina.fkiIdEmpresa = 1 And fkiIdPeriodo = " & cboperiodo.SelectedValue
            sql &= " and Nomina.iEstatus=1 and iEstatusEmpleado=" & cboserie.SelectedIndex
            sql &= " and iTipoNomina=" & tiponom
            sql &= " order by " & campoordenamiento 'cNombreLargo"
            'sql = "EXEC getNominaXEmpresaXPeriodo " & gIdEmpresa & "," & cboperiodo.SelectedValue & ",1"

            bCalcular = True
            Dim rwNominaGuardada As DataRow() = nConsulta(sql)

            'If rwNominaGuardadaFinal Is Nothing = False Then
            If rwNominaGuardada Is Nothing = False Then
                'Cargamos los datos de guardados como final
                For x As Integer = 0 To rwNominaGuardada.Count - 1

                    Dim fila As DataRow = dsPeriodo.Tables("Tabla").NewRow

                    fila.Item("Consecutivo") = (x + 1).ToString
                    fila.Item("Id_empleado") = rwNominaGuardada(x)("fkiIdEmpleadoC").ToString





                    fila.Item("CodigoEmpleado") = rwNominaGuardada(x)("cCodigoEmpleado").ToString
                    fila.Item("Nombre") = rwNominaGuardada(x)("cNombreLargo").ToString.ToUpper()
                    fila.Item("Status") = IIf(rwNominaGuardada(x)("iOrigen").ToString = "1", "INTERINO", "PLANTA")
                    fila.Item("RFC") = rwNominaGuardada(x)("cRFC").ToString
                    fila.Item("CURP") = rwNominaGuardada(x)("cCURP").ToString
                    fila.Item("Num_IMSS") = rwNominaGuardada(x)("cIMSS").ToString

                    fila.Item("Fecha_Nac") = Date.Parse(rwNominaGuardada(x)("dFechaNac").ToString).ToShortDateString()
                    'Dim tiempo As TimeSpan = Date.Now - Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString)

                    fila.Item("Edad") = CalcularEdad(Date.Parse(rwNominaGuardada(x)("dFechaNac").ToString).Day, Date.Parse(rwNominaGuardada(x)("dFechaNac").ToString).Month, Date.Parse(rwNominaGuardada(x)("dFechaNac").ToString).Year)
                    fila.Item("Puesto") = rwNominaGuardada(x)("Puesto").ToString
                    fila.Item("Buque") = rwNominaGuardada(x)("Buque").ToString

                    fila.Item("Tipo_Infonavit") = rwNominaGuardada(x)("TipoInfonavit").ToString
                    fila.Item("Valor_Infonavit") = rwNominaGuardada(x)("fValor").ToString
                    '
                    fila.Item("Sueldo_Base") = rwNominaGuardada(x)("fSalarioBase").ToString
                    fila.Item("Salario_Diario") = rwNominaGuardada(x)("fSalarioDiario").ToString
                    fila.Item("Salario_Cotización") = rwNominaGuardada(x)("fSalarioBC").ToString


                    fila.Item("Dias_Trabajados") = rwNominaGuardada(x)("iDiasTrabajados").ToString
                    fila.Item("Tipo_Incapacidad") = rwNominaGuardada(x)("TipoIncapacidad").ToString
                    fila.Item("Número_días") = rwNominaGuardada(x)("iNumeroDias").ToString
                    fila.Item("Sueldo_Bruto") = rwNominaGuardada(x)("fSueldoBruto").ToString
                    fila.Item("Aguinaldo_gravado") = rwNominaGuardada(x)("fAguinaldoGravado").ToString
                    fila.Item("Aguinaldo_exento") = rwNominaGuardada(x)("fAguinaldoExento").ToString
                    fila.Item("Total_Aguinaldo") = Math.Round(Double.Parse(rwNominaGuardada(x)("fAguinaldoGravado").ToString) + Double.Parse(rwNominaGuardada(x)("fAguinaldoExento").ToString), 2)
                    fila.Item("Prima_vac_gravado") = rwNominaGuardada(x)("fPrimaVacacionalGravado").ToString
                    fila.Item("Prima_vac_exento") = rwNominaGuardada(x)("fPrimaVacacionalExento").ToString
                    fila.Item("Total_Prima_vac") = Math.Round(Double.Parse(rwNominaGuardada(x)("fPrimaVacacionalGravado").ToString) + Double.Parse(rwNominaGuardada(x)("fPrimaVacacionalExento").ToString), 2)
                    fila.Item("Vacaciones_proporcionales") = rwNominaGuardada(x)("fVacacionesProporcionales").ToString
                    fila.Item("Bono_Puntualidad") = rwNominaGuardada(x)("fBonoPuntualidad").ToString
                    fila.Item("Bono_Asistencia") = rwNominaGuardada(x)("fBonoAsistencia").ToString
                    fila.Item("Fomento_Deporte") = rwNominaGuardada(x)("fFomentoDeporte").ToString
                    fila.Item("Bono_Proceso") = rwNominaGuardada(x)("fBonoProceso").ToString
                    fila.Item("Total_percepciones") = rwNominaGuardada(x)("fTotalPercepciones").ToString
                    fila.Item("Total_percepciones_p/isr") = rwNominaGuardada(x)("fTotalPercepcionesISR").ToString
                    fila.Item("Incapacidad") = rwNominaGuardada(x)("fIncapacidad").ToString
                    fila.Item("ISR") = rwNominaGuardada(x)("fIsr").ToString
                    fila.Item("IMSS") = rwNominaGuardada(x)("fImss").ToString
                    fila.Item("Infonavit") = rwNominaGuardada(x)("fInfonavit").ToString
                    fila.Item("Infonavit_bim_anterior") = rwNominaGuardada(x)("fInfonavitBanterior").ToString
                    fila.Item("Ajuste_infonavit") = rwNominaGuardada(x)("fAjusteInfonavit").ToString
                    fila.Item("Pension_Alimenticia") = rwNominaGuardada(x)("fPensionAlimenticia").ToString
                    fila.Item("Prestamo") = rwNominaGuardada(x)("fPrestamo").ToString
                    fila.Item("Fonacot") = rwNominaGuardada(x)("fFonacot").ToString
                    fila.Item("Subsidio_Generado") = rwNominaGuardada(x)("fSubsidioGenerado").ToString
                    fila.Item("Subsidio_Aplicado") = rwNominaGuardada(x)("fSubsidioAplicado").ToString
                    fila.Item("Operadora") = rwNominaGuardada(x)("fOperadora").ToString
                    fila.Item("Prestamo_Personal_A") = rwNominaGuardada(x)("fPrestamoPerA").ToString
                    fila.Item("Adeudo_Infonavit_A") = rwNominaGuardada(x)("fAdeudoInfonavitA").ToString
                    fila.Item("Diferencia_Infonavit_A") = rwNominaGuardada(x)("fDiferenciaInfonavitA").ToString
                    fila.Item("Asimilados") = rwNominaGuardada(x)("fAsimilados").ToString
                    fila.Item("Retenciones_Operadora") = rwNominaGuardada(x)("fRetencionOperadora").ToString
                    fila.Item("%_Comisión") = rwNominaGuardada(x)("fPorComision").ToString
                    fila.Item("Comisión_Operadora") = rwNominaGuardada(x)("fComisionOperadora").ToString
                    fila.Item("Comisión_Asimilados") = rwNominaGuardada(x)("fComisionAsimilados").ToString
                    fila.Item("IMSS_CS") = rwNominaGuardada(x)("fImssCS").ToString
                    fila.Item("RCV_CS") = rwNominaGuardada(x)("fRcvCS").ToString
                    fila.Item("Infonavit_CS") = rwNominaGuardada(x)("fInfonavitCS").ToString
                    fila.Item("ISN_CS") = rwNominaGuardada(x)("fInsCS").ToString
                    fila.Item("Total_Costo_Social") = rwNominaGuardada(x)("fTotalCostoSocial").ToString
                    fila.Item("Subtotal") = rwNominaGuardada(x)("fSubtotal").ToString
                    fila.Item("IVA") = rwNominaGuardada(x)("fIVA").ToString
                    fila.Item("TOTAL_DEPOSITO") = rwNominaGuardada(x)("fTotalDeposito").ToString


                    dsPeriodo.Tables("Tabla").Rows.Add(fila)
                Next

                dtgDatos.DataSource = dsPeriodo.Tables("Tabla")

                dtgDatos.Columns(0).Width = 30
                dtgDatos.Columns(0).ReadOnly = True
                dtgDatos.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                'consecutivo
                dtgDatos.Columns(1).Width = 60
                dtgDatos.Columns(1).ReadOnly = True
                dtgDatos.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'idempleado
                dtgDatos.Columns(2).Width = 100
                dtgDatos.Columns(2).ReadOnly = True
                dtgDatos.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'codigo empleado
                dtgDatos.Columns(3).Width = 100
                dtgDatos.Columns(3).ReadOnly = True
                dtgDatos.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'Nombre
                dtgDatos.Columns(4).Width = 250
                dtgDatos.Columns(4).ReadOnly = True
                'Estatus
                dtgDatos.Columns(5).Width = 100
                dtgDatos.Columns(5).ReadOnly = True
                'RFC
                dtgDatos.Columns(6).Width = 100
                dtgDatos.Columns(6).ReadOnly = True
                'dtgDatos.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                'CURP
                dtgDatos.Columns(7).Width = 150
                dtgDatos.Columns(7).ReadOnly = True
                'IMSS 

                dtgDatos.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(8).ReadOnly = True
                'Fecha_Nac
                dtgDatos.Columns(9).Width = 150
                dtgDatos.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(9).ReadOnly = True

                'Edad
                dtgDatos.Columns(10).ReadOnly = True
                dtgDatos.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                'Puesto
                dtgDatos.Columns(11).ReadOnly = True
                dtgDatos.Columns(11).Width = 200
                dtgDatos.Columns.Remove("Puesto")

                Dim combo As New DataGridViewComboBoxColumn

                sql = "select * from puestos where iTipo=1 order by cNombre"

                'Dim rwPuestos As DataRow() = nConsulta(sql)
                'If rwPuestos Is Nothing = False Then
                '    combo.Items.Add("uno")
                '    combo.Items.Add("dos")
                '    combo.Items.Add("tres")
                'End If

                nCargaCBO(combo, sql, "cNombre", "iIdPuesto")

                combo.HeaderText = "Puesto"

                combo.Width = 150
                dtgDatos.Columns.Insert(11, combo)
                'DirectCast(dtgDatos.Columns(11), DataGridViewComboBoxColumn).Sorted = True
                'Dim combo2 As New DataGridViewComboBoxCell
                'combo2 = CType(Me.dtgDatos.Rows(2).Cells(11), DataGridViewComboBoxCell)
                'combo2.Value = combo.Items(11)



                'dtgDatos.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                'Buque
                'dtgDatos.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(12).ReadOnly = True
                dtgDatos.Columns(12).Width = 150
                dtgDatos.Columns.Remove("Buque")

                Dim combo2 As New DataGridViewComboBoxColumn

                sql = "select * from departamentos where iEstatus=1 order by cNombre"

                'Dim rwPuestos As DataRow() = nConsulta(sql)
                'If rwPuestos Is Nothing = False Then
                '    combo.Items.Add("uno")
                '    combo.Items.Add("dos")
                '    combo.Items.Add("tres")
                'End If

                nCargaCBO(combo2, sql, "cNombre", "iIdDepartamento")

                combo2.HeaderText = "Buque"
                combo2.Width = 150
                dtgDatos.Columns.Insert(12, combo2)

                'Tipo_Infonavit
                dtgDatos.Columns(13).ReadOnly = True
                dtgDatos.Columns(13).Width = 150
                'dtgDatos.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight



                'Valor_Infonavit
                dtgDatos.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(14).ReadOnly = True
                dtgDatos.Columns(14).Width = 150
                'Sueldo_Base
                dtgDatos.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(15).ReadOnly = True
                dtgDatos.Columns(15).Width = 150
                'Salario_Diario
                dtgDatos.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(16).ReadOnly = True
                dtgDatos.Columns(16).Width = 150
                'Salario_Cotización
                dtgDatos.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(17).ReadOnly = True
                dtgDatos.Columns(17).Width = 150
                'Dias_Trabajados
                dtgDatos.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(18).Width = 150
                'Tipo_Incapacidad
                dtgDatos.Columns(19).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(19).ReadOnly = True
                dtgDatos.Columns(19).Width = 150
                'Número_días
                dtgDatos.Columns(20).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(20).ReadOnly = True
                dtgDatos.Columns(20).Width = 150
                'Sueldo_Bruto
                dtgDatos.Columns(21).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(21).ReadOnly = True
                dtgDatos.Columns(21).Width = 150
                'Tiempo_Extra_Fijo_Gravado
                dtgDatos.Columns(22).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(22).ReadOnly = True
                dtgDatos.Columns(22).Width = 150

                'Tiempo_Extra_Fijo_Exento
                dtgDatos.Columns(23).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(23).ReadOnly = True
                dtgDatos.Columns(23).Width = 150

                'Tiempo_Extra_Ocasional
                dtgDatos.Columns(24).Width = 150
                dtgDatos.Columns(24).ReadOnly = True
                dtgDatos.Columns(24).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'Desc_Sem_Obligatorio
                dtgDatos.Columns(25).Width = 150
                dtgDatos.Columns(25).ReadOnly = True
                dtgDatos.Columns(25).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'Vacaciones_proporcionales
                dtgDatos.Columns(26).Width = 150
                dtgDatos.Columns(26).ReadOnly = True
                dtgDatos.Columns(26).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'Aguinaldo_gravado
                dtgDatos.Columns(27).Width = 150
                dtgDatos.Columns(27).ReadOnly = True
                dtgDatos.Columns(27).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'Aguinaldo_exento
                dtgDatos.Columns(28).Width = 150
                dtgDatos.Columns(28).ReadOnly = True
                dtgDatos.Columns(28).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'Total_Aguinaldo
                dtgDatos.Columns(29).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(29).Width = 150
                dtgDatos.Columns(29).ReadOnly = True

                'Prima_vac_gravado
                dtgDatos.Columns(30).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(30).ReadOnly = True
                dtgDatos.Columns(30).Width = 150
                'Prima_vac_exento 
                dtgDatos.Columns(31).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(31).ReadOnly = True
                dtgDatos.Columns(31).Width = 150

                'Total_Prima_vac
                dtgDatos.Columns(32).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(32).ReadOnly = True
                dtgDatos.Columns(32).Width = 150


                'Total_percepciones
                dtgDatos.Columns(33).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(33).ReadOnly = True
                dtgDatos.Columns(33).Width = 150
                'Total_percepciones_p/isr
                dtgDatos.Columns(34).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(34).ReadOnly = True
                dtgDatos.Columns(34).Width = 150

                'Incapacidad
                dtgDatos.Columns(35).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(35).ReadOnly = True
                dtgDatos.Columns(35).Width = 150

                'ISR
                dtgDatos.Columns(36).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(36).ReadOnly = True
                dtgDatos.Columns(36).Width = 150


                'IMSS
                dtgDatos.Columns(37).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(37).ReadOnly = True
                dtgDatos.Columns(37).Width = 150

                'Infonavit
                dtgDatos.Columns(38).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(38).ReadOnly = True
                dtgDatos.Columns(38).Width = 150
                'Infonavit_bim_anterior
                dtgDatos.Columns(39).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(39).ReadOnly = True
                dtgDatos.Columns(39).Width = 150
                'Ajuste_infonavit
                dtgDatos.Columns(40).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(40).ReadOnly = True
                dtgDatos.Columns(40).Width = 150
                'Pension_Alimenticia
                dtgDatos.Columns(41).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(40).ReadOnly = True
                dtgDatos.Columns(41).Width = 150
                'Prestamo
                dtgDatos.Columns(42).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(42).ReadOnly = True
                dtgDatos.Columns(42).Width = 150
                'Fonacot
                dtgDatos.Columns(43).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(43).ReadOnly = True
                dtgDatos.Columns(43).Width = 150
                'Subsidio_Generado
                dtgDatos.Columns(44).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(44).ReadOnly = True
                dtgDatos.Columns(44).Width = 150
                'Subsidio_Aplicado
                dtgDatos.Columns(45).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(45).ReadOnly = True
                dtgDatos.Columns(45).Width = 150
                'Operadora
                dtgDatos.Columns(46).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(46).ReadOnly = True
                dtgDatos.Columns(46).Width = 150

                'Prestamo Personal Asimilado
                dtgDatos.Columns(47).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(48).ReadOnly = True
                dtgDatos.Columns(47).Width = 150

                'Adeudo_Infonavit_Asimilado
                dtgDatos.Columns(48).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(49).ReadOnly = True
                dtgDatos.Columns(48).Width = 150

                'Difencia infonavit Asimilado
                dtgDatos.Columns(49).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(50).ReadOnly = True
                dtgDatos.Columns(49).Width = 150

                'Complemento Asimilado
                dtgDatos.Columns(50).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(50).ReadOnly = True
                dtgDatos.Columns(50).Width = 150

                'Retenciones_Operadora
                dtgDatos.Columns(51).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(51).ReadOnly = True
                dtgDatos.Columns(51).Width = 150

                '% Comision
                dtgDatos.Columns(52).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(52).ReadOnly = True
                dtgDatos.Columns(52).Width = 150

                'Comision_Operadora
                dtgDatos.Columns(53).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(53).ReadOnly = True
                dtgDatos.Columns(53).Width = 150

                'Comision asimilados
                dtgDatos.Columns(54).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(54).ReadOnly = True
                dtgDatos.Columns(54).Width = 150

                'IMSS_CS
                dtgDatos.Columns(55).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(55).ReadOnly = True
                dtgDatos.Columns(55).Width = 150

                'RCV_CS
                dtgDatos.Columns(56).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(56).ReadOnly = True
                dtgDatos.Columns(56).Width = 150

                'Infonavit_CS
                dtgDatos.Columns(57).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(57).ReadOnly = True
                dtgDatos.Columns(57).Width = 150

                'ISN_CS
                dtgDatos.Columns(58).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(58).ReadOnly = True
                dtgDatos.Columns(58).Width = 150

                'Total Costo Social
                dtgDatos.Columns(59).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(59).ReadOnly = True
                dtgDatos.Columns(59).Width = 150

                'Subtotal
                dtgDatos.Columns(60).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(60).ReadOnly = True
                dtgDatos.Columns(60).Width = 150

                'IVA
                dtgDatos.Columns(61).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(61).ReadOnly = True
                dtgDatos.Columns(61).Width = 150

                'TOTAL DEPOSITO
                dtgDatos.Columns(62).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(62).ReadOnly = True
                dtgDatos.Columns(62).Width = 150

                'calcular()

                'Cambiamos index del combo en el grid

                'For x As Integer = 0 To dtgDatos.Rows.Count - 1

                '    sql = "select * from nomina where fkiIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                '    sql &= " and fkiIdPeriodo=" & cboperiodo.SelectedValue
                '    sql &= " and iEstatusEmpleado=" & cboserie.SelectedIndex
                '    sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex
                '    Dim rwFila As DataRow() = nConsulta(sql)



                '    CType(Me.dtgDatos.Rows(x).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("Puesto").ToString()
                '    CType(Me.dtgDatos.Rows(x).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("Buque").ToString()
                'Next


                'verificar costo social

                Dim contador, Posicion1, Posicion2, Posicion3, Posicion4, Posicion5 As Integer


                For x As Integer = 0 To dtgDatos.Rows.Count - 1
                    contador = 0


                    For y As Integer = 0 To dtgDatos.Rows.Count - 1
                        If dtgDatos.Rows(x).Cells(2).Value = dtgDatos.Rows(y).Cells(2).Value Then
                            contador = contador + 1
                            If contador = 1 Then
                                Posicion1 = y
                            End If
                            If contador = 2 Then
                                Posicion2 = y
                            End If
                            If contador = 3 Then
                                Posicion3 = y
                            End If
                            If contador = 4 Then
                                Posicion4 = y
                            End If
                            If contador = 5 Then
                                Posicion5 = y
                            End If
                        End If



                    Next
                    sql = "select * from Nomina inner join EmpleadosC on fkiIdEmpleadoC=iIdEmpleadoC  where fkiIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                    'sql = "select * from nomina inner join EmpleadosC on nomin where fkiIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                    sql &= " and fkiIdPeriodo=" & cboperiodo.SelectedValue
                    sql &= " and iEstatusEmpleado=" & cboserie.SelectedIndex
                    sql &= " and iTipoNomina=" & tiponom
                    sql &= " order by " & campoordenamiento

                    Dim rwFila As DataRow() = nConsulta(sql)

                    If rwFila.Length = 1 Then
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("Buque").ToString()

                    End If

                    If rwFila.Length = 2 Then
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion2).Cells(11), DataGridViewComboBoxCell).Value = rwFila(1)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion2).Cells(12), DataGridViewComboBoxCell).Value = rwFila(1)("Buque").ToString()

                    End If
                    If rwFila.Length = 3 Then
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion2).Cells(11), DataGridViewComboBoxCell).Value = rwFila(1)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion2).Cells(12), DataGridViewComboBoxCell).Value = rwFila(1)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion3).Cells(11), DataGridViewComboBoxCell).Value = rwFila(2)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion3).Cells(12), DataGridViewComboBoxCell).Value = rwFila(2)("Buque").ToString()
                    End If
                    If rwFila.Length = 4 Then
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion2).Cells(11), DataGridViewComboBoxCell).Value = rwFila(1)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion2).Cells(12), DataGridViewComboBoxCell).Value = rwFila(1)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion3).Cells(11), DataGridViewComboBoxCell).Value = rwFila(2)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion3).Cells(12), DataGridViewComboBoxCell).Value = rwFila(2)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion4).Cells(11), DataGridViewComboBoxCell).Value = rwFila(3)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion4).Cells(12), DataGridViewComboBoxCell).Value = rwFila(3)("Buque").ToString()
                    End If
                    If rwFila.Length = 5 Then
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion1).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion2).Cells(11), DataGridViewComboBoxCell).Value = rwFila(1)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion2).Cells(12), DataGridViewComboBoxCell).Value = rwFila(1)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion3).Cells(11), DataGridViewComboBoxCell).Value = rwFila(2)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion3).Cells(12), DataGridViewComboBoxCell).Value = rwFila(2)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion4).Cells(11), DataGridViewComboBoxCell).Value = rwFila(3)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion4).Cells(12), DataGridViewComboBoxCell).Value = rwFila(3)("Buque").ToString()
                        CType(Me.dtgDatos.Rows(Posicion5).Cells(11), DataGridViewComboBoxCell).Value = rwFila(4)("Puesto").ToString()
                        CType(Me.dtgDatos.Rows(Posicion5).Cells(12), DataGridViewComboBoxCell).Value = rwFila(4)("Buque").ToString()
                    End If
                Next



                'Cambiamos el index del combro de departamentos

                'For x As Integer = 0 To dtgDatos.Rows.Count - 1

                '    sql = "select * from empleadosC where iIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                '    Dim rwFila As DataRow() = nConsulta(sql)




                'Next

                MessageBox.Show("Datos cargados", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)


            Else

                If cboTipoNomina.SelectedIndex = 0 Then
                    If cboserie.SelectedIndex = 0 Then
                        'Buscamos los datos de sindicato solamente
                        sql = "select  * from empleadosC where fkiIdClienteInter=-1"
                        'sql = "select iIdEmpleadoC,NumCuenta, (cApellidoP + ' ' + cApellidoM + ' ' + cNombre) as nombre, fkiIdEmpresa,fSueldoOrd,fCosto from empleadosC"
                        'sql &= " where empleadosC.iOrigen=2 and empleadosC.iEstatus=1"
                        'sql &= " and empleadosC.fkiIdEmpresa =" & gIdEmpresa
                        sql &= " order by cFuncionesPuesto,cNombreLargo"

                    ElseIf cboserie.SelectedIndex > 0 Or cboserie.SelectedIndex - 1 Then
                        sql = "select * from Nomina inner join EmpleadosC on fkiIdEmpleadoC=iIdEmpleadoC"
                        sql &= " where Nomina.fkiIdEmpresa = 1 And fkiIdPeriodo = " & cboperiodo.SelectedValue
                        sql &= " and Nomina.iEstatus=1 and iEstatusEmpleado=20"
                        sql &= " order by cNombreLargo"

                    End If


                    Dim rwDatosEmpleados As DataRow() = nConsulta(sql)
                    If rwDatosEmpleados Is Nothing = False Then
                        For x As Integer = 0 To rwDatosEmpleados.Length - 1


                            Dim fila As DataRow = dsPeriodo.Tables("Tabla").NewRow

                            fila.Item("Consecutivo") = (x + 1).ToString
                            fila.Item("Id_empleado") = rwDatosEmpleados(x)("iIdEmpleadoC").ToString
                            fila.Item("CodigoEmpleado") = rwDatosEmpleados(x)("cCodigoEmpleado").ToString
                            fila.Item("Nombre") = rwDatosEmpleados(x)("cNombreLargo").ToString.ToUpper()
                            fila.Item("Status") = IIf(rwDatosEmpleados(x)("iOrigen").ToString = "1", "INTERINO", "PLANTA")
                            fila.Item("RFC") = rwDatosEmpleados(x)("cRFC").ToString
                            fila.Item("CURP") = rwDatosEmpleados(x)("cCURP").ToString
                            fila.Item("Num_IMSS") = rwDatosEmpleados(x)("cIMSS").ToString

                            fila.Item("Fecha_Nac") = Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString).ToShortDateString()
                            'Dim tiempo As TimeSpan = Date.Now - Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString)
                            fila.Item("Edad") = CalcularEdad(Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString).Day, Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString).Month, Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString).Year)
                            fila.Item("Puesto") = rwDatosEmpleados(x)("cPuesto").ToString
                            fila.Item("Buque") = "ECO III"

                            fila.Item("Tipo_Infonavit") = rwDatosEmpleados(x)("cTipoFactor").ToString
                            fila.Item("Valor_Infonavit") = rwDatosEmpleados(x)("fFactor").ToString
                            fila.Item("Sueldo_Base") = "0.00"
                            fila.Item("Salario_Diario") = rwDatosEmpleados(x)("fSueldoBase").ToString
                            fila.Item("Salario_Cotización") = rwDatosEmpleados(x)("fSueldoIntegrado").ToString
                            fila.Item("Dias_Trabajados") = "30"
                            fila.Item("Tipo_Incapacidad") = TipoIncapacidad(rwDatosEmpleados(x)("iIdEmpleadoC").ToString, cboperiodo.SelectedValue)
                            fila.Item("Número_días") = NumDiasIncapacidad(rwDatosEmpleados(x)("iIdEmpleadoC").ToString, cboperiodo.SelectedValue)
                            fila.Item("Sueldo_Bruto") = ""
                            fila.Item("Tiempo_Extra_Fijo_Gravado") = ""
                            fila.Item("Tiempo_Extra_Fijo_Exento") = ""
                            fila.Item("Tiempo_Extra_Ocasional") = ""
                            fila.Item("Desc_Sem_Obligatorio") = ""
                            fila.Item("Vacaciones_proporcionales") = ""
                            fila.Item("Aguinaldo_gravado") = ""
                            fila.Item("Aguinaldo_exento") = ""
                            fila.Item("Total_Aguinaldo") = ""
                            fila.Item("Prima_vac_gravado") = ""
                            fila.Item("Prima_vac_exento") = ""

                            fila.Item("Total_Prima_vac") = ""
                            fila.Item("Total_percepciones") = ""
                            fila.Item("Total_percepciones_p/isr") = ""
                            fila.Item("Incapacidad") = ""
                            fila.Item("ISR") = ""
                            fila.Item("IMSS") = ""
                            fila.Item("Infonavit") = ""
                            fila.Item("Infonavit_bim_anterior") = ""
                            fila.Item("Ajuste_infonavit") = ""
                            fila.Item("Pension_Alimenticia") = ""
                            fila.Item("Prestamo") = ""
                            fila.Item("Fonacot") = ""
                            fila.Item("Subsidio_Generado") = ""
                            fila.Item("Subsidio_Aplicado") = ""
                            fila.Item("Operadora") = ""
                            fila.Item("Prestamo_Personal_A") = ""
                            fila.Item("Adeudo_Infonavit_A") = ""
                            fila.Item("Diferencia_Infonavit_A") = ""
                            fila.Item("Asimilados") = ""
                            fila.Item("Retenciones_Operadora") = ""
                            fila.Item("%_Comisión") = ""
                            fila.Item("Comisión_Operadora") = ""
                            fila.Item("Comisión_Asimilados") = ""
                            fila.Item("IMSS_CS") = ""
                            fila.Item("RCV_CS") = ""
                            fila.Item("Infonavit_CS") = ""
                            fila.Item("ISN_CS") = ""
                            fila.Item("Total_Costo_Social") = ""
                            fila.Item("Subtotal") = ""
                            fila.Item("IVA") = ""
                            fila.Item("TOTAL_DEPOSITO") = ""


                            dsPeriodo.Tables("Tabla").Rows.Add(fila)




                        Next




                        dtgDatos.DataSource = dsPeriodo.Tables("Tabla")

                        dtgDatos.Columns(0).Width = 30
                        dtgDatos.Columns(0).ReadOnly = True
                        dtgDatos.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                        'consecutivo
                        dtgDatos.Columns(1).Width = 60
                        dtgDatos.Columns(1).ReadOnly = True
                        dtgDatos.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'idempleado
                        dtgDatos.Columns(2).Width = 100
                        dtgDatos.Columns(2).ReadOnly = True
                        dtgDatos.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'codigo empleado
                        dtgDatos.Columns(3).Width = 100
                        dtgDatos.Columns(3).ReadOnly = True
                        dtgDatos.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Nombre
                        dtgDatos.Columns(4).Width = 250
                        dtgDatos.Columns(4).ReadOnly = True
                        'Estatus
                        dtgDatos.Columns(5).Width = 100
                        dtgDatos.Columns(5).ReadOnly = True
                        'RFC
                        dtgDatos.Columns(6).Width = 100
                        dtgDatos.Columns(6).ReadOnly = True
                        'dtgDatos.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                        'CURP
                        dtgDatos.Columns(7).Width = 150
                        dtgDatos.Columns(7).ReadOnly = True
                        'IMSS 

                        dtgDatos.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(8).ReadOnly = True
                        'Fecha_Nac
                        dtgDatos.Columns(9).Width = 150
                        dtgDatos.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(9).ReadOnly = True

                        'Edad
                        dtgDatos.Columns(10).ReadOnly = True
                        dtgDatos.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                        'Puesto
                        dtgDatos.Columns(11).ReadOnly = True
                        dtgDatos.Columns(11).Width = 200
                        dtgDatos.Columns.Remove("Puesto")

                        Dim combo As New DataGridViewComboBoxColumn

                        sql = "select * from puestos where iTipo=1 order by cNombre"

                        'Dim rwPuestos As DataRow() = nConsulta(sql)
                        'If rwPuestos Is Nothing = False Then
                        '    combo.Items.Add("uno")
                        '    combo.Items.Add("dos")
                        '    combo.Items.Add("tres")
                        'End If

                        nCargaCBO(combo, sql, "cNombre", "iIdPuesto")

                        combo.HeaderText = "Puesto"

                        combo.Width = 150
                        dtgDatos.Columns.Insert(11, combo)
                        'DirectCast(dtgDatos.Columns(11), DataGridViewComboBoxColumn).Sorted = True
                        'Dim combo2 As New DataGridViewComboBoxCell
                        'combo2 = CType(Me.dtgDatos.Rows(2).Cells(11), DataGridViewComboBoxCell)
                        'combo2.Value = combo.Items(11)



                        'dtgDatos.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                        'Buque
                        'dtgDatos.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(12).ReadOnly = True
                        dtgDatos.Columns(12).Width = 150
                        dtgDatos.Columns.Remove("Buque")

                        Dim combo2 As New DataGridViewComboBoxColumn

                        sql = "select * from departamentos where iEstatus=1 order by cNombre"

                        'Dim rwPuestos As DataRow() = nConsulta(sql)
                        'If rwPuestos Is Nothing = False Then
                        '    combo.Items.Add("uno")
                        '    combo.Items.Add("dos")
                        '    combo.Items.Add("tres")
                        'End If

                        nCargaCBO(combo2, sql, "cNombre", "iIdDepartamento")

                        combo2.HeaderText = "Buque"
                        combo2.Width = 150
                        dtgDatos.Columns.Insert(12, combo2)

                        'Tipo_Infonavit
                        dtgDatos.Columns(13).ReadOnly = True
                        dtgDatos.Columns(13).Width = 150
                        'dtgDatos.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight



                        'Valor_Infonavit
                        dtgDatos.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(14).ReadOnly = True
                        dtgDatos.Columns(14).Width = 150
                        'Sueldo_Base
                        dtgDatos.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(15).ReadOnly = True
                        dtgDatos.Columns(15).Width = 150
                        'Salario_Diario
                        dtgDatos.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(16).ReadOnly = True
                        dtgDatos.Columns(16).Width = 150
                        'Salario_Cotización
                        dtgDatos.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(17).ReadOnly = True
                        dtgDatos.Columns(17).Width = 150
                        'Dias_Trabajados
                        dtgDatos.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(18).Width = 150
                        'Tipo_Incapacidad
                        dtgDatos.Columns(19).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(19).ReadOnly = True
                        dtgDatos.Columns(19).Width = 150
                        'Número_días
                        dtgDatos.Columns(20).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(20).ReadOnly = True
                        dtgDatos.Columns(20).Width = 150
                        'Sueldo_Bruto
                        dtgDatos.Columns(21).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(21).ReadOnly = True
                        dtgDatos.Columns(21).Width = 150
                        'Tiempo_Extra_Fijo_Gravado
                        dtgDatos.Columns(22).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(22).ReadOnly = True
                        dtgDatos.Columns(22).Width = 150

                        'Tiempo_Extra_Fijo_Exento
                        dtgDatos.Columns(23).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(23).ReadOnly = True
                        dtgDatos.Columns(23).Width = 150

                        'Tiempo_Extra_Ocasional
                        dtgDatos.Columns(24).Width = 150
                        dtgDatos.Columns(24).ReadOnly = True
                        dtgDatos.Columns(24).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Desc_Sem_Obligatorio
                        dtgDatos.Columns(25).Width = 150
                        dtgDatos.Columns(25).ReadOnly = True
                        dtgDatos.Columns(25).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Vacaciones_proporcionales
                        dtgDatos.Columns(26).Width = 150
                        dtgDatos.Columns(26).ReadOnly = True
                        dtgDatos.Columns(26).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Aguinaldo_gravado
                        dtgDatos.Columns(27).Width = 150
                        dtgDatos.Columns(27).ReadOnly = True
                        dtgDatos.Columns(27).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Aguinaldo_exento
                        dtgDatos.Columns(28).Width = 150
                        dtgDatos.Columns(28).ReadOnly = True
                        dtgDatos.Columns(28).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'Total_Aguinaldo
                        dtgDatos.Columns(29).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(29).Width = 150
                        dtgDatos.Columns(29).ReadOnly = True

                        'Prima_vac_gravado
                        dtgDatos.Columns(30).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(30).ReadOnly = True
                        dtgDatos.Columns(30).Width = 150
                        'Prima_vac_exento 
                        dtgDatos.Columns(31).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(31).ReadOnly = True
                        dtgDatos.Columns(31).Width = 150

                        'Total_Prima_vac
                        dtgDatos.Columns(32).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(32).ReadOnly = True
                        dtgDatos.Columns(32).Width = 150


                        'Total_percepciones
                        dtgDatos.Columns(33).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(33).ReadOnly = True
                        dtgDatos.Columns(33).Width = 150
                        'Total_percepciones_p/isr
                        dtgDatos.Columns(34).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(34).ReadOnly = True
                        dtgDatos.Columns(34).Width = 150

                        'Incapacidad
                        dtgDatos.Columns(35).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(35).ReadOnly = True
                        dtgDatos.Columns(35).Width = 150

                        'ISR
                        dtgDatos.Columns(36).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(36).ReadOnly = True
                        dtgDatos.Columns(36).Width = 150


                        'IMSS
                        dtgDatos.Columns(37).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(37).ReadOnly = True
                        dtgDatos.Columns(37).Width = 150

                        'Infonavit
                        dtgDatos.Columns(38).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(38).ReadOnly = True
                        dtgDatos.Columns(38).Width = 150
                        'Infonavit_bim_anterior
                        dtgDatos.Columns(39).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(39).ReadOnly = True
                        dtgDatos.Columns(39).Width = 150
                        'Ajuste_infonavit
                        dtgDatos.Columns(40).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(40).ReadOnly = True
                        dtgDatos.Columns(40).Width = 150
                        'Pension_Alimenticia
                        dtgDatos.Columns(41).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(40).ReadOnly = True
                        dtgDatos.Columns(41).Width = 150
                        'Prestamo
                        dtgDatos.Columns(42).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(42).ReadOnly = True
                        dtgDatos.Columns(42).Width = 150
                        'Fonacot
                        dtgDatos.Columns(43).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(43).ReadOnly = True
                        dtgDatos.Columns(43).Width = 150
                        'Subsidio_Generado
                        dtgDatos.Columns(44).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(44).ReadOnly = True
                        dtgDatos.Columns(44).Width = 150
                        'Subsidio_Aplicado
                        dtgDatos.Columns(45).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(45).ReadOnly = True
                        dtgDatos.Columns(45).Width = 150
                        'Operadora
                        dtgDatos.Columns(46).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(46).ReadOnly = True
                        dtgDatos.Columns(46).Width = 150

                        'Prestamo Personal Asimilado
                        dtgDatos.Columns(47).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(48).ReadOnly = True
                        dtgDatos.Columns(47).Width = 150

                        'Adeudo_Infonavit_Asimilado
                        dtgDatos.Columns(48).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(49).ReadOnly = True
                        dtgDatos.Columns(48).Width = 150

                        'Difencia infonavit Asimilado
                        dtgDatos.Columns(49).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        'dtgDatos.Columns(50).ReadOnly = True
                        dtgDatos.Columns(49).Width = 150

                        'Complemento Asimilado
                        dtgDatos.Columns(50).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(50).ReadOnly = True
                        dtgDatos.Columns(50).Width = 150

                        'Retenciones_Operadora
                        dtgDatos.Columns(51).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(51).ReadOnly = True
                        dtgDatos.Columns(51).Width = 150

                        '% Comision
                        dtgDatos.Columns(52).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(52).ReadOnly = True
                        dtgDatos.Columns(52).Width = 150

                        'Comision_Operadora
                        dtgDatos.Columns(53).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(53).ReadOnly = True
                        dtgDatos.Columns(53).Width = 150

                        'Comision asimilados
                        dtgDatos.Columns(54).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(54).ReadOnly = True
                        dtgDatos.Columns(54).Width = 150

                        'IMSS_CS
                        dtgDatos.Columns(55).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(55).ReadOnly = True
                        dtgDatos.Columns(55).Width = 150

                        'RCV_CS
                        dtgDatos.Columns(56).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(56).ReadOnly = True
                        dtgDatos.Columns(56).Width = 150

                        'Infonavit_CS
                        dtgDatos.Columns(57).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(57).ReadOnly = True
                        dtgDatos.Columns(57).Width = 150

                        'ISN_CS
                        dtgDatos.Columns(58).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(58).ReadOnly = True
                        dtgDatos.Columns(58).Width = 150

                        'Total Costo Social
                        dtgDatos.Columns(59).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(59).ReadOnly = True
                        dtgDatos.Columns(59).Width = 150

                        'Subtotal
                        dtgDatos.Columns(60).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(60).ReadOnly = True
                        dtgDatos.Columns(60).Width = 150

                        'IVA
                        dtgDatos.Columns(61).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(61).ReadOnly = True
                        dtgDatos.Columns(61).Width = 150

                        'TOTAL DEPOSITO
                        dtgDatos.Columns(62).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(62).ReadOnly = True
                        dtgDatos.Columns(62).Width = 150
                        'calcular()

                        'Cambiamos index del combo en el grid

                        For x As Integer = 0 To dtgDatos.Rows.Count - 1

                            sql = "select * from empleadosC where iIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                            Dim rwFila As DataRow() = nConsulta(sql)



                            CType(Me.dtgDatos.Rows(x).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("cPuesto").ToString()
                            CType(Me.dtgDatos.Rows(x).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("cFuncionesPuesto").ToString()
                        Next


                        'Cambiamos el index del combro de departamentos

                        'For x As Integer = 0 To dtgDatos.Rows.Count - 1

                        '    sql = "select * from empleadosC where iIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                        '    Dim rwFila As DataRow() = nConsulta(sql)




                        'Next


                        '  MessageBox.Show("Datos cargados", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("No hay datos en este período", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If




                    'No hay datos en este período
                Else
                    MessageBox.Show("Para la nomina Descanso, solo se mostraran datos guardados, no se podrá calcular de 0", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If




            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub


    Function MonthString(ByRef month As Integer) As String

        Select Case month
            Case 1 : Return "Enero"
            Case 2 : Return "Febrero"
            Case 3 : Return "Marzo"
            Case 4 : Return "Abril"
            Case 5 : Return "Mayo"
            Case 6 : Return "Junio"
            Case 7 : Return "Julio"
            Case 8 : Return "Agosto"
            Case 9 : Return "Septiembre"
            Case 10 : Return "Octubre"
            Case 11 : Return "Noviembre"
            Case 12, 0 : Return "Diciembre"

        End Select

    End Function

    Private Sub btnReporte_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReporte.Click

    End Sub
End Class