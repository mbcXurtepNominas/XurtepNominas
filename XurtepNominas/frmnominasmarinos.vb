Imports ClosedXML.Excel
Imports System.IO

Public Class frmnominasmarinos
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

    Private Sub dvgCombo_SelectedIndexChanged(sender As Object, e As EventArgs)
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

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub frmcontpaqnominas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

    
    

    Private Sub cmdverdatos_Click(sender As Object, e As EventArgs) Handles cmdverdatos.Click
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
            dsPeriodo.Tables("Tabla").Columns.Add("Tiempo_Extra_Fijo_Gravado")
            dsPeriodo.Tables("Tabla").Columns.Add("Tiempo_Extra_Fijo_Exento")
            dsPeriodo.Tables("Tabla").Columns.Add("Tiempo_Extra_Ocasional")
            dsPeriodo.Tables("Tabla").Columns.Add("Desc_Sem_Obligatorio")
            dsPeriodo.Tables("Tabla").Columns.Add("Vacaciones_proporcionales")
            dsPeriodo.Tables("Tabla").Columns.Add("Aguinaldo_gravado")
            dsPeriodo.Tables("Tabla").Columns.Add("Aguinaldo_exento")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_Aguinaldo")
            dsPeriodo.Tables("Tabla").Columns.Add("Prima_vac_gravado")
            dsPeriodo.Tables("Tabla").Columns.Add("Prima_vac_exento")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_Prima_vac")
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
            dsPeriodo.Tables("Tabla").Columns.Add("fecha_inicio")
            dsPeriodo.Tables("Tabla").Columns.Add("fecha_fin")


            'verificamos que no sea una nomina ya guardada como final
            sql = "select * from Nomina inner join EmpleadosC on fkiIdEmpleadoC=iIdEmpleadoC"
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
                    fila.Item("Tiempo_Extra_Fijo_Gravado") = rwNominaGuardada(x)("fTExtraFijoGravado").ToString
                    fila.Item("Tiempo_Extra_Fijo_Exento") = rwNominaGuardada(x)("fTExtraFijoExento").ToString
                    fila.Item("Tiempo_Extra_Ocasional") = rwNominaGuardada(x)("fTExtraOcasional").ToString
                    fila.Item("Desc_Sem_Obligatorio") = rwNominaGuardada(x)("fDescSemObligatorio").ToString
                    fila.Item("Vacaciones_proporcionales") = rwNominaGuardada(x)("fVacacionesProporcionales").ToString
                    fila.Item("Aguinaldo_gravado") = rwNominaGuardada(x)("fAguinaldoGravado").ToString
                    fila.Item("Aguinaldo_exento") = rwNominaGuardada(x)("fAguinaldoExento").ToString
                    fila.Item("Total_Aguinaldo") = Math.Round(Double.Parse(rwNominaGuardada(x)("fAguinaldoGravado").ToString) + Double.Parse(rwNominaGuardada(x)("fAguinaldoExento").ToString), 2)
                    fila.Item("Prima_vac_gravado") = rwNominaGuardada(x)("fPrimaVacacionalGravado").ToString
                    fila.Item("Prima_vac_exento") = rwNominaGuardada(x)("fPrimaVacacionalExento").ToString

                    fila.Item("Total_Prima_vac") = Math.Round(Double.Parse(rwNominaGuardada(x)("fPrimaVacacionalGravado").ToString) + Double.Parse(rwNominaGuardada(x)("fPrimaVacacionalExento").ToString), 2)
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
                    fila.Item("Operadora") = rwNominaGuardada(x)("fXurtep").ToString
                    fila.Item("Prestamo_Personal_A") = rwNominaGuardada(x)("fPrestamoPerA").ToString
                    fila.Item("Adeudo_Infonavit_A") = rwNominaGuardada(x)("fAdeudoInfonavitA").ToString
                    fila.Item("Diferencia_Infonavit_A") = rwNominaGuardada(x)("fDiferenciaInfonavitA").ToString
                    fila.Item("Asimilados") = rwNominaGuardada(x)("fAsimilados").ToString
                    fila.Item("Retenciones_Operadora") = rwNominaGuardada(x)("fRetencionOperadora").ToString
                    fila.Item("%_Comisión") = rwNominaGuardada(x)("fPorComision").ToString
                    fila.Item("Comisión_Operadora") = rwNominaGuardada(x)("fComisionXurtep").ToString
                    fila.Item("Comisión_Asimilados") = rwNominaGuardada(x)("fComisionAsimilados").ToString
                    fila.Item("IMSS_CS") = rwNominaGuardada(x)("fImssCS").ToString
                    fila.Item("RCV_CS") = rwNominaGuardada(x)("fRcvCS").ToString
                    fila.Item("Infonavit_CS") = rwNominaGuardada(x)("fInfonavitCS").ToString
                    fila.Item("ISN_CS") = rwNominaGuardada(x)("fInsCS").ToString
                    fila.Item("Total_Costo_Social") = rwNominaGuardada(x)("fTotalCostoSocial").ToString
                    fila.Item("Subtotal") = rwNominaGuardada(x)("fSubtotal").ToString
                    fila.Item("IVA") = rwNominaGuardada(x)("fIVA").ToString
                    fila.Item("TOTAL_DEPOSITO") = rwNominaGuardada(x)("fTotalDeposito").ToString
                    fila.Item("fecha_inicio") = Date.Parse(rwNominaGuardada(x)("dFechaInicial").ToString).ToShortDateString
                    fila.Item("fecha_fin") = Date.Parse(rwNominaGuardada(x)("dFechaFinal").ToString).ToShortDateString

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
                'dtgDatos.Columns(5).ReadOnly = True
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
                'dtgDatos.Columns(13).ReadOnly = True
                dtgDatos.Columns(13).Width = 150
                'dtgDatos.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight



                'Valor_Infonavit
                dtgDatos.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                'dtgDatos.Columns(14).ReadOnly = True
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


                'FECHA INICIAL
                dtgDatos.Columns(63).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(63).ReadOnly = True
                dtgDatos.Columns(63).Width = 150

                'FECHA FINAL
                dtgDatos.Columns(64).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dtgDatos.Columns(64).ReadOnly = True
                dtgDatos.Columns(64).Width = 150
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
                            fila.Item("fecha_inicio") = ""
                            fila.Item("fecha_fin") = ""

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

                        'FECHA INICIAL
                        dtgDatos.Columns(63).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(63).ReadOnly = True
                        dtgDatos.Columns(63).Width = 150

                        'FECHA FINAL
                        dtgDatos.Columns(64).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dtgDatos.Columns(64).ReadOnly = True
                        dtgDatos.Columns(64).Width = 150
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

    Private Function TipoIncapacidad(idempleado As String, periodo As Integer) As String
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

    Private Function NumDiasIncapacidad(idempleado As String, periodo As Integer) As String
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

    Private Function Identificadorincapacidad(identificador As String) As String
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


    Private Sub cmdguardarnomina_Click(sender As Object, e As EventArgs) Handles cmdguardarnomina.Click
        Try
            Dim sql As String
            Dim sql2 As String
            sql = "select * from Nomina where fkiIdEmpresa=1 and fkiIdPeriodo=" & cboperiodo.SelectedValue
            sql &= " and iEstatusNomina=1 and iEstatus=1 and iEstatusEmpleado=" & cboserie.SelectedIndex
            sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex
            'Dim sueldobase, salariodiario, salariointegrado, sueldobruto, TiempoExtraFijoGravado, TiempoExtraFijoExento As Double
            'Dim TiempoExtraOcasional, DesSemObligatorio, VacacionesProporcionales, AguinaldoGravado, AguinaldoExento As Double
            'Dim PrimaVacGravada, PrimaVacExenta, TotalPercepciones, TotalPercepcionesISR As Double
            'Dim incapacidad, ISR, IMSS, Infonavit, InfonavitAnterior, InfonavitAjuste, PensionAlimenticia As Double
            'Dim Prestamo, Fonacot, NetoaPagar, Excedente, Total, ImssCS, RCVCS, InfonavitCS, ISNCS
            'sql = "EXEC getNominaXEmpresaXPeriodo " & gIdEmpresa & "," & cboperiodo.SelectedValue & ",1"

            Dim rwNominaGuardadaFinal As DataRow() = nConsulta(sql)

            If rwNominaGuardadaFinal Is Nothing = False Then
                MessageBox.Show("La nomina ya esta marcada como final, no  se pueden guardar cambios", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else


                sql = "delete from Nomina"
                sql &= " where fkiIdEmpresa=1 and fkiIdPeriodo=" & cboperiodo.SelectedValue
                sql &= " and iEstatusNomina=0 and iEstatus=1 and iEstatusEmpleado=" & cboserie.SelectedIndex
                sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex
                If nExecute(sql) = False Then
                    MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    'pnlProgreso.Visible = False
                    Exit Sub
                End If

                sql = "delete from DetalleDescInfonavit"
                sql &= " where fkiIdPeriodo=" & cboperiodo.SelectedValue
                sql &= " and iSerie=" & cboserie.SelectedIndex
                'sql &= " and iSerie=" & cboserie.SelectedIndex
                sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex

                If nExecute(sql) = False Then
                    MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    'pnlProgreso.Visible = False
                    Exit Sub
                End If

                pnlProgreso.Visible = True

                Application.DoEvents()
                'pnlCatalogo.Enabled = False
                pgbProgreso.Minimum = 0
                pgbProgreso.Value = 0
                pgbProgreso.Maximum = dtgDatos.Rows.Count

                For x As Integer = 0 To dtgDatos.Rows.Count - 1



                    sql = "EXEC [setNominaInsertar ] 0"
                    'periodo
                    sql &= "," & cboperiodo.SelectedValue
                    'idempleado
                    sql &= "," & dtgDatos.Rows(x).Cells(2).Value
                    'idempresa
                    sql &= ",1"
                    'Puesto
                    'buscamos el valor en la tabla
                    sql2 = "select * from puestos where cNombre='" & dtgDatos.Rows(x).Cells(11).FormattedValue & "'"

                    Dim rwPuesto As DataRow() = nConsulta(sql2)

                    sql &= "," & rwPuesto(0)("iIdPuesto")


                    'departamento
                    'buscamos el valor en la tabla
                    sql2 = "select * from departamentos where cNombre='" & dtgDatos.Rows(x).Cells(12).FormattedValue & "'"

                    Dim rwDepto As DataRow() = nConsulta(sql2)

                    sql &= "," & rwDepto(0)("iIdDepartamento")

                    'estatus empleado
                    sql &= "," & cboserie.SelectedIndex
                    'edad
                    sql &= "," & dtgDatos.Rows(x).Cells(10).Value
                    'puesto
                    sql &= ",'" & dtgDatos.Rows(x).Cells(11).FormattedValue & "'"
                    'buque
                    sql &= ",'" & dtgDatos.Rows(x).Cells(12).FormattedValue & "'"
                    'iTipo Infonavit
                    sql &= ",'" & dtgDatos.Rows(x).Cells(13).Value & "'"
                    'valor infonavit
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(14).Value = "", "0", dtgDatos.Rows(x).Cells(14).Value.ToString.Replace(",", ""))
                    'salario base
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(15).Value = "", "0", dtgDatos.Rows(x).Cells(15).Value.ToString.Replace(",", ""))
                    'salario diario
                    sql &= "," & dtgDatos.Rows(x).Cells(16).Value
                    'salario integrado
                    sql &= "," & dtgDatos.Rows(x).Cells(17).Value
                    'Dias trabajados
                    sql &= "," & dtgDatos.Rows(x).Cells(18).Value
                    'tipo incapacidad

                    sql &= ",'" & dtgDatos.Rows(x).Cells(19).Value & "'"
                    'numero dias incapacidad
                    sql &= "," & dtgDatos.Rows(x).Cells(20).Value
                    'sueldobruto
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(21).Value = "", "0", dtgDatos.Rows(x).Cells(21).Value.ToString.Replace(",", ""))
                    'tiempo extra fijo gravado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(22).Value = "", "0", dtgDatos.Rows(x).Cells(22).Value.ToString.Replace(",", ""))
                    'tiempo extra fijo exento
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(23).Value = "", "0", dtgDatos.Rows(x).Cells(23).Value.ToString.Replace(",", ""))
                    'Tiempo extra ocasional
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(24).Value = "", "0", dtgDatos.Rows(x).Cells(24).Value.ToString.Replace(",", ""))
                    'descanso semanal obligatorio
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(25).Value = "", "0", dtgDatos.Rows(x).Cells(25).Value.ToString.Replace(",", ""))
                    'vacaciones proporcionales
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(26).Value = "", "0", dtgDatos.Rows(x).Cells(26).Value.ToString.Replace(",", ""))
                    'aguinaldo gravado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(27).Value = "", "0", dtgDatos.Rows(x).Cells(27).Value.ToString.Replace(",", ""))
                    'aguinaldo exento
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(28).Value = "", "0", dtgDatos.Rows(x).Cells(28).Value.ToString.Replace(",", ""))
                    'prima vacacional gravado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(30).Value = "", "0", dtgDatos.Rows(x).Cells(30).Value.ToString.Replace(",", ""))
                    'prima vacacional exento
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(31).Value = "", "0", dtgDatos.Rows(x).Cells(31).Value.ToString.Replace(",", ""))

                    'totalpercepciones
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(33).Value = "", "0", dtgDatos.Rows(x).Cells(33).Value.ToString.Replace(",", ""))
                    'totalpercepcionesISR
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(34).Value = "", "0", dtgDatos.Rows(x).Cells(34).Value.ToString.Replace(",", ""))
                    'Incapacidad
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(35).Value = "", "0", dtgDatos.Rows(x).Cells(35).Value.ToString.Replace(",", ""))
                    'isr
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(36).Value = "", "0", dtgDatos.Rows(x).Cells(36).Value.ToString.Replace(",", ""))
                    'imss
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(37).Value = "", "0", dtgDatos.Rows(x).Cells(37).Value.ToString.Replace(",", ""))
                    'infonavit
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(38).Value = "", "0", dtgDatos.Rows(x).Cells(38).Value.ToString.Replace(",", ""))
                    'infonavit anterior
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(39).Value = "", "0", dtgDatos.Rows(x).Cells(39).Value.ToString.Replace(",", ""))
                    'ajuste infonavit
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(40).Value = "", "0", dtgDatos.Rows(x).Cells(40).Value.ToString.Replace(",", ""))
                    'Pension alimenticia
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(41).Value = "", "0", dtgDatos.Rows(x).Cells(41).Value.ToString.Replace(",", ""))
                    'Prestamo
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(42).Value = "", "0", dtgDatos.Rows(x).Cells(42).Value.ToString.Replace(",", ""))
                    'Fonacot
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(43).Value = "", "0", dtgDatos.Rows(x).Cells(43).Value.ToString.Replace(",", ""))
                    'Subsidio Generado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(44).Value = "", "0", dtgDatos.Rows(x).Cells(44).Value.ToString.Replace(",", ""))
                    'Subsidio Aplicado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(45).Value = "", "0", dtgDatos.Rows(x).Cells(45).Value.ToString.Replace(",", ""))
                    'Operadora
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(46).Value = "", "0", dtgDatos.Rows(x).Cells(46).Value.ToString.Replace(",", ""))
                    'Prestamo Personal Asimilado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(47).Value = "", "0", dtgDatos.Rows(x).Cells(47).Value.ToString.Replace(",", ""))
                    'Adeudo_Infonavit_Asimilado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(48).Value = "", "0", dtgDatos.Rows(x).Cells(48).Value.ToString.Replace(",", ""))
                    'Difencia infonavit Asimilado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(49).Value = "", "0", dtgDatos.Rows(x).Cells(49).Value.ToString.Replace(",", ""))
                    'Complemento Asimilado
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(50).Value = "", "0", dtgDatos.Rows(x).Cells(50).Value.ToString.Replace(",", ""))
                    'Retenciones_Operadora
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(51).Value = "", "0", dtgDatos.Rows(x).Cells(51).Value.ToString.Replace(",", ""))
                    '% Comision
                    sql &= ",0.02" '& IIf(dtgDatos.Rows(x).Cells(52).Value = "", "0", dtgDatos.Rows(x).Cells(52).Value.ToString.Replace(",", ""))
                    'Comision_Operadora
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(53).Value = "", "0", dtgDatos.Rows(x).Cells(53).Value.ToString.Replace(",", ""))
                    'Comision asimilados
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(54).Value = "", "0", dtgDatos.Rows(x).Cells(54).Value.ToString.Replace(",", ""))
                    'IMSS_CS
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(55).Value = "", "0", dtgDatos.Rows(x).Cells(55).Value.ToString.Replace(",", ""))
                    'RCV_CS
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(56).Value = "", "0", dtgDatos.Rows(x).Cells(56).Value.ToString.Replace(",", ""))
                    'Infonavit_CS
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(57).Value = "", "0", dtgDatos.Rows(x).Cells(57).Value.ToString.Replace(",", ""))
                    'ISN_CS
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(58).Value = "", "0", dtgDatos.Rows(x).Cells(58).Value.ToString.Replace(",", ""))
                    'Total Costo Social
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(59).Value = "", "0", dtgDatos.Rows(x).Cells(59).Value.ToString.Replace(",", ""))
                    'Subtotal
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(60).Value = "", "0", dtgDatos.Rows(x).Cells(60).Value.ToString.Replace(",", ""))
                    'IVA
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(61).Value = "", "0", dtgDatos.Rows(x).Cells(61).Value.ToString.Replace(",", ""))
                    'TOTAL DEPOSITO
                    sql &= "," & IIf(dtgDatos.Rows(x).Cells(62).Value = "", "0", dtgDatos.Rows(x).Cells(62).Value.ToString.Replace(",", ""))
                    'Estatus
                    sql &= ",1"
                    'Estatus Nomina
                    sql &= ",0"
                    'Tipo Nomina
                    sql &= "," & cboTipoNomina.SelectedIndex
                    'fechainicial
                    sql &= ",'" & IIf(dtgDatos.Rows(x).Cells(63).Value = "", "0", dtgDatos.Rows(x).Cells(63).Value.ToString.Replace(",", ""))
                    'fechafinal
                    sql &= "','" & IIf(dtgDatos.Rows(x).Cells(64).Value = "", "0", dtgDatos.Rows(x).Cells(64).Value.ToString.Replace(",", "")) & "'"







                    If nExecute(sql) = False Then
                        MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        'pnlProgreso.Visible = False
                        Exit Sub
                    End If

                    '########GUARDAR INFONAVIT

                    Dim numbimestre As Integer
                    If Month(FechaInicioPeriodoGlobal) Mod 2 = 0 Then
                        numbimestre = Month(FechaInicioPeriodoGlobal) / 2
                    Else
                        numbimestre = (Month(FechaInicioPeriodoGlobal) + 1) / 2
                    End If


                    If Double.Parse(IIf(dtgDatos.Rows(x).Cells(38).Value = "", "0", dtgDatos.Rows(x).Cells(38).Value)) > 0 Then

                        Dim MontoInfonavit As Double = MontoInfonavitF(cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value))

                        sql = "EXEC setDetalleDescInfonavitInsertar  0"
                        'fk Calculo infonavit
                        sql &= "," & IIf(MontoInfonavit > 0, IDCalculoInfonavit, 0)
                        'Cantidad
                        sql &= "," & dtgDatos.Rows(x).Cells(38).Value
                        ' fk Empleado
                        sql &= "," & dtgDatos.Rows(x).Cells(2).Value
                        'Numbimestre
                        sql &= "," & numbimestre
                        'Anio
                        sql &= "," & FechaInicioPeriodoGlobal.Year
                        'fk Periodo
                        sql &= "," & cboperiodo.SelectedValue
                        'Serie
                        sql &= "," & cboserie.SelectedIndex
                        'Tipo Nomina
                        sql &= "," & cboTipoNomina.SelectedIndex
                        'Tipo Pagadora
                        sql &= ",101"
                        'iEstatu
                        sql &= ",1"

                        If nExecute(sql) = False Then
                            MessageBox.Show("Ocurrio un error insertar pago prestamo ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            'pnlProgreso.Visible = False
                            Exit Sub
                        End If
                    End If

                    If Double.Parse(IIf(dtgDatos.Rows(x).Cells(39).Value = "", "0", dtgDatos.Rows(x).Cells(39).Value)) > 0 Then

                        Dim MontoInfonavit As Double = MontoInfonavitF(cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value))

                        sql = "EXEC setDetalleDescInfonavitInsertar  0"
                        'fk Calculo infonavit
                        sql &= "," & IIf(MontoInfonavit > 0, IDCalculoInfonavit, 0)
                        'Cantidad
                        sql &= "," & dtgDatos.Rows(x).Cells(39).Value
                        ' fk Empleado
                        sql &= "," & dtgDatos.Rows(x).Cells(2).Value
                        'Numbimestre
                        sql &= "," & numbimestre
                        'Anio
                        sql &= "," & FechaInicioPeriodoGlobal.Year
                        'fk Periodo
                        sql &= "," & cboperiodo.SelectedValue
                        'Serie
                        sql &= "," & cboserie.SelectedIndex
                        'Tipo Nomina
                        sql &= "," & cboTipoNomina.SelectedIndex
                        'Tipo Pagadora
                        sql &= ",102"
                        'iEstatu
                        sql &= ",1"

                        If nExecute(sql) = False Then
                            MessageBox.Show("Ocurrio un error insertar pago prestamo ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            'pnlProgreso.Visible = False
                            Exit Sub
                        End If
                    End If

                    If Double.Parse(IIf(dtgDatos.Rows(x).Cells(40).Value = "", "0", dtgDatos.Rows(x).Cells(40).Value)) > 0 Then

                        Dim MontoInfonavit As Double = MontoInfonavitF(cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value))

                        sql = "EXEC setDetalleDescInfonavitInsertar  0"
                        'fk Calculo infonavit
                        sql &= "," & IIf(MontoInfonavit > 0, IDCalculoInfonavit, 0)
                        'Cantidad
                        sql &= "," & dtgDatos.Rows(x).Cells(40).Value
                        ' fk Empleado
                        sql &= "," & dtgDatos.Rows(x).Cells(2).Value
                        'Numbimestre
                        sql &= "," & numbimestre
                        'Anio
                        sql &= "," & FechaInicioPeriodoGlobal.Year
                        'fk Periodo
                        sql &= "," & cboperiodo.SelectedValue
                        'Serie
                        sql &= "," & cboserie.SelectedIndex
                        'Tipo Nomina
                        sql &= "," & cboTipoNomina.SelectedIndex
                        'Tipo Pagadora
                        sql &= ",103"
                        'iEstatu
                        sql &= ",1"

                        If nExecute(sql) = False Then
                            MessageBox.Show("Ocurrio un error insertar pago prestamo ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            'pnlProgreso.Visible = False
                            Exit Sub
                        End If
                    End If
                    '########GUARDAR SEGURO INFONAVIT
                    'sql = "select * from periodos where iIdPeriodo= " & cboperiodo.SelectedValue
                    'Dim rwPeriodo As DataRow() = nConsulta(sql)

                    'Dim FechaInicioPeriodo1 As Date


                    'Dim numbimestre As Integer
                    'If rwPeriodo Is Nothing = False Then
                    '    FechaInicioPeriodo1 = Date.Parse(rwPeriodo(0)("dFechaInicio"))

                    '    If Month(FechaInicioPeriodo1) Mod 2 = 0 Then
                    '        numbimestre = Month(FechaInicioPeriodo1) / 2
                    '    Else
                    '        numbimestre = (Month(FechaInicioPeriodo1) + 1) / 2
                    '    End If

                    'End If


                    'sql = "select * from PagoSeguroInfonavit where fkiIdEmpleadoC= " & dtgDatos.Rows(x).Cells(2).Value
                    'sql &= " And NumBimestre= " & numbimestre & " And Anio=" & FechaInicioPeriodo1.Year.ToString
                    'Dim rwSeguro1 As DataRow() = nConsulta(sql)

                    'If rwSeguro1 Is Nothing = True Then
                    '    'Insertar seguro
                    '    sql = "EXEC setPagoSeguroInfonavitInsertar  0"
                    '    ' fk Empleado
                    '    sql &= "," & dtgDatos.Rows(x).Cells(2).Value
                    '    'bimestre
                    '    sql &= "," & numbimestre
                    '    ' anio
                    '    sql &= ",'" & FechaInicioPeriodo1.Year.ToString


                    'End If






                    'sql = "update empleadosC set fSueldoOrd=" & dtgDatos.Rows(x).Cells(6).Value & ", fCosto =" & dtgDatos.Rows(x).Cells(18).Value
                    'sql &= " where iIdEmpleadoC = " & dtgDatos.Rows(x).Cells(2).Value

                    'If nExecute(sql) = False Then
                    '    MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    '    'pnlProgreso.Visible = False
                    '    Exit Sub
                    'End If

                    pgbProgreso.Value += 1
                    Application.DoEvents()
                Next
                pnlProgreso.Visible = False
                pnlCatalogo.Enabled = True

                If cboTipoNomina.SelectedIndex = 0 Then
                    MessageBox.Show("Datos guardados correctamente, se generara la nomina descanso", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    NominaB()
                    MessageBox.Show("Nomina Descanso generado, si no hay cambios proceda a guardar", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Datos guardados correctamente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If


            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    Private Sub NominaB()
        cboTipoNomina.SelectedIndex = 1
        For x As Integer = 0 To dtgDatos.Rows.Count - 1
            'Dim cadena As String = dgvCombo.Text
            If dtgDatos.Rows(x).Cells(11).FormattedValue = "OFICIALES EN PRACTICAS: PILOTIN / ASPIRANTE" Then
                dtgDatos.Rows(x).Cells(15).Value = "0.00"
                dtgDatos.Rows(x).Cells(18).Value = "0.00"
                dtgDatos.Rows(x).Cells(21).Value = "0.00"
                dtgDatos.Rows(x).Cells(22).Value = "0.00"
                dtgDatos.Rows(x).Cells(23).Value = "0.00"
                dtgDatos.Rows(x).Cells(24).Value = "0.00"
                dtgDatos.Rows(x).Cells(25).Value = "0.00"
                dtgDatos.Rows(x).Cells(26).Value = "0.00"
                dtgDatos.Rows(x).Cells(27).Value = "0.00"
                dtgDatos.Rows(x).Cells(28).Value = "0.00"
                dtgDatos.Rows(x).Cells(29).Value = "0.00"
                dtgDatos.Rows(x).Cells(30).Value = "0.00"
                dtgDatos.Rows(x).Cells(31).Value = "0.00"
                dtgDatos.Rows(x).Cells(32).Value = "0.00"
                dtgDatos.Rows(x).Cells(33).Value = "0.00"
                dtgDatos.Rows(x).Cells(34).Value = "0.00"
                dtgDatos.Rows(x).Cells(35).Value = "0.00"
                'ISR
                dtgDatos.Rows(x).Cells(36).Value = "0.00"
                'IMSS
                dtgDatos.Rows(x).Cells(37).Value = "0.00"
                'INFONAVIT
                dtgDatos.Rows(x).Cells(38).Value = "0.00"
                'INFONAVIT BIMESTRE ANTERIOR
                dtgDatos.Rows(x).Cells(38).Value = "0.00"
                'AJUSTE INFONAVIT
                dtgDatos.Rows(x).Cells(38).Value = "0.00"
                'PENSION
                dtgDatos.Rows(x).Cells(38).Value = "0.00"
                'PRESTAMO
                dtgDatos.Rows(x).Cells(38).Value = "0.00"
                'FONACOT
                dtgDatos.Rows(x).Cells(38).Value = "0.00"
                'SUBSIDIO GENERADO
                dtgDatos.Rows(x).Cells(44).Value = "0.00"
                'SUBSIDIO APLICADO
                dtgDatos.Rows(x).Cells(45).Value = "0.00"
                'NETO
                dtgDatos.Rows(x).Cells(46).Value = "0.00"
                'Prestamo Personal Asimilado
                dtgDatos.Rows(x).Cells(47).Value = "0.00"
                'Adeudo_Infonavit_Asimilado
                dtgDatos.Rows(x).Cells(48).Value = "0.00"
                'Difencia infonavit Asimilado
                dtgDatos.Rows(x).Cells(49).Value = "0.00"
                'Complemento Asimilado
                dtgDatos.Rows(x).Cells(50).Value = "0.00"
                'Retenciones_Operadora
                dtgDatos.Rows(x).Cells(51).Value = "0.00"
                '% Comision
                dtgDatos.Rows(x).Cells(52).Value = "0.00"
                'Comision_Operadora
                dtgDatos.Rows(x).Cells(53).Value = "0.00"
                'Comision asimilados
                dtgDatos.Rows(x).Cells(54).Value = "0.00"
                'IMSS_CS
                dtgDatos.Rows(x).Cells(55).Value = "0.00"
                'RCV_CS
                dtgDatos.Rows(x).Cells(56).Value = "0.00"
                'Infonavit_CS
                dtgDatos.Rows(x).Cells(57).Value = "0.00"
                'ISN_CS
                dtgDatos.Rows(x).Cells(58).Value = "0.00"
                'Total Costo Social
                dtgDatos.Rows(x).Cells(59).Value = "0.00"
                'Subtotal
                dtgDatos.Rows(x).Cells(60).Value = "0.00"
                'IVA
                dtgDatos.Rows(x).Cells(61).Value = "0.00"
                'TOTAL DEPOSITO
                dtgDatos.Rows(x).Cells(62).Value = "0.00"

            Else
                dtgDatos.Rows(x).Cells(15).Value = "0.00"

            End If




        Next
        calcular()

    End Sub

    Private Sub cmdcalcular_Click(sender As Object, e As EventArgs) Handles cmdcalcular.Click
        Try
            Dim sql As String
            Dim sql2 As String
            Dim sql3 As String
            Dim sql4 As String
            Dim sql5 As String

            sql = "select * from Nomina where fkiIdEmpresa=1 and fkiIdPeriodo=" & cboperiodo.SelectedValue
            sql &= " and iEstatusNomina=1 and iEstatus=1 and iEstatusEmpleado=" & cboserie.SelectedIndex
            sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex
            'Dim sueldobase, salariodiario, salariointegrado, sueldobruto, TiempoExtraFijoGravado, TiempoExtraFijoExento As Double
            'Dim TiempoExtraOcasional, DesSemObligatorio, VacacionesProporcionales, AguinaldoGravado, AguinaldoExento As Double
            'Dim PrimaVacGravada, PrimaVacExenta, TotalPercepciones, TotalPercepcionesISR As Double
            'Dim incapacidad, ISR, IMSS, Infonavit, InfonavitAnterior, InfonavitAjuste, PensionAlimenticia As Double
            'Dim Prestamo, Fonacot, NetoaPagar, Excedente, Total, ImssCS, RCVCS, InfonavitCS, ISNCS
            'sql = "EXEC getNominaXEmpresaXPeriodo " & gIdEmpresa & "," & cboperiodo.SelectedValue & ",1"

            Dim rwNominaGuardadaFinal As DataRow() = nConsulta(sql)

            If rwNominaGuardadaFinal Is Nothing = False Then
                MessageBox.Show("La nomina ya esta marcada como final, no  se puede calcular", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If cboTipoNomina.SelectedIndex = 0 Then
                    sql = "delete from DetalleDescInfonavit"
                    sql &= " where fkiIdPeriodo=" & cboperiodo.SelectedValue
                    sql &= " and iSerie=" & cboserie.SelectedIndex



                    sql2 = " delete from DetallePensionAlimenticia"
                    sql2 &= " where fkiIdPeriodo=" & cboperiodo.SelectedValue
                    sql2 &= " and iSerie=" & cboserie.SelectedIndex

                    'sql &= " and iSerie=" & cboserie.SelectedIndex
                    'sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex


                    '' borrar el seguro si solo tiene un registro
                    'For x As Integer = 0 To dtgDatos.Rows.Count - 1
                    '    Dim ValorInfo As Double
                    '    ValorInfo = IIf(dtgDatos.Rows(x).Cells(14).Value = "", "0", dtgDatos.Rows(x).Cells(14).Value)
                    '    If ValorInfo > 0 Then
                    '        Dim numbimestre As Integer

                    '        If Month(FechaInicioPeriodoGlobal) Mod 2 = 0 Then
                    '            numbimestre = Month(FechaInicioPeriodoGlobal) / 2
                    '        Else
                    '            numbimestre = (Month(FechaInicioPeriodoGlobal) + 1) / 2
                    '        End If

                    '        sql = "select * from DetalleDescInfonavit inner join nomina on DetalleDescInfonavit.fkiIdEmpleado"
                    '        sql &= " where fkiIdPeriodo=" & cboperiodo.SelectedValue & "or fkiIdPeriodo = "
                    '        sql &= " and fkiIdEmpleado=" & dtgDatos.Rows(x).Cells(2).Value

                    '    End If
                    'Next

                Else
                    sql = "delete from DetalleDescInfonavit"
                    sql &= " where fkiIdPeriodo=" & cboperiodo.SelectedValue
                    sql &= " and iSerie=" & cboserie.SelectedIndex
                    'sql &= " and iSerie=" & cboserie.SelectedIndex
                    sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex


                    sql2 = " delete from DetallePensionAlimenticia"
                    sql2 &= " where fkiIdPeriodo=" & cboperiodo.SelectedValue
                    sql2 &= " and iSerie=" & cboserie.SelectedIndex
                    sql2 &= " and iTipo=" & cboTipoNomina.SelectedIndex
                End If


                If nExecute(sql) = False Then
                    MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    'pnlProgreso.Visible = False
                    Exit Sub
                End If

                If nExecute(sql2) = False Then
                    MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    'pnlProgreso.Visible = False
                    Exit Sub
                End If

                calcular()
            End If



        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub calcular()
        Dim Sueldo As Double
        Dim SueldoBase As Double
        Dim ValorIncapacidad As Double
        Dim TotalPercepciones As Double
        Dim Incapacidad As Double
        Dim isr As Double
        Dim imss As Double
        Dim infonavitvalor As Double
        Dim infonavitanterior As Double
        Dim ajusteinfonavit As Double
        Dim pension As Double
        Dim prestamo As Double
        Dim fonacot As Double
        Dim subsidiogenerado As Double
        Dim subsidioaplicado As Double
        Dim RetencionOperadora As Double
        Dim InfonavitNormal As Double
        Dim PrestamoPersonalAsimilados As Double
        Dim AdeudoINfonavitAsimilados As Double
        Dim DiferenciaInfonavitAsimilados As Double
        Dim PensionAlimenticia As Double

        Dim Operadora As Double
        Dim ComplementoAsimilados As Double

        Dim SueldoBaseTMM As Double
        Dim CostoSocialTotal As Double
        Dim ComisionOperadora As Double
        Dim ComisionAsimilados As Double
        Dim subtotal As Double
        Dim iva As Double



        Dim sql As String
        Dim ValorUMA As Double
        Dim primavacacionesgravada As Double
        Dim primavacacionesexenta As Double
        Dim diastrabajados As Double
        Dim Sueldobruto As Double
        Dim TEFG As Double
        Dim TEFE As Double
        Dim TEO As Double
        Dim DSO As Double
        Dim VACAPRO As Double
        Dim numbimestre As Integer

        Try
            'verificamos que tenga dias a calcular
            'For x As Integer = 0 To dtgDatos.Rows.Count - 1
            '    If Double.Parse(IIf(dtgDatos.Rows(x).Cells(18).Value = "", "0", dtgDatos.Rows(x).Cells(18).Value)) <= 0 Then
            '        MessageBox.Show("Existen trabajadores que no tiene dias trabajados, favor de verificar", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            '        Exit Sub
            '    End If
            'Next



            sql = "select * from Salario "
            sql &= " where Anio=" & aniocostosocial
            sql &= " and iEstatus=1"
            Dim rwValorUMA As DataRow() = nConsulta(sql)
            If rwValorUMA Is Nothing = False Then
                ValorUMA = Double.Parse(rwValorUMA(0)("uma").ToString)
            Else
                ValorUMA = 0
                MessageBox.Show("No se encontro valor para UMA en el año: " & aniocostosocial)
            End If


            pnlProgreso.Visible = True

            Application.DoEvents()
            pnlCatalogo.Enabled = False
            pgbProgreso.Minimum = 0
            pgbProgreso.Value = 0
            pgbProgreso.Maximum = dtgDatos.Rows.Count




            For x As Integer = 0 To dtgDatos.Rows.Count - 1


                'verificamos los sueldos
                sql = "Select salariod,sbc,salariodTopado,sbcTopado from costosocial inner join puestos on costosocial.fkiIdPuesto=puestos.iIdPuesto "
                sql &= " where cNombre = '" & dtgDatos.Rows(x).Cells(11).FormattedValue & "' and anio=" & aniocostosocial

                Dim rwDatosSalario As DataRow() = nConsulta(sql)

                If rwDatosSalario Is Nothing = False Then
                    If dtgDatos.Rows(x).Cells(10).Value >= 55 Then
                        dtgDatos.Rows(x).Cells(16).Value = rwDatosSalario(0)("salariodTopado")
                        dtgDatos.Rows(x).Cells(17).Value = rwDatosSalario(0)("sbcTopado")
                    Else
                        dtgDatos.Rows(x).Cells(16).Value = rwDatosSalario(0)("salariod")
                        dtgDatos.Rows(x).Cells(17).Value = rwDatosSalario(0)("sbc")
                    End If

                Else
                    MessageBox.Show("No se encontraron datos")
                End If


                'Dim cadena As String = dgvCombo.Text
                If dtgDatos.Rows(x).Cells(11).FormattedValue = "OFICIALES EN PRACTICAS: PILOTIN / ASPIRANTE" Then
                    Sueldo = Double.Parse(dtgDatos.Rows(x).Cells(17).Value) * Double.Parse(IIf(dtgDatos.Rows(x).Cells(18).Value = "", "0", dtgDatos.Rows(x).Cells(18).Value))
                    dtgDatos.Rows(x).Cells(21).Value = Math.Round(Sueldo, 2).ToString("###,##0.00")
                    dtgDatos.Rows(x).Cells(22).Value = "0.00"
                    dtgDatos.Rows(x).Cells(23).Value = "0.00"
                    dtgDatos.Rows(x).Cells(24).Value = "0.00"
                    dtgDatos.Rows(x).Cells(25).Value = "0.00"
                    dtgDatos.Rows(x).Cells(26).Value = "0.00"
                    dtgDatos.Rows(x).Cells(27).Value = "0.00"
                    dtgDatos.Rows(x).Cells(28).Value = "0.00"
                    dtgDatos.Rows(x).Cells(29).Value = "0.00"
                    dtgDatos.Rows(x).Cells(30).Value = "0.00"
                    dtgDatos.Rows(x).Cells(31).Value = "0.00"
                    dtgDatos.Rows(x).Cells(32).Value = "0.00"
                    dtgDatos.Rows(x).Cells(33).Value = Math.Round(Sueldo, 2).ToString("###,##0.00")
                    dtgDatos.Rows(x).Cells(34).Value = Math.Round(Sueldo, 2).ToString("###,##0.00")
                    'Incapacidad
                    ValorIncapacidad = 0.0
                    If dtgDatos.Rows(x).Cells(19).Value <> "Ninguno" Then

                        ValorIncapacidad = Incapacidades(dtgDatos.Rows(x).Cells(19).Value, dtgDatos.Rows(x).Cells(20).Value, dtgDatos.Rows(x).Cells(16).Value)

                    End If
                    dtgDatos.Rows(x).Cells(35).Value = Math.Round(ValorIncapacidad, 2).ToString("###,##0.00")
                    'ISR
                    dtgDatos.Rows(x).Cells(36).Value = Math.Round(Double.Parse((baseisrtotal(dtgDatos.Rows(x).Cells(11).Value, 30, dtgDatos.Rows(x).Cells(17).Value, ValorIncapacidad)) / 30 * dtgDatos.Rows(x).Cells(18).Value), 2).ToString("###,##0.00")
                    'IMSS
                    dtgDatos.Rows(x).Cells(37).Value = "0.00"
                    'INFONAVIT
                    '##### VERIFICAR SI ESTA YA CALCULADO EL INFONAVIT DEL BIMESTRE

                    If dtgDatos.Rows(x).Tag = "" Then
                        Dim CalculoInfonavit As Integer = VerificarCalculoInfonavit(cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value))

                        Select Case CalculoInfonavit
                            Case 0
                                'No es necesario calcular
                                dtgDatos.Rows(x).Cells(38).Value = "0.00"
                            Case 1
                                'Ya esta Calculado
                                'Verificar cuanto le toca para el pago
                                Dim MontoInfonavit As Double = MontoInfonavitF(cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value))

                                If MontoInfonavit > 0 Then

                                    sql = "select isnull(sum(Cantidad),0) as monto from DetalleDescInfonavit where fkiIdCalculoInfonavit=" & IDCalculoInfonavit
                                    Dim rwMontoInfonavit As DataRow() = nConsulta(sql)
                                    If rwMontoInfonavit Is Nothing = False Then

                                        If Double.Parse(rwMontoInfonavit(0)("monto").ToString) < MontoInfonavit Then
                                            'Diferencia
                                            Dim FaltanteInfonavit As Double = MontoInfonavit - Double.Parse(rwMontoInfonavit(0)("monto").ToString)

                                            TotalPercepciones = Double.Parse(IIf(dtgDatos.Rows(x).Cells(33).Value = "", "0", dtgDatos.Rows(x).Cells(33).Value.ToString.Replace(",", "")))
                                            Incapacidad = Double.Parse(IIf(dtgDatos.Rows(x).Cells(35).Value = "", "0", dtgDatos.Rows(x).Cells(35).Value))
                                            isr = Double.Parse(IIf(dtgDatos.Rows(x).Cells(36).Value = "", "0", dtgDatos.Rows(x).Cells(36).Value))
                                            imss = Double.Parse(IIf(dtgDatos.Rows(x).Cells(37).Value = "", "0", dtgDatos.Rows(x).Cells(37).Value))

                                            Dim SubtotalAntesInfonavit As Double = TotalPercepciones - Incapacidad - isr - imss

                                            If SubtotalAntesInfonavit > (FaltanteInfonavit / 2) Then
                                                dtgDatos.Rows(x).Cells(38).Value = Math.Round((FaltanteInfonavit / 2), 2)

                                            Else
                                                dtgDatos.Rows(x).Cells(38).Value = Math.Round((SubtotalAntesInfonavit - 1), 2)
                                            End If



                                        Else
                                            dtgDatos.Rows(x).Cells(38).Value = "0.00"
                                        End If


                                    End If
                                Else
                                    dtgDatos.Rows(x).Cells(38).Value = "0.00"

                                End If
                            Case 2
                                'No esta calculado
                                If CalcularInfonavit(dtgDatos.Rows(x).Cells(13).Value, Double.Parse(dtgDatos.Rows(x).Cells(14).Value), Double.Parse(dtgDatos.Rows(x).Cells(17).Value), Date.Parse("01/01/1900"), cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value)) Then
                                    'Verificar cuanto le toca para el pago
                                    Dim MontoInfonavit As Double = MontoInfonavitF(cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value))

                                    If MontoInfonavit > 0 Then
                                        sql = "select isnull(sum(Cantidad),0) as monto from DetalleDescInfonavit where fkiIdCalculoInfonavit=" & IDCalculoInfonavit
                                        Dim rwMontoInfonavit As DataRow() = nConsulta(sql)
                                        If rwMontoInfonavit Is Nothing = False Then

                                            If Double.Parse(rwMontoInfonavit(0)("monto").ToString) < MontoInfonavit Then
                                                'Diferencia
                                                Dim FaltanteInfonavit As Double = MontoInfonavit - Double.Parse(rwMontoInfonavit(0)("monto").ToString)

                                                TotalPercepciones = Double.Parse(IIf(dtgDatos.Rows(x).Cells(33).Value = "", "0", dtgDatos.Rows(x).Cells(33).Value.ToString.Replace(",", "")))
                                                Incapacidad = Double.Parse(IIf(dtgDatos.Rows(x).Cells(35).Value = "", "0", dtgDatos.Rows(x).Cells(35).Value))
                                                isr = Double.Parse(IIf(dtgDatos.Rows(x).Cells(36).Value = "", "0", dtgDatos.Rows(x).Cells(36).Value))
                                                imss = Double.Parse(IIf(dtgDatos.Rows(x).Cells(37).Value = "", "0", dtgDatos.Rows(x).Cells(37).Value))

                                                Dim SubtotalAntesInfonavit As Double = TotalPercepciones - Incapacidad - isr - imss

                                                If SubtotalAntesInfonavit > (FaltanteInfonavit / 2) Then
                                                    dtgDatos.Rows(x).Cells(38).Value = Math.Round((FaltanteInfonavit / 2), 2)

                                                Else
                                                    dtgDatos.Rows(x).Cells(38).Value = Math.Round((SubtotalAntesInfonavit - 1), 2)
                                                End If



                                            Else
                                                dtgDatos.Rows(x).Cells(38).Value = "0.00"
                                            End If


                                        End If
                                    Else
                                        dtgDatos.Rows(x).Cells(38).Value = "0.00"

                                    End If


                                End If
                        End Select
                    Else

                    End If



                    '############# CALCULO POR DIAS INFONAVIT

                    'dtgDatos.Rows(x).Cells(38).Value = Math.Round(infonavit(dtgDatos.Rows(x).Cells(13).Value, Double.Parse(dtgDatos.Rows(x).Cells(14).Value), Double.Parse(dtgDatos.Rows(x).Cells(17).Value), Date.Parse("01/01/1900"), cboperiodo.SelectedValue, Double.Parse(dtgDatos.Rows(x).Cells(18).Value), Integer.Parse(dtgDatos.Rows(x).Cells(2).Value)), 2).ToString("###,##0.00")
                    '#############

                    'INFONAVIT BIMESTRE ANTERIOR
                    'AJUSTE INFONAVIT
                    'PENSION
                    'PRESTAMO
                    'FONACOT
                    'SUBSIDIO GENERADO
                    dtgDatos.Rows(x).Cells(44).Value = baseSubsidio(dtgDatos.Rows(x).Cells(11).FormattedValue, 30, Double.Parse(dtgDatos.Rows(x).Cells(17).Value), ValorIncapacidad)
                    'SUBSIDIO APLICADO
                    dtgDatos.Rows(x).Cells(45).Value = Math.Round(baseSubsidiototal(dtgDatos.Rows(x).Cells(11).FormattedValue, 30, Double.Parse(dtgDatos.Rows(x).Cells(17).Value), ValorIncapacidad) / 30 * Double.Parse(dtgDatos.Rows(x).Cells(18).Value), 2)
                    'NETO


                    TotalPercepciones = Double.Parse(IIf(dtgDatos.Rows(x).Cells(33).Value = "", "0", dtgDatos.Rows(x).Cells(33).Value.ToString.Replace(",", "")))
                    Incapacidad = Double.Parse(IIf(dtgDatos.Rows(x).Cells(35).Value = "", "0", dtgDatos.Rows(x).Cells(35).Value))
                    isr = Double.Parse(IIf(dtgDatos.Rows(x).Cells(36).Value = "", "0", dtgDatos.Rows(x).Cells(36).Value))
                    imss = Double.Parse(IIf(dtgDatos.Rows(x).Cells(37).Value = "", "0", dtgDatos.Rows(x).Cells(37).Value))
                    infonavitvalor = Double.Parse(IIf(dtgDatos.Rows(x).Cells(38).Value = "", "0", dtgDatos.Rows(x).Cells(38).Value))
                    infonavitanterior = Double.Parse(IIf(dtgDatos.Rows(x).Cells(39).Value = "", "0", dtgDatos.Rows(x).Cells(39).Value))
                    ajusteinfonavit = Double.Parse(IIf(dtgDatos.Rows(x).Cells(40).Value = "", "0", dtgDatos.Rows(x).Cells(40).Value))
                    pension = Double.Parse(IIf(dtgDatos.Rows(x).Cells(41).Value = "", "0", dtgDatos.Rows(x).Cells(41).Value))
                    prestamo = Double.Parse(IIf(dtgDatos.Rows(x).Cells(42).Value = "", "0", dtgDatos.Rows(x).Cells(42).Value))
                    fonacot = Double.Parse(IIf(dtgDatos.Rows(x).Cells(43).Value = "", "0", dtgDatos.Rows(x).Cells(43).Value))
                    subsidiogenerado = Double.Parse(IIf(dtgDatos.Rows(x).Cells(44).Value = "", "0", dtgDatos.Rows(x).Cells(44).Value))
                    subsidioaplicado = Double.Parse(IIf(dtgDatos.Rows(x).Cells(45).Value = "", "0", dtgDatos.Rows(x).Cells(45).Value))
                    Operadora = Math.Round(TotalPercepciones - Incapacidad - isr - imss - infonavitvalor - infonavitanterior - ajusteinfonavit - pension - prestamo - fonacot + subsidioaplicado, 2)
                    dtgDatos.Rows(x).Cells(46).Value = Operadora





                Else
                    diastrabajados = Double.Parse(IIf(dtgDatos.Rows(x).Cells(18).Value = "", "0", dtgDatos.Rows(x).Cells(18).Value))
                    If diastrabajados = 0 Then
                        dtgDatos.Rows(x).Cells(21).Value = "0.00"
                        dtgDatos.Rows(x).Cells(22).Value = "0.00"
                        dtgDatos.Rows(x).Cells(23).Value = "0.00"
                        dtgDatos.Rows(x).Cells(24).Value = "0.00"
                        dtgDatos.Rows(x).Cells(25).Value = "0.00"
                        dtgDatos.Rows(x).Cells(26).Value = "0.00"
                        dtgDatos.Rows(x).Cells(27).Value = "0.00"
                        dtgDatos.Rows(x).Cells(28).Value = "0.00"
                        dtgDatos.Rows(x).Cells(29).Value = "0.00"
                        dtgDatos.Rows(x).Cells(30).Value = "0.00"
                        dtgDatos.Rows(x).Cells(31).Value = "0.00"
                        dtgDatos.Rows(x).Cells(32).Value = "0.00"
                        dtgDatos.Rows(x).Cells(33).Value = "0.00"
                        dtgDatos.Rows(x).Cells(34).Value = "0.00"
                        'Incapacidad
                        dtgDatos.Rows(x).Cells(35).Value = "0.00"
                        'ISR
                        dtgDatos.Rows(x).Cells(36).Value = "0.00"
                        'IMSS
                        dtgDatos.Rows(x).Cells(37).Value = "0.00"
                        'INFONAVIT
                        '##### VERIFICAR SI ESTA YA CALCULADO EL INFONAVIT DEL BIMESTRE
                        dtgDatos.Rows(x).Cells(38).Value = "0.00"


                        '############# CALCULO POR DIAS INFONAVIT
                        'dtgDatos.Rows(x).Cells(38).Value = Math.Round(infonavit(dtgDatos.Rows(x).Cells(13).Value, Double.Parse(dtgDatos.Rows(x).Cells(14).Value), Double.Parse(dtgDatos.Rows(x).Cells(17).Value), Date.Parse("01/01/1900"), cboperiodo.SelectedValue, Double.Parse(dtgDatos.Rows(x).Cells(18).Value), Integer.Parse(dtgDatos.Rows(x).Cells(2).Value)), 2).ToString("###,##0.00")
                        '############# CALCULO POR DIAS INFONAVIT

                        'INFONAVIT BIMESTRE ANTERIOR
                        'AJUSTE INFONAVIT
                        'PENSION
                        'PRESTAMO
                        'FONACOT
                        'SUBSIDIO GENERADO
                        dtgDatos.Rows(x).Cells(44).Value = "0.00"
                        'SUBSIDIO APLICADO
                        dtgDatos.Rows(x).Cells(45).Value = "0.00"
                        'NETO
                        TotalPercepciones = Double.Parse(IIf(dtgDatos.Rows(x).Cells(33).Value = "", "0", dtgDatos.Rows(x).Cells(33).Value.ToString.Replace(",", "")))
                        Incapacidad = Double.Parse(IIf(dtgDatos.Rows(x).Cells(35).Value = "", "0", dtgDatos.Rows(x).Cells(35).Value))
                        isr = Double.Parse(IIf(dtgDatos.Rows(x).Cells(36).Value = "", "0", dtgDatos.Rows(x).Cells(36).Value))
                        imss = Double.Parse(IIf(dtgDatos.Rows(x).Cells(37).Value = "", "0", dtgDatos.Rows(x).Cells(37).Value))
                        infonavitvalor = Double.Parse(IIf(dtgDatos.Rows(x).Cells(38).Value = "", "0", dtgDatos.Rows(x).Cells(38).Value))
                        infonavitanterior = Double.Parse(IIf(dtgDatos.Rows(x).Cells(39).Value = "", "0", dtgDatos.Rows(x).Cells(39).Value))
                        ajusteinfonavit = Double.Parse(IIf(dtgDatos.Rows(x).Cells(40).Value = "", "0", dtgDatos.Rows(x).Cells(40).Value))
                        pension = Double.Parse(IIf(dtgDatos.Rows(x).Cells(41).Value = "", "0", dtgDatos.Rows(x).Cells(41).Value))
                        prestamo = Double.Parse(IIf(dtgDatos.Rows(x).Cells(42).Value = "", "0", dtgDatos.Rows(x).Cells(42).Value))
                        fonacot = Double.Parse(IIf(dtgDatos.Rows(x).Cells(43).Value = "", "0", dtgDatos.Rows(x).Cells(43).Value))
                        subsidiogenerado = Double.Parse(IIf(dtgDatos.Rows(x).Cells(44).Value = "", "0", dtgDatos.Rows(x).Cells(44).Value))
                        subsidioaplicado = Double.Parse(IIf(dtgDatos.Rows(x).Cells(45).Value = "", "0", dtgDatos.Rows(x).Cells(45).Value))

                        Operadora = 0
                        dtgDatos.Rows(x).Cells(46).Value = Operadora

                    Else
                        Sueldo = Double.Parse(dtgDatos.Rows(x).Cells(17).Value) * diastrabajados
                        dtgDatos.Rows(x).Cells(21).Value = Math.Round(Sueldo * (26.19568006 / 100), 2).ToString("###,##0.00")
                        Sueldobruto = Math.Round(Sueldo * (26.19568006 / 100), 2)
                        dtgDatos.Rows(x).Cells(22).Value = Math.Round((Sueldo * (8.5070471 / 100)) / 2, 2).ToString("###,##0.00")
                        TEFG = Math.Round((Sueldo * (8.5070471 / 100)) / 2, 2)
                        dtgDatos.Rows(x).Cells(23).Value = Math.Round((Sueldo * (8.5070471 / 100)) / 2, 2).ToString("###,##0.00")
                        TEFE = Math.Round((Sueldo * (8.5070471 / 100)) / 2, 2)
                        dtgDatos.Rows(x).Cells(24).Value = Math.Round(Sueldo * (42.89215164 / 100), 2).ToString("###,##0.00")
                        TEO = Math.Round(Sueldo * (42.89215164 / 100), 2)
                        dtgDatos.Rows(x).Cells(25).Value = Math.Round(Sueldo * (9.677848468 / 100), 2).ToString("###,##0.00")
                        DSO = Math.Round(Sueldo * (9.677848468 / 100), 2)
                        dtgDatos.Rows(x).Cells(26).Value = Math.Round(Sueldo * (7.272727273 / 100), 2).ToString("###,##0.00")
                        VACAPRO = Math.Round(Sueldo * (7.272727273 / 100), 2)
                        SueldoBase = Sueldobruto + TEFG + TEFE + TEO + DSO


                        'Aguinaldo gravado 

                        If ((SueldoBase / diastrabajados) * 15 / 12 * (diastrabajados / 30)) > ((ValorUMA * 30 / 12) * (diastrabajados / 30)) Then
                            'Aguinaldo gravado
                            dtgDatos.Rows(x).Cells(27).Value = Math.Round(((SueldoBase / diastrabajados) * 15 / 12 * (diastrabajados / 30)) - ((ValorUMA * 30 / 12) * (diastrabajados / 30)), 2)
                            'Aguinaldo exento
                            dtgDatos.Rows(x).Cells(28).Value = Math.Round(((ValorUMA * 30 / 12) * (diastrabajados / 30)), 2)


                        Else
                            'Aguinaldo gravado

                            dtgDatos.Rows(x).Cells(27).Value = "0.00"
                            'Aguinaldo exento
                            dtgDatos.Rows(x).Cells(28).Value = Math.Round(((SueldoBase / diastrabajados) * 15 / 12 * (diastrabajados / 30)), 2)

                        End If


                        'Aguinaldo total
                        dtgDatos.Rows(x).Cells(29).Value = Math.Round(Double.Parse(dtgDatos.Rows(x).Cells(27).Value) + Double.Parse(dtgDatos.Rows(x).Cells(28).Value), 2)

                        'Prima de vacaciones

                        'Calculos prima

                        primavacacionesgravada = (SueldoBase * 0.25 / 12 * (diastrabajados / 30)) - ((ValorUMA * 15 / 12) * (diastrabajados / 30))
                        primavacacionesexenta = ((ValorUMA * 15 / 12) * (diastrabajados / 30))

                        If primavacacionesgravada > 0 Then
                            dtgDatos.Rows(x).Cells(30).Value = Math.Round(primavacacionesgravada, 2)
                            dtgDatos.Rows(x).Cells(31).Value = Math.Round(primavacacionesexenta, 2)
                        Else
                            primavacacionesexenta = (SueldoBase * 0.25 / 12 * (diastrabajados / 30))
                            dtgDatos.Rows(x).Cells(30).Value = "0.00"
                            dtgDatos.Rows(x).Cells(31).Value = Math.Round(primavacacionesexenta, 2)
                        End If

                        'Total Prima de vacaciones                    
                        dtgDatos.Rows(x).Cells(32).Value = Math.Round(IIf(primavacacionesgravada > 0, primavacacionesgravada, 0) + primavacacionesexenta, 2)
                        'Total percepciones
                        dtgDatos.Rows(x).Cells(33).Value = Math.Round(SueldoBase + VACAPRO + dtgDatos.Rows(x).Cells(29).Value + dtgDatos.Rows(x).Cells(32).Value, 2)
                        'Total percepsiones para isr
                        dtgDatos.Rows(x).Cells(34).Value = Math.Round(SueldoBase - TEFE + VACAPRO + dtgDatos.Rows(x).Cells(27).Value + dtgDatos.Rows(x).Cells(30).Value, 2)
                        'Incapacidad
                        ValorIncapacidad = 0.0
                        If dtgDatos.Rows(x).Cells(19).Value <> "Ninguno" Then

                            ValorIncapacidad = Incapacidades(dtgDatos.Rows(x).Cells(19).Value, dtgDatos.Rows(x).Cells(20).Value, dtgDatos.Rows(x).Cells(16).Value)

                        End If
                        dtgDatos.Rows(x).Cells(35).Value = Math.Round(ValorIncapacidad, 2).ToString("###,##0.00")
                        'ISR
                        dtgDatos.Rows(x).Cells(36).Value = Math.Round(Double.Parse((baseisrtotal(dtgDatos.Rows(x).Cells(11).FormattedValue, 30, dtgDatos.Rows(x).Cells(17).Value, ValorIncapacidad)) / 30 * dtgDatos.Rows(x).Cells(18).Value), 2).ToString("###,##0.00")
                        'IMSS
                        dtgDatos.Rows(x).Cells(37).Value = "0.00"
                        'INFONAVIT
                        '##### VERIFICAR SI ESTA YA CALCULADO EL INFONAVIT DEL BIMESTRE


                        If chkNoinfonavit.Checked = False Then
                            If dtgDatos.Rows(x).Tag = "" Then
                                Dim CalculoInfonavit As Integer = VerificarCalculoInfonavit(cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value))

                                Select Case CalculoInfonavit
                                    Case 0
                                        'No es necesario calcular
                                        dtgDatos.Rows(x).Cells(38).Value = "0.00"
                                    Case 1
                                        'Ya esta Calculado
                                        'Verificar cuanto le toca para el pago
                                        Dim MontoInfonavit As Double = MontoInfonavitF(cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value))

                                        If MontoInfonavit > 0 Then
                                            'Dim numbimestre As Integer

                                            If Month(FechaInicioPeriodoGlobal) Mod 2 = 0 Then
                                                numbimestre = Month(FechaInicioPeriodoGlobal) / 2
                                            Else
                                                numbimestre = (Month(FechaInicioPeriodoGlobal) + 1) / 2
                                            End If
                                            sql = "select isnull(sum(Cantidad),0) as monto from DetalleDescInfonavit where fkiIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value & " and Numbimestre= " & numbimestre & " and Anio=" & FechaInicioPeriodoGlobal.Year & "and (iTipopagadora=101 or iTipopagadora=102)"
                                            Dim rwMontoInfonavit As DataRow() = nConsulta(sql)
                                            If rwMontoInfonavit Is Nothing = False Then

                                                'Verificamos el monto del infonavit a calcular

                                                InfonavitNormal = Math.Round(infonavit(dtgDatos.Rows(x).Cells(13).Value, Double.Parse(dtgDatos.Rows(x).Cells(14).Value), Double.Parse(dtgDatos.Rows(x).Cells(17).Value), Date.Parse("01/01/1900"), cboperiodo.SelectedValue, Double.Parse(dtgDatos.Rows(x).Cells(18).Value), Integer.Parse(dtgDatos.Rows(x).Cells(2).Value), Integer.Parse(dtgDatos.Rows(x).Cells(1).Value) - 1), 2).ToString("###,##0.00")

                                                '########


                                                If Double.Parse(rwMontoInfonavit(0)("monto").ToString) < MontoInfonavit Then
                                                    'Diferencia
                                                    Dim FaltanteInfonavit As Double = MontoInfonavit - Double.Parse(rwMontoInfonavit(0)("monto").ToString)

                                                    TotalPercepciones = Double.Parse(IIf(dtgDatos.Rows(x).Cells(33).Value = "", "0", dtgDatos.Rows(x).Cells(33).Value.ToString.Replace(",", "")))
                                                    Incapacidad = Double.Parse(IIf(dtgDatos.Rows(x).Cells(35).Value = "", "0", dtgDatos.Rows(x).Cells(35).Value))
                                                    isr = Double.Parse(IIf(dtgDatos.Rows(x).Cells(36).Value = "", "0", dtgDatos.Rows(x).Cells(36).Value))
                                                    imss = Double.Parse(IIf(dtgDatos.Rows(x).Cells(37).Value = "", "0", dtgDatos.Rows(x).Cells(37).Value))

                                                    Dim SubtotalAntesInfonavit As Double = TotalPercepciones - Incapacidad - isr - imss


                                                    'VErificamos el infonavit

                                                    If FaltanteInfonavit > InfonavitNormal Then

                                                        If SubtotalAntesInfonavit > InfonavitNormal Then
                                                            dtgDatos.Rows(x).Cells(38).Value = Math.Round((InfonavitNormal), 2)

                                                        Else
                                                            dtgDatos.Rows(x).Cells(38).Value = Math.Round((SubtotalAntesInfonavit - 1), 2)
                                                        End If
                                                    Else
                                                        If SubtotalAntesInfonavit > FaltanteInfonavit Then
                                                            dtgDatos.Rows(x).Cells(38).Value = Math.Round((FaltanteInfonavit), 2)

                                                        Else
                                                            dtgDatos.Rows(x).Cells(38).Value = Math.Round((SubtotalAntesInfonavit - 1), 2)
                                                        End If

                                                    End If




                                                    'If SubtotalAntesInfonavit > (FaltanteInfonavit / 2) Then
                                                    '    dtgDatos.Rows(x).Cells(38).Value = Math.Round((FaltanteInfonavit / 2), 2)

                                                    'Else
                                                    '    dtgDatos.Rows(x).Cells(38).Value = Math.Round((SubtotalAntesInfonavit - 1), 2)
                                                    'End If



                                                Else
                                                    dtgDatos.Rows(x).Cells(38).Value = "0.00"
                                                End If


                                            End If
                                        Else
                                            dtgDatos.Rows(x).Cells(38).Value = "0.00"

                                        End If
                                    Case 2
                                        'No esta calculado
                                        If CalcularInfonavit(dtgDatos.Rows(x).Cells(13).Value, Double.Parse(dtgDatos.Rows(x).Cells(14).Value), Double.Parse(dtgDatos.Rows(x).Cells(17).Value), Date.Parse("01/01/1900"), cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value)) Then
                                            'Verificar cuanto le toca para el pago
                                            Dim MontoInfonavit As Double = MontoInfonavitF(cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value))

                                            If MontoInfonavit > 0 Then
                                                'Dim numbimestre As Integer

                                                If Month(FechaInicioPeriodoGlobal) Mod 2 = 0 Then
                                                    numbimestre = Month(FechaInicioPeriodoGlobal) / 2
                                                Else
                                                    numbimestre = (Month(FechaInicioPeriodoGlobal) + 1) / 2
                                                End If

                                                sql = "select isnull(sum(Cantidad),0) as monto from DetalleDescInfonavit where fkiIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value & " and Numbimestre= " & numbimestre & " and Anio=" & FechaInicioPeriodoGlobal.Year & "and (iTipopagadora=101 or iTipopagadora=102)"
                                                Dim rwMontoInfonavit As DataRow() = nConsulta(sql)
                                                If rwMontoInfonavit Is Nothing = False Then
                                                    'Verificamos el monto del infonavit a calcular

                                                    InfonavitNormal = Math.Round(infonavit(dtgDatos.Rows(x).Cells(13).Value, Double.Parse(dtgDatos.Rows(x).Cells(14).Value), Double.Parse(dtgDatos.Rows(x).Cells(17).Value), Date.Parse("01/01/1900"), cboperiodo.SelectedValue, Double.Parse(dtgDatos.Rows(x).Cells(18).Value), Integer.Parse(dtgDatos.Rows(x).Cells(2).Value), Integer.Parse(dtgDatos.Rows(x).Cells(1).Value)) - 1, 2).ToString("###,##0.00")

                                                    '########
                                                    If Double.Parse(rwMontoInfonavit(0)("monto").ToString) < MontoInfonavit Then
                                                        'Diferencia
                                                        Dim FaltanteInfonavit As Double = MontoInfonavit - Double.Parse(rwMontoInfonavit(0)("monto").ToString)

                                                        TotalPercepciones = Double.Parse(IIf(dtgDatos.Rows(x).Cells(33).Value = "", "0", dtgDatos.Rows(x).Cells(33).Value.ToString.Replace(",", "")))
                                                        Incapacidad = Double.Parse(IIf(dtgDatos.Rows(x).Cells(35).Value = "", "0", dtgDatos.Rows(x).Cells(35).Value))
                                                        isr = Double.Parse(IIf(dtgDatos.Rows(x).Cells(36).Value = "", "0", dtgDatos.Rows(x).Cells(36).Value))
                                                        imss = Double.Parse(IIf(dtgDatos.Rows(x).Cells(37).Value = "", "0", dtgDatos.Rows(x).Cells(37).Value))

                                                        Dim SubtotalAntesInfonavit As Double = TotalPercepciones - Incapacidad - isr - imss

                                                        If FaltanteInfonavit > InfonavitNormal Then

                                                            If SubtotalAntesInfonavit > InfonavitNormal Then
                                                                dtgDatos.Rows(x).Cells(38).Value = Math.Round((InfonavitNormal), 2)

                                                            Else
                                                                dtgDatos.Rows(x).Cells(38).Value = Math.Round((SubtotalAntesInfonavit - 1), 2)
                                                            End If
                                                        Else
                                                            If SubtotalAntesInfonavit > FaltanteInfonavit Then
                                                                dtgDatos.Rows(x).Cells(38).Value = Math.Round((FaltanteInfonavit), 2)

                                                            Else
                                                                dtgDatos.Rows(x).Cells(38).Value = Math.Round((SubtotalAntesInfonavit - 1), 2)
                                                            End If

                                                        End If



                                                    Else
                                                        dtgDatos.Rows(x).Cells(38).Value = "0.00"
                                                    End If


                                                End If
                                            Else
                                                dtgDatos.Rows(x).Cells(38).Value = "0.00"

                                            End If


                                        End If
                                End Select
                            Else

                            End If
                        End If




                        'Guardamos el infonavit en el sistema para  tenerlo actualizado para el siguiente trabajador

                        '########GUARDAR INFONAVIT

                        'Dim numbimestre As Integer
                        If Month(FechaInicioPeriodoGlobal) Mod 2 = 0 Then
                            numbimestre = Month(FechaInicioPeriodoGlobal) / 2
                        Else
                            numbimestre = (Month(FechaInicioPeriodoGlobal) + 1) / 2
                        End If


                        If Double.Parse(IIf(dtgDatos.Rows(x).Cells(38).Value = "", "0", dtgDatos.Rows(x).Cells(38).Value)) Then

                            Dim MontoInfonavit As Double = MontoInfonavitF(cboperiodo.SelectedValue, Integer.Parse(dtgDatos.Rows(x).Cells(2).Value))

                            sql = "EXEC setDetalleDescInfonavitInsertar  0"
                            'fk Calculo infonavit
                            sql &= "," & IIf(MontoInfonavit > 0, IDCalculoInfonavit, 0)
                            'Cantidad
                            sql &= "," & dtgDatos.Rows(x).Cells(38).Value
                            ' fk Empleado
                            sql &= "," & dtgDatos.Rows(x).Cells(2).Value
                            'Numbimestre
                            sql &= "," & numbimestre
                            'Anio
                            sql &= "," & FechaInicioPeriodoGlobal.Year
                            'fk Periodo
                            sql &= "," & cboperiodo.SelectedValue
                            'Serie
                            sql &= "," & cboserie.SelectedIndex
                            'Tipo Nomina
                            sql &= "," & cboTipoNomina.SelectedIndex
                            'Tipo Pagadora
                            sql &= ",101"
                            'iEstatu
                            sql &= ",1"

                            If nExecute(sql) = False Then
                                MessageBox.Show("Ocurrio un error insertar pago prestamo ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                'pnlProgreso.Visible = False
                                Exit Sub
                            End If
                        End If




                        '############# CALCULO POR DIAS INFONAVIT

                        'dtgDatos.Rows(x).Cells(38).Value = Math.Round(infonavit(dtgDatos.Rows(x).Cells(13).Value, Double.Parse(dtgDatos.Rows(x).Cells(14).Value), Double.Parse(dtgDatos.Rows(x).Cells(17).Value), Date.Parse("01/01/1900"), cboperiodo.SelectedValue, Double.Parse(dtgDatos.Rows(x).Cells(18).Value), Integer.Parse(dtgDatos.Rows(x).Cells(2).Value)), 2).ToString("###,##0.00")
                        '############# CALCULO POR DIAS INFONAVIT


                        'SUBSIDIO GENERADO
                        dtgDatos.Rows(x).Cells(44).Value = Math.Round((baseSubsidio(dtgDatos.Rows(x).Cells(11).FormattedValue, 30, Double.Parse(dtgDatos.Rows(x).Cells(17).Value), ValorIncapacidad)), 2).ToString("###,##0.00")
                        'SUBSIDIO APLICADO
                        dtgDatos.Rows(x).Cells(45).Value = Math.Round((baseSubsidiototal(dtgDatos.Rows(x).Cells(11).FormattedValue, 30, Double.Parse(dtgDatos.Rows(x).Cells(17).Value), ValorIncapacidad)) / 30 * Double.Parse(dtgDatos.Rows(x).Cells(18).Value), 2).ToString("###,##0.00")

                        TotalPercepciones = Double.Parse(IIf(dtgDatos.Rows(x).Cells(33).Value = "", "0", dtgDatos.Rows(x).Cells(33).Value.ToString.Replace(",", "")))
                        Incapacidad = Double.Parse(IIf(dtgDatos.Rows(x).Cells(35).Value = "", "0", dtgDatos.Rows(x).Cells(35).Value))
                        isr = Double.Parse(IIf(dtgDatos.Rows(x).Cells(36).Value = "", "0", dtgDatos.Rows(x).Cells(36).Value))
                        imss = Double.Parse(IIf(dtgDatos.Rows(x).Cells(37).Value = "", "0", dtgDatos.Rows(x).Cells(37).Value))
                        infonavitvalor = Double.Parse(IIf(dtgDatos.Rows(x).Cells(38).Value = "", "0", dtgDatos.Rows(x).Cells(38).Value))
                        infonavitanterior = Double.Parse(IIf(dtgDatos.Rows(x).Cells(39).Value = "", "0", dtgDatos.Rows(x).Cells(39).Value))
                        ajusteinfonavit = Double.Parse(IIf(dtgDatos.Rows(x).Cells(40).Value = "", "0", dtgDatos.Rows(x).Cells(40).Value))

                        prestamo = Double.Parse(IIf(dtgDatos.Rows(x).Cells(42).Value = "", "0", dtgDatos.Rows(x).Cells(42).Value))
                        fonacot = Double.Parse(IIf(dtgDatos.Rows(x).Cells(43).Value = "", "0", dtgDatos.Rows(x).Cells(43).Value))
                        subsidiogenerado = Double.Parse(IIf(dtgDatos.Rows(x).Cells(44).Value = "", "0", dtgDatos.Rows(x).Cells(44).Value))
                        subsidioaplicado = Double.Parse(IIf(dtgDatos.Rows(x).Cells(45).Value = "", "0", dtgDatos.Rows(x).Cells(45).Value))



                        'INFONAVIT BIMESTRE ANTERIOR
                        'AJUSTE INFONAVIT
                        'PENSION
                        PensionAlimenticia = TotalPercepciones - Incapacidad - isr - imss - infonavitvalor - infonavitanterior - ajusteinfonavit - prestamo - fonacot + subsidioaplicado
                        'Buscamos la Pension

                        sql = "select * from PensionAlimenticia where fkiIdEmpleadoC=" & Integer.Parse(dtgDatos.Rows(x).Cells(2).Value)


                        Dim rwPensionEmpleado As DataRow() = nConsulta(sql)

                        pension = 0

                        If rwPensionEmpleado Is Nothing = False Then
                            For y As Integer = 0 To rwPensionEmpleado.Length - 1
                                pension = pension + Math.Round(PensionAlimenticia * (Double.Parse(rwPensionEmpleado(y)("fPorcentaje")) / 100), 2)


                                'dtgDatos.Rows(x).Cells(41).Value = PensionAlimenticia * (Double.Parse(rwPensionEmpleado(y)("fPorcentaje")) / 100)

                                'Insertar la pension
                                'Insertamos los datos

                                sql = "EXEC [setDetallePensionAlimenticiaInsertar] 0"
                                'Id Empleado
                                sql &= "," & Integer.Parse(dtgDatos.Rows(x).Cells(2).Value)
                                'id Pension
                                sql &= "," & Integer.Parse(rwPensionEmpleado(y)("iIdPensionAlimenticia"))
                                'id Periodo
                                sql &= ",'" & cboperiodo.SelectedValue
                                'serie
                                sql &= "'," & cboserie.SelectedIndex
                                'tipo
                                sql &= "," & cboTipoNomina.SelectedIndex
                                'Monto
                                sql &= "," & Math.Round(PensionAlimenticia * (Double.Parse(rwPensionEmpleado(y)("fPorcentaje")) / 100), 2)
                                'Estatus
                                sql &= ",1"


                                If nExecute(sql) = False Then
                                    MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)


                                End If

                            Next
                            dtgDatos.Rows(x).Cells(41).Value = pension
                            'dtgDatos.Rows(x).Cells(41).Value = PensionAlimenticia * (Double.Parse(rwPensionEmpleado(0)("fPorcentaje")) / 100)
                        Else
                            pension = 0
                            dtgDatos.Rows(x).Cells(41).Value = "0"
                        End If


                        'PRESTAMO
                        'FONACOT

                        'NETO


                        pension = Double.Parse(IIf(dtgDatos.Rows(x).Cells(41).Value = "", "0", dtgDatos.Rows(x).Cells(41).Value))
                        Operadora = Math.Round(TotalPercepciones - Incapacidad - isr - imss - infonavitvalor - infonavitanterior - ajusteinfonavit - pension - prestamo - fonacot + subsidioaplicado, 2)
                        dtgDatos.Rows(x).Cells(46).Value = Operadora

                    End If




                End If

                'Sueldo Base TMM
                SueldoBaseTMM = (Double.Parse(IIf(dtgDatos.Rows(x).Cells(15).Value = "", "0", dtgDatos.Rows(x).Cells(15).Value))) / 2

                'Prestamo Personal Asimilado
                PrestamoPersonalAsimilados = Double.Parse(IIf(dtgDatos.Rows(x).Cells(47).Value = "", "0", dtgDatos.Rows(x).Cells(47).Value))
                'Adeudo_Infonavit_Asimilado
                AdeudoINfonavitAsimilados = Double.Parse(IIf(dtgDatos.Rows(x).Cells(48).Value = "", "0", dtgDatos.Rows(x).Cells(48).Value))
                'Difencia infonavit Asimilado
                DiferenciaInfonavitAsimilados = Double.Parse(IIf(dtgDatos.Rows(x).Cells(49).Value = "", "0", dtgDatos.Rows(x).Cells(49).Value))
                'Complemento Asimilado
                ComplementoAsimilados = Math.Round(SueldoBaseTMM - infonavitvalor - infonavitanterior - ajusteinfonavit - pension - prestamo - fonacot - PrestamoPersonalAsimilados - AdeudoINfonavitAsimilados - DiferenciaInfonavitAsimilados - Operadora, 2)
                dtgDatos.Rows(x).Cells(50).Value = ComplementoAsimilados
                'Retenciones_Operadora
                RetencionOperadora = Math.Round(Incapacidad + isr + imss + infonavitvalor + infonavitanterior + ajusteinfonavit + pension + prestamo + fonacot, 2)
                dtgDatos.Rows(x).Cells(51).Value = RetencionOperadora

                '%Comision
                dtgDatos.Rows(x).Cells(52).Value = "2%"
                'Comision Maecco
                ComisionOperadora = Math.Round((Operadora + RetencionOperadora) * 0.02, 2)
                dtgDatos.Rows(x).Cells(53).Value = ComisionOperadora

                'Comision Complemento
                ComisionAsimilados = Math.Round((ComplementoAsimilados + PrestamoPersonalAsimilados + AdeudoINfonavitAsimilados + DiferenciaInfonavitAsimilados) * 0.02, 2)
                dtgDatos.Rows(x).Cells(54).Value = ComisionAsimilados


                'Calcular el costo social

                'Obtenemos los datos del empleado,id puesto
                'de acuerdo a la edad y el status

                If dtgDatos.Rows(x).Cells(2).Tag = "" And diastrabajados <> 0 Then
                    sql = "select * from empleadosC where iIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                    Dim rwEmpleado As DataRow() = nConsulta(sql)
                    If rwEmpleado Is Nothing = False Then
                        sql = "select * from puestos inner join costosocial on puestos.iidPuesto= costosocial.fkiIdPuesto where puestos.cnombre='" & dtgDatos.Rows(x).Cells(11).FormattedValue & "' and anio=" & aniocostosocial
                        Dim rwCostoSocial As DataRow() = nConsulta(sql)
                        If rwCostoSocial Is Nothing = False Then
                            If dtgDatos.Rows(x).Cells(10).Value >= 55 Then
                                If dtgDatos.Rows(x).Cells(5).Value = "PLANTA" Then
                                    dtgDatos.Rows(x).Cells(55).Value = rwCostoSocial(0)("imsstopado")
                                    dtgDatos.Rows(x).Cells(56).Value = rwCostoSocial(0)("RCVtopado")
                                    dtgDatos.Rows(x).Cells(57).Value = rwCostoSocial(0)("infonavittopado")
                                    dtgDatos.Rows(x).Cells(58).Value = rwCostoSocial(0)("ISNtopado")
                                    dtgDatos.Rows(x).Cells(59).Value = Math.Round(Double.Parse(dtgDatos.Rows(x).Cells(55).Value) + Double.Parse(dtgDatos.Rows(x).Cells(56).Value) + Double.Parse(dtgDatos.Rows(x).Cells(57).Value) + Double.Parse(dtgDatos.Rows(x).Cells(58).Value), 2)
                                Else
                                    dtgDatos.Rows(x).Cells(55).Value = Math.Round(Double.Parse(rwCostoSocial(0)("imsstopado")) / 30 * dtgDatos.Rows(x).Cells(18).Value, 2)
                                    dtgDatos.Rows(x).Cells(56).Value = Math.Round(Double.Parse(rwCostoSocial(0)("RCVtopado")) / 30 * dtgDatos.Rows(x).Cells(18).Value, 2)
                                    dtgDatos.Rows(x).Cells(57).Value = Math.Round(Double.Parse(rwCostoSocial(0)("infonavittopado")) / 30 * dtgDatos.Rows(x).Cells(18).Value, 2)
                                    dtgDatos.Rows(x).Cells(58).Value = Math.Round(Double.Parse(rwCostoSocial(0)("ISNtopado")) / 30 * dtgDatos.Rows(x).Cells(18).Value, 2)
                                    dtgDatos.Rows(x).Cells(59).Value = Math.Round(Double.Parse(dtgDatos.Rows(x).Cells(55).Value) + Double.Parse(dtgDatos.Rows(x).Cells(56).Value) + Double.Parse(dtgDatos.Rows(x).Cells(57).Value) + Double.Parse(dtgDatos.Rows(x).Cells(58).Value), 2)
                                End If

                            Else
                                If dtgDatos.Rows(x).Cells(5).Value = "PLANTA" Then
                                    dtgDatos.Rows(x).Cells(55).Value = rwCostoSocial(0)("imss")
                                    dtgDatos.Rows(x).Cells(56).Value = rwCostoSocial(0)("RCV")
                                    dtgDatos.Rows(x).Cells(57).Value = rwCostoSocial(0)("Infonavit")
                                    dtgDatos.Rows(x).Cells(58).Value = rwCostoSocial(0)("ISN")
                                    dtgDatos.Rows(x).Cells(59).Value = Math.Round(Double.Parse(dtgDatos.Rows(x).Cells(55).Value) + Double.Parse(dtgDatos.Rows(x).Cells(56).Value) + Double.Parse(dtgDatos.Rows(x).Cells(57).Value) + Double.Parse(dtgDatos.Rows(x).Cells(58).Value), 2)
                                Else
                                    dtgDatos.Rows(x).Cells(55).Value = Math.Round(Double.Parse(rwCostoSocial(0)("imss")) / 30 * dtgDatos.Rows(x).Cells(18).Value, 2)
                                    dtgDatos.Rows(x).Cells(56).Value = Math.Round(Double.Parse(rwCostoSocial(0)("RCV")) / 30 * dtgDatos.Rows(x).Cells(18).Value, 2)
                                    dtgDatos.Rows(x).Cells(57).Value = Math.Round(Double.Parse(rwCostoSocial(0)("Infonavit")) / 30 * dtgDatos.Rows(x).Cells(18).Value, 2)
                                    dtgDatos.Rows(x).Cells(58).Value = Math.Round(Double.Parse(rwCostoSocial(0)("ISN")) / 30 * dtgDatos.Rows(x).Cells(18).Value, 2)
                                    dtgDatos.Rows(x).Cells(59).Value = Math.Round(Double.Parse(dtgDatos.Rows(x).Cells(55).Value) + Double.Parse(dtgDatos.Rows(x).Cells(56).Value) + Double.Parse(dtgDatos.Rows(x).Cells(57).Value) + Double.Parse(dtgDatos.Rows(x).Cells(58).Value), 2)
                                End If
                            End If
                        End If
                    End If
                Else
                    dtgDatos.Rows(x).Cells(55).Value = Math.Round(Double.Parse("0.00"), 2)
                    dtgDatos.Rows(x).Cells(56).Value = Math.Round(Double.Parse("0.00"), 2)
                    dtgDatos.Rows(x).Cells(57).Value = Math.Round(Double.Parse("0.00"), 2)
                    dtgDatos.Rows(x).Cells(58).Value = Math.Round(Double.Parse("0.00"), 2)
                    dtgDatos.Rows(x).Cells(59).Value = Math.Round(Double.Parse(dtgDatos.Rows(x).Cells(55).Value) + Double.Parse(dtgDatos.Rows(x).Cells(56).Value) + Double.Parse(dtgDatos.Rows(x).Cells(57).Value) + Double.Parse(dtgDatos.Rows(x).Cells(58).Value), 2)
                End If

                'TOTAL COSTO SOCIAL
                CostoSocialTotal = Math.Round(Double.Parse(dtgDatos.Rows(x).Cells(55).Value) + Double.Parse(dtgDatos.Rows(x).Cells(56).Value) + Double.Parse(dtgDatos.Rows(x).Cells(57).Value) + Double.Parse(dtgDatos.Rows(x).Cells(58).Value), 2)
                dtgDatos.Rows(x).Cells(59).Value = CostoSocialTotal

                'SUBTOTAL
                subtotal = Math.Round(ComplementoAsimilados + PrestamoPersonalAsimilados + AdeudoINfonavitAsimilados + DiferenciaInfonavitAsimilados + Operadora + RetencionOperadora + ComisionOperadora + ComisionAsimilados + CostoSocialTotal, 2)
                dtgDatos.Rows(x).Cells(60).Value = subtotal

                'IVA
                iva = Math.Round(subtotal * 0.16)
                dtgDatos.Rows(x).Cells(61).Value = iva
                'TOTAL DEPOSITO
                dtgDatos.Rows(x).Cells(62).Value = subtotal + iva




                pgbProgreso.Value += 1
                Application.DoEvents()
            Next

            'verificar costo social

            Dim contador, Posicion1, Posicion2, Posicion3, Posicion4 As Integer


            For x As Integer = 0 To dtgDatos.Rows.Count - 1

                'calcular ISN

                Dim TotalParaISN As Double = Double.Parse(dtgDatos.Rows(x).Cells(33).Value) - Double.Parse(dtgDatos.Rows(x).Cells(22).Value) - Double.Parse(dtgDatos.Rows(x).Cells(23).Value)


                If chkSoloCostoSocial.Checked = True Then
                    If dtgDatos.Rows(x).Cells(59).Tag = "" Then

                        If dtgDatos.Rows(x).Cells(2).Tag = "" Then
                            If dtgDatos.Rows(x).Cells(4).Tag = "1" Then
                                dtgDatos.Rows(x).Cells(56).Value = Math.Round(0, 2)
                                dtgDatos.Rows(x).Cells(57).Value = Math.Round(0, 2)
                                dtgDatos.Rows(x).Cells(58).Value = Math.Round(0, 2)
                                dtgDatos.Rows(x).Cells(59).Value = Math.Round(TotalParaISN * 0.03 + (TotalParaISN * 0.03 * 0.15), 2)
                                dtgDatos.Rows(x).Cells(60).Value = Math.Round(Double.Parse(dtgDatos.Rows(x).Cells(56).Value) + Double.Parse(dtgDatos.Rows(x).Cells(57).Value) + Double.Parse(dtgDatos.Rows(x).Cells(58).Value) + Double.Parse(dtgDatos.Rows(x).Cells(59).Value), 2)
                            Else
                                contador = 0
                                'Posicion1 = x
                                For y As Integer = x To dtgDatos.Rows.Count - 1
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
                                    End If



                                Next

                                

                                If contador > 1 Then
                                    If contador = 2 Then

                                        Dim Valores() As Date = New Date(1) {Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value), Date.Parse(dtgDatos.Rows(Posicion2).Cells(63).Value)}
                                        'Dim Auxiliar As Date
                                        Dim Longitud As Integer = Valores.Length - 1
                                        Array.Sort(Valores)
                                        Dim iniciomes As Date = "01/" & Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value).Month & "/" & Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value).Year
                                        'aqui a la fecha inicial del segundo le restamos uno para saber asegurar cobrar correctamente
                                        Dim final1 As Date = Date.Parse(dtgDatos.Rows(Posicion2).Cells(63).Value).AddDays(-1)
                                        Dim dias As Integer = (DateDiff("y", iniciomes, final1)) + 1

                                        'MessageBox.Show(Valores(2))

                                        dtgDatos.Rows(Posicion1).Cells(59).Style.BackColor = Color.Chocolate
                                        dtgDatos.Rows(Posicion1).Cells(59).Tag = "1"
                                        dtgDatos.Rows(Posicion2).Cells(59).Style.BackColor = Color.Chocolate
                                        dtgDatos.Rows(Posicion2).Cells(59).Tag = "1"


                                        Dim dsarreglo As New DataSet
                                        dsarreglo.Tables.Add("Tabla")
                                        dsarreglo.Tables("Tabla").Columns.Add("Posicion")
                                        dsarreglo.Tables("Tabla").Columns.Add("fechainicial")
                                        dsarreglo.Tables("Tabla").Columns.Add("fechafina")
                                        'dsarreglo.Tables("Tabla").Columns.Add("Nombre")
                                        For z = 0 To Longitud
                                            Dim fila As DataRow = dsarreglo.Tables("Tabla").NewRow

                                            fila.Item("Posicion") = ""
                                            fila.Item("fechainicial") = Valores(z)
                                            fila.Item("fechafina") = ""
                                            dsarreglo.Tables("Tabla").Rows.Add(fila)
                                        Next
                                        For z = 0 To Longitud

                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value) And z = 0 Then
                                                dsarreglo.Tables(0).Rows(0)("Posicion") = Posicion1
                                                dsarreglo.Tables(0).Rows(0)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion1).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion2).Cells(63).Value) And z = 0 Then
                                                dsarreglo.Tables(0).Rows(0)("Posicion") = Posicion2
                                                dsarreglo.Tables(0).Rows(0)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion2).Cells(64).Value)
                                            End If

                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value) And z = 1 Then
                                                dsarreglo.Tables(0).Rows(1)("Posicion") = Posicion1
                                                dsarreglo.Tables(0).Rows(1)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion1).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion2).Cells(63).Value) And z = 1 Then
                                                dsarreglo.Tables(0).Rows(1)("Posicion") = Posicion2
                                                dsarreglo.Tables(0).Rows(1)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion2).Cells(64).Value)
                                            End If

                                            'If  = Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value) Then

                                        Next
                                        'recorrer el ds

                                        For q = 0 To dsarreglo.Tables(0).Rows.Count - 1
                                            If q = 0 Then
                                                iniciomes = "01/" & Date.Parse(dsarreglo.Tables(0).Rows(0)("fechainicial")).Month & "/" & Date.Parse(dsarreglo.Tables(0).Rows(0)("fechainicial")).Year
                                                'aqui a la fecha inicial del segundo le restamos uno para saber asegurar cobrar correctamente
                                                final1 = Date.Parse(dsarreglo.Tables(0).Rows(1)("fechainicial")).AddDays(-1)
                                                dias = (DateDiff("y", iniciomes, final1)) + 1
                                                calcularcostosocial(dsarreglo.Tables(0).Rows(0)("Posicion"), dias)
                                            End If
                                            If q = 1 Then
                                                iniciomes = Date.Parse(dsarreglo.Tables(0).Rows(1)("fechainicial"))
                                                'aqui a la fecha inicial del segundo le restamos uno para saber asegurar cobrar correctamente
                                                final1 = Date.Parse(DiasMes(cboperiodo.SelectedValue) & "/" & Date.Parse(dsarreglo.Tables(0).Rows(1)("fechafina")).Month & "/" & Date.Parse(dsarreglo.Tables(0).Rows(1)("fechafina")).Year)
                                                dias = (DateDiff("y", iniciomes, final1)) + 1
                                                calcularcostosocial(dsarreglo.Tables(0).Rows(1)("Posicion"), dias)
                                            End If

                                            'dsarreglo.Tables(0).Rows
                                        Next



                                    End If


                                    If contador = 3 Then

                                        Dim Valores() As Date = New Date(2) {Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value), Date.Parse(dtgDatos.Rows(Posicion2).Cells(63).Value), Date.Parse(dtgDatos.Rows(Posicion3).Cells(63).Value)}
                                        'Dim Auxiliar As Date
                                        Dim Longitud As Integer = Valores.Length - 1
                                        Array.Sort(Valores)
                                        Dim iniciomes As Date = "01/" & Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value).Month & "/" & Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value).Year
                                        'aqui a la fecha inicial del segundo le restamos uno para saber asegurar cobrar correctamente
                                        Dim final1 As Date = Date.Parse(dtgDatos.Rows(Posicion2).Cells(63).Value).AddDays(-1)
                                        Dim dias As Integer = (DateDiff("y", iniciomes, final1)) + 1

                                        'MessageBox.Show(Valores(2))

                                        dtgDatos.Rows(Posicion1).Cells(59).Style.BackColor = Color.Chocolate
                                        dtgDatos.Rows(Posicion1).Cells(59).Tag = "1"
                                        dtgDatos.Rows(Posicion2).Cells(59).Style.BackColor = Color.Chocolate
                                        dtgDatos.Rows(Posicion2).Cells(59).Tag = "1"
                                        dtgDatos.Rows(Posicion3).Cells(59).Style.BackColor = Color.Chocolate
                                        dtgDatos.Rows(Posicion3).Cells(59).Tag = "1"

                                        Dim dsarreglo As New DataSet
                                        dsarreglo.Tables.Add("Tabla")
                                        dsarreglo.Tables("Tabla").Columns.Add("Posicion")
                                        dsarreglo.Tables("Tabla").Columns.Add("fechainicial")
                                        dsarreglo.Tables("Tabla").Columns.Add("fechafina")
                                        'dsarreglo.Tables("Tabla").Columns.Add("Nombre")
                                        For z = 0 To Longitud
                                            Dim fila As DataRow = dsarreglo.Tables("Tabla").NewRow

                                            fila.Item("Posicion") = ""
                                            fila.Item("fechainicial") = Valores(z)
                                            fila.Item("fechafina") = ""
                                            dsarreglo.Tables("Tabla").Rows.Add(fila)
                                        Next
                                        For z = 0 To Longitud

                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value) And z = 0 Then
                                                dsarreglo.Tables(0).Rows(0)("Posicion") = Posicion1
                                                dsarreglo.Tables(0).Rows(0)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion1).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion2).Cells(63).Value) And z = 0 Then
                                                dsarreglo.Tables(0).Rows(0)("Posicion") = Posicion2
                                                dsarreglo.Tables(0).Rows(0)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion2).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion3).Cells(63).Value) And z = 0 Then
                                                dsarreglo.Tables(0).Rows(0)("Posicion") = Posicion3
                                                dsarreglo.Tables(0).Rows(0)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion3).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value) And z = 1 Then
                                                dsarreglo.Tables(0).Rows(1)("Posicion") = Posicion1
                                                dsarreglo.Tables(0).Rows(1)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion1).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion2).Cells(63).Value) And z = 1 Then
                                                dsarreglo.Tables(0).Rows(1)("Posicion") = Posicion2
                                                dsarreglo.Tables(0).Rows(1)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion2).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion3).Cells(63).Value) And z = 1 Then
                                                dsarreglo.Tables(0).Rows(1)("Posicion") = Posicion3
                                                dsarreglo.Tables(0).Rows(1)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion3).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value) And z = 2 Then
                                                dsarreglo.Tables(0).Rows(2)("Posicion") = Posicion1
                                                dsarreglo.Tables(0).Rows(2)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion1).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion2).Cells(63).Value) And z = 2 Then
                                                dsarreglo.Tables(0).Rows(2)("Posicion") = Posicion2
                                                dsarreglo.Tables(0).Rows(2)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion2).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion3).Cells(63).Value) And z = 2 Then
                                                dsarreglo.Tables(0).Rows(2)("Posicion") = Posicion3
                                                dsarreglo.Tables(0).Rows(2)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion3).Cells(64).Value)
                                            End If

                                            'If  = Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value) Then

                                        Next
                                        'recorrer el ds

                                        For q = 0 To dsarreglo.Tables(0).Rows.Count - 1
                                            If q = 0 Then
                                                iniciomes = "01/" & Date.Parse(dsarreglo.Tables(0).Rows(0)("fechainicial")).Month & "/" & Date.Parse(dsarreglo.Tables(0).Rows(0)("fechainicial")).Year
                                                'aqui a la fecha inicial del segundo le restamos uno para saber asegurar cobrar correctamente
                                                final1 = Date.Parse(dsarreglo.Tables(0).Rows(1)("fechainicial")).AddDays(-1)
                                                dias = (DateDiff("y", iniciomes, final1)) + 1
                                                calcularcostosocial(dsarreglo.Tables(0).Rows(0)("Posicion"), dias)
                                            End If
                                            If q = 1 Then
                                                iniciomes = Date.Parse(dsarreglo.Tables(0).Rows(1)("fechainicial"))
                                                'aqui a la fecha inicial del segundo le restamos uno para saber asegurar cobrar correctamente
                                                final1 = Date.Parse(dsarreglo.Tables(0).Rows(2)("fechainicial")).AddDays(-1)
                                                dias = (DateDiff("y", iniciomes, final1)) + 1
                                                calcularcostosocial(dsarreglo.Tables(0).Rows(1)("Posicion"), dias)
                                            End If
                                            If q = 2 Then
                                                iniciomes = Date.Parse(dsarreglo.Tables(0).Rows(2)("fechainicial"))
                                                'aqui a la fecha inicial del segundo le restamos uno para saber asegurar cobrar correctamente
                                                final1 = Date.Parse(DiasMes(cboperiodo.SelectedValue) & "/" & Date.Parse(dsarreglo.Tables(0).Rows(2)("fechafina")).Month & "/" & Date.Parse(dsarreglo.Tables(0).Rows(2)("fechafina")).Year)
                                                dias = (DateDiff("y", iniciomes, final1)) + 1
                                                calcularcostosocial(dsarreglo.Tables(0).Rows(2)("Posicion"), dias)
                                            End If
                                            'dsarreglo.Tables(0).Rows
                                        Next

                                    End If
                                    If contador = 4 Then
                                        Dim Valores() As Date = New Date(3) {Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value), Date.Parse(dtgDatos.Rows(Posicion2).Cells(63).Value), Date.Parse(dtgDatos.Rows(Posicion3).Cells(63).Value), Date.Parse(dtgDatos.Rows(Posicion3).Cells(63).Value)}
                                        'Dim Auxiliar As Date
                                        Dim Longitud As Integer = Valores.Length - 1
                                        Array.Sort(Valores)
                                        Dim iniciomes As Date = "01/" & Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value).Month & "/" & Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value).Year
                                        'aqui a la fecha inicial del segundo le restamos uno para saber asegurar cobrar correctamente
                                        Dim final1 As Date = Date.Parse(dtgDatos.Rows(Posicion2).Cells(63).Value).AddDays(-1)
                                        Dim dias As Integer = (DateDiff("y", iniciomes, final1)) + 1

                                        'MessageBox.Show(Valores(2))

                                        'dtgDatos.Rows(Posicion1).Cells(59).Style.BackColor = Color.Chocolate
                                        'dtgDatos.Rows(Posicion1).Cells(59).Tag = "1"
                                        'dtgDatos.Rows(Posicion2).Cells(59).Style.BackColor = Color.Chocolate
                                        'dtgDatos.Rows(Posicion2).Cells(59).Tag = "1"
                                        'dtgDatos.Rows(Posicion3).Cells(59).Style.BackColor = Color.Chocolate
                                        'dtgDatos.Rows(Posicion3).Cells(59).Tag = "1"

                                        Dim dsarreglo As New DataSet
                                        dsarreglo.Tables.Add("Tabla")
                                        dsarreglo.Tables("Tabla").Columns.Add("Posicion")
                                        dsarreglo.Tables("Tabla").Columns.Add("fechainicial")
                                        dsarreglo.Tables("Tabla").Columns.Add("fechafina")
                                        'dsarreglo.Tables("Tabla").Columns.Add("Nombre")
                                        For z = 0 To Longitud
                                            Dim fila As DataRow = dsarreglo.Tables("Tabla").NewRow

                                            fila.Item("Posicion") = ""
                                            fila.Item("fechainicial") = Valores(z)
                                            fila.Item("fechafina") = ""
                                            dsarreglo.Tables("Tabla").Rows.Add(fila)
                                        Next
                                        For z = 0 To Longitud

                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value) And z = 0 Then
                                                dsarreglo.Tables(0).Rows(0)("Posicion") = Posicion1
                                                dsarreglo.Tables(0).Rows(0)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion1).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion2).Cells(63).Value) And z = 0 Then
                                                dsarreglo.Tables(0).Rows(0)("Posicion") = Posicion2
                                                dsarreglo.Tables(0).Rows(0)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion2).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion3).Cells(63).Value) And z = 0 Then
                                                dsarreglo.Tables(0).Rows(0)("Posicion") = Posicion3
                                                dsarreglo.Tables(0).Rows(0)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion3).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion4).Cells(63).Value) And z = 0 Then
                                                dsarreglo.Tables(0).Rows(0)("Posicion") = Posicion4
                                                dsarreglo.Tables(0).Rows(0)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion4).Cells(64).Value)
                                            End If

                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value) And z = 1 Then
                                                dsarreglo.Tables(0).Rows(1)("Posicion") = Posicion1
                                                dsarreglo.Tables(0).Rows(1)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion1).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion2).Cells(63).Value) And z = 1 Then
                                                dsarreglo.Tables(0).Rows(1)("Posicion") = Posicion2
                                                dsarreglo.Tables(0).Rows(1)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion2).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion3).Cells(63).Value) And z = 1 Then
                                                dsarreglo.Tables(0).Rows(1)("Posicion") = Posicion3
                                                dsarreglo.Tables(0).Rows(1)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion3).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion4).Cells(63).Value) And z = 1 Then
                                                dsarreglo.Tables(0).Rows(1)("Posicion") = Posicion4
                                                dsarreglo.Tables(0).Rows(1)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion4).Cells(64).Value)
                                            End If

                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value) And z = 2 Then
                                                dsarreglo.Tables(0).Rows(2)("Posicion") = Posicion1
                                                dsarreglo.Tables(0).Rows(2)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion1).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion2).Cells(63).Value) And z = 2 Then
                                                dsarreglo.Tables(0).Rows(2)("Posicion") = Posicion2
                                                dsarreglo.Tables(0).Rows(2)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion2).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion3).Cells(63).Value) And z = 2 Then
                                                dsarreglo.Tables(0).Rows(2)("Posicion") = Posicion3
                                                dsarreglo.Tables(0).Rows(2)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion3).Cells(64).Value)
                                            End If
                                            If Valores(z) = Date.Parse(dtgDatos.Rows(Posicion4).Cells(63).Value) And z = 2 Then
                                                dsarreglo.Tables(0).Rows(2)("Posicion") = Posicion4
                                                dsarreglo.Tables(0).Rows(2)("fechafina") = Date.Parse(dtgDatos.Rows(Posicion4).Cells(64).Value)
                                            End If


                                            'If  = Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value) Then

                                        Next
                                        'recorrer el ds

                                        For q = 0 To dsarreglo.Tables(0).Rows.Count - 1
                                            If q = 0 Then
                                                iniciomes = "01/" & Date.Parse(dsarreglo.Tables(0).Rows(0)("fechainicial")).Month & "/" & Date.Parse(dsarreglo.Tables(0).Rows(0)("fechainicial")).Year
                                                'aqui a la fecha inicial del segundo le restamos uno para saber asegurar cobrar correctamente
                                                final1 = Date.Parse(dsarreglo.Tables(0).Rows(1)("fechainicial")).AddDays(-1)
                                                dias = (DateDiff("y", iniciomes, final1)) + 1
                                                calcularcostosocial(dsarreglo.Tables(0).Rows(0)("Posicion"), dias)
                                            End If
                                            If q = 1 Then
                                                iniciomes = Date.Parse(dsarreglo.Tables(0).Rows(1)("fechainicial"))
                                                'aqui a la fecha inicial del segundo le restamos uno para saber asegurar cobrar correctamente
                                                final1 = Date.Parse(dsarreglo.Tables(0).Rows(2)("fechainicial")).AddDays(-1)
                                                dias = (DateDiff("y", iniciomes, final1)) + 1
                                                calcularcostosocial(dsarreglo.Tables(0).Rows(1)("Posicion"), dias)
                                            End If

                                            If q = 2 Then
                                                iniciomes = Date.Parse(dsarreglo.Tables(0).Rows(2)("fechainicial"))
                                                'aqui a la fecha inicial del segundo le restamos uno para saber asegurar cobrar correctamente
                                                final1 = Date.Parse(dsarreglo.Tables(0).Rows(3)("fechainicial")).AddDays(-1)
                                                dias = (DateDiff("y", iniciomes, final1)) + 1
                                                calcularcostosocial(dsarreglo.Tables(0).Rows(2)("Posicion"), dias)
                                            End If
                                            If q = 3 Then
                                                iniciomes = Date.Parse(dsarreglo.Tables(0).Rows(3)("fechainicial"))
                                                'aqui a la fecha inicial del segundo le restamos uno para saber asegurar cobrar correctamente
                                                final1 = Date.Parse(DiasMes(cboperiodo.SelectedValue) & "/" & Date.Parse(dsarreglo.Tables(0).Rows(3)("fechafina")).Month & "/" & Date.Parse(dsarreglo.Tables(0).Rows(3)("fechafina")).Year)
                                                dias = (DateDiff("y", iniciomes, final1)) + 1
                                                calcularcostosocial(dsarreglo.Tables(0).Rows(3)("Posicion"), dias)
                                            End If
                                            'dsarreglo.Tables(0).Rows
                                        Next
                                    End If
                                Else
                                    Dim iniciomes As Date = "01/" & Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value).Month & "/" & Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value).Year
                                    'aqui a la fecha inicial del segundo le restamos uno para saber asegurar cobrar correctamente
                                    Dim final1 As Date = Date.Parse(dtgDatos.Rows(Posicion2).Cells(63).Value).AddDays(-1)
                                    Dim dias As Integer = (DateDiff("y", iniciomes, final1)) + 1
                                    iniciomes = "01/" & Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value).Month & "/" & Date.Parse(dtgDatos.Rows(Posicion1).Cells(63).Value).Year
                                    'aqui a la fecha inicial del segundo le restamos uno para saber asegurar cobrar correctamente
                                    final1 = Date.Parse(DiasMes(cboperiodo.SelectedValue) & "/" & Date.Parse(dtgDatos.Rows(Posicion1).Cells(64).Value).Month & "/" & Date.Parse(dtgDatos.Rows(Posicion1).Cells(64).Value).Year)
                                    dias = (DateDiff("y", iniciomes, final1)) + 1
                                    calcularcostosocial(Posicion1, dias)
                                End If
                                '#### ordenamiento de los repetidos interno


                            End If

                            
                        End If

                    End If

                    
                Else
                    contador = 0

                    For y As Integer = 0 To dtgDatos.Rows.Count - 1
                        If dtgDatos.Rows(x).Cells(2).Value = dtgDatos.Rows(y).Cells(2).Value Then
                            contador = contador + 1
                            If contador = 2 Then
                                Posicion2 = y
                            End If
                            If contador = 3 Then
                                Posicion3 = y
                            End If
                            If contador = 4 Then
                                Posicion4 = y
                            End If
                        End If



                    Next
                    If contador = 2 Then
                        If dtgDatos.Rows(Posicion2).Cells(5).Value = "PLANTA" Then
                            dtgDatos.Rows(Posicion2).Cells(55).Value = "0.00"
                            dtgDatos.Rows(Posicion2).Cells(56).Value = "0.00"
                            dtgDatos.Rows(Posicion2).Cells(57).Value = "0.00"
                            dtgDatos.Rows(Posicion2).Cells(58).Value = "0.00"
                            dtgDatos.Rows(Posicion2).Cells(59).Value = "0.00"
                        End If

                    End If
                    If contador = 3 Then
                        If dtgDatos.Rows(Posicion2).Cells(5).Value = "PLANTA" Then
                            dtgDatos.Rows(Posicion2).Cells(55).Value = "0.00"
                            dtgDatos.Rows(Posicion2).Cells(56).Value = "0.00"
                            dtgDatos.Rows(Posicion2).Cells(57).Value = "0.00"
                            dtgDatos.Rows(Posicion2).Cells(58).Value = "0.00"
                            dtgDatos.Rows(Posicion2).Cells(59).Value = "0.00"
                        End If
                        If dtgDatos.Rows(Posicion3).Cells(5).Value = "PLANTA" Then
                            dtgDatos.Rows(Posicion3).Cells(55).Value = "0.00"
                            dtgDatos.Rows(Posicion3).Cells(56).Value = "0.00"
                            dtgDatos.Rows(Posicion3).Cells(57).Value = "0.00"
                            dtgDatos.Rows(Posicion3).Cells(58).Value = "0.00"
                            dtgDatos.Rows(Posicion3).Cells(59).Value = "0.00"
                        End If
                    End If
                    If contador = 4 Then
                        If dtgDatos.Rows(Posicion2).Cells(5).Value = "PLANTA" Then
                            dtgDatos.Rows(Posicion2).Cells(55).Value = "0.00"
                            dtgDatos.Rows(Posicion2).Cells(56).Value = "0.00"
                            dtgDatos.Rows(Posicion2).Cells(57).Value = "0.00"
                            dtgDatos.Rows(Posicion2).Cells(58).Value = "0.00"
                            dtgDatos.Rows(Posicion2).Cells(59).Value = "0.00"
                        End If
                        If dtgDatos.Rows(Posicion3).Cells(5).Value = "PLANTA" Then
                            dtgDatos.Rows(Posicion3).Cells(55).Value = "0.00"
                            dtgDatos.Rows(Posicion3).Cells(56).Value = "0.00"
                            dtgDatos.Rows(Posicion3).Cells(57).Value = "0.00"
                            dtgDatos.Rows(Posicion3).Cells(58).Value = "0.00"
                            dtgDatos.Rows(Posicion3).Cells(59).Value = "0.00"
                        End If
                        If dtgDatos.Rows(Posicion4).Cells(5).Value = "PLANTA" Then
                            dtgDatos.Rows(Posicion4).Cells(55).Value = "0.00"
                            dtgDatos.Rows(Posicion4).Cells(56).Value = "0.00"
                            dtgDatos.Rows(Posicion4).Cells(57).Value = "0.00"
                            dtgDatos.Rows(Posicion4).Cells(58).Value = "0.00"
                            dtgDatos.Rows(Posicion3).Cells(59).Value = "0.00"
                        End If
                    End If

                End If



            Next

            pnlProgreso.Visible = False
            pnlCatalogo.Enabled = True
            MessageBox.Show("Datos calculados ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Function Bisiesto(Num As Integer) As Boolean
        If Num Mod 4 = 0 And (Num Mod 100 Or Num Mod 400 = 0) Then
            Bisiesto = True
        Else
            Bisiesto = False
        End If
    End Function

    Private Function infonavit(tipo As String, valor As Double, sdi As Double, fechapago As Date, periodo As String, diastrabajados As Integer, idempleado As Integer, consecutivo As Integer) As Double
        Try
            Dim numbimestre As Integer
            Dim numbimestre2 As Integer
            Dim numdias As Integer
            Dim numdias2 As Integer
            Dim DiasCadaPeriodo As Integer
            Dim DiasCadaPeriodo2 As Integer
            Dim diasfebrero As Integer
            Dim valorinfonavit As Double
            Dim sql As String
            Dim FechaInicioPeriodo1 As Date
            Dim FechaFinPeriodo1 As Date
            Dim FechaInicioPeriodo2 As Date
            Dim FechaFinPeriodo2 As Date
            Dim Seguro1 As Double
            Dim Seguro2 As Double
            Dim ValorInfonavitTabla As Double
            Dim contador As Integer

            'Validamos si el trabajador tiene o no activo el infonavit
            sql = "select iPermanente from empleadosC where iIdEmpleadoC=" & idempleado
            Dim rwCalcularInfonavit As DataRow() = nConsulta(sql)
            If rwCalcularInfonavit Is Nothing = False Then
                If rwCalcularInfonavit(0)("iPermanente") = "1" Then
                    sql = "select * from periodos where iIdPeriodo= " & periodo
                    Dim rwPeriodo As DataRow() = nConsulta(sql)

                    If rwPeriodo Is Nothing = False Then

                        If diastrabajados = 30 Then
                            FechaInicioPeriodo1 = Date.Parse(rwPeriodo(0)("dFechaInicio"))
                            FechaFinPeriodo1 = Date.Parse("01/" & FechaInicioPeriodo1.Month & "/" & FechaInicioPeriodo1.Year).AddMonths(1).AddDays(-1)
                            FechaFinPeriodo2 = Date.Parse(rwPeriodo(0)("dFechaFin"))
                            FechaInicioPeriodo2 = Date.Parse("01/" & FechaFinPeriodo2.Month & "/" & FechaFinPeriodo2.Year)
                            If (FechaInicioPeriodo1 = FechaInicioPeriodo2) Then
                                FechaInicioPeriodo2 = Date.Parse("01/01/1900")
                            End If

                            If (FechaFinPeriodo1 = FechaFinPeriodo2) Then
                                FechaFinPeriodo2 = Date.Parse("01/01/1900")
                            End If
                        Else
                            'Verificamos si tiene un embarque dentro de periodo
                            sql = "select * from DatosEmbarque where FechaEmbarque Between '" & Date.Parse(rwPeriodo(0)("dFechaInicio")).ToShortDateString & "' and '" & Date.Parse(rwPeriodo(0)("dFechaFin")).ToShortDateString & "'"
                            Dim rwDatosEmbarque As DataRow() = nConsulta(sql)
                            If rwDatosEmbarque Is Nothing = False Then
                                FechaInicioPeriodo1 = rwDatosEmbarque(0)("FechaEmbarque")
                                FechaFinPeriodo2 = FechaInicioPeriodo1.AddDays(diastrabajados)
                                FechaFinPeriodo2 = FechaFinPeriodo2.AddDays(-1)

                                If FechaInicioPeriodo1.Month = FechaFinPeriodo2.Month Then
                                    FechaFinPeriodo1 = FechaInicioPeriodo1.AddDays(diastrabajados - 1)
                                    FechaInicioPeriodo2 = Date.Parse("01/01/1900")
                                    FechaFinPeriodo2 = Date.Parse("01/01/1900")

                                Else

                                    FechaFinPeriodo1 = Date.Parse("01/" & FechaFinPeriodo1.Month & "/" & FechaInicioPeriodo1.Year).AddMonths(1).AddDays(-1)
                                    FechaInicioPeriodo2 = Date.Parse("01/" & FechaFinPeriodo2.Month & "/" & FechaFinPeriodo2.Year)
                                End If


                            Else
                                'Si no lo tiene sumamos de inicio del periodo hasta el numero de dias
                                'Verificamos si esta dentro del mismo mes
                                FechaInicioPeriodo1 = Date.Parse(rwPeriodo(0)("dFechaInicio"))
                                FechaFinPeriodo2 = FechaInicioPeriodo1.AddDays(diastrabajados)
                                FechaFinPeriodo2 = FechaFinPeriodo2.AddDays(-1)
                                If FechaInicioPeriodo1.Month = FechaFinPeriodo2.Month Then
                                    FechaFinPeriodo1 = FechaInicioPeriodo1.AddDays(diastrabajados - 1)
                                    FechaInicioPeriodo2 = Date.Parse("01/01/1900")
                                    FechaFinPeriodo2 = Date.Parse("01/01/1900")

                                Else
                                    FechaFinPeriodo1 = Date.Parse("01/" & FechaInicioPeriodo1.Month & "/" & FechaInicioPeriodo1.Year).AddMonths(1).AddDays(-1)
                                    FechaInicioPeriodo2 = Date.Parse("01/" & FechaFinPeriodo2.Month & "/" & FechaFinPeriodo2.Year)
                                End If
                            End If
                        End If





                        If Month(FechaInicioPeriodo1) Mod 2 = 0 Then
                            numbimestre = Month(FechaInicioPeriodo1) / 2
                        Else
                            numbimestre = (Month(FechaInicioPeriodo1) + 1) / 2
                        End If

                        If numbimestre = 1 Then
                            If Bisiesto(Year(FechaInicioPeriodo1)) = True Then
                                diasfebrero = 29
                            Else
                                diasfebrero = 28
                            End If
                            'diasfebrero = Day(DateSerial(Year(fechapago), 3, 0))
                            numdias = 31 + diasfebrero
                        End If

                        If numbimestre = 2 Then
                            numdias = 61
                        End If

                        If numbimestre = 3 Then
                            numdias = 61
                        End If

                        If numbimestre = 4 Then
                            numdias = 62
                        End If

                        If numbimestre = 5 Then
                            numdias = 61
                        End If

                        If numbimestre = 6 Then
                            numdias = 61
                        End If



                        If Month(FechaInicioPeriodo2) Mod 2 = 0 Then
                            numbimestre2 = Month(FechaInicioPeriodo2) / 2
                        Else
                            numbimestre2 = (Month(FechaInicioPeriodo2) + 1) / 2
                        End If

                        If numbimestre2 = 1 Then
                            If Bisiesto(Year(FechaInicioPeriodo1)) = True Then
                                diasfebrero = 29
                            Else
                                diasfebrero = 28
                            End If
                            'diasfebrero = Day(DateSerial(Year(fechapago), 3, 0))
                            numdias2 = 31 + diasfebrero
                        End If

                        If numbimestre2 = 2 Then
                            numdias2 = 61
                        End If

                        If numbimestre2 = 3 Then
                            numdias2 = 61
                        End If

                        If numbimestre2 = 4 Then
                            numdias2 = 62
                        End If

                        If numbimestre2 = 5 Then
                            numdias2 = 61
                        End If

                        If numbimestre2 = 6 Then
                            numdias2 = 61
                        End If



                        DiasCadaPeriodo = DateDiff(DateInterval.Day, FechaInicioPeriodo1, FechaFinPeriodo1) + 1

                        'Verificamos si ya existe el seguro en ese bimestre

                        sql = "select * from PagoSeguroInfonavit where fkiIdEmpleadoC= " & idempleado
                        sql &= " And NumBimestre= " & numbimestre & " And Anio=" & FechaInicioPeriodo1.Year.ToString
                        Dim rwSeguro1 As DataRow() = nConsulta(sql)

                        If rwSeguro1 Is Nothing = False Then
                            Seguro1 = 0
                        Else
                            If cboTipoNomina.SelectedIndex = 0 Then
                                contador = 0
                                For x As Integer = 0 To consecutivo

                                    If idempleado = 50 Then
                                        'llego
                                        contador = contador
                                    End If
                                    If dtgDatos.Rows(x).Cells(2).Value = idempleado Then
                                        contador = contador + 1
                                    End If

                                Next

                                If contador = 1 Then
                                    Seguro1 = 15
                                End If



                            Else
                                Seguro1 = 0
                            End If

                        End If

                        If FechaInicioPeriodo2 = Date.Parse("01/01/1900") Then
                            DiasCadaPeriodo2 = 0
                            Seguro2 = 0

                        Else
                            DiasCadaPeriodo2 = DateDiff(DateInterval.Day, FechaInicioPeriodo2, FechaFinPeriodo2) + 1
                            sql = "select * from PagoSeguroInfonavit where fkiIdEmpleadoC= " & idempleado
                            sql &= " And NumBimestre= " & numbimestre2 & " And Anio=" & FechaInicioPeriodo2.Year.ToString
                            Dim rwSeguro2 As DataRow() = nConsulta(sql)

                            If rwSeguro2 Is Nothing = False Then
                                Seguro2 = 0
                            Else
                                If cboTipoNomina.SelectedIndex = 0 And numbimestre <> numbimestre2 Then
                                    Seguro2 = 15
                                Else
                                    Seguro2 = 0
                                End If
                                'Seguro2 = 15
                            End If

                        End If


                        'Obtener el valor para VSM segun tabla
                        If FechaInicioPeriodo2 = Date.Parse("01/01/1900") Then

                        Else

                        End If


                        sql = "select * from Salario "
                        sql &= " where Anio=" & IIf(FechaFinPeriodo2 = Date.Parse("01/01/1900"), FechaFinPeriodo1.Year.ToString, FechaInicioPeriodo2.Year.ToString)
                        sql &= " and iEstatus=1"
                        Dim rwValorInfonavit As DataRow() = nConsulta(sql)

                        If rwValorInfonavit Is Nothing = False Then
                            ValorInfonavitTabla = rwValorInfonavit(0)("infonavit")
                        Else
                            sql = "select * from Salario "
                            sql &= " where Anio=" & IIf(FechaFinPeriodo2 = Date.Parse("01/01/1900"), FechaFinPeriodo1.Year.ToString, FechaInicioPeriodo2.Year.ToString)
                            sql &= " and iEstatus=1"
                            Dim rwValorInfonavitAntes As DataRow() = nConsulta(sql)
                            If rwValorInfonavitAntes Is Nothing = False Then
                                ValorInfonavitTabla = rwValorInfonavit(0)("infonavit")
                            End If
                        End If



                        If tipo = "VSM" And valor > 0 Then
                            valorinfonavit = (((ValorInfonavitTabla * valor * 2) / numdias) * DiasCadaPeriodo) + Seguro1
                            valorinfonavit = valorinfonavit + ((((ValorInfonavitTabla * valor * 2) / numdias2) * DiasCadaPeriodo2) + IIf(DiasCadaPeriodo2 = 0, 0, Seguro2))
                        End If

                        If tipo = "CUOTA FIJA" And valor > 0 Then


                            valorinfonavit = (((valor * 2) / numdias) * DiasCadaPeriodo) + Seguro1
                            valorinfonavit = valorinfonavit + ((((valor * 2) / numdias2) * DiasCadaPeriodo2) + IIf(DiasCadaPeriodo2 = 0, 0, Seguro2))

                        End If

                        If tipo = "PORCENTAJE" And valor > 0 Then

                            valorinfonavit = ((sdi * (valor / 100) * numdias) + 15) / numdias
                        End If


                        Return valorinfonavit

                    End If

                End If

            End If


            Return 0



        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 0
        End Try
    End Function
    Private Function CalcularInfonavit(tipo As String, valor As Double, sdi As Double, fechapago As Date, periodo As String, idempleado As Integer) As Boolean
        Try
            Dim numbimestre As Integer

            Dim numdias As Integer

            Dim DiasCadaPeriodo As Integer

            Dim diasfebrero As Integer
            Dim valorinfonavit As Double
            Dim sql As String
            Dim FechaInicioPeriodo1 As Date
            Dim FechaFinPeriodo1 As Date
            Dim FechaInicioPeriodo2 As Date
            Dim FechaFinPeriodo2 As Date

            Dim ValorInfonavitTabla As Double

            'Validamos si el trabajador tiene o no activo el infonavit
            sql = "select iPermanente from empleadosC where iIdEmpleadoC=" & idempleado
            Dim rwCalcularInfonavit As DataRow() = nConsulta(sql)
            If rwCalcularInfonavit Is Nothing = False Then
                If rwCalcularInfonavit(0)("iPermanente") = "1" Then
                    sql = "select * from periodos where iIdPeriodo= " & periodo
                    Dim rwPeriodo As DataRow() = nConsulta(sql)

                    If rwPeriodo Is Nothing = False Then
                        FechaInicioPeriodo1 = Date.Parse(rwPeriodo(0)("dFechaInicio"))




                        If Month(FechaInicioPeriodo1) Mod 2 = 0 Then
                            numbimestre = Month(FechaInicioPeriodo1) / 2
                        Else
                            numbimestre = (Month(FechaInicioPeriodo1) + 1) / 2
                        End If

                        If numbimestre = 1 Then
                            If Bisiesto(Year(FechaInicioPeriodo1)) = True Then
                                diasfebrero = 29
                            Else
                                diasfebrero = 28
                            End If
                            'diasfebrero = Day(DateSerial(Year(fechapago), 3, 0))
                            numdias = 31 + diasfebrero
                        End If

                        If numbimestre = 2 Then
                            numdias = 61
                        End If

                        If numbimestre = 3 Then
                            numdias = 61
                        End If

                        If numbimestre = 4 Then
                            numdias = 62
                        End If

                        If numbimestre = 5 Then
                            numdias = 61
                        End If

                        If numbimestre = 6 Then
                            numdias = 61
                        End If



                        sql = "select * from Salario "
                        sql &= " where Anio=" & IIf(FechaInicioPeriodo1 = Date.Parse("01/01/1900"), FechaInicioPeriodo1.Year.ToString, FechaInicioPeriodo1.Year.ToString)
                        sql &= " and iEstatus=1"
                        Dim rwValorInfonavit As DataRow() = nConsulta(sql)

                        If rwValorInfonavit Is Nothing = False Then
                            ValorInfonavitTabla = rwValorInfonavit(0)("infonavit")
                        Else

                        End If



                        If tipo = "VSM" And valor > 0 Then
                            valorinfonavit = (((ValorInfonavitTabla * valor * 2) / numdias) * numdias) + 15

                        End If

                        If tipo = "CUOTA FIJA" And valor > 0 Then


                            valorinfonavit = (((valor * 2) / numdias) * numdias) + 15


                        End If

                        If tipo = "PORCENTAJE" And valor > 0 Then

                            valorinfonavit = ((sdi * (valor / 100) * numdias) + 15)
                        End If


                        'Insertamos los datos

                        sql = "EXEC [setCalculoInfonavitInsertar  ] 0"
                        'Bimestre
                        sql &= "," & numbimestre
                        'Anio
                        sql &= "," & Year(FechaInicioPeriodo1)
                        'TipoFactor
                        sql &= ",'" & tipo
                        'Factor
                        sql &= "'," & valor
                        'idEmpleado
                        sql &= "," & idempleado
                        'Monto
                        sql &= "," & valorinfonavit
                        'Estatus
                        sql &= ",1"






                        If nExecute(sql) = False Then
                            MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            Return False

                        End If

                        Return True
                    End If

                End If

            End If


            Return False



        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try
    End Function

    Private Function VerificarCalculoInfonavit(periodo As String, idempleado As Integer) As Integer

        Try
            Dim numbimestre As Integer

            Dim numdias As Integer

            Dim diasfebrero As Integer

            Dim sql As String
            Dim FechaInicioPeriodo1 As Date


            'Validamos si el trabajador tiene o no activo el infonavit
            sql = "select iPermanente from empleadosC where iIdEmpleadoC=" & idempleado
            Dim rwCalcularInfonavit As DataRow() = nConsulta(sql)
            If rwCalcularInfonavit Is Nothing = False Then
                If rwCalcularInfonavit(0)("iPermanente") = "1" Then
                    sql = "select * from periodos where iIdPeriodo= " & periodo
                    Dim rwPeriodo As DataRow() = nConsulta(sql)

                    If rwPeriodo Is Nothing = False Then
                        FechaInicioPeriodo1 = Date.Parse(rwPeriodo(0)("dFechaInicio"))

                        If Month(FechaInicioPeriodo1) Mod 2 = 0 Then
                            numbimestre = Month(FechaInicioPeriodo1) / 2
                        Else
                            numbimestre = (Month(FechaInicioPeriodo1) + 1) / 2
                        End If

                        If numbimestre = 1 Then
                            If Bisiesto(Year(FechaInicioPeriodo1)) = True Then
                                diasfebrero = 29
                            Else
                                diasfebrero = 28
                            End If
                            'diasfebrero = Day(DateSerial(Year(fechapago), 3, 0))
                            numdias = 31 + diasfebrero
                        End If

                        If numbimestre = 2 Then
                            numdias = 61
                        End If

                        If numbimestre = 3 Then
                            numdias = 61
                        End If

                        If numbimestre = 4 Then
                            numdias = 62
                        End If

                        If numbimestre = 5 Then
                            numdias = 61
                        End If

                        If numbimestre = 6 Then
                            numdias = 61
                        End If





                        'Realizamos la busqueda

                        sql = "select * from CalculoInfonavit where iBimestre=" & numbimestre
                        sql &= " And iAnio= " & Year(FechaInicioPeriodo1) & " And fkiIdEmpleadoC=" & idempleado
                        Dim rwCalculoInfonavit As DataRow() = nConsulta(sql)
                        If rwCalculoInfonavit Is Nothing = False Then
                            Return 1
                        Else
                            Return 2
                        End If

                    Else
                        Return 0
                    End If
                Else
                    Return 0
                End If
            Else
                Return 0
            End If


            Return 0



        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 0
        End Try
    End Function

    Private Function MontoInfonavitF(periodo As String, idempleado As Integer) As Double

        Try
            Dim numbimestre As Integer
            Dim sql As String
            Dim FechaInicioPeriodo1 As Date


            'Validamos si el trabajador tiene o no activo el infonavit
            sql = "select iPermanente from empleadosC where iIdEmpleadoC=" & idempleado
            Dim rwCalcularInfonavit As DataRow() = nConsulta(sql)
            If rwCalcularInfonavit Is Nothing = False Then
                If rwCalcularInfonavit(0)("iPermanente") = "1" Then
                    sql = "select * from periodos where iIdPeriodo= " & periodo
                    Dim rwPeriodo As DataRow() = nConsulta(sql)

                    If rwPeriodo Is Nothing = False Then
                        FechaInicioPeriodo1 = Date.Parse(rwPeriodo(0)("dFechaInicio"))

                        If Month(FechaInicioPeriodo1) Mod 2 = 0 Then
                            numbimestre = Month(FechaInicioPeriodo1) / 2
                        Else
                            numbimestre = (Month(FechaInicioPeriodo1) + 1) / 2
                        End If


                        'Realizamos la busqueda

                        sql = "select * from CalculoInfonavit where iBimestre=" & numbimestre
                        sql &= " And iAnio= " & Year(FechaInicioPeriodo1) & " And fkiIdEmpleadoC=" & idempleado
                        Dim rwCalculoInfonavit As DataRow() = nConsulta(sql)
                        If rwCalculoInfonavit Is Nothing = False Then
                            Return Double.Parse(rwCalculoInfonavit(0)("Monto"))
                            IDCalculoInfonavit = rwCalculoInfonavit(0)("iIdCalculoInfonavit")
                        Else
                            Return 0
                        End If

                    Else
                        Return 0
                    End If
                Else
                    Return 0
                End If
            Else
                Return 0
            End If


            Return 0



        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 0
        End Try
    End Function

    Private Function baseisrtotal(puesto As String, dias As Integer, sdi As Double, incapacidad As Double) As Double
        Dim sueldo As Double
        Dim sueldobase As Double
        Dim baseisr As Double
        Dim isrcalculado As Double
        Dim aguinaldog As Double
        Dim primag As Double
        Dim sql As String
        Dim ValorUMA As Double
        Try

            Sql = "select * from Salario "
            Sql &= " where Anio=" & aniocostosocial
            Sql &= " and iEstatus=1"
            Dim rwValorUMA As DataRow() = nConsulta(Sql)
            If rwValorUMA Is Nothing = False Then
                ValorUMA = Double.Parse(rwValorUMA(0)("uma").ToString)
            Else
                ValorUMA = 0
                MessageBox.Show("No se encontro valor para UMA en el año: " & aniocostosocial)
            End If

            If puesto = "OFICIALES EN PRACTICAS: PILOTIN / ASPIRANTE" Then
                sueldo = sdi * dias
                sueldobase = sueldo
                baseisr = sueldobase - incapacidad
                isrcalculado = isrmensual(baseisr)
            Else
                sueldo = sdi * dias
                sueldobase = (sueldo * (26.19568006 / 100)) + ((sueldo * (8.5070471 / 100)) / 2) + ((sueldo * (8.5070471 / 100)) / 2) + (sueldo * (42.89215164 / 100)) + (sueldo * (9.677848468 / 100))

                ''Aguinaldo gravado
                'aguinaldog = Math.Round(((sueldobase / dias) * 15 / 12 * (dias / 30)) - ((ValorUMA * 30 / 12) * (dias / 30)), 2)


                'primag = (sueldobase * 0.25 / 12 * (dias / 30)) - ((ValorUMA * 15 / 12) * (dias / 30))


                'Aguinaldo gravado 

                If ((sueldobase / dias) * 15 / 12 * (dias / 30)) > ((ValorUMA * 30 / 12) * (dias / 30)) Then
                    'Aguinaldo gravado
                    aguinaldog = Math.Round(((sueldobase / dias) * 15 / 12 * (dias / 30)) - ((ValorUMA * 30 / 12) * (dias / 30)), 2)
                Else
                    'Aguinaldo gravado
                    aguinaldog = "0.00"
                End If

                'Prima de vacaciones

                'Calculos prima
                Dim primavacacionesgravada As Double
                Dim primavacacionesexenta As Double

                primavacacionesgravada = (sueldobase * 0.25 / 12 * (dias / 30)) - ((ValorUMA * 15 / 12) * (dias / 30))
                primavacacionesexenta = ((ValorUMA * 15 / 12) * (dias / 30))

                If primavacacionesgravada > 0 Then
                    primag = primavacacionesgravada

                Else
                    primag = 0
                End If


                baseisr = (sueldobase - ((sueldo * (8.5070471 / 100)) / 2)) + (sueldo * (7.272727273 / 100)) + aguinaldog + primag - incapacidad
                isrcalculado = isrmensual(baseisr)

            End If
            Return isrcalculado
        Catch ex As Exception

        End Try
    End Function


    Private Function isrmensual(monto As Double) As Double

        Dim excendente As Double
        Dim isr As Double
        Dim subsidio As Double



        Dim SQL As String

        Try


            'calculos

            'Calculamos isr

            '1.- buscamos datos para el calculo
            isr = 0
            SQL = "select * from isr where ((" & monto & ">=isr.limiteinf and " & monto & "<=isr.limitesup)"
            SQL &= " or (" & monto & ">=isr.limiteinf and isr.limitesup=0)) and fkiIdTipoPeriodo2=1"


            Dim rwISRCALCULO As DataRow() = nConsulta(SQL)
            If rwISRCALCULO Is Nothing = False Then
                excendente = monto - Double.Parse(rwISRCALCULO(0)("limiteinf").ToString)
                isr = (excendente * (Double.Parse(rwISRCALCULO(0)("porcentaje").ToString) / 100)) + Double.Parse(rwISRCALCULO(0)("cuotafija").ToString)

            End If
            subsidio = 0
            SQL = "select * from subsidio where ((" & monto & ">=subsidio.limiteinf and " & monto & "<=subsidio.limitesup)"
            SQL &= " or (" & monto & ">=subsidio.limiteinf and subsidio.limitesup=0)) and fkiIdTipoPeriodo2=1"


            Dim rwSubsidio As DataRow() = nConsulta(SQL)
            If rwSubsidio Is Nothing = False Then
                subsidio = Double.Parse(rwSubsidio(0)("credito").ToString)

            End If
            If isr > subsidio Then
                Return isr - subsidio
            Else
                Return 0
            End If


        Catch ex As Exception

        End Try
    End Function

    Function subsidiomensual(monto As Double) As Double
        Dim excendente As Double
        Dim isr As Double
        Dim subsidio As Double



        Dim SQL As String

        Try


            'calculos

            'Calculamos isr

            '1.- buscamos datos para el calculo
            isr = 0
            SQL = "select * from isr where ((" & monto & ">=isr.limiteinf and " & monto & "<=isr.limitesup)"
            SQL &= " or (" & monto & ">=isr.limiteinf and isr.limitesup=0)) and fkiIdTipoPeriodo2=1"


            Dim rwISRCALCULO As DataRow() = nConsulta(SQL)
            If rwISRCALCULO Is Nothing = False Then
                excendente = monto - Double.Parse(rwISRCALCULO(0)("limiteinf").ToString)
                isr = (excendente * (Double.Parse(rwISRCALCULO(0)("porcentaje").ToString) / 100)) + Double.Parse(rwISRCALCULO(0)("cuotafija").ToString)

            End If
            subsidio = 0
            SQL = "select * from subsidio where ((" & monto & ">=subsidio.limiteinf and " & monto & "<=subsidio.limitesup)"
            SQL &= " or (" & monto & ">=subsidio.limiteinf and subsidio.limitesup=0)) and fkiIdTipoPeriodo2=1"


            Dim rwSubsidio As DataRow() = nConsulta(SQL)
            If rwSubsidio Is Nothing = False Then
                subsidio = Double.Parse(rwSubsidio(0)("credito").ToString)

            End If

            If isr >= subsidio Then
                subsidiomensual = 0
            Else
                subsidiomensual = subsidio - isr
            End If


        Catch ex As Exception

        End Try



    End Function

    Private Function baseSubsidiototal(puesto As String, dias As Double, sdi As Double, incapacidad As Double) As Double



        Dim sueldo As Double
        Dim sueldobase As Double
        Dim baseisr As Double
        Dim isrcalculado As Double
        Dim aguinaldog As Double
        Dim primag As Double
        Dim sql As String
        Dim ValorUMA As Double
        Try

            sql = "select * from Salario "
            sql &= " where Anio=" & aniocostosocial
            sql &= " and iEstatus=1"
            Dim rwValorUMA As DataRow() = nConsulta(sql)
            If rwValorUMA Is Nothing = False Then
                ValorUMA = Double.Parse(rwValorUMA(0)("uma").ToString)
            Else
                ValorUMA = 0
                MessageBox.Show("No se encontro valor para UMA en el año: " & aniocostosocial)
            End If

            If puesto = "OFICIALES EN PRACTICAS: PILOTIN / ASPIRANTE" Then
                sueldo = sdi * dias
                sueldobase = sueldo
                baseisr = sueldobase - incapacidad
                baseSubsidiototal = subsidiomensual(baseisr)
            Else
                sueldo = sdi * dias
                sueldobase = (sueldo * (26.19568006 / 100)) + ((sueldo * (8.5070471 / 100)) / 2) + ((sueldo * (8.5070471 / 100)) / 2) + (sueldo * (42.89215164 / 100)) + (sueldo * (9.677848468 / 100))

                'Aguinaldo gravado 

                If ((sueldobase / dias) * 15 / 12 * (dias / 30)) > ((ValorUMA * 30 / 12) * (dias / 30)) Then
                    'Aguinaldo gravado
                    aguinaldog = Math.Round(((sueldobase / dias) * 15 / 12 * (dias / 30)) - ((ValorUMA * 30 / 12) * (dias / 30)), 2)
                Else
                    'Aguinaldo gravado
                    aguinaldog = "0.00"
                End If

                'Prima de vacaciones

                'Calculos prima
                Dim primavacacionesgravada As Double
                Dim primavacacionesexenta As Double

                primavacacionesgravada = (sueldobase * 0.25 / 12 * (dias / 30)) - ((ValorUMA * 15 / 12) * (dias / 30))
                primavacacionesexenta = ((ValorUMA * 15 / 12) * (dias / 30))

                If primavacacionesgravada > 0 Then
                    primag = primavacacionesgravada

                Else
                    primag = 0
                End If


                baseisr = (sueldobase - ((sueldo * (8.5070471 / 100)) / 2)) + (sueldo * (7.272727273 / 100)) + aguinaldog + primag - incapacidad
                baseSubsidiototal = subsidiomensual(baseisr)

            End If
            Return baseSubsidiototal
        Catch ex As Exception

        End Try



    End Function


    Function subsidiomensualCausado(monto As Double) As Double
        Dim excendente As Double
        Dim isr As Double
        Dim subsidio As Double



        Dim SQL As String

        Try


            'calculos

            'Calculamos isr

            '1.- buscamos datos para el calculo
            isr = 0
            SQL = "select * from isr where ((" & monto & ">=isr.limiteinf and " & monto & "<=isr.limitesup)"
            SQL &= " or (" & monto & ">=isr.limiteinf and isr.limitesup=0)) and fkiIdTipoPeriodo2=1"


            Dim rwISRCALCULO As DataRow() = nConsulta(SQL)
            If rwISRCALCULO Is Nothing = False Then
                excendente = monto - Double.Parse(rwISRCALCULO(0)("limiteinf").ToString)
                isr = (excendente * (Double.Parse(rwISRCALCULO(0)("porcentaje").ToString) / 100)) + Double.Parse(rwISRCALCULO(0)("cuotafija").ToString)

            End If
            subsidio = 0
            SQL = "select * from subsidio where ((" & monto & ">=subsidio.limiteinf and " & monto & "<=subsidio.limitesup)"
            SQL &= " or (" & monto & ">=subsidio.limiteinf and subsidio.limitesup=0)) and fkiIdTipoPeriodo2=1"


            Dim rwSubsidio As DataRow() = nConsulta(SQL)
            If rwSubsidio Is Nothing = False Then
                subsidio = Double.Parse(rwSubsidio(0)("credito").ToString)

            End If

            If isr >= subsidio Then
                subsidiomensualCausado = 0
            Else
                subsidiomensualCausado = subsidio
            End If


        Catch ex As Exception

        End Try



    End Function


    Function baseSubsidio(puesto As String, dias As Double, sdi As Double, incapacidad As Double) As Double
        Dim sueldo As Double
        Dim sueldobase As Double
        Dim baseisr As Double
        Dim isrcalculado As Double
        Dim aguinaldog As Double
        Dim primag As Double
        Dim sql As String
        Dim ValorUMA As Double
        Try

            sql = "select * from Salario "
            sql &= " where Anio=" & aniocostosocial
            sql &= " and iEstatus=1"
            Dim rwValorUMA As DataRow() = nConsulta(sql)
            If rwValorUMA Is Nothing = False Then
                ValorUMA = Double.Parse(rwValorUMA(0)("uma").ToString)
            Else
                ValorUMA = 0
                MessageBox.Show("No se encontro valor para UMA en el año: " & aniocostosocial)
            End If

            If puesto = "OFICIALES EN PRACTICAS: PILOTIN / ASPIRANTE" Then
                sueldo = sdi * dias
                sueldobase = sueldo
                baseisr = sueldobase - incapacidad
                baseSubsidio = subsidiomensualCausado(baseisr)
            Else
                sueldo = sdi * dias
                sueldobase = (sueldo * (26.19568006 / 100)) + ((sueldo * (8.5070471 / 100)) / 2) + ((sueldo * (8.5070471 / 100)) / 2) + (sueldo * (42.89215164 / 100)) + (sueldo * (9.677848468 / 100))

                'Aguinaldo gravado 

                If ((sueldobase / dias) * 15 / 12 * (dias / 30)) > ((ValorUMA * 30 / 12) * (dias / 30)) Then
                    'Aguinaldo gravado
                    aguinaldog = Math.Round(((sueldobase / dias) * 15 / 12 * (dias / 30)) - ((ValorUMA * 30 / 12) * (dias / 30)), 2)
                Else
                    'Aguinaldo gravado
                    aguinaldog = "0.00"
                End If

                'Prima de vacaciones

                'Calculos prima
                Dim primavacacionesgravada As Double
                Dim primavacacionesexenta As Double

                primavacacionesgravada = (sueldobase * 0.25 / 12 * (dias / 30)) - ((ValorUMA * 15 / 12) * (dias / 30))
                primavacacionesexenta = ((ValorUMA * 15 / 12) * (dias / 30))

                If primavacacionesgravada > 0 Then
                    primag = primavacacionesgravada

                Else
                    primag = 0
                End If

                baseisr = (sueldobase - ((sueldo * (8.5070471 / 100)) / 2)) + (sueldo * (7.272727273 / 100)) + aguinaldog + primag - incapacidad
                baseSubsidio = subsidiomensualCausado(baseisr)

            End If
            Return baseSubsidio
        Catch ex As Exception

        End Try



    End Function


    Private Function Incapacidades(tipo As String, valor As Double, sd As Double) As Double
        Dim incapacidad As Double
        incapacidad = 0.0
        Try
            If tipo = "Riesgo de trabajo" Then
                Incapacidades = 0
            ElseIf tipo = "Enfermedad general" Then
                Incapacidades = valor * sd
            ElseIf tipo = "Maternidad" Then
                Incapacidades = 0
            End If
            Return incapacidad
        Catch ex As Exception

        End Try
    End Function



    Private Sub cmdguardarfinal_Click(sender As Object, e As EventArgs) Handles cmdguardarfinal.Click
        Try
            Dim sql As String
            Dim sql2 As String
            sql = "select * from Nomina where fkiIdEmpresa=1 and fkiIdPeriodo=" & cboperiodo.SelectedValue
            sql &= " and iEstatusNomina=1 and iEstatus=1 and iEstatusEmpleado=" & cboserie.SelectedIndex
            sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex
            'Dim sueldobase, salariodiario, salariointegrado, sueldobruto, TiempoExtraFijoGravado, TiempoExtraFijoExento As Double
            'Dim TiempoExtraOcasional, DesSemObligatorio, VacacionesProporcionales, AguinaldoGravado, AguinaldoExento As Double
            'Dim PrimaVacGravada, PrimaVacExenta, TotalPercepciones, TotalPercepcionesISR As Double
            'Dim incapacidad, ISR, IMSS, Infonavit, InfonavitAnterior, InfonavitAjuste, PensionAlimenticia As Double
            'Dim Prestamo, Fonacot, NetoaPagar, Excedente, Total, ImssCS, RCVCS, InfonavitCS, ISNCS
            'sql = "EXEC getNominaXEmpresaXPeriodo " & gIdEmpresa & "," & cboperiodo.SelectedValue & ",1"

            Dim rwNominaGuardadaFinal As DataRow() = nConsulta(sql)

            If rwNominaGuardadaFinal Is Nothing = False Then
                MessageBox.Show("La nomina ya esta marcada como final, no  se pueden guardar cambios", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                'MessageBox.Show("Se borraran los datos tanto de la nomina abordo como la de descanso", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

                sql = "update Nomina set iEstatusNomina=1 "
                sql &= " where fkiIdEmpresa=1 and fkiIdPeriodo=" & cboperiodo.SelectedValue
                sql &= " and iEstatusNomina=0 and iEstatus=1 and iEstatusEmpleado=" & cboserie.SelectedIndex
                'sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex
                If nExecute(sql) = False Then
                    MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    'pnlProgreso.Visible = False
                    Exit Sub
                End If


                pnlProgreso.Visible = True

                Application.DoEvents()
                pnlCatalogo.Enabled = False
                pgbProgreso.Minimum = 0
                pgbProgreso.Value = 0
                pgbProgreso.Maximum = dtgDatos.Rows.Count


                For x As Integer = 0 To dtgDatos.Rows.Count - 1
                    '########GUARDAR INFONAVIT

                    Dim numbimestre As Integer
                    If Month(FechaInicioPeriodoGlobal) Mod 2 = 0 Then
                        numbimestre = Month(FechaInicioPeriodoGlobal) / 2
                    Else
                        numbimestre = (Month(FechaInicioPeriodoGlobal) + 1) / 2
                    End If

                    '########GUARDAR SEGURO INFONAVIT
                    sql = "select * from periodos where iIdPeriodo= " & cboperiodo.SelectedValue
                    Dim rwPeriodo As DataRow() = nConsulta(sql)

                    Dim FechaInicioPeriodo1 As Date


                    'Dim numbimestre As Integer
                    If rwPeriodo Is Nothing = False Then
                        FechaInicioPeriodo1 = Date.Parse(rwPeriodo(0)("dFechaInicio"))

                        If Month(FechaInicioPeriodo1) Mod 2 = 0 Then
                            numbimestre = Month(FechaInicioPeriodo1) / 2
                        Else
                            numbimestre = (Month(FechaInicioPeriodo1) + 1) / 2
                        End If

                    End If

                    If Double.Parse(dtgDatos.Rows(x).Cells(38).Value) > 0 Then

                        sql = "select * from PagoSeguroInfonavit where fkiIdEmpleadoC= " & dtgDatos.Rows(x).Cells(2).Value
                        sql &= " And NumBimestre= " & numbimestre & " And Anio=" & FechaInicioPeriodo1.Year.ToString
                        Dim rwSeguro1 As DataRow() = nConsulta(sql)

                        If rwSeguro1 Is Nothing = True Then
                            'Insertar seguro
                            sql = "EXEC setPagoSeguroInfonavitInsertar  0"
                            ' fk Empleado
                            sql &= "," & dtgDatos.Rows(x).Cells(2).Value
                            'bimestre
                            sql &= "," & numbimestre
                            ' anio
                            sql &= "," & FechaInicioPeriodo1.Year.ToString


                            If nExecute(sql) = False Then
                                MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                'pnlProgreso.Visible = False
                                Exit Sub
                            End If
                        End If

                    End If




                    pgbProgreso.Value += 1
                    Application.DoEvents()
                Next
                pnlProgreso.Visible = False
                pnlCatalogo.Enabled = True
                MessageBox.Show("Datos guardados correctamente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    Private Sub cboperiodo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboperiodo.SelectedIndexChanged
        Try
            dtgDatos.DataSource = ""
            dtgDatos.Columns.Clear()
            Dim Sql As String = "select * from periodos where iIdPeriodo= " & cboperiodo.SelectedValue
            Dim rwPeriodo As DataRow() = nConsulta(Sql)

            If rwPeriodo Is Nothing = False Then
                FechaInicioPeriodoGlobal = Date.Parse(rwPeriodo(0)("dFechaInicio"))
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub cmdreciboss_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub dtgDatos_CellMouseDown(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dtgDatos.CellMouseDown
        Try
            If e.RowIndex > -1 And e.ColumnIndex > -1 Then
                dtgDatos.CurrentCell = dtgDatos.Rows(e.RowIndex).Cells(e.ColumnIndex)


            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub dtgDatos_CellMouseUp(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dtgDatos.CellMouseUp

    End Sub



    Private Sub dtgDatos_KeyPress(sender As Object, e As KeyPressEventArgs) Handles dtgDatos.KeyPress
        Try

            SoloNumero.NumeroDec(e, sender)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btnReporte_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReporte.Click
        Try


            Dim filaExcel As Integer = 0
            Dim dialogo As New SaveFileDialog()
            Dim periodo, fechadepago As String
            Dim mes, anio As String
            Dim fechapagoletra() As String

            Dim sueldobase, tef, teo, desc, vacpro, taginaldo, tpvac, tpercepciones, comAsim, prestamo, comXurtep, comComple, costoSocial As Double
            Dim IMSS, SAR, INFONAVIT, ISN As Double

            Dim sueldobaseD, tefD, teoD, descD, vacproD, taginaldoD, tpvacD, tpercepcionesD, comAsimD, prestamoD, comXurtepD, comCompleD, costoSocialD As Double
            Dim IMSSD, SARD, INFONAVITD, ISND As Double

            If dtgDatos.Rows.Count > 0 Then

                'Clonar el XLSX
                Dim ruta As String
                ruta = My.Application.Info.DirectoryPath() & "\Archivos\contadormarinos.xlsx"

                Dim book As New ClosedXML.Excel.XLWorkbook(ruta)
                Dim libro As New ClosedXML.Excel.XLWorkbook

                book.Worksheet(1).CopyTo(libro, "ECO III")
                Dim hoja As IXLWorksheet = libro.Worksheets(0)

                '<<<<<<<<<<<ECO III>>>>>>>>>>>
                Dim rwPeriodo0 As DataRow() = nConsulta("Select * from periodos where iIdPeriodo=" & cboperiodo.SelectedValue)
                If rwPeriodo0 Is Nothing = False Then
                    periodo = MonthString(rwPeriodo0(0).Item("iMes")).ToUpper & " " & (rwPeriodo0(0).Item("iEjercicio"))
                    'fecha = MonthString(rwPeriodo0(0).Item("iMes")).ToUpper
                    'hoja.Cell(10, 2).Style.Font.SetBold(True)
                    'hoja.Cell(10, 2).Style.NumberFormat.Format = "@"
                    'hoja.Cell(10, 2).Value = periodo

                End If

                'Se valida si es Abordo
                If cboTipoNomina.SelectedIndex = 1 Then
                    llenargridD("0")
                End If

                'Sumar Abordo
                For x As Integer = 0 To dtgDatos.Rows.Count - 1
                    sueldobase += dtgDatos.Rows(x).Cells(21).Value
                    tef += (CDbl(dtgDatos.Rows(x).Cells(22).Value) + CDbl(dtgDatos.Rows(x).Cells(23).Value))
                    teo += dtgDatos.Rows(x).Cells(24).Value
                    desc += dtgDatos.Rows(x).Cells(25).Value
                    vacpro += dtgDatos.Rows(x).Cells(26).Value
                    taginaldo += dtgDatos.Rows(x).Cells(29).Value
                    tpvac += dtgDatos.Rows(x).Cells(32).Value
                    tpercepciones += dtgDatos.Rows(x).Cells(33).Value
                    comAsim += IIf(dtgDatos.Rows(x).Cells(11).FormattedValue = "OFICIALES EN PRACTICAS: PILOTIN / ASPIRANTE", 0, CDbl(dtgDatos.Rows(x).Cells(50).Value) * 2)
                    prestamo += dtgDatos.Rows(x).Cells(47).Value
                    comXurtep += IIf(dtgDatos.Rows(x).Cells(11).FormattedValue = "OFICIALES EN PRACTICAS: PILOTIN / ASPIRANTE", CDbl(dtgDatos.Rows(x).Cells(53).Value), CDbl(dtgDatos.Rows(x).Cells(53).Value) * 2)
                    comComple += IIf(dtgDatos.Rows(x).Cells(11).FormattedValue = "OFICIALES EN PRACTICAS: PILOTIN / ASPIRANTE", 0, CDbl(dtgDatos.Rows(x).Cells(54).Value) * 2)
                    costoSocial += dtgDatos.Rows(x).Cells(59).Value

                    IMSS += dtgDatos.Rows(x).Cells(55).Value
                    SAR += dtgDatos.Rows(x).Cells(56).Value
                    INFONAVIT += dtgDatos.Rows(x).Cells(57).Value
                    ISN += dtgDatos.Rows(x).Cells(58).Value

                Next x

                'Sumar Descanso
                llenargridD("1")

                For x As Integer = 0 To dtgDatos.Rows.Count - 1
                    sueldobaseD += dtgDatos.Rows(x).Cells(21).Value
                    tefD += (CDbl(dtgDatos.Rows(x).Cells(22).Value) + CDbl(dtgDatos.Rows(x).Cells(23).Value))
                    teoD += dtgDatos.Rows(x).Cells(24).Value
                    descD += dtgDatos.Rows(x).Cells(25).Value
                    vacproD += dtgDatos.Rows(x).Cells(26).Value
                    taginaldoD += dtgDatos.Rows(x).Cells(29).Value
                    tpvacD += dtgDatos.Rows(x).Cells(32).Value
                    tpercepcionesD += dtgDatos.Rows(x).Cells(33).Value
                    comAsimD += dtgDatos.Rows(x).Cells(50).Value
                    prestamoD += dtgDatos.Rows(x).Cells(47).Value
                    comXurtepD += (dtgDatos.Rows(x).Cells(53).Value) ' * 2%)
                    comCompleD += (dtgDatos.Rows(x).Cells(54).Value) '* 2%)
                    costoSocialD += dtgDatos.Rows(x).Cells(59).Value

                    IMSSD += dtgDatos.Rows(x).Cells(55).Value
                    SARD += dtgDatos.Rows(x).Cells(56).Value
                    INFONAVITD += dtgDatos.Rows(x).Cells(57).Value
                    ISND += dtgDatos.Rows(x).Cells(58).Value

                Next x

                'Agregar los valores
                hoja.Cell("B5").Value = sueldobase + sueldobaseD
                hoja.Cell("C5").Value = tef + tefD
                hoja.Cell("D5").Value = teo + teoD
                hoja.Cell("E5").Value = desc + descD
                hoja.Cell("F5").Value = vacpro + vacproD
                hoja.Cell("G5").Value = taginaldo + taginaldoD
                hoja.Cell("H5").Value = tpvac + tpvacD
                hoja.Cell("I5").Value = tpercepciones + tpercepcionesD
                hoja.Cell("J5").Value = comAsim  '+ comXurtepD
                hoja.Cell("K5").Value = "0.0" 'prestamo
                hoja.Cell("L5").Value = (comXurtep) '* 2%
                hoja.Cell("M5").Value = (comComple) '+ comCompleD) * 2%
                hoja.Cell("N5").Value = costoSocial '+ costoSocialD

                hoja.Cell("E19").Value = IMSS ' + IMSSD
                hoja.Cell("E20").Value = SAR '+ SARD
                hoja.Cell("E21").Value = INFONAVIT ' + INFONAVITD
                hoja.Cell("E22").Value = ISN '+ ISND

                'Regresa a Abordo
                llenargridD("0")

                'Guardar Excel
                'Titulo
                Dim moment As Date = Date.Now()
                Dim month As Integer = moment.Month
                Dim year As Integer = moment.Year


                dialogo.FileName = "REPORTE XURTEP MARINOS" & periodo & " SERIE " & cboserie.SelectedItem
                dialogo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
               
                If dialogo.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

                    libro.SaveAs(dialogo.FileName)
                    libro = Nothing

                    MessageBox.Show("Archivo generado correctamente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("No se guardo el archivo", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                End If
            End If
        Catch ex As Exception

            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Try


    End Sub



    Private Sub tsbImportar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles tsbImportar.Click

    End Sub

    Private Sub cmdincidencias_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdincidencias.Click

    End Sub

    Private Sub cmdreiniciar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdreiniciar.Click
        Try
            Dim sql As String
            Dim resultado As Integer = MessageBox.Show("Se borraran los datos tanto de la nomina Abordo como Descanso,¿Desea reiniciar la nomina?", "Pregunta", MessageBoxButtons.YesNo)
            If resultado = DialogResult.Yes Then

                sql = "select * from Nomina where fkiIdEmpresa=1 and fkiIdPeriodo=" & cboperiodo.SelectedValue
                sql &= " and iEstatusNomina=1 and iEstatus=1 and iEstatusEmpleado=" & cboserie.SelectedIndex
                sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex

                Dim rwNominaGuardadaFinal As DataRow() = nConsulta(sql)



                If rwNominaGuardadaFinal Is Nothing = False Then
                    MessageBox.Show("La nomina ya esta marcada como final, no  se pueden guardar cambios.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    MessageBox.Show("Se borraran los datos tanto de la nomina abordo como la de descanso", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

                    sql = "delete from Nomina"
                    sql &= " where fkiIdEmpresa=1 and fkiIdPeriodo=" & cboperiodo.SelectedValue
                    sql &= " and iEstatusNomina=0 and iEstatus=1 and iEstatusEmpleado=" & cboserie.SelectedIndex
                    'sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex

                    If nExecute(sql) = False Then
                        MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        'pnlProgreso.Visible = False
                        Exit Sub
                    End If

                    'borrar el detalle del infonavit


                    sql = "delete from DetalleDescInfonavit"
                    sql &= " where fkiIdPeriodo=" & cboperiodo.SelectedValue
                    sql &= " and iSerie=" & cboserie.SelectedIndex
                    'sql &= " and iSerie=" & cboserie.SelectedIndex
                    'sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex

                    If nExecute(sql) = False Then
                        MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        'pnlProgreso.Visible = False
                        Exit Sub
                    End If

                    MessageBox.Show("Nomina reiniciada correctamente, vuelva a cargar los datos", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    dtgDatos.DataSource = ""
                    dtgDatos.Columns.Clear()
                End If



            End If




        Catch ex As Exception

        End Try


    End Sub

    Private Sub tsbIEmpleados_Click(ByVal sender As Object, ByVal e As EventArgs) Handles tsbIEmpleados.Click
        Try
            Dim Forma As New frmEmpleados
            Forma.gIdEmpresa = gIdEmpresa
            Forma.gIdPeriodo = cboperiodo.SelectedValue
            Forma.gIdTipoPuesto = 1
            Forma.ShowDialog()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub dtgDatos_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles dtgDatos.CellClick
        Try
            If e.ColumnIndex = 0 Then
                dtgDatos.Rows(e.RowIndex).Cells(0).Value = Not dtgDatos.Rows(e.RowIndex).Cells(0).Value
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    Private Sub dtgDatos_CellEnter(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles dtgDatos.CellEnter
        'MessageBox.Show("Ocurrio un error ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    End Sub

    Private Sub TextboxNumeric_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        Try
            'Dim columna As Integer
            'Dim fila As Integer

            'columna = CInt(DirectCast(sender, System.Windows.Forms.DataGridView).CurrentCell.ColumnIndex)
            'Fila = CInt(DirectCast(sender, System.Windows.Forms.DataGridView).CurrentCell.RowIndex)


            Dim nonNumberEntered As Boolean

            nonNumberEntered = True

            If (Convert.ToInt32(e.KeyChar) >= 48 AndAlso Convert.ToInt32(e.KeyChar) <= 57) OrElse Convert.ToInt32(e.KeyChar) = 8 OrElse Convert.ToInt32(e.KeyChar) = 46 Then

                'If Convert.ToInt32(e.KeyChar) = 46 Then
                '    If InStr(dtgDatos.Rows(Fila).Cells(columna).Value, ".") = 0 Then
                '        nonNumberEntered = False
                '    Else
                '        nonNumberEntered = False
                '    End If
                'Else
                '    nonNumberEntered = False
                'End If
                nonNumberEntered = False
            End If

            If nonNumberEntered = True Then
                ' Stop the character from being entered into the control since it is non-numerical.
                e.Handled = True
            Else
                e.Handled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try



    End Sub

    Private Sub dtgDatos_CellEndEdit(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles dtgDatos.CellEndEdit
        Try
            If Not m_currentControl Is Nothing Then
                RemoveHandler m_currentControl.KeyPress, AddressOf TextboxNumeric_KeyPress
            End If
            If Not dgvCombo Is Nothing Then
                RemoveHandler dgvCombo.SelectedIndexChanged, AddressOf dvgCombo_SelectedIndexChanged
            End If
            If dgvCombo IsNot Nothing Then
                RemoveHandler dgvCombo.SelectedIndexChanged, New EventHandler(AddressOf dvgCombo_SelectedIndexChanged)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    Private Sub dtgDatos_EditingControlShowing(ByVal sender As Object, ByVal e As DataGridViewEditingControlShowingEventArgs) Handles dtgDatos.EditingControlShowing
        Try
            Dim columna As Integer
            m_currentControl = Nothing
            columna = CInt(DirectCast(sender, System.Windows.Forms.DataGridView).CurrentCell.ColumnIndex)
            If columna = 15 Or columna = 18 Or columna = 39 Or columna = 40 Or columna = 41 Or columna = 42 Or columna = 43 Or columna = 10 Then
                AddHandler e.Control.KeyPress, AddressOf TextboxNumeric_KeyPress
                m_currentControl = e.Control
            End If


            dgvCombo = TryCast(e.Control, DataGridViewComboBoxEditingControl)

            If dgvCombo IsNot Nothing Then
                AddHandler dgvCombo.SelectedIndexChanged, New EventHandler(AddressOf dvgCombo_SelectedIndexChanged)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try



    End Sub

    Private Sub dtgDatos_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) Handles dtgDatos.ColumnHeaderMouseClick
        Try
            Dim newColumn As DataGridViewColumn = dtgDatos.Columns(e.ColumnIndex)

            Dim sql As String
            If e.ColumnIndex = 0 Then
                dtgDatos.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
            Else
                If e.ColumnIndex = 11 Then
                    'DirectCast(dtgDatos.Columns(11), DataGridViewComboBoxColumn).Sorted = True
                    Dim resultado As Integer = MessageBox.Show("Para realizar este ordenamiento es necesario guardar la nomina primeramente, ¿desea continuar?", "Pregunta", MessageBoxButtons.YesNo)
                    If resultado = DialogResult.Yes Then

                        cmdguardarnomina_Click(sender, e)
                        campoordenamiento = "nomina.Puesto,cNombreLargo"
                        llenargrid()
                    End If

                End If

                If e.ColumnIndex = 12 Then
                    Dim resultado As Integer = MessageBox.Show("Para realizar este ordenamiento es necesario guardar la nomina primeramente, ¿desea continuar?", "Pregunta", MessageBoxButtons.YesNo)
                    If resultado = DialogResult.Yes Then

                        cmdguardarnomina_Click(sender, e)
                        campoordenamiento = "Nomina.Buque,cNombreLargo"
                        llenargrid()
                    End If
                End If
                'dtgDatos.Columns(e.ColumnIndex).SortMode = DataGridViewColumnSortMode.Automatic
            End If

            For x As Integer = 0 To dtgDatos.Rows.Count - 1

                sql = "select * from empleadosC where iIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                Dim rwFila As DataRow() = nConsulta(sql)



                CType(Me.dtgDatos.Rows(x).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("cPuesto").ToString()

                CType(Me.dtgDatos.Rows(x).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("cFuncionesPuesto").ToString()
                dtgDatos.Rows(x).Cells(1).Value = x + 1
            Next


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try



    End Sub

    Private Sub chkAll_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkAll.CheckedChanged
        For x As Integer = 0 To dtgDatos.Rows.Count - 1
            dtgDatos.Rows(x).Cells(0).Value = Not dtgDatos.Rows(x).Cells(0).Value
        Next
        chkAll.Text = IIf(chkAll.Checked, "Desmarcar todos", "Marcar todos")
    End Sub

    Private Sub cmdlayouts_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdlayouts.Click

    End Sub

    Function RemoverBasura(ByVal nombre As String) As String
        Dim COMPSTR As String = "áéíóúÁÉÍÓÚ.ñÑ"
        Dim REPLSTR As String = "aeiouAEIOU nN"
        Dim Posicion As Integer
        Dim cadena As String = ""
        Dim arreglo As Char() = nombre.ToCharArray()
        For x As Integer = 0 To arreglo.Length - 1
            Posicion = COMPSTR.IndexOf(arreglo(x))
            If Posicion <> -1 Then
                arreglo(x) = REPLSTR(Posicion)

            End If
            cadena = cadena & arreglo(x)
        Next
        Return cadena
    End Function

    Function TipoCuentaBanco(ByVal idempleado As String, ByVal idempresa As String) As String
        'Agregar el banco y el tipo de cuenta ya sea a terceros o interbancaria
        'Buscamos el banco y verificarmos el tipo de cuenta a tercero o interbancaria
        Dim Sql As String
        Dim cadenabanco As String
        cadenabanco = ""

        Sql = "select iIdempleadoC,NumCuenta,Clabe,cuenta2,clabe2,fkiIdBanco,fkiIdBanco2"
        Sql &= " from empleadosC"
        Sql &= " where fkiIdEmpresa=" & gIdEmpresa & " and iIdempleadoC=" & idempleado

        Dim rwDatosBanco As DataRow() = nConsulta(Sql)

        cadenabanco = "@"

        If rwDatosBanco Is Nothing = False Then
            If rwDatosBanco(0)("NumCuenta") = "" Then
                cadenabanco &= "I"
            Else
                cadenabanco &= "T"
            End If

            If rwDatosBanco(0)("fkiIdBanco") = "1" Then
                cadenabanco &= "-BANAMEX"
            ElseIf rwDatosBanco(0)("fkiIdBanco") = "4" Then
                cadenabanco &= "-BANCOMER"
            ElseIf rwDatosBanco(0)("fkiIdBanco") = "13" Then
                cadenabanco &= "-SCOTIABANK"
            ElseIf rwDatosBanco(0)("fkiIdBanco") = "18" Then
                cadenabanco &= "-BANORTE"
            Else
                cadenabanco &= "-OTRO"
            End If

            cadenabanco &= "/"

            If rwDatosBanco(0)("cuenta2") = "" Then
                cadenabanco &= "I"
            Else
                cadenabanco &= "T"
            End If

            If rwDatosBanco(0)("fkiIdBanco2") = "1" Then
                cadenabanco &= "-BANAMEX"
            ElseIf rwDatosBanco(0)("fkiIdBanco2") = "4" Then
                cadenabanco &= "-BANCOMER"
            ElseIf rwDatosBanco(0)("fkiIdBanco2") = "13" Then
                cadenabanco &= "-SCOTIABANK"
            ElseIf rwDatosBanco(0)("fkiIdBanco2") = "18" Then
                cadenabanco &= "-BANORTE"
            Else
                cadenabanco &= "-OTRO"
            End If


        End If

        Return cadenabanco
    End Function

    Function CalculoPrimaSindicato(ByVal idempleado As String, ByVal idempresa As String) As String
        'Agregar el banco y el tipo de cuenta ya sea a terceros o interbancaria
        'Buscamos el banco y verificarmos el tipo de cuenta a tercero o interbancaria
        Dim Sql As String
        Dim cadenabanco As String
        Dim dia As String
        Dim mes As String
        Dim anio As String
        Dim anios As Integer
        Dim sueldodiario As Double
        Dim dias As Integer

        Dim Prima As String


        cadenabanco = ""


        Sql = "select *"
        Sql &= " from empleadosC"
        Sql &= " where fkiIdEmpresa=" & gIdEmpresa & " and iIdempleadoC=" & idempleado

        Dim rwDatosBanco As DataRow() = nConsulta(Sql)

        cadenabanco = "@"
        Prima = "0"
        If rwDatosBanco Is Nothing = False Then

            If Double.Parse(rwDatosBanco(0)("fsueldoOrd")) > 0 Then
                dia = Date.Parse(rwDatosBanco(0)("dFechaAntiguedad").ToString).Day.ToString("00")
                mes = Date.Parse(rwDatosBanco(0)("dFechaAntiguedad").ToString).Month.ToString("00")
                anio = Date.Today.Year
                'verificar el periodo para saber si queda entre el rango de fecha

                sueldodiario = Double.Parse(rwDatosBanco(0)("fsueldoOrd")) / diasperiodo

                Sql = "select * from periodos where iIdPeriodo= " & cboperiodo.SelectedValue
                Dim rwPeriodo As DataRow() = nConsulta(Sql)

                If rwPeriodo Is Nothing = False Then
                    Dim FechaBuscar As Date = Date.Parse(dia & "/" & mes & "/" & anio)
                    Dim FechaInicial As Date = Date.Parse(rwPeriodo(0)("dFechaInicio"))
                    Dim FechaFinal As Date = Date.Parse(rwPeriodo(0)("dFechaFin"))
                    Dim FechaAntiguedad As Date = Date.Parse(rwDatosBanco(0)("dFechaAntiguedad"))

                    If FechaBuscar.CompareTo(FechaInicial) >= 0 And FechaBuscar.CompareTo(FechaFinal) <= 0 Then
                        'Estamos dentro del rango 
                        'Calculamos la prima

                        anios = DateDiff("yyyy", FechaAntiguedad, FechaBuscar)

                        dias = CalculoDiasVacaciones(anios)

                        'Calcular prima

                        Prima = Math.Round(sueldodiario * dias * 0.25, 2).ToString()




                    End If


                End If


            End If


        End If


        Return Prima


    End Function


    Function CalculoDiasVacaciones(ByVal anios As Integer) As Integer
        Dim dias As Integer

        If anios = 1 Then
            dias = 6
        End If

        If anios = 2 Then
            dias = 8
        End If

        If anios = 3 Then
            dias = 10
        End If

        If anios = 4 Then
            dias = 12
        End If

        If anios >= 5 And anios <= 9 Then
            dias = 14
        End If

        If anios >= 10 And anios <= 14 Then
            dias = 16
        End If

        If anios >= 15 And anios <= 19 Then
            dias = 18
        End If

        If anios >= 20 And anios <= 24 Then
            dias = 20
        End If

        If anios >= 25 And anios <= 29 Then
            dias = 22
        End If

        If anios >= 30 And anios <= 34 Then
            dias = 24
        End If

        Return dias
    End Function

    Private Sub dtgDatos_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dtgDatos.CellContentClick
        Try
            If e.RowIndex = -1 And e.ColumnIndex = 0 Then
                Return
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub tsbEmpleados_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbEmpleados.Click
        Dim frm As New frmImportarEmpleadosAlta
        frm.ShowDialog()
    End Sub

    Private Sub dtgDatos_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dtgDatos.DataError
        Try
            e.Cancel = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub EliminarDeLaListaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EliminarDeLaListaToolStripMenuItem.Click
        If dtgDatos.CurrentRow Is Nothing = False Then
            Dim resultado As Integer = MessageBox.Show("¿Desea eliminar a este trabajador de la lista?", "Pregunta", MessageBoxButtons.YesNo)
            If resultado = DialogResult.Yes Then

                dtgDatos.Rows.Remove(dtgDatos.CurrentRow)
            End If
        End If

    End Sub

    Private Sub cbodias_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub cboserie_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboserie.SelectedIndexChanged
        dtgDatos.Columns.Clear()
        dtgDatos.DataSource = ""


    End Sub

    Private Sub cmdrecibosA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdrecibosA.Click

    End Sub


    Private Sub cboTipoNomina_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTipoNomina.SelectedIndexChanged
        'dtgDatos.Columns.Clear()
        'dtgDatos.DataSource = ""

    End Sub

    Private Sub AgregarTrabajadoresToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AgregarTrabajadoresToolStripMenuItem.Click
        Try
            Dim Forma As New frmAgregarEmpleado
            Dim ids As String()
            Dim sql As String
            Dim cadenaempleados As String
            If Forma.ShowDialog = Windows.Forms.DialogResult.OK Then
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
                dsPeriodo.Tables("Tabla").Columns.Add("Tiempo_Extra_Fijo_Gravado")
                dsPeriodo.Tables("Tabla").Columns.Add("Tiempo_Extra_Fijo_Exento")
                dsPeriodo.Tables("Tabla").Columns.Add("Tiempo_Extra_Ocasional")
                dsPeriodo.Tables("Tabla").Columns.Add("Desc_Sem_Obligatorio")
                dsPeriodo.Tables("Tabla").Columns.Add("Vacaciones_proporcionales")
                dsPeriodo.Tables("Tabla").Columns.Add("Aguinaldo_gravado")
                dsPeriodo.Tables("Tabla").Columns.Add("Aguinaldo_exento")
                dsPeriodo.Tables("Tabla").Columns.Add("Total_Aguinaldo")
                dsPeriodo.Tables("Tabla").Columns.Add("Prima_vac_gravado")
                dsPeriodo.Tables("Tabla").Columns.Add("Prima_vac_exento")
                dsPeriodo.Tables("Tabla").Columns.Add("Total_Prima_vac")
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


                ids = Forma.gidEmpleados.Split(",")
                If dtgDatos.Rows.Count > 0 Then
                    'Dim dt As DataTable = DirectCast(dtgDatos.DataSource, DataTable)
                    'dsPeriodo.Tables("Tabla") = dtgDatos.DataSource, DataTable
                    'Dim dt As DataTable = dsPeriodo.Tables("Tabla")


                    'For y As Integer = 0 To dt.Rows.Count - 1
                    '    dsPeriodo.Tables("Tabla").ImportRow(dt.Rows[y])
                    'Next

                    'Pasamos del datagrid al dataset ya creado
                    For y As Integer = 0 To dtgDatos.Rows.Count - 1

                        Dim fila As DataRow = dsPeriodo.Tables("Tabla").NewRow

                        fila.Item("Consecutivo") = (y + 1).ToString
                        fila.Item("Id_empleado") = dtgDatos.Rows(y).Cells(2).Value
                        fila.Item("CodigoEmpleado") = dtgDatos.Rows(y).Cells(3).Value
                        fila.Item("Nombre") = dtgDatos.Rows(y).Cells(4).Value
                        fila.Item("Status") = dtgDatos.Rows(y).Cells(5).Value
                        fila.Item("RFC") = dtgDatos.Rows(y).Cells(6).Value
                        fila.Item("CURP") = dtgDatos.Rows(y).Cells(7).Value
                        fila.Item("Num_IMSS") = dtgDatos.Rows(y).Cells(8).Value
                        fila.Item("Fecha_Nac") = dtgDatos.Rows(y).Cells(9).Value
                        'Dim tiempo As TimeSpan = Date.Now - Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString)





                        fila.Item("Edad") = dtgDatos.Rows(y).Cells(10).Value
                        fila.Item("Puesto") = dtgDatos.Rows(y).Cells(11).FormattedValue
                        fila.Item("Buque") = dtgDatos.Rows(y).Cells(12).FormattedValue

                        fila.Item("Tipo_Infonavit") = dtgDatos.Rows(y).Cells(13).Value
                        fila.Item("Valor_Infonavit") = IIf(dtgDatos.Rows(y).Cells(14).Value = "", "0", dtgDatos.Rows(y).Cells(14).Value.ToString.Replace(",", ""))
                        'salario base
                        fila.Item("Sueldo_Base") = dtgDatos.Rows(y).Cells(15).Value
                        fila.Item("Salario_Diario") = dtgDatos.Rows(y).Cells(16).Value
                        fila.Item("Salario_Cotización") = dtgDatos.Rows(y).Cells(17).Value
                        fila.Item("Dias_Trabajados") = dtgDatos.Rows(y).Cells(18).Value
                        fila.Item("Tipo_Incapacidad") = dtgDatos.Rows(y).Cells(19).Value
                        fila.Item("Número_días") = dtgDatos.Rows(y).Cells(20).Value
                        fila.Item("Sueldo_Bruto") = IIf(dtgDatos.Rows(y).Cells(21).Value = "", "0", dtgDatos.Rows(y).Cells(21).Value.ToString.Replace(",", ""))
                        fila.Item("Tiempo_Extra_Fijo_Gravado") = IIf(dtgDatos.Rows(y).Cells(22).Value = "", "0", dtgDatos.Rows(y).Cells(22).Value.ToString.Replace(",", ""))
                        fila.Item("Tiempo_Extra_Fijo_Exento") = IIf(dtgDatos.Rows(y).Cells(23).Value = "", "0", dtgDatos.Rows(y).Cells(23).Value.ToString.Replace(",", ""))
                        fila.Item("Tiempo_Extra_Ocasional") = IIf(dtgDatos.Rows(y).Cells(24).Value = "", "0", dtgDatos.Rows(y).Cells(24).Value.ToString.Replace(",", ""))
                        fila.Item("Desc_Sem_Obligatorio") = IIf(dtgDatos.Rows(y).Cells(25).Value = "", "0", dtgDatos.Rows(y).Cells(25).Value.ToString.Replace(",", ""))
                        fila.Item("Vacaciones_proporcionales") = IIf(dtgDatos.Rows(y).Cells(26).Value = "", "0", dtgDatos.Rows(y).Cells(26).Value.ToString.Replace(",", ""))
                        fila.Item("Aguinaldo_gravado") = IIf(dtgDatos.Rows(y).Cells(27).Value = "", "0", dtgDatos.Rows(y).Cells(27).Value.ToString.Replace(",", ""))
                        fila.Item("Aguinaldo_exento") = IIf(dtgDatos.Rows(y).Cells(28).Value = "", "0", dtgDatos.Rows(y).Cells(28).Value.ToString.Replace(",", ""))
                        fila.Item("Total_Aguinaldo") = IIf(dtgDatos.Rows(y).Cells(29).Value = "", "0", dtgDatos.Rows(y).Cells(29).Value.ToString.Replace(",", ""))
                        fila.Item("Prima_vac_gravado") = IIf(dtgDatos.Rows(y).Cells(30).Value = "", "0", dtgDatos.Rows(y).Cells(30).Value.ToString.Replace(",", ""))
                        fila.Item("Prima_vac_exento") = IIf(dtgDatos.Rows(y).Cells(31).Value = "", "0", dtgDatos.Rows(y).Cells(31).Value.ToString.Replace(",", ""))
                        fila.Item("Total_Prima_vac") = IIf(dtgDatos.Rows(y).Cells(32).Value = "", "0", dtgDatos.Rows(y).Cells(32).Value.ToString.Replace(",", ""))
                        fila.Item("Total_percepciones") = IIf(dtgDatos.Rows(y).Cells(33).Value = "", "0", dtgDatos.Rows(y).Cells(33).Value.ToString.Replace(",", ""))
                        fila.Item("Total_percepciones_p/isr") = IIf(dtgDatos.Rows(y).Cells(34).Value = "", "0", dtgDatos.Rows(y).Cells(34).Value.ToString.Replace(",", ""))
                        fila.Item("Incapacidad") = IIf(dtgDatos.Rows(y).Cells(35).Value = "", "0", dtgDatos.Rows(y).Cells(35).Value.ToString.Replace(",", ""))
                        fila.Item("ISR") = IIf(dtgDatos.Rows(y).Cells(36).Value = "", "0", dtgDatos.Rows(y).Cells(36).Value.ToString.Replace(",", ""))
                        fila.Item("IMSS") = IIf(dtgDatos.Rows(y).Cells(37).Value = "", "0", dtgDatos.Rows(y).Cells(37).Value.ToString.Replace(",", ""))
                        fila.Item("Infonavit") = IIf(dtgDatos.Rows(y).Cells(38).Value = "", "0", dtgDatos.Rows(y).Cells(38).Value.ToString.Replace(",", ""))
                        fila.Item("Infonavit_bim_anterior") = IIf(dtgDatos.Rows(y).Cells(39).Value = "", "0", dtgDatos.Rows(y).Cells(39).Value.ToString.Replace(",", ""))
                        fila.Item("Ajuste_infonavit") = IIf(dtgDatos.Rows(y).Cells(40).Value = "", "0", dtgDatos.Rows(y).Cells(40).Value.ToString.Replace(",", ""))
                        fila.Item("Pension_Alimenticia") = IIf(dtgDatos.Rows(y).Cells(41).Value = "", "0", dtgDatos.Rows(y).Cells(41).Value.ToString.Replace(",", ""))
                        fila.Item("Prestamo") = IIf(dtgDatos.Rows(y).Cells(42).Value = "", "0", dtgDatos.Rows(y).Cells(42).Value.ToString.Replace(",", ""))
                        fila.Item("Fonacot") = IIf(dtgDatos.Rows(y).Cells(43).Value = "", "0", dtgDatos.Rows(y).Cells(43).Value.ToString.Replace(",", ""))
                        fila.Item("Subsidio_Generado") = IIf(dtgDatos.Rows(y).Cells(44).Value = "", "0", dtgDatos.Rows(y).Cells(44).Value.ToString.Replace(",", ""))
                        fila.Item("Subsidio_Aplicado") = IIf(dtgDatos.Rows(y).Cells(45).Value = "", "0", dtgDatos.Rows(y).Cells(45).Value.ToString.Replace(",", ""))
                        fila.Item("Operadora") = IIf(dtgDatos.Rows(y).Cells(46).Value = "", "0", dtgDatos.Rows(y).Cells(46).Value.ToString.Replace(",", ""))





                        fila.Item("Prestamo_Personal_A") = IIf(dtgDatos.Rows(y).Cells(47).Value = "", "0", dtgDatos.Rows(y).Cells(47).Value.ToString.Replace(",", ""))
                        fila.Item("Adeudo_Infonavit_A") = IIf(dtgDatos.Rows(y).Cells(48).Value = "", "0", dtgDatos.Rows(y).Cells(48).Value.ToString.Replace(",", ""))
                        fila.Item("Diferencia_Infonavit_A") = IIf(dtgDatos.Rows(y).Cells(49).Value = "", "0", dtgDatos.Rows(y).Cells(49).Value.ToString.Replace(",", ""))
                        fila.Item("Asimilados") = IIf(dtgDatos.Rows(y).Cells(50).Value = "", "0", dtgDatos.Rows(y).Cells(50).Value.ToString.Replace(",", ""))
                        fila.Item("Retenciones_Operadora") = IIf(dtgDatos.Rows(y).Cells(51).Value = "", "0", dtgDatos.Rows(y).Cells(51).Value.ToString.Replace(",", ""))
                        fila.Item("%_Comisión") = IIf(dtgDatos.Rows(y).Cells(52).Value = "", "0", dtgDatos.Rows(y).Cells(52).Value.ToString.Replace(",", ""))
                        fila.Item("Comisión_Operadora") = IIf(dtgDatos.Rows(y).Cells(53).Value = "", "0", dtgDatos.Rows(y).Cells(53).Value.ToString.Replace(",", ""))
                        fila.Item("Comisión_Asimilados") = IIf(dtgDatos.Rows(y).Cells(54).Value = "", "0", dtgDatos.Rows(y).Cells(54).Value.ToString.Replace(",", ""))
                        fila.Item("IMSS_CS") = IIf(dtgDatos.Rows(y).Cells(55).Value = "", "0", dtgDatos.Rows(y).Cells(55).Value.ToString.Replace(",", ""))
                        fila.Item("RCV_CS") = IIf(dtgDatos.Rows(y).Cells(56).Value = "", "0", dtgDatos.Rows(y).Cells(56).Value.ToString.Replace(",", ""))
                        fila.Item("Infonavit_CS") = IIf(dtgDatos.Rows(y).Cells(57).Value = "", "0", dtgDatos.Rows(y).Cells(57).Value.ToString.Replace(",", ""))
                        fila.Item("ISN_CS") = IIf(dtgDatos.Rows(y).Cells(58).Value = "", "0", dtgDatos.Rows(y).Cells(58).Value.ToString.Replace(",", ""))
                        fila.Item("Total_Costo_Social") = IIf(dtgDatos.Rows(y).Cells(59).Value = "", "0", dtgDatos.Rows(y).Cells(59).Value.ToString.Replace(",", ""))
                        fila.Item("Subtotal") = IIf(dtgDatos.Rows(y).Cells(60).Value = "", "0", dtgDatos.Rows(y).Cells(60).Value.ToString.Replace(",", ""))
                        fila.Item("IVA") = IIf(dtgDatos.Rows(y).Cells(61).Value = "", "0", dtgDatos.Rows(y).Cells(61).Value.ToString.Replace(",", ""))
                        fila.Item("TOTAL_DEPOSITO") = IIf(dtgDatos.Rows(y).Cells(62).Value = "", "0", dtgDatos.Rows(y).Cells(62).Value.ToString.Replace(",", ""))




                        dsPeriodo.Tables("Tabla").Rows.Add(fila)
                    Next

                    'Agregar a la tabla los datos que vienen de la busqueda de empleados
                    For x As Integer = 0 To ids.Length - 1

                        Dim fila As DataRow = dsPeriodo.Tables("Tabla").NewRow
                        'Dim fila As DataRow = dt.NewRow
                        'Dim fila As DataRow = dsPeriodo.Tables("Tabla").NewRow
                        sql = "select  * from empleadosC where " 'fkiIdClienteInter=-1"
                        sql &= " iIdEmpleadoC=" & ids(x)
                        sql &= " order by cFuncionesPuesto,cNombreLargo"
                        Dim rwEmpleado As DataRow() = nConsulta(sql)
                        If rwEmpleado Is Nothing = False Then
                            fila.Item("Consecutivo") = (dtgDatos.Rows.Count + x + 1).ToString
                            fila.Item("Id_empleado") = rwEmpleado(0)("iIdEmpleadoC").ToString
                            fila.Item("CodigoEmpleado") = rwEmpleado(0)("cCodigoEmpleado").ToString
                            fila.Item("Nombre") = rwEmpleado(0)("cNombreLargo").ToString.ToUpper()
                            fila.Item("Status") = IIf(rwEmpleado(0)("iOrigen").ToString = "1", "INTERINO", "PLANTA")
                            fila.Item("RFC") = rwEmpleado(0)("cRFC").ToString
                            fila.Item("CURP") = rwEmpleado(0)("cCURP").ToString
                            fila.Item("Num_IMSS") = rwEmpleado(0)("cIMSS").ToString

                            fila.Item("Fecha_Nac") = Date.Parse(rwEmpleado(0)("dFechaNac").ToString).ToShortDateString()
                            'Dim tiempo As TimeSpan = Date.Now - Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString)
                            fila.Item("Edad") = CalcularEdad(Date.Parse(rwEmpleado(0)("dFechaNac").ToString).Day, Date.Parse(rwEmpleado(0)("dFechaNac").ToString).Month, Date.Parse(rwEmpleado(0)("dFechaNac").ToString).Year)
                            fila.Item("Puesto") = rwEmpleado(0)("cPuesto").ToString
                            fila.Item("Buque") = "ECO III"

                            fila.Item("Tipo_Infonavit") = rwEmpleado(0)("cTipoFactor").ToString
                            fila.Item("Valor_Infonavit") = rwEmpleado(0)("fFactor").ToString
                            fila.Item("Sueldo_Base") = "0.00"
                            fila.Item("Salario_Diario") = rwEmpleado(0)("fSueldoBase").ToString
                            fila.Item("Salario_Cotización") = rwEmpleado(0)("fSueldoIntegrado").ToString
                            fila.Item("Dias_Trabajados") = "30"
                            fila.Item("Tipo_Incapacidad") = TipoIncapacidad(rwEmpleado(0)("iIdEmpleadoC").ToString, cboperiodo.SelectedValue)
                            fila.Item("Número_días") = NumDiasIncapacidad(rwEmpleado(0)("iIdEmpleadoC").ToString, cboperiodo.SelectedValue)
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
                            'dt.Rows.Add(fila)



                        End If

                    Next
                    'dtgDatos.DataSource = dt
                    dtgDatos.Columns.Clear()
                    Dim chk As New DataGridViewCheckBoxColumn()
                    dtgDatos.Columns.Add(chk)
                    chk.HeaderText = ""
                    chk.Name = "chk"
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

                    'calcular()

                    'Cambiamos index del combo en el grid




                    For x As Integer = 0 To dtgDatos.Rows.Count - 1

                        sql = "select * from nomina where fkiIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                        sql &= " and fkiIdPeriodo=" & cboperiodo.SelectedValue
                        sql &= " and iEstatusEmpleado=" & cboserie.SelectedIndex
                        sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex
                        Dim rwFila As DataRow() = nConsulta(sql)

                        If rwFila Is Nothing = False Then
                            CType(Me.dtgDatos.Rows(x).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("Puesto").ToString()
                            CType(Me.dtgDatos.Rows(x).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("Buque").ToString()
                        Else
                            sql = "select * from empleadosC where iIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                            Dim rwEmpleado As DataRow() = nConsulta(sql)



                            CType(Me.dtgDatos.Rows(x).Cells(11), DataGridViewComboBoxCell).Value = rwEmpleado(0)("cPuesto").ToString()
                            CType(Me.dtgDatos.Rows(x).Cells(12), DataGridViewComboBoxCell).Value = rwEmpleado(0)("cFuncionesPuesto").ToString()
                        End If



                    Next

                    MessageBox.Show("Datos cargados", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                Else



                    cadenaempleados = ""

                    For x As Integer = 0 To ids.Length - 1
                        If x = 0 Then
                            cadenaempleados = " iIdEmpleadoC=" & ids(x)
                        Else
                            cadenaempleados &= "  or iIdEmpleadoC=" & ids(x)
                        End If
                    Next






                    sql = "select  * from empleadosC where " 'fkiIdClienteInter=-1"
                    sql &= cadenaempleados
                    sql &= " order by cFuncionesPuesto,cNombreLargo"

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

                        dtgDatos.Columns.Clear()
                        Dim chk As New DataGridViewCheckBoxColumn()
                        dtgDatos.Columns.Add(chk)
                        chk.HeaderText = ""
                        chk.Name = "chk"
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




                        MessageBox.Show("Datos cargados", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("No hay datos en este período", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If




                    'No hay datos en este período


                End If




                'MessageBox.Show("Trabajadores asignados", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                'If cboempresa.SelectedIndex > -1 Then
                '    cargarlista()
                'End If
                'lsvLista.SelectedItems(0).Tag = ""
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

    Private Sub cmdSubirDatos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSubirDatos.Click
        Try
            Dim Forma As New frmSubirDatos
            Dim ids As String()
            Dim sql As String
            Dim cadenaempleados As String
            If Forma.ShowDialog = Windows.Forms.DialogResult.OK Then


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
                dsPeriodo.Tables("Tabla").Columns.Add("Tiempo_Extra_Fijo_Gravado")
                dsPeriodo.Tables("Tabla").Columns.Add("Tiempo_Extra_Fijo_Exento")
                dsPeriodo.Tables("Tabla").Columns.Add("Tiempo_Extra_Ocasional")
                dsPeriodo.Tables("Tabla").Columns.Add("Desc_Sem_Obligatorio")
                dsPeriodo.Tables("Tabla").Columns.Add("Vacaciones_proporcionales")
                dsPeriodo.Tables("Tabla").Columns.Add("Aguinaldo_gravado")
                dsPeriodo.Tables("Tabla").Columns.Add("Aguinaldo_exento")
                dsPeriodo.Tables("Tabla").Columns.Add("Total_Aguinaldo")
                dsPeriodo.Tables("Tabla").Columns.Add("Prima_vac_gravado")
                dsPeriodo.Tables("Tabla").Columns.Add("Prima_vac_exento")
                dsPeriodo.Tables("Tabla").Columns.Add("Total_Prima_vac")
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
                dsPeriodo.Tables("Tabla").Columns.Add("fecha_inicio")
                dsPeriodo.Tables("Tabla").Columns.Add("fecha_fin")




                dtgDatos.Columns.Clear()
                dtgDatos.DataSource = Nothing

                'ids = Forma.gidEmpleados.Split(",")
                If dtgDatos.Rows.Count > 0 Then
                    'Dim dt As DataTable = DirectCast(dtgDatos.DataSource, DataTable)
                    'dsPeriodo.Tables("Tabla") = dtgDatos.DataSource, DataTable
                    'Dim dt As DataTable = dsPeriodo.Tables("Tabla")


                    'For y As Integer = 0 To dt.Rows.Count - 1
                    '    dsPeriodo.Tables("Tabla").ImportRow(dt.Rows[y])
                    'Next

                    'Pasamos del datagrid al dataset ya creado
                    For y As Integer = 0 To dtgDatos.Rows.Count - 1

                        Dim fila As DataRow = dsPeriodo.Tables("Tabla").NewRow

                        fila.Item("Consecutivo") = (y + 1).ToString
                        fila.Item("Id_empleado") = dtgDatos.Rows(y).Cells(2).Value
                        fila.Item("CodigoEmpleado") = dtgDatos.Rows(y).Cells(3).Value
                        fila.Item("Nombre") = dtgDatos.Rows(y).Cells(4).Value
                        fila.Item("Status") = dtgDatos.Rows(y).Cells(5).Value
                        fila.Item("RFC") = dtgDatos.Rows(y).Cells(6).Value
                        fila.Item("CURP") = dtgDatos.Rows(y).Cells(7).Value
                        fila.Item("Num_IMSS") = dtgDatos.Rows(y).Cells(8).Value
                        fila.Item("Fecha_Nac") = dtgDatos.Rows(y).Cells(9).Value
                        'Dim tiempo As TimeSpan = Date.Now - Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString)

                        fila.Item("Edad") = dtgDatos.Rows(y).Cells(10).Value
                        fila.Item("Puesto") = dtgDatos.Rows(y).Cells(11).FormattedValue
                        fila.Item("Buque") = dtgDatos.Rows(y).Cells(12).FormattedValue

                        fila.Item("Tipo_Infonavit") = dtgDatos.Rows(y).Cells(13).Value
                        fila.Item("Valor_Infonavit") = IIf(dtgDatos.Rows(y).Cells(14).Value = "", "0", dtgDatos.Rows(y).Cells(14).Value.ToString.Replace(",", ""))
                        'salario base
                        fila.Item("Sueldo_Base") = dtgDatos.Rows(y).Cells(15).Value
                        fila.Item("Salario_Diario") = dtgDatos.Rows(y).Cells(16).Value
                        fila.Item("Salario_Cotización") = dtgDatos.Rows(y).Cells(17).Value
                        fila.Item("Dias_Trabajados") = dtgDatos.Rows(y).Cells(18).Value
                        fila.Item("Tipo_Incapacidad") = dtgDatos.Rows(y).Cells(19).Value
                        fila.Item("Número_días") = dtgDatos.Rows(y).Cells(20).Value
                        fila.Item("Sueldo_Bruto") = IIf(dtgDatos.Rows(y).Cells(21).Value = "", "0", dtgDatos.Rows(y).Cells(21).Value.ToString.Replace(",", ""))
                        fila.Item("Tiempo_Extra_Fijo_Gravado") = IIf(dtgDatos.Rows(y).Cells(22).Value = "", "0", dtgDatos.Rows(y).Cells(22).Value.ToString.Replace(",", ""))
                        fila.Item("Tiempo_Extra_Fijo_Exento") = IIf(dtgDatos.Rows(y).Cells(23).Value = "", "0", dtgDatos.Rows(y).Cells(23).Value.ToString.Replace(",", ""))
                        fila.Item("Tiempo_Extra_Ocasional") = IIf(dtgDatos.Rows(y).Cells(24).Value = "", "0", dtgDatos.Rows(y).Cells(24).Value.ToString.Replace(",", ""))
                        fila.Item("Desc_Sem_Obligatorio") = IIf(dtgDatos.Rows(y).Cells(25).Value = "", "0", dtgDatos.Rows(y).Cells(25).Value.ToString.Replace(",", ""))
                        fila.Item("Vacaciones_proporcionales") = IIf(dtgDatos.Rows(y).Cells(26).Value = "", "0", dtgDatos.Rows(y).Cells(26).Value.ToString.Replace(",", ""))
                        fila.Item("Aguinaldo_gravado") = IIf(dtgDatos.Rows(y).Cells(27).Value = "", "0", dtgDatos.Rows(y).Cells(27).Value.ToString.Replace(",", ""))
                        fila.Item("Aguinaldo_exento") = IIf(dtgDatos.Rows(y).Cells(28).Value = "", "0", dtgDatos.Rows(y).Cells(28).Value.ToString.Replace(",", ""))
                        fila.Item("Total_Aguinaldo") = IIf(dtgDatos.Rows(y).Cells(29).Value = "", "0", dtgDatos.Rows(y).Cells(29).Value.ToString.Replace(",", ""))
                        fila.Item("Prima_vac_gravado") = IIf(dtgDatos.Rows(y).Cells(30).Value = "", "0", dtgDatos.Rows(y).Cells(30).Value.ToString.Replace(",", ""))
                        fila.Item("Prima_vac_exento") = IIf(dtgDatos.Rows(y).Cells(31).Value = "", "0", dtgDatos.Rows(y).Cells(31).Value.ToString.Replace(",", ""))

                        fila.Item("Total_Prima_vac") = IIf(dtgDatos.Rows(y).Cells(32).Value = "", "0", dtgDatos.Rows(y).Cells(32).Value.ToString.Replace(",", ""))
                        fila.Item("Total_percepciones") = IIf(dtgDatos.Rows(y).Cells(33).Value = "", "0", dtgDatos.Rows(y).Cells(33).Value.ToString.Replace(",", ""))
                        fila.Item("Total_percepciones_p/isr") = IIf(dtgDatos.Rows(y).Cells(34).Value = "", "0", dtgDatos.Rows(y).Cells(34).Value.ToString.Replace(",", ""))
                        fila.Item("Incapacidad") = IIf(dtgDatos.Rows(y).Cells(35).Value = "", "0", dtgDatos.Rows(y).Cells(35).Value.ToString.Replace(",", ""))
                        fila.Item("ISR") = IIf(dtgDatos.Rows(y).Cells(36).Value = "", "0", dtgDatos.Rows(y).Cells(36).Value.ToString.Replace(",", ""))
                        fila.Item("IMSS") = IIf(dtgDatos.Rows(y).Cells(37).Value = "", "0", dtgDatos.Rows(y).Cells(37).Value.ToString.Replace(",", ""))
                        fila.Item("Infonavit") = IIf(dtgDatos.Rows(y).Cells(38).Value = "", "0", dtgDatos.Rows(y).Cells(38).Value.ToString.Replace(",", ""))
                        fila.Item("Infonavit_bim_anterior") = IIf(dtgDatos.Rows(y).Cells(39).Value = "", "0", dtgDatos.Rows(y).Cells(39).Value.ToString.Replace(",", ""))
                        fila.Item("Ajuste_infonavit") = IIf(dtgDatos.Rows(y).Cells(40).Value = "", "0", dtgDatos.Rows(y).Cells(40).Value.ToString.Replace(",", ""))
                        fila.Item("Pension_Alimenticia") = IIf(dtgDatos.Rows(y).Cells(41).Value = "", "0", dtgDatos.Rows(y).Cells(41).Value.ToString.Replace(",", ""))
                        fila.Item("Prestamo") = IIf(dtgDatos.Rows(y).Cells(42).Value = "", "0", dtgDatos.Rows(y).Cells(42).Value.ToString.Replace(",", ""))
                        fila.Item("Fonacot") = IIf(dtgDatos.Rows(y).Cells(43).Value = "", "0", dtgDatos.Rows(y).Cells(43).Value.ToString.Replace(",", ""))
                        fila.Item("Subsidio_Generado") = IIf(dtgDatos.Rows(y).Cells(44).Value = "", "0", dtgDatos.Rows(y).Cells(44).Value.ToString.Replace(",", ""))
                        fila.Item("Subsidio_Aplicado") = IIf(dtgDatos.Rows(y).Cells(45).Value = "", "0", dtgDatos.Rows(y).Cells(45).Value.ToString.Replace(",", ""))
                        fila.Item("Operadora") = IIf(dtgDatos.Rows(y).Cells(46).Value = "", "0", dtgDatos.Rows(y).Cells(46).Value.ToString.Replace(",", ""))
                        fila.Item("Prestamo_Personal_A") = IIf(dtgDatos.Rows(y).Cells(47).Value = "", "0", dtgDatos.Rows(y).Cells(47).Value.ToString.Replace(",", ""))
                        fila.Item("Adeudo_Infonavit_A") = IIf(dtgDatos.Rows(y).Cells(48).Value = "", "0", dtgDatos.Rows(y).Cells(48).Value.ToString.Replace(",", ""))
                        fila.Item("Diferencia_Infonavit_A") = IIf(dtgDatos.Rows(y).Cells(49).Value = "", "0", dtgDatos.Rows(y).Cells(49).Value.ToString.Replace(",", ""))
                        fila.Item("Asimilados") = IIf(dtgDatos.Rows(y).Cells(50).Value = "", "0", dtgDatos.Rows(y).Cells(50).Value.ToString.Replace(",", ""))
                        fila.Item("Retenciones_Operadora") = IIf(dtgDatos.Rows(y).Cells(51).Value = "", "0", dtgDatos.Rows(y).Cells(51).Value.ToString.Replace(",", ""))
                        fila.Item("%_Comisión") = IIf(dtgDatos.Rows(y).Cells(52).Value = "", "0", dtgDatos.Rows(y).Cells(52).Value.ToString.Replace(",", ""))
                        fila.Item("Comisión_Operadora") = IIf(dtgDatos.Rows(y).Cells(53).Value = "", "0", dtgDatos.Rows(y).Cells(53).Value.ToString.Replace(",", ""))
                        fila.Item("Comisión_Asimilados") = IIf(dtgDatos.Rows(y).Cells(54).Value = "", "0", dtgDatos.Rows(y).Cells(54).Value.ToString.Replace(",", ""))
                        fila.Item("IMSS_CS") = IIf(dtgDatos.Rows(y).Cells(55).Value = "", "0", dtgDatos.Rows(y).Cells(55).Value.ToString.Replace(",", ""))
                        fila.Item("RCV_CS") = IIf(dtgDatos.Rows(y).Cells(56).Value = "", "0", dtgDatos.Rows(y).Cells(56).Value.ToString.Replace(",", ""))
                        fila.Item("Infonavit_CS") = IIf(dtgDatos.Rows(y).Cells(57).Value = "", "0", dtgDatos.Rows(y).Cells(57).Value.ToString.Replace(",", ""))
                        fila.Item("ISN_CS") = IIf(dtgDatos.Rows(y).Cells(58).Value = "", "0", dtgDatos.Rows(y).Cells(58).Value.ToString.Replace(",", ""))
                        fila.Item("Total_Costo_Social") = IIf(dtgDatos.Rows(y).Cells(59).Value = "", "0", dtgDatos.Rows(y).Cells(59).Value.ToString.Replace(",", ""))
                        fila.Item("Subtotal") = IIf(dtgDatos.Rows(y).Cells(60).Value = "", "0", dtgDatos.Rows(y).Cells(60).Value.ToString.Replace(",", ""))
                        fila.Item("IVA") = IIf(dtgDatos.Rows(y).Cells(61).Value = "", "0", dtgDatos.Rows(y).Cells(61).Value.ToString.Replace(",", ""))
                        fila.Item("TOTAL_DEPOSITO") = IIf(dtgDatos.Rows(y).Cells(62).Value = "", "0", dtgDatos.Rows(y).Cells(62).Value.ToString.Replace(",", ""))

                        fila.Item("Neto_Pagar") = IIf(dtgDatos.Rows(y).Cells(46).Value = "", "0", dtgDatos.Rows(y).Cells(46).Value.ToString.Replace(",", ""))
                        fila.Item("Excendente") = IIf(dtgDatos.Rows(y).Cells(47).Value = "", "0", dtgDatos.Rows(y).Cells(47).Value.ToString.Replace(",", ""))
                        fila.Item("Total") = IIf(dtgDatos.Rows(y).Cells(48).Value = "", "0", dtgDatos.Rows(y).Cells(48).Value.ToString.Replace(",", ""))
                        fila.Item("IMSS_CS") = IIf(dtgDatos.Rows(y).Cells(49).Value = "", "0", dtgDatos.Rows(y).Cells(49).Value.ToString.Replace(",", ""))
                        fila.Item("RCV_CS") = IIf(dtgDatos.Rows(y).Cells(50).Value = "", "0", dtgDatos.Rows(y).Cells(50).Value.ToString.Replace(",", ""))
                        fila.Item("Infonavit_CS") = IIf(dtgDatos.Rows(y).Cells(51).Value = "", "0", dtgDatos.Rows(y).Cells(51).Value.ToString.Replace(",", ""))
                        fila.Item("ISN_CS") = IIf(dtgDatos.Rows(y).Cells(52).Value = "", "0", dtgDatos.Rows(y).Cells(52).Value.ToString.Replace(",", ""))
                        fila.Item("Prestamo_Personal") = IIf(dtgDatos.Rows(y).Cells(53).Value = "", "0", dtgDatos.Rows(y).Cells(53).Value.ToString.Replace(",", ""))
                        fila.Item("Adeudo_Infonavit") = IIf(dtgDatos.Rows(y).Cells(54).Value = "", "0", dtgDatos.Rows(y).Cells(54).Value.ToString.Replace(",", ""))
                        fila.Item("Diferencia_Infonavit") = IIf(dtgDatos.Rows(y).Cells(55).Value = "", "0", dtgDatos.Rows(y).Cells(55).Value.ToString.Replace(",", ""))
                        fila.Item("Complemento_Asimilados") = IIf(dtgDatos.Rows(y).Cells(56).Value = "", "0", dtgDatos.Rows(y).Cells(56).Value.ToString.Replace(",", ""))


                        dsPeriodo.Tables("Tabla").Rows.Add(fila)
                    Next

                    'Agregar a la tabla los datos que vienen de la busqueda de empleados
                    For x As Integer = 0 To Forma.dsReporte.Tables(0).Rows.Count - 1
                        Dim fila As DataRow = dsPeriodo.Tables("Tabla").NewRow
                        'Dim fila As DataRow = dt.NewRow
                        'Dim fila As DataRow = dsPeriodo.Tables("Tabla").NewRow
                        sql = "select  * from empleadosC where " 'fkiIdClienteInter=-1"
                        sql &= " iIdEmpleadoC=" & ids(x)
                        sql &= " order by cFuncionesPuesto,cNombreLargo"
                        Dim rwEmpleado As DataRow() = nConsulta(sql)
                        If rwEmpleado Is Nothing = False Then
                            fila.Item("Consecutivo") = (dtgDatos.Rows.Count + x + 1).ToString
                            fila.Item("Id_empleado") = rwEmpleado(0)("iIdEmpleadoC").ToString
                            fila.Item("CodigoEmpleado") = rwEmpleado(0)("cCodigoEmpleado").ToString
                            fila.Item("Nombre") = rwEmpleado(0)("cNombreLargo").ToString.ToUpper()
                            fila.Item("Status") = IIf(rwEmpleado(0)("iOrigen").ToString = "1", "INTERINO", "PLANTA")
                            fila.Item("RFC") = rwEmpleado(0)("cRFC").ToString
                            fila.Item("CURP") = rwEmpleado(0)("cCURP").ToString
                            fila.Item("Num_IMSS") = rwEmpleado(0)("cIMSS").ToString

                            fila.Item("Fecha_Nac") = Date.Parse(rwEmpleado(0)("dFechaNac").ToString).ToShortDateString()
                            'Dim tiempo As TimeSpan = Date.Now - Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString)
                            fila.Item("Edad") = CalcularEdad(Date.Parse(rwEmpleado(0)("dFechaNac").ToString).Day, Date.Parse(rwEmpleado(0)("dFechaNac").ToString).Month, Date.Parse(rwEmpleado(0)("dFechaNac").ToString).Year)
                            fila.Item("Puesto") = rwEmpleado(0)("cPuesto").ToString
                            fila.Item("Buque") = "ECO III"

                            fila.Item("Tipo_Infonavit") = rwEmpleado(0)("cTipoFactor").ToString
                            fila.Item("Valor_Infonavit") = rwEmpleado(0)("fFactor").ToString
                            fila.Item("Sueldo_Base") = "0.00"
                            fila.Item("Salario_Diario") = rwEmpleado(0)("fSueldoBase").ToString
                            fila.Item("Salario_Cotización") = rwEmpleado(0)("fSueldoIntegrado").ToString
                            fila.Item("Dias_Trabajados") = "30"
                            fila.Item("Tipo_Incapacidad") = TipoIncapacidad(rwEmpleado(0)("iIdEmpleadoC").ToString, cboperiodo.SelectedValue)
                            fila.Item("Número_días") = NumDiasIncapacidad(rwEmpleado(0)("iIdEmpleadoC").ToString, cboperiodo.SelectedValue)
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
                            'dt.Rows.Add(fila)



                        End If
                    Next

                    'dtgDatos.DataSource = dt
                    dtgDatos.Columns.Clear()
                    Dim chk As New DataGridViewCheckBoxColumn()
                    dtgDatos.Columns.Add(chk)
                    chk.HeaderText = ""
                    chk.Name = "chk"
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

                    'calcular()

                    'Cambiamos index del combo en el grid




                    For x As Integer = 0 To dtgDatos.Rows.Count - 1

                        sql = "select * from nomina where fkiIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                        sql &= " and fkiIdPeriodo=" & cboperiodo.SelectedValue
                        sql &= " and iEstatusEmpleado=" & cboserie.SelectedIndex
                        sql &= " and iTipoNomina=" & cboTipoNomina.SelectedIndex
                        Dim rwFila As DataRow() = nConsulta(sql)

                        If rwFila Is Nothing = False Then
                            CType(Me.dtgDatos.Rows(x).Cells(11), DataGridViewComboBoxCell).Value = rwFila(0)("Puesto").ToString()
                            CType(Me.dtgDatos.Rows(x).Cells(12), DataGridViewComboBoxCell).Value = rwFila(0)("Buque").ToString()
                        Else
                            sql = "select * from empleadosC where iIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                            Dim rwEmpleado As DataRow() = nConsulta(sql)



                            CType(Me.dtgDatos.Rows(x).Cells(11), DataGridViewComboBoxCell).Value = rwEmpleado(0)("cPuesto").ToString()
                            CType(Me.dtgDatos.Rows(x).Cells(12), DataGridViewComboBoxCell).Value = rwEmpleado(0)("cFuncionesPuesto").ToString()
                        End If



                    Next

                    MessageBox.Show("Datos cargados", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                Else



                    cadenaempleados = ""

                    For x As Integer = 0 To Forma.dsReporte.Tables(0).Rows.Count - 1
                        sql = "select  * from empleadosC where " 'fkiIdClienteInter=-1"
                        sql &= "iIdEmpleadoC=" & Forma.dsReporte.Tables(0).Rows(x)("Id_empleado")
                        sql &= " order by cFuncionesPuesto,cNombreLargo"
                        Dim rwDatosEmpleado As DataRow() = nConsulta(sql)
                        If rwDatosEmpleado Is Nothing = False Then
                            Dim fila As DataRow = dsPeriodo.Tables("Tabla").NewRow

                            fila.Item("Consecutivo") = (x + 1).ToString
                            fila.Item("Id_empleado") = rwDatosEmpleado(0)("iIdEmpleadoC").ToString
                            fila.Item("CodigoEmpleado") = rwDatosEmpleado(0)("cCodigoEmpleado").ToString
                            fila.Item("Nombre") = rwDatosEmpleado(0)("cNombreLargo").ToString.ToUpper()
                            fila.Item("Status") = IIf(rwDatosEmpleado(0)("iOrigen").ToString = "1", "INTERINO", "PLANTA")
                            fila.Item("RFC") = rwDatosEmpleado(0)("cRFC").ToString
                            fila.Item("CURP") = rwDatosEmpleado(0)("cCURP").ToString
                            fila.Item("Num_IMSS") = rwDatosEmpleado(0)("cIMSS").ToString

                            fila.Item("Fecha_Nac") = Date.Parse(rwDatosEmpleado(0)("dFechaNac").ToString).ToShortDateString()
                            'Dim tiempo As TimeSpan = Date.Now - Date.Parse(rwDatosEmpleados(x)("dFechaNac").ToString)
                            fila.Item("Edad") = CalcularEdad(Date.Parse(rwDatosEmpleado(0)("dFechaNac").ToString).Day, Date.Parse(rwDatosEmpleado(0)("dFechaNac").ToString).Month, Date.Parse(rwDatosEmpleado(0)("dFechaNac").ToString).Year)
                            fila.Item("Puesto") = rwDatosEmpleado(0)("cPuesto").ToString
                            fila.Item("Buque") = "ECO III"

                            fila.Item("Tipo_Infonavit") = rwDatosEmpleado(0)("cTipoFactor").ToString
                            fila.Item("Valor_Infonavit") = rwDatosEmpleado(0)("fFactor").ToString
                            fila.Item("Sueldo_Base") = Forma.dsReporte.Tables(0).Rows(x)("SalarioTMM")
                            fila.Item("Salario_Diario") = rwDatosEmpleado(0)("fSueldoBase").ToString
                            fila.Item("Salario_Cotización") = rwDatosEmpleado(0)("fSueldoIntegrado").ToString
                            fila.Item("Dias_Trabajados") = Forma.dsReporte.Tables(0).Rows(x)("dias")
                            fila.Item("Tipo_Incapacidad") = TipoIncapacidad(rwDatosEmpleado(0)("iIdEmpleadoC").ToString, cboperiodo.SelectedValue)
                            fila.Item("Número_días") = NumDiasIncapacidad(rwDatosEmpleado(0)("iIdEmpleadoC").ToString, cboperiodo.SelectedValue)
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

                            fila.Item("fecha_inicio") = Forma.dsReporte.Tables(0).Rows(x)("Fechainicio")
                            fila.Item("fecha_fin") = Forma.dsReporte.Tables(0).Rows(x)("Fechafin")
                            dsPeriodo.Tables("Tabla").Rows.Add(fila)
                        End If




                    Next




                    dtgDatos.Columns.Clear()
                    Dim chk As New DataGridViewCheckBoxColumn()
                    dtgDatos.Columns.Add(chk)
                    chk.HeaderText = ""
                    chk.Name = "chk"
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

                    'FECHA INICIO
                    dtgDatos.Columns(63).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(63).ReadOnly = True
                    dtgDatos.Columns(63).Width = 150

                    'FECHA FIN
                    dtgDatos.Columns(64).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    dtgDatos.Columns(64).ReadOnly = True
                    dtgDatos.Columns(64).Width = 150

                    'calcular()

                    'Cambiamos index del combo en el grid

                    For x As Integer = 0 To Forma.dsReporte.Tables(0).Rows.Count - 1

                        'sql = "select * from empleadosC where iIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value
                        'Dim rwFila As DataRow() = nConsulta(sql)


                        'buscar el nombre del puesto

                        sql = "select * from puestos where iIdPuesto=" & Forma.dsReporte.Tables(0).Rows(x)("CodigoPuesto")
                        Dim rwPuesto As DataRow() = nConsulta(sql)


                        CType(Me.dtgDatos.Rows(x).Cells(11), DataGridViewComboBoxCell).Value = rwPuesto(0)("cNombre").ToString()


                        'buscar el nombre del buque de acuerdo a lo guardado

                        sql = "select * from departamentos where iIdDepartamento=" & Forma.dsReporte.Tables(0).Rows(x)("CodigoBuque")
                        Dim rwBuque As DataRow() = nConsulta(sql)

                        CType(Me.dtgDatos.Rows(x).Cells(12), DataGridViewComboBoxCell).Value = rwBuque(0)("cNombre").ToString()

                    Next




                    MessageBox.Show("Datos cargados", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)





                    'No hay datos en este período


                End If

            End If
            'Forma.gIdEmpresa = gIdEmpresa
            'Forma.gIdPeriodo = cboperiodo.SelectedValue
            'Forma.gIdTipoPuesto = 1
            'Forma.ShowDialog()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cmdexcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdexcel.Click
        Try

            Dim filaExcel As Integer = 0
            Dim filatmp As Integer = 0
            Dim dialogo As New SaveFileDialog()
            Dim periodo, iejercicio As String

            pnlProgreso.Visible = True
            pnlCatalogo.Enabled = False
            Application.DoEvents()

            pgbProgreso.Minimum = 0
            pgbProgreso.Value = 0
            pgbProgreso.Maximum = dtgDatos.Rows.Count


            If dtgDatos.Rows.Count > 0 Then
                Dim ruta As String
                ruta = My.Application.Info.DirectoryPath() & "\Archivos\nominasmarinos.xlsx"

                Dim book As New ClosedXML.Excel.XLWorkbook(ruta)
                Dim libro As New ClosedXML.Excel.XLWorkbook

                book.Worksheet(1).CopyTo(libro, "MARINOS")
                book.Worksheet(2).CopyTo(libro, "XURTEP ABORDO")
                book.Worksheet(3).CopyTo(libro, "XURTEP DESCANSO")
                book.Worksheet(4).CopyTo(libro, "ASIMILADOS DESCANSO")
                book.Worksheets(5).CopyTo(libro, "RESUMEN")
                book.Worksheets(6).CopyTo(libro, "PENSION ALIMENTICIA")


                Dim hoja As IXLWorksheet = libro.Worksheets(0)
                Dim hoja2 As IXLWorksheet = libro.Worksheets(1)
                Dim hoja3 As IXLWorksheet = libro.Worksheets(2)
                Dim hoja4 As IXLWorksheet = libro.Worksheets(3)
                Dim hoja5 As IXLWorksheet = libro.Worksheets(4)
                Dim hoja6 As IXLWorksheet = libro.Worksheets(5)

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

                '<<<<<<<<<<<<<<<<<<<<<<<<<<Marinos>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                recorrerFilasColumnas(hoja, 11, dtgDatos.Rows.Count + 10, 50, "clear")

                Dim rwPeriodo0 As DataRow() = nConsulta("Select * from periodos where iIdPeriodo=" & cboperiodo.SelectedValue)
                If rwPeriodo0 Is Nothing = False Then
                    Dim Fechafin As Date = rwPeriodo0(0).Item("dFechaFin")
                    periodo = "1 " & MonthString(rwPeriodo0(0).Item("iMes")).ToUpper & " AL " & Fechafin.Day & " " & MonthString(rwPeriodo0(0).Item("iMes")).ToUpper & " " & rwPeriodo0(0).Item("iEjercicio")
                    iejercicio = (rwPeriodo0(0).Item("iEjercicio"))
                    fecha = MonthString(rwPeriodo0(0).Item("iMes")).ToUpper
                    hoja.Cell(8, 1).Style.Font.SetBold(True)
                    hoja.Cell(8, 1).Style.NumberFormat.Format = "@"
                    hoja.Cell(8, 1).Value = periodo
                    hoja.Cell(8, 1).Style.Font.FontSize = 12

                End If
                Dim formula1 As String
                For x As Integer = 0 To dtgDatos.Rows.Count - 1

                    ' hoja.Cell(filaExcel + x, 4).Style.NumberFormat.Format = "@"

                    hoja.Cell(filaExcel + x, 1).Value = dtgDatos.Rows(x).Cells(3).Value ' NO TRABAJADOR
                    hoja.Cell(filaExcel + x, 2).Value = dtgDatos.Rows(x).Cells(10).Value ' EDAD 
                    hoja.Cell(filaExcel + x, 3).Value = dtgDatos.Rows(x).Cells(4).Value ' TRABAJADOR 
                    hoja.Cell(filaExcel + x, 4).Value = dtgDatos.Rows(x).Cells(5).Value 'STATUS/Tipo de nomina
                    hoja.Cell(filaExcel + x, 5).Value = "ECO III"
                    hoja.Cell(filaExcel + x, 6).Value = dtgDatos.Rows(x).Cells(11).FormattedValue 'PUESTO dtgDatos.Rows(x).Cells(11).Value ' CATEGORIA
                    hoja.Cell(filaExcel + x, 7).Value = dtgDatos.Rows(x).Cells(18).Value ' DIAS ABORDO
                    hoja.Cell(filaExcel + x, 8).Value = IIf(dtgDatos.Rows(x).Cells(11).FormattedValue = "OFICIALES EN PRACTICAS: PILOTIN / ASPIRANTE", 0, dtgDatos.Rows(x).Cells(18).Value) ' DIAS DESCANSO
                    hoja.Cell(filaExcel + x, 9).FormulaA1 = IIf(dtgDatos.Rows(x).Cells(11).FormattedValue = "OFICIALES EN PRACTICAS: PILOTIN / ASPIRANTE", "=K" & filaExcel + x, "=K" & filaExcel + x & "/2")  'NOMINA ABORDO
                    hoja.Cell(filaExcel + x, 10).FormulaA1 = IIf(dtgDatos.Rows(x).Cells(11).FormattedValue = "OFICIALES EN PRACTICAS: PILOTIN / ASPIRANTE", 0, "=K" & filaExcel + x & "/2") 'NOMINA DESCANSO
                    hoja.Cell(filaExcel + x, 11).Value = dtgDatos.Rows(x).Cells(15).Value ' SUELDO ORDINARIO
                    hoja.Cell(filaExcel + x, 12).FormulaA1 = "='XURTEP ABORDO'!AE" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AE" & filaExcel + x + 1 & "+'XURTEP ABORDO'!AG" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AG" & filaExcel + x + 1 & "+'XURTEP ABORDO'!AF" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AF" & filaExcel + x + 1 ' INFONAVIT
                    hoja.Cell(filaExcel + x, 13).FormulaA1 = "='XURTEP ABORDO'!AJ" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AJ" & filaExcel + x + 1 'FONACOT
                    hoja.Cell(filaExcel + x, 14).FormulaA1 = "='XURTEP ABORDO'!AH" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AH" & filaExcel + x + 1 ' PENSION ALIMENTICIA
                    hoja.Cell(filaExcel + x, 15).FormulaA1 = "='XURTEP ABORDO'!AI" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AI" & filaExcel + x + 1 'Anticipo SA
                    hoja.Cell(filaExcel + x, 16).FormulaA1 = dtgDatos.Rows(x).Cells(47).Value + CDbl(getDatosNomina(0, dtgDatos.Rows(x).Cells(3).Value, dtgDatos.Rows(x).Cells(18).Value, "prestamoA")) 'Anticipo Asimiliados
                    hoja.Cell(filaExcel + x, 17).Value = dtgDatos.Rows(x).Cells(48).Value + CDbl(getDatosNomina(0, dtgDatos.Rows(x).Cells(3).Value, dtgDatos.Rows(x).Cells(18).Value, "descuentoInfoA")) 'Desc Infona
                    hoja.Cell(filaExcel + x, 18).FormulaA1 = "=K" & filaExcel + x & "-L" & filaExcel + x & "-M" & filaExcel + x & "-N" & filaExcel + x & "-O" & filaExcel + x & "-P" & filaExcel + x & "-Q" & filaExcel + x  'SUELDO ORDINARIO
                    hoja.Cell(filaExcel + x, 19).FormulaA1 = "" '
                    hoja.Cell(filaExcel + x, 20).FormulaA1 = "='XURTEP ABORDO'!AL" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AL" & filaExcel + x + 1  ' XURTEP
                    hoja.Cell(filaExcel + x, 21).FormulaA1 = "=R" & filaExcel + x & "-T" & filaExcel + x 'COMPLEMENTO (ASIM NETOS)
                    formula1 = "='XURTEP ABORDO'!AB" & filaExcel + x + 1 & "+'XURTEP ABORDO'!AC" & filaExcel + x + 1 & "+'XURTEP ABORDO'!AD" & filaExcel + x + 1 & "+'XURTEP ABORDO'!AE" & filaExcel + x + 1 & "+'XURTEP ABORDO'!AF" & filaExcel + x + 1 & "+'XURTEP ABORDO'!AG" & filaExcel + x + 1 & "+'XURTEP ABORDO'!AH" & filaExcel + x + 1 & "+'XURTEP ABORDO'!AI" & filaExcel + x + 1 & "+'XURTEP ABORDO'!AJ" & filaExcel + x + 1
                    formula1 &= "+'XURTEP DESCANSO'!AB" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AC" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AD" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AE" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AF" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AG" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AH" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AI" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AJ" & filaExcel + x + 1
                    hoja.Cell(filaExcel + x, 22).FormulaA1 = formula1
                    hoja.Cell(filaExcel + x, 23).FormulaA1 = ""
                    hoja.Cell(filaExcel + x, 24).FormulaA1 = "2%" '%COMISION
                    hoja.Cell(filaExcel + x, 25).FormulaA1 = "=+((T" & filaExcel + x & "+V" & filaExcel + x & ")*X" & filaExcel + x & ")" 'COMISION XURTEP
                    hoja.Cell(filaExcel + x, 26).FormulaA1 = "=+(P" & filaExcel + x & "+Q" & filaExcel + x & "+U" & filaExcel + x & ")*X" & filaExcel + x  ' COMPLEMENTO COMISION
                    hoja.Cell(filaExcel + x, 27).FormulaA1 = "=+'XURTEP ABORDO'!AN" & filaExcel + x + 1 ' dtgDatos.Rows(x).Cells(55).Value 'IMSS
                    hoja.Cell(filaExcel + x, 28).FormulaA1 = "=+'XURTEP ABORDO'!AO" & filaExcel + x + 1 'dtgDatos.Rows(x).Cells(56).Value 'RCV
                    hoja.Cell(filaExcel + x, 29).FormulaA1 = "=+'XURTEP ABORDO'!AP" & filaExcel + x + 1 ' dtgDatos.Rows(x).Cells(57).Value 'INFONAVIT
                    hoja.Cell(filaExcel + x, 30).FormulaA1 = "=+'XURTEP ABORDO'!AQ" & filaExcel + x + 1 & "+'XURTEP DESCANSO'!AQ" & filaExcel + x + 1   ' dtgDatos.Rows(x).Cells(58).Value 'isn
                    hoja.Cell(filaExcel + x, 31).FormulaA1 = "=AA" & filaExcel + x & "+AB" & filaExcel + x & "+AC" & filaExcel + x & "+AD" & filaExcel + x
                    hoja.Cell(filaExcel + x, 32).FormulaA1 = "=AE" & filaExcel + x 'COSTO SOCIAL 
                    hoja.Cell(filaExcel + x, 33).FormulaA1 = "=+P" & filaExcel + x & "+Q" & filaExcel + x & "+T" & filaExcel + x & "+U" & filaExcel + x & "+V" & filaExcel + x & "+Y" & filaExcel + x & "+AF" & filaExcel + x & "+Z" & filaExcel + x
                    hoja.Cell(filaExcel + x, 34).FormulaA1 = "=+AG" & filaExcel + x & "*0.16" 'IVA 16%
                    hoja.Cell(filaExcel + x, 35).FormulaA1 = ("=+AG" & filaExcel + x & "+AH" & filaExcel + x) 'tOTAL DEPOSUTO

                    hoja.Cell(filaExcel + x, 37).FormulaA1 = dtgDatos.Rows(x).Cells(59).Value
                    hoja.Cell(filaExcel + x, 38).FormulaA1 = "30"
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

                'Tabla de facturas
                Dim sep As Integer = filaExcel + total + 5
                'Style
                hoja.Cell("C" & sep).Style.Font.SetBold(True)
                hoja.Range("C" & sep + 4, "D" & sep + 4).Style.Font.SetBold(True)
                hoja.Range("C" & sep + 8, "D" & sep + 8).Style.Font.SetBold(True)
                hoja.Range("C" & sep + 4, "D" & sep + 4).Style.Fill.BackgroundColor = XLColor.PowderBlue
                hoja.Range("C" & sep + 8, "D" & sep + 8).Style.Fill.BackgroundColor = XLColor.PowderBlue


                hoja.Cell("C" & sep).Value = "ECO III MARINOS"
                hoja.Cell("C" & sep + 2).Value = "DEPOSITO ULAM"
                hoja.Cell("C" & sep + 3).Value = "IVA"
                hoja.Cell("C" & sep + 4).Value = "TOTAL DEPOSITO ULAM"

                hoja.Cell("C" & sep + 6).Value = "DEPOSITO SNEIDER"
                hoja.Cell("C" & sep + 7).Value = "IVA"
                hoja.Cell("C" & sep + 8).Value = "TOTAL DEPOSITO SNEIDER"

                hoja.Cell("D" & sep + 2).FormulaA1 = "=T" & filaExcel + total & "+V" & filaExcel + total & "+Y" & filaExcel + total & "+AF" & filaExcel + total
                hoja.Cell("D" & sep + 3).FormulaA1 = "=D" & sep + 2 & "*16%"
                hoja.Cell("D" & sep + 4).FormulaA1 = "=D" & sep + 2 & "+D" & sep + 3

                hoja.Cell("D" & sep + 6).FormulaA1 = "=P" & filaExcel + total & "+Q" & filaExcel + total & "+U" & filaExcel + total & "+Z" & filaExcel + total
                hoja.Cell("D" & sep + 7).FormulaA1 = "=D" & sep + 6 & "*16%"
                hoja.Cell("D" & sep + 8).FormulaA1 = "=D" & sep + 6 & "+D" & sep + 7

                hoja.Cell("D" & sep + 10).FormulaA1 = "=D" & sep + 4 & "+D" & sep + 8

                Dim cuenta, banco, clabe As String

                '<<<<<<<<<<<<<<ASIMILADOS >>>>>>>>>
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
                    hoja4.Cell(filaExcel, 4).FormulaA1 = "=MARINOS!U" & filatmp 'Asimilado
                    hoja4.Cell(filaExcel, 5).Value = dtgDatos.Rows(x).Cells(8).Value '# Afiliacion IMSS
                    hoja4.Cell(filaExcel, 6).Value = dtgDatos.Rows(x).Cells(18).Value 'Dias Trabajandos
                    hoja4.Cell(filaExcel, 7).Value = banco
                    hoja4.Cell(filaExcel, 8).Value = cuenta ' IIf(cuenta = 0, "", cuenta)
                    hoja4.Cell(filaExcel, 9).Value = clabe
                    hoja4.Cell(filaExcel, 10).Value = "SIN TARJT" ' Tarjeta
                    hoja4.Cell(filaExcel, 11).Value = dtgDatos.Rows(x).Cells(6).Value ' RFC
                    hoja4.Cell(filaExcel, 12).Value = dtgDatos.Rows(x).Cells(7).Value ' CURP

                    filaExcel = filaExcel + 1
                    filatmp = filatmp + 1
                Next x


                '<<<<<<<<<<<<<<<RESUMEN>>>>>>>>>>>>>>>>>>
                recorrerFilasColumnas(hoja5, 5, dtgDatos.Rows.Count + 30, 10, "clear")
                filaExcel = 5
                filatmp = 11
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
                    hoja5.Cell(filaExcel, 10).FormulaA1 = "=MARINOS!T" & filatmp ' XURTEP
                    hoja5.Cell(filaExcel, 11).FormulaA1 = "=MARINOS!U" & filatmp ' ASIMILADOS


                    filaExcel = filaExcel + 1
                    filatmp = filatmp + 1

                    pgbProgreso.Value += 1
                    Application.DoEvents()

                Next x


                'Formulas
                hoja5.Range(filaExcel + 2, 10, filaExcel + 4, 11).Style.Font.SetBold(True)
                hoja5.Cell(filaExcel + 2, 10).FormulaA1 = "=SUM(J5:J" & filaExcel & ")"
                hoja5.Cell(filaExcel + 2, 11).FormulaA1 = "=SUM(K5:K" & filaExcel & ")"

                '<<<<<<<<<<<<<<<Pension Alimenticia>>>>>>>>>>>>>>>>
                filaExcel = 2
                filatmp = 12
                inicio = 0
                Dim contador As Integer = 0
                Dim contadorpensionado As Integer
                Dim contadorpensionadofinal As Integer

                Dim TotalPercepciones As Double
                Dim Incapacidad As Double
                Dim isr As Double
                Dim imss As Double
                Dim infonavitvalor As Double
                Dim infonavitanterior As Double
                Dim ajusteinfonavit As Double
                Dim pension As Double
                Dim prestamo As Double
                Dim fonacot As Double
                Dim subsidiogenerado As Double
                Dim subsidioaplicado As Double
                Dim PensionAlimenticia As Double

                'limpiar
                recorrerFilasColumnas(hoja6, filaExcel, dtgDatos.Rows.Count + 50, 50, "clear")


                For x As Integer = 0 To dtgDatos.Rows.Count - 1
                    Dim nombrebenficiario As String
                    Dim porcentaje As String
                    If dtgDatos.Rows(x).Cells(41).Value > 0 Then

                        'Revisa si hay repetidos
                        TotalPercepciones = Double.Parse(IIf(dtgDatos.Rows(x).Cells(33).Value = "", "0", dtgDatos.Rows(x).Cells(33).Value.ToString.Replace(",", "")))
                        Incapacidad = Double.Parse(IIf(dtgDatos.Rows(x).Cells(35).Value = "", "0", dtgDatos.Rows(x).Cells(35).Value))
                        isr = Double.Parse(IIf(dtgDatos.Rows(x).Cells(36).Value = "", "0", dtgDatos.Rows(x).Cells(36).Value))
                        imss = Double.Parse(IIf(dtgDatos.Rows(x).Cells(37).Value = "", "0", dtgDatos.Rows(x).Cells(37).Value))
                        infonavitvalor = Double.Parse(IIf(dtgDatos.Rows(x).Cells(38).Value = "", "0", dtgDatos.Rows(x).Cells(38).Value))
                        infonavitanterior = Double.Parse(IIf(dtgDatos.Rows(x).Cells(39).Value = "", "0", dtgDatos.Rows(x).Cells(39).Value))
                        ajusteinfonavit = Double.Parse(IIf(dtgDatos.Rows(x).Cells(40).Value = "", "0", dtgDatos.Rows(x).Cells(40).Value))
                        ' pension = Double.Parse(IIf(dtgDatos.Rows(x).Cells(41).Value = "", "0", dtgDatos.Rows(x).Cells(41).Value))
                        prestamo = Double.Parse(IIf(dtgDatos.Rows(x).Cells(42).Value = "", "0", dtgDatos.Rows(x).Cells(42).Value))
                        fonacot = Double.Parse(IIf(dtgDatos.Rows(x).Cells(43).Value = "", "0", dtgDatos.Rows(x).Cells(43).Value))
                        subsidioaplicado = Double.Parse(IIf(dtgDatos.Rows(x).Cells(45).Value = "", "0", dtgDatos.Rows(x).Cells(45).Value))

                        hoja6.Cell(filaExcel, 11).Style.NumberFormat.Format = "@"
                        Dim rwPension As DataRow() = nConsulta("SELECT * FROM PensionAlimenticia WHERE fkiIdEmpleadoC=" & dtgDatos.Rows(x).Cells(2).Value & "ORDER BY Nombrebeneficiario")
                        If rwPension Is Nothing = False Then

                            If rwPension.Length > 1 Then

                                For Each pensionado In rwPension
                                    PensionAlimenticia = TotalPercepciones - Incapacidad - isr - imss - infonavitvalor - infonavitanterior - ajusteinfonavit - prestamo - fonacot + subsidioaplicado
                                    pension = Math.Round(PensionAlimenticia * (Double.Parse(pensionado.Item("fPorcentaje")) / 100), 2)
                                    Dim bank As DataRow() = nConsulta("select * from bancos where iIdBanco =" & pensionado.Item("fkiIdBanco"))
                                    If bank Is Nothing = False Then
                                        banco = bank(0).Item("cBANCO")
                                    End If
                                    hoja6.Cell("F" & filaExcel).Style.NumberFormat.Format = "@"
                                    hoja6.Cell("B" & filaExcel).Value = dtgDatos.Rows(x).Cells(4).Value ' Nombre
                                    hoja6.Cell("C" & filaExcel).Value = pensionado.Item("Nombrebeneficiario")
                                    hoja6.Cell("D" & filaExcel).Value = banco
                                    hoja6.Cell("E" & filaExcel).Value = pensionado.Item("Cuenta")
                                    hoja6.Cell("F" & filaExcel).Value = pensionado.Item("Clabe")
                                    hoja6.Cell("G" & filaExcel).Value = pensionado.Item("fPorcentaje")
                                    hoja6.Cell("H" & filaExcel).Value = pension ' MONTO ABORDO
                                    hoja6.Cell("I" & filaExcel).Value = pension ' MONTO DESCANSO
                                    hoja6.Cell("J" & filaExcel).FormulaA1 = "=H" & filaExcel & "+I" & filaExcel 'SUMA



                                    filaExcel = filaExcel + 1
                                Next

                            Else
                                nombrebenficiario = rwPension(0).Item("Nombrebeneficiario")
                                porcentaje = rwPension(0).Item("fPorcentaje")
                                Dim bank As DataRow() = nConsulta("select * from bancos where iIdBanco =" & rwPension(0).Item("fkiIdBanco"))
                                If bank Is Nothing = False Then
                                    banco = bank(0).Item("cBANCO")
                                End If

                                hoja6.Cell("F" & filaExcel).Style.NumberFormat.Format = "@"
                                hoja6.Cell("B" & filaExcel).Value = dtgDatos.Rows(x).Cells(4).Value ' Nombre
                                hoja6.Cell("C" & filaExcel).Value = nombrebenficiario
                                hoja6.Cell("D" & filaExcel).Value = banco
                                hoja6.Cell("E" & filaExcel).Value = rwPension(0).Item("Cuenta")
                                hoja6.Cell("F" & filaExcel).Value = rwPension(0).Item("Clabe")
                                hoja6.Cell("G" & filaExcel).Value = porcentaje
                                hoja6.Cell("H" & filaExcel).FormulaA1 = "='XURTEP ABORDO'!AH" & filatmp ' MONTO ABORDO
                                hoja6.Cell("I" & filaExcel).FormulaA1 = "='XURTEP DESCANSO'!AH" & filatmp ' MONTO DESCANSO
                                hoja6.Cell("J" & filaExcel).FormulaA1 = "=H" & filaExcel & "+I" & filaExcel 'SUMA

                                filaExcel = filaExcel + 1
                            End If

                        End If

                    End If

                    filatmp = filatmp + 1

                Next

                hoja6.Cell(filaExcel + 2, 10).FormulaA1 = "=SUM(J2:J" & filaExcel & ")"

                '<<<<<<<<<<<<<<<<<Xurtep Abordo>>>>>>>>>>>>>>>>>>>>>>>>

                'Limpiar encabezado y relleno
                recorrerFilasColumnas(hoja2, 1, 10, 56, "clear", 13)
                recorrerFilasColumnas(hoja2, 12, dtgDatos.Rows.Count + 30, 56, "clear", 1)


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
                    hoja2.Cell(filaExcel, 32).Value = dtgDatos.Rows(x).Cells(39).Value ' DIF INFONAVIT
                    hoja2.Cell(filaExcel, 33).Value = dtgDatos.Rows(x).Cells(40).Value ' AJUST INFONAVIT
                    hoja2.Cell(filaExcel, 34).Value = dtgDatos.Rows(x).Cells(41).Value ' PENSION ALIMENTICIA
                    hoja2.Cell(filaExcel, 35).Value = dtgDatos.Rows(x).Cells(42).Value ' PRESTAMOS/ANTICIPO NOMINA?
                    hoja2.Cell(filaExcel, 36).Value = dtgDatos.Rows(x).Cells(43).Value ' FONACOT
                    hoja2.Cell(filaExcel, 37).FormulaA1 = "=AB" & filaExcel & "+AC" & filaExcel & "+AD" & filaExcel & "+AE" & filaExcel & "+AF" & filaExcel & "+AG" & filaExcel & "+AH" & filaExcel & "+AI" & filaExcel & "+AJ" & filaExcel
                    hoja2.Cell(filaExcel, 38).Value = dtgDatos.Rows(x).Cells(46).Value ' NETO A PAGAR


                    hoja2.Cell(filaExcel, 40).Value = dtgDatos.Rows(x).Cells(55).Value
                    hoja2.Cell(filaExcel, 41).Value = dtgDatos.Rows(x).Cells(56).Value
                    hoja2.Cell(filaExcel, 42).Value = dtgDatos.Rows(x).Cells(57).Value
                    hoja2.Cell(filaExcel, 43).Value = dtgDatos.Rows(x).Cells(58).Value
                    hoja2.Cell(filaExcel, 44).FormulaA1 = "=SUM(AN" & filaExcel & ":AQ" & filaExcel & ")"
                    hoja2.Cell(filaExcel, 45).FormulaA1 = "=AR" & filaExcel 'dtgDatos.Rows(x).Cells(59).Value

                    filaExcel = filaExcel + 1


                Next x

                'STYLE
                hoja2.Range(filaExcel + 2, 13, filaExcel + 2, 40).Style.Font.SetFontColor(XLColor.Black)
                hoja2.Range(filaExcel + 2, 13, filaExcel + 2, 40).Style.NumberFormat.NumberFormatId = 4
                hoja2.Range(filaExcel + 2, 13, filaExcel + 2, 40).Style.Font.SetBold(True)
                hoja2.Range(filaExcel + 2, 13, filaExcel + 2, 39).Style.Border.BottomBorder = XLBorderStyleValues.Double
                hoja2.Range(filaExcel + 2, 13, filaExcel + 2, 39).Style.Border.TopBorder = XLBorderStyleValues.Double

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
                hoja2.Cell(filaExcel + 2, 38).FormulaA1 = "=SUM(AL12:AL" & filaExcel & ")"
                'hoja2.Cell(filaExcel + 2, 39).FormulaA1 = "=SUM(AM12:AM" & filaExcel & ")"


                limpiarCell(hoja2, 50)


                '<<<<<<<<<<<<<<<xurtep Descanso>>>>>>>>>>>>>>>>>>

                'Limpiar encabezado y relleno
                recorrerFilasColumnas(hoja3, 1, 10, 56, "clear", 13)
                recorrerFilasColumnas(hoja3, 12, dtgDatos.Rows.Count + 30, 56, "clear", 1)

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
                    hoja3.Cell(filaExcel, 32).Value = dtgDatos.Rows(x).Cells(39).Value ' DIF INFONAVIT
                    hoja3.Cell(filaExcel, 33).Value = dtgDatos.Rows(x).Cells(40).Value ' DIF INFONAVIT
                    hoja3.Cell(filaExcel, 34).Value = dtgDatos.Rows(x).Cells(41).Value ' PENSION ALIMENTICIA
                    hoja3.Cell(filaExcel, 35).Value = dtgDatos.Rows(x).Cells(42).Value ' PRESTAMOS/ANTICIPO NOMINA?
                    hoja3.Cell(filaExcel, 36).Value = dtgDatos.Rows(x).Cells(43).Value ' FONACOT
                    hoja3.Cell(filaExcel, 37).FormulaA1 = "=AB" & filaExcel & "+AC" & filaExcel & "+AD" & filaExcel & "+AE" & filaExcel & "+AF" & filaExcel & "+AG" & filaExcel & "+AH" & filaExcel & "+AI" & filaExcel & "+AJ" & filaExcel
                    hoja3.Cell(filaExcel, 38).Value = dtgDatos.Rows(x).Cells(46).Value ' NETO A PAGAR


                    hoja3.Cell(filaExcel, 40).Value = "0.0" 'dtgDatos.Rows(x).Cells(55).Value
                    hoja3.Cell(filaExcel, 41).Value = "0.0" 'dtgDatos.Rows(x).Cells(56).Value
                    hoja3.Cell(filaExcel, 42).Value = "0.0" 'dtgDatos.Rows(x).Cells(57).Value
                    hoja3.Cell(filaExcel, 43).Value = dtgDatos.Rows(x).Cells(58).Value
                    hoja3.Cell(filaExcel, 44).FormulaA1 = "=SUM(AN" & filaExcel & ":AQ" & filaExcel & ")"
                    hoja3.Cell(filaExcel, 45).FormulaA1 = "=AR" & filaExcel

                    filaExcel = filaExcel + 1


                Next x



                'STYLE
                hoja3.Range(filaExcel + 4, 18, filaExcel + 4, 40).Style.Font.SetFontColor(XLColor.Black)
                hoja3.Range(filaExcel + 4, 18, filaExcel + 4, 40).Style.NumberFormat.NumberFormatId = 4
                hoja3.Range(filaExcel + 4, 18, filaExcel + 4, 40).Style.Font.SetBold(True)
                hoja3.Range(filaExcel + 2, 13, filaExcel + 2, 39).Style.Border.BottomBorder = XLBorderStyleValues.Double
                hoja3.Range(filaExcel + 2, 13, filaExcel + 2, 39).Style.Border.TopBorder = XLBorderStyleValues.Double
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
                hoja3.Cell(filaExcel + 2, 38).FormulaA1 = "=SUM(AL12:AL" & filaExcel & ")"
                'hoja3.Cell(filaExcel + 2, 39).FormulaA1 = "=SUM(AM12:AM" & filaExcel & ")"











                'Titulo
                Dim moment As Date = Date.Now()
                Dim month As Integer = moment.Month
                Dim year As Integer = moment.Year

                pnlProgreso.Visible = False
                pnlCatalogo.Enabled = True

                dialogo.FileName = "XURTEP MARINOS " + fecha.ToUpper + " " + iejercicio & " SERIE " & cboserie.SelectedItem
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

    Public Sub limpiarCell(ByVal hoja As IXLWorksheet, ByVal celda As Integer) ', ByVal fila As Integer, ByVal filatotal As Integer)

        For x As Integer = celda To 200

            'For y As Integer = fila + 1 To filatotal + 20
            '    hoja.Cell(x, y).Clear()

            'Next y
            hoja.Cell(1, x).Clear()
            hoja.Cell(2, x).Clear()
            hoja.Cell(3, x).Clear()
            hoja.Cell(4, x).Clear()
            hoja.Cell(5, x).Clear()
            hoja.Cell(6, x).Clear()
            hoja.Cell(7, x).Clear()
            hoja.Cell(8, x).Clear()
            hoja.Cell(9, x).Clear()
            hoja.Cell(10, x).Clear()
            hoja.Cell(11, x).Clear()
            hoja.Cell(12, x).Clear()
            hoja.Cell(13, x).Clear()
            hoja.Cell(14, x).Clear()
        Next x
    End Sub

    Private Sub layoutTimbrado_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles layoutTimbrado.Click
        '<<<<FUNCION1>>>>
        ExisteEnLista()
        '<<<<<<<<>>>>>>

    End Sub




    Public Function ExisteEnLista(Optional ByVal tipo As String = "sa")

        'Revisa si existen duplicados
        Dim dialogo As New SaveFileDialog()

        Dim filas, filas2 As Integer
        Dim contador As Integer = 0
        Dim pos, pos2 As Integer
        Dim dtgDupl As New DataGridView

        Dim mesperiodo As String
        Dim mesid As String
        Dim fechapagoletra As String


        dtgDupl.ColumnCount = dtgDatos.ColumnCount

        Dim rwPeriodo0 As DataRow() = nConsulta("Select (CONVERT(nvarchar(12),dFechaInicio,103) + ' al ' + CONVERT(nvarchar(12),dFechaFin,103)) as periodo, iMes, iEjercicio, iNumeroPeriodo, iIdPeriodo, dFechaFin  from periodos where iIdPeriodo=" & cboperiodo.SelectedValue)
        If rwPeriodo0 Is Nothing = False Then

            mesperiodo = MonthString(rwPeriodo0(0).Item("iMes")).ToUpper
            mesid = rwPeriodo0(0).Item("iMes")

            fechapagoletra = Date.Parse(rwPeriodo0(0).Item("dFechaFin")).ToLongDateString()

            'fechadepago = rwPeriodo0(0).Item("dFechaFin")
        End If
        dialogo.DefaultExt = "*.xlsx"
        Dim fechita() As String = fechapagoletra.Split(",")
        If tipo = "sa" Then
            dialogo.FileName = fechita(1).ToUpper() & " MARINOS " & IIf(cboTipoNomina.SelectedIndex = 0, "ABORDO", "DESCANSO")
        Else
            dialogo.FileName = "Asimilados " & fechita(1).ToUpper() & " " & IIf(cboTipoNomina.SelectedIndex = 0, "NA", "ND")
        End If
        dialogo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx"

        If dialogo.ShowDialog() = System.Windows.Forms.DialogResult.OK Then



            'Revisa el datagrid
            For filas = 0 To dtgDatos.Rows.Count - 1
                For filas2 = 1 + filas To dtgDatos.Rows.Count - 1
                    'Compara si existe un duplicado
                    If dtgDatos.Rows(filas).Cells(3).Value = dtgDatos.Rows(filas2).Cells(3).Value Then
                        'Selecciona el valor duplicado
                        dtgDatos.Rows(filas2).Selected = True
                        contador = contador + 1

                    End If
                    If filas2 = dtgDatos.Rows.Count Then
                        Exit For
                    End If
                Next
                If filas = dtgDatos.Rows.Count Then
                    Exit Function
                End If
            Next

            Dim path As String = dialogo.FileName
            'Revisa las filas seleccionadas
            For Each Seleccion As DataGridViewRow In dtgDatos.SelectedRows
                'Se guarda en una nueva DataGrind
                dtgDupl.Rows.Add(ObtenerValoresFila(Seleccion))
                'Se quitan los seleccionados 
                dtgDatos.Rows.Remove(Seleccion)
                dtgDupl.ClearSelection()

            Next

            'Se verifica si la nueva tabla tiene datos
            'Se guardaron repetidos
            If dtgDupl.Rows.Count - 1 <= 0 Then

                '1 Nomina
                If tipo = "sa" Then
                    generarLayout2(dtgDatos, path)
                Else
                    generarLayoutAsimilados(dtgDatos, path)
                End If

            Else

                'Nomina A
                If tipo = "sa" Then
                    generarLayout2(dtgDatos, path.Replace(".xlsx", " A.xlsx"))
                Else
                    generarLayoutAsimilados(dtgDatos, path.Replace(".xlsx", " A.xlsx"))
                End If
                'Verfica si en el nuevo datagrid
                If ExisteEnLista2(dtgDupl, path, tipo) = False Then
                    If tipo = "sa" Then
                        generarLayout2(dtgDupl, path.Replace(".xlsx", " B.xlsx"))
                    Else
                        generarLayoutAsimilados(dtgDupl, path.Replace(".xlsx", " B.xlsx"))
                    End If
                End If
            End If

            dtgDupl.Rows.Clear()

            pnlProgreso.Visible = False
            pnlCatalogo.Enabled = True

            llenargrid()

        Else
            MessageBox.Show("No se guardo el archivo", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
    End Function
    Public Function ExisteEnLista2(ByRef dtgTercer As DataGridView, ByVal path As String, Optional ByVal tipo As String = "sa") As Boolean
        Dim filas, filas2 As Integer
        Dim contador As Integer = 0
        Dim pos, pos2 As Integer
        Dim dtgDupl2 As New DataGridView
        Dim MiDataSet As New DataSet()

        dtgDupl2.ColumnCount = dtgTercer.ColumnCount

        For filas = 0 To dtgTercer.Rows.Count - 1
            For filas2 = 1 + filas To dtgTercer.Rows.Count - 1

                If dtgTercer.Rows(filas).Cells(3).Value = dtgTercer.Rows(filas2).Cells(3).Value Then

                    dtgTercer.Rows(filas2).Selected = True
                    contador = contador + 1
                End If
                If filas2 = dtgTercer.Rows.Count Then
                    Exit For
                End If
            Next
            If filas = dtgTercer.Rows.Count Then

                Exit Function
            End If
        Next
        For Each Seleccion As DataGridViewRow In dtgTercer.SelectedRows

            dtgDupl2.Rows.Add(ObtenerValoresFila(Seleccion))
            dtgTercer.Rows.Remove(Seleccion)
            dtgDupl2.ClearSelection()
        Next


        If dtgDupl2.Rows.Count - 1 <= 0 Then
            If tipo = "sa" Then
                generarLayout2(dtgTercer, path.Replace(".xlsx", " B.xlsx"))
            Else
                generarLayoutAsimilados(dtgTercer, path.Replace(".xlsx", " B.xlsx"))
            End If
            Return False

        Else
            ' este seria para Segundo
            If tipo = "sa" Then
                generarLayout2(dtgTercer, path.Replace(".xlsx", " B.xlsx"))
            Else
                generarLayoutAsimilados(dtgTercer, path.Replace(".xlsx", " B.xlsx"))
            End If


            If ExisteEnLista3(dtgDupl2, path, tipo) = False Then
                If tipo = "sa" Then
                    generarLayout2(dtgDupl2, path.Replace(".xlsx", " C.xlsx"))
                Else
                    generarLayoutAsimilados(dtgDupl2, path.Replace(".xlsx", " C.xlsx"))
                End If

            End If

            Return True

        End If

        dtgDupl2.Rows.Clear()


    End Function
    Public Function ExisteEnLista3(ByRef dtgCuarto As DataGridView, ByVal path As String, Optional ByVal tipo As String = "sa") As Boolean
        Dim filas, filas2 As Integer
        Dim contador As Integer = 0
        Dim pos, pos2 As Integer
        Dim dtgDupl3 As New DataGridView
        Dim MiDataSet As New DataSet()

        dtgDupl3.ColumnCount = dtgCuarto.ColumnCount

        For filas = 0 To dtgCuarto.Rows.Count - 1
            For filas2 = 1 + filas To dtgCuarto.Rows.Count - 1

                If dtgCuarto.Rows(filas).Cells(3).Value = dtgCuarto.Rows(filas2).Cells(3).Value Then

                    dtgCuarto.Rows(filas2).Selected = True
                    contador = contador + 1
                End If
                If filas2 = dtgCuarto.Rows.Count Then
                    Exit For
                End If
            Next
            If filas = dtgCuarto.Rows.Count Then

                Exit Function
            End If
        Next
        For Each Seleccion As DataGridViewRow In dtgCuarto.SelectedRows

            dtgDupl3.Rows.Add(ObtenerValoresFila(Seleccion))
            dtgCuarto.Rows.Remove(Seleccion)
            dtgDupl3.ClearSelection()
        Next


        If dtgDupl3.Rows.Count - 1 <= 0 Then
            If tipo = "sa" Then
                generarLayout2(dtgCuarto, path.Replace(".xlsx", " C.xlsx"))
            Else
                generarLayoutAsimilados(dtgCuarto, path.Replace(".xlsx", " C.xlsx"))
            End If
            Return False
            'MsgBox(contador.ToString & " Datos repetidos")
        Else
            ' este seria para un tercero
            If tipo = "sa" Then
                generarLayout2(dtgCuarto, path.Replace(".xlsx", " C.xlsx"))
            Else
                generarLayoutAsimilados(dtgCuarto, path.Replace(".xlsx", " C.xlsx"))
            End If

            If ExisteEnLista4(dtgDupl3, path, tipo) = False Then
                If tipo = "sa" Then
                    generarLayout2(dtgDupl3, path.Replace(".xlsx", " D.xlsx"))
                Else
                    generarLayoutAsimilados(dtgDupl3, path.Replace(".xlsx", " D.xlsx"))
                End If

            End If

            Return True

        End If

        dtgDupl3.Rows.Clear()

    End Function
    Public Function ExisteEnLista4(ByRef dtgCinco As DataGridView, ByVal path As String, Optional ByVal tipo As String = "sa") As Boolean
        Dim filas, filas2 As Integer
        Dim contador As Integer = 0
        Dim pos, pos2 As Integer
        Dim dtgDupl4 As New DataGridView
        Dim MiDataSet As New DataSet()

        dtgDupl4.ColumnCount = dtgCinco.ColumnCount

        For filas = 0 To dtgCinco.Rows.Count - 1
            For filas2 = 1 + filas To dtgCinco.Rows.Count - 1

                If dtgCinco.Rows(filas).Cells(3).Value = dtgCinco.Rows(filas2).Cells(3).Value Then

                    dtgCinco.Rows(filas2).Selected = True
                    contador = contador + 1
                End If
                If filas2 = dtgCinco.Rows.Count Then
                    Exit For
                End If
            Next
            If filas = dtgCinco.Rows.Count Then

                Exit Function
            End If
        Next
        For Each Seleccion As DataGridViewRow In dtgCinco.SelectedRows

            dtgDupl4.Rows.Add(ObtenerValoresFila(Seleccion))
            dtgCinco.Rows.Remove(Seleccion)
            dtgDupl4.ClearSelection()
        Next


        If dtgDupl4.Rows.Count - 1 <= 0 Then
            If tipo = "sa" Then
                generarLayout2(dtgCinco, path.Replace(".xlsx", " D.xlsx"))
            Else
                generarLayoutAsimilados(dtgCinco, path.Replace(".xlsx", " D.xlsx"))
            End If
            Return False

        Else
            ' este seria para un cuarto
            If tipo = "sa" Then
                generarLayout2(dtgCinco, path.Replace(".xlsx", " D.xlsx"))
                generarLayout2(dtgDupl4, path.Replace(".xlsx", " E.xlsx"))
            Else
                generarLayoutAsimilados(dtgCinco, path.Replace(".xlsx", " D.xlsx"))
                generarLayoutAsimilados(dtgDupl4, path.Replace(".xlsx", " E.xlsx"))
            End If

            Return True

        End If

        dtgDupl4.Rows.Clear()

    End Function

    Function generarLayout2(ByVal dtgD As DataGridView, ByVal path As String)
        Try
            Dim ejercicio As String
            Dim mesperiodo As String
            Dim mesid As String
            Dim fechapagoletra As String
            Dim filaExcel As Integer = 2
            Dim dialogo As New SaveFileDialog()

            pnlProgreso.Visible = True
            pnlCatalogo.Enabled = False
            Application.DoEvents()

            pgbProgreso.Minimum = 0
            pgbProgreso.Value = 0
            pgbProgreso.Maximum = dtgD.Rows.Count

            Dim rwPeriodo0 As DataRow() = nConsulta("Select (CONVERT(nvarchar(12),dFechaInicio,103) + ' al ' + CONVERT(nvarchar(12),dFechaFin,103)) as periodo, iMes, iEjercicio, iNumeroPeriodo, iIdPeriodo, dFechaFin  from periodos where iIdPeriodo=" & cboperiodo.SelectedValue)
            If rwPeriodo0 Is Nothing = False Then

                mesperiodo = MonthString(rwPeriodo0(0).Item("iMes")).ToUpper
                mesid = rwPeriodo0(0).Item("iMes")
                ejercicio = rwPeriodo0(0).Item("iEjercicio")
                fechapagoletra = Date.Parse(rwPeriodo0(0).Item("dFechaFin")).ToLongDateString()
                'fechadepago = rwPeriodo0(0).Item("dFechaFin")
            End If

            If dtgD.Rows.Count > 0 Then


                'Abrimos el machote
                Dim ruta As String
                ruta = My.Application.Info.DirectoryPath() & "\Archivos\marinos1.xlsx"

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



                '' filaExcel = 6
                For x As Integer = 0 To dtgD.Rows.Count - 1

                    Dim cuenta, clavebanco, fechainiciorelaboral As String

                    If (dtgD.Rows(x).Cells(3).Value Is Nothing = False) Then
                        Dim rwEmpleado As DataRow() = nConsulta("SELECT * FROM empleadosC where cCodigoEmpleado=" & dtgD.Rows(x).Cells(3).Value)
                        If rwEmpleado Is Nothing = False Then

                            cuenta = rwEmpleado(0).Item("Clabe")
                            fechainiciorelaboral = rwEmpleado(0).Item("dFechaPatrona")

                            Dim rwBanco As DataRow() = nConsulta("SELECT* FROM bancos where iIdBanco=" & rwEmpleado(0).Item("fkiIdBanco"))

                            clavebanco = rwBanco(0).Item("clave")
                        End If

                    End If
                    hoja.Range(2, 1, filaExcel, 1).Style.NumberFormat.Format = "@"
                    hoja.Range(2, 5, filaExcel, 5).Style.NumberFormat.Format = "@"
                    hoja.Range(2, 6, filaExcel, 6).Style.NumberFormat.Format = "@"
                    hoja.Range(2, 26, filaExcel, 26).Style.NumberFormat.Format = "@"


                    If dtgD.Rows(x).Cells(3).Value <> "" Then


                        ''Generales
                        hoja.Cell(filaExcel, 1).Value = dtgD.Rows(x).Cells(3).Value 'No Empleado
                        hoja.Cell(filaExcel, 2).Value = dtgD.Rows(x).Cells(6).Value 'RFC
                        hoja.Cell(filaExcel, 3).Value = dtgD.Rows(x).Cells(4).Value 'Nombre
                        hoja.Cell(filaExcel, 4).Value = dtgD.Rows(x).Cells(7).Value 'CURP
                        hoja.Cell(filaExcel, 5).Value = dtgD.Rows(x).Cells(8).Value 'SSA
                        hoja.Cell(filaExcel, 6).Value = cuenta ' Cuenta Bancaria
                        hoja.Cell(filaExcel, 7).Value = dtgD.Rows(x).Cells(17).Value 'SBC //O 17 SALARIO_COTIZACION
                        hoja.Cell(filaExcel, 8).Value = dtgD.Rows(x).Cells(16).Value 'SDI
                        hoja.Cell(filaExcel, 9).Value = "G0666980109" 'Reg. Patronal 
                        hoja.Cell(filaExcel, 10).Value = "VER" 'Ent. Federativa  
                        hoja.Cell(filaExcel, 11).Value = dtgD.Rows(x).Cells(18).Value 'Días Pagados
                        hoja.Cell(filaExcel, 12).Value = fechainiciorelaboral 'FechaInicioRelaboral
                        hoja.Cell(filaExcel, 13).Value = "3" 'Tipo Contrato 
                        hoja.Cell(filaExcel, 14).Value = ""
                        hoja.Cell(filaExcel, 15).Value = ""  'Sndicalizado
                        hoja.Cell(filaExcel, 16).Value = "1"  'Tipo Jornada
                        hoja.Cell(filaExcel, 17).Value = ""
                        hoja.Cell(filaExcel, 18).Value = "2"  'Tipo Regimen
                        hoja.Cell(filaExcel, 19).Value = ""
                        hoja.Cell(filaExcel, 20).Value = ""
                        hoja.Cell(filaExcel, 21).Value = dtgD.Rows(x).Cells(11).FormattedValue   'Puesto
                        hoja.Cell(filaExcel, 22).Value = "4"  'Riesgo Puesto
                        hoja.Cell(filaExcel, 23).Value = ""
                        hoja.Cell(filaExcel, 24).Value = "5"  'Periodicidad Pago
                        hoja.Cell(filaExcel, 25).Value = ""
                        hoja.Cell(filaExcel, 26).Value = clavebanco  'Banco
                        hoja.Cell(filaExcel, 27).Value = ""
                        hoja.Cell(filaExcel, 28).Value = "" 'Subcontratacion
                        hoja.Cell(filaExcel, 29).Value = IIf(cboTipoNomina.SelectedIndex = 0, "NA", "ND") 'Tipo de Recibo
                        hoja.Cell(filaExcel, 30).Value = mesid ' Mes Pago
                        hoja.Cell(filaExcel, 31).Value = dtgD.Rows(x).Cells(12).FormattedValue ' Buque
                        filaExcel = filaExcel + 1
                    End If
                    pgbProgreso.Value += 1
                    Application.DoEvents()
                Next


                filaExcel = 4
                For x As Integer = 0 To dtgD.Rows.Count - 1

                    'Percepciones
                    hoja2.Cell(filaExcel, 1).Value = dtgD.Rows(x).Cells(6).Value 'RFC
                    hoja2.Cell(filaExcel, 2).Value = dtgD.Rows(x).Cells(4).Value 'Nombre
                    hoja2.Cell(filaExcel, 3).Value = dtgD.Rows(x).Cells(21).Value ' Sueldo Base Gravado 
                    hoja2.Cell(filaExcel, 4).Value = "" ' Sueldo Base Exento 
                    hoja2.Cell(filaExcel, 5).Value = dtgD.Rows(x).Cells(22).Value ' Tiempo Extra Fijo Gravado  
                    hoja2.Cell(filaExcel, 6).Value = dtgD.Rows(x).Cells(23).Value ' Tiempo Extra Fijo Exento  
                    hoja2.Cell(filaExcel, 7).Value = dtgD.Rows(x).Cells(24).Value 'Tiempo Extra Ocasional Gravado 
                    hoja2.Cell(filaExcel, 8).Value = "" ' Tiempo Extra Ocasional Exento
                    hoja2.Cell(filaExcel, 9).Value = dtgD.Rows(x).Cells(25).Value ' Desc. Sem. Obligatorio Gravado
                    hoja2.Cell(filaExcel, 10).Value = "" ' Desc. Sem. Obligatorio Gravado
                    hoja2.Cell(filaExcel, 11).Value = dtgD.Rows(x).Cells(26).Value 'Vac Proporcionales Gravado
                    hoja2.Cell(filaExcel, 12).Value = "" 'Vac Proporcionales Exento
                    hoja2.Cell(filaExcel, 13).Value = dtgD.Rows(x).Cells(27).Value ' Aguinaldo Gravado
                    hoja2.Cell(filaExcel, 14).Value = dtgD.Rows(x).Cells(28).Value ' Aguinaldo Exento
                    hoja2.Cell(filaExcel, 15).Value = dtgD.Rows(x).Cells(30).Value ' Prima Vacional Gravado
                    hoja2.Cell(filaExcel, 16).Value = dtgD.Rows(x).Cells(31).Value ' Prima Vacional Exento
                    hoja2.Cell(filaExcel, 17).Value = ""
                    hoja2.Cell(filaExcel, 18).Value = ""
                    hoja2.Cell(filaExcel, 19).Value = ""
                    hoja2.Cell(filaExcel, 20).Value = ""
                    hoja2.Cell(filaExcel, 21).Value = ""
                    hoja2.Cell(filaExcel, 22).Value = ""
                    hoja2.Cell(filaExcel, 23).Value = ""

                    ''Deducciones
                    hoja3.Cell(filaExcel, 1).Value = dtgD.Rows(x).Cells(6).Value 'RFC
                    hoja3.Cell(filaExcel, 2).Value = dtgD.Rows(x).Cells(4).Value 'Nombre
                    hoja3.Cell(filaExcel, 3).Value = dtgD.Rows(x).Cells(37).Value 'IMSS
                    hoja3.Cell(filaExcel, 4).Value = dtgD.Rows(x).Cells(36).Value 'ISR
                    hoja3.Cell(filaExcel, 5).Value = dtgD.Rows(x).Cells(42).Value 'Prestamo
                    hoja3.Cell(filaExcel, 6).Value = "" 'INCAPACIDAD, DIAS
                    hoja3.Cell(filaExcel, 7).Value = "" ' TIPO
                    hoja3.Cell(filaExcel, 8).Value = dtgD.Rows(x).Cells(35).Value 'IMPORTE
                    hoja3.Cell(filaExcel, 9).Value = dtgD.Rows(x).Cells(41).Value 'PENSION ALIMENTICIA IMPORTE

                    If (dtgD.Rows(x).Cells(38).Value = "") Then
                        hoja3.Cell(filaExcel, 11).Value = dtgD.Rows(x).Cells(38).Value ' INFONAVIT IMPORTE
                    Else
                        hoja3.Cell(filaExcel, 11).Value = validateInfonavit(dtgD.Rows(x).Cells(39).Value, dtgD.Rows(x).Cells(38).Value)
                    End If
                    hoja3.Cell(filaExcel, 12).Value = dtgD.Rows(x).Cells(40).Value 'Ajuste infonavit
                    hoja3.Cell(filaExcel, 13).Value = dtgD.Rows(x).Cells(43).Value ' Fonacot

                    ' ''Otros Pagos
                    'hoja4.Columns("A").Width = 20
                    'hoja4.Columns("B").Width = 20
                    'hoja4.Cell(filaExcel, 1).Value = dtgD.Rows(x).Cells(6).Value ' RFC
                    'hoja4.Cell(filaExcel, 2).Value = dtgD.Rows(x).Cells(4).Value 'NOMBRE
                    'hoja4.Cell(filaExcel, 3).Value = dtgD.Rows(x).Cells(44).Value ' SUBSIDIO IMPORTE
                    'hoja4.Cell(filaExcel, 4).Value = dtgD.Rows(x).Cells(45).Value ' SUBSIDIO CUSADO

                    filaExcel = filaExcel + 1

                Next

                'Se guarda
                libro.SaveAs(path)
                libro = Nothing

                pnlProgreso.Visible = False
                pnlCatalogo.Enabled = True
            Else

                MessageBox.Show("Por favor seleccione al menos una registro para importar.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message.ToString(), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Try
    End Function



    Public Function validateInfonavit(ByVal diferencia As Object, ByVal infonavit As Object) As String
        Dim negativo As Integer = diferencia.ToString.IndexOf("-")
        Dim infonavitcalculado As Double
        If negativo <> -1 Then
            Dim diferenciaInfonavit As Double = diferencia


            infonavitcalculado = CDbl(infonavit) + CDbl(diferencia)
            ' Return infonavitcalculado.ToString()
        Else
            'Return infonavit.ToString
            infonavitcalculado = CDbl(infonavit) + CDbl(diferencia)
        End If

        Return infonavitcalculado.ToString()


    End Function
    Function ObtenerValoresFila(ByVal fila As DataGridViewRow) As String()

        Dim Contenido(dtgDatos.ColumnCount - 1) As String

        For Ndx As Integer = 0 To Contenido.Length - 1
            If Ndx = 0 Then
                Contenido(Ndx) = "1"
            Else
                Contenido(Ndx) = fila.Cells(Ndx).Value
            End If

        Next
        Return Contenido

    End Function

    Private Sub cmdInfonavit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInfonavit.Click
        Try
            Dim filaExcel As Integer = 0
            Dim dialogo As New SaveFileDialog()
            Dim periodo As String

            If dtgDatos.Rows.Count > 0 Then


                Dim ruta As String
                ruta = My.Application.Info.DirectoryPath() & "\Archivos\msexcel.xlsx"

                Dim book As New ClosedXML.Excel.XLWorkbook(ruta)


                Dim libro As New ClosedXML.Excel.XLWorkbook

                book.Worksheet(1).CopyTo(libro, "IAS (93713)")

                Dim hoja As IXLWorksheet = libro.Worksheets(0)

                '<<<<<<<<<<<<<<<<<<<<<<IAS>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                filaExcel = 17

                Dim nombrebuque As String
                Dim inicio As Integer = 0
                Dim contadorexcelbuqueinicial As Integer = 0
                Dim contadorexcelbuquefinal As Integer = 0
                Dim total As Integer = dtgDatos.Rows.Count - 1
                'Dim filatmp As Integer = 13 - 4
                Dim fecha As String

                hoja.Cell(filaExcel + 1, 1).InsertCellsAbove(total + filaExcel + 2)

                recorrerFilasColumnas(hoja, filaExcel, filaExcel + total, 11, "clear", 2)

                For x As Integer = 0 To dtgDatos.Rows.Count - 1


                    Dim cuenta, banco, clabe As String
                    Dim nom, app, apm As String

                    If (dtgDatos.Rows(x).Cells(3).Value Is Nothing = False) Then

                        Dim rwEmpleado As DataRow() = nConsulta("SELECT * FROM empleadosC where cCodigoEmpleado=" & dtgDatos.Rows(x).Cells(3).Value)
                        If rwEmpleado Is Nothing = False Then

                            clabe = rwEmpleado(0).Item("Clabe")
                            cuenta = rwEmpleado(0).Item("NumCuenta")
                            nom = rwEmpleado(0).Item("cNombre")
                            app = rwEmpleado(0).Item("cApellidoP")
                            apm = rwEmpleado(0).Item("cApellidoM")

                            Dim rwBanco As DataRow() = nConsulta("SELECT* FROM bancos where iIdBanco=" & rwEmpleado(0).Item("fkiIdBanco"))

                            banco = rwBanco(0).Item("cBanco")
                        End If

                    End If

                    Dim asimilado As Double
                    hoja.Cell(filaExcel + x, 8).Style.NumberFormat.Format = "@"

                    asimilado = Double.Parse(dtgDatos.Rows(x).Cells(50).Value) + Double.Parse(dtgDatos.Rows(x).Cells(50).Value)


                    hoja.Cell(filaExcel + x, 2).Value = app 'AP PATERNO
                    hoja.Cell(filaExcel + x, 3).Value = apm 'AP MATERNO
                    hoja.Cell(filaExcel + x, 4).Value = nom ' NOMBRE
                    hoja.Cell(filaExcel + x, 5).Value = banco ' BANCO
                    hoja.Cell(filaExcel + x, 6).Value = IIf(cuenta = 0, "SIN CTA", cuenta) 'CUENTA
                    hoja.Cell(filaExcel + x, 7).Value = "SIN TJT" ' TARJETA
                    hoja.Cell(filaExcel + x, 8).Value = clabe ' CLABE BANARIA
                    hoja.Cell(filaExcel + x, 9).Value = asimilado 'ASIMILADOS
                    hoja.Cell(filaExcel + x, 10).Value = dtgDatos.Rows(x).Cells(7).Value 'CURP
                    hoja.Cell(filaExcel + x, 11).Value = dtgDatos.Rows(x).Cells(6).Value 'RFC
                Next x

                hoja.Cell("I" & total + filaExcel + 1).FormulaA1 = "=SUM(I17:I" & total & ")"

                'STYLE
                hoja.Range("A1", "L1").Style.Fill.BackgroundColor = XLColor.BlueGray


                Dim fechacreacion As Date = Date.Now

                dialogo.FileName = "Nomina_93713_msexcel_" & fechacreacion.ToString("ddMMyy")
                dialogo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
                ''  dialogo.ShowDialog()

                If dialogo.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    ' OK button pressed
                    libro.SaveAs(dialogo.FileName)
                    libro = Nothing
                    MessageBox.Show("Archivo generado correctamente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("No se guardo el archivo", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                End If

            End If


        Catch ex As Exception

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
            dsPeriodo.Tables("Tabla").Columns.Add("Tiempo_Extra_Fijo_Gravado")
            dsPeriodo.Tables("Tabla").Columns.Add("Tiempo_Extra_Fijo_Exento")
            dsPeriodo.Tables("Tabla").Columns.Add("Tiempo_Extra_Ocasional")
            dsPeriodo.Tables("Tabla").Columns.Add("Desc_Sem_Obligatorio")
            dsPeriodo.Tables("Tabla").Columns.Add("Vacaciones_proporcionales")
            dsPeriodo.Tables("Tabla").Columns.Add("Aguinaldo_gravado")
            dsPeriodo.Tables("Tabla").Columns.Add("Aguinaldo_exento")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_Aguinaldo")
            dsPeriodo.Tables("Tabla").Columns.Add("Prima_vac_gravado")
            dsPeriodo.Tables("Tabla").Columns.Add("Prima_vac_exento")
            dsPeriodo.Tables("Tabla").Columns.Add("Total_Prima_vac")
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
            sql = "select * from Nomina inner join EmpleadosC on fkiIdEmpleadoC=iIdEmpleadoC"
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
                    fila.Item("Tiempo_Extra_Fijo_Gravado") = rwNominaGuardada(x)("fTExtraFijoGravado").ToString
                    fila.Item("Tiempo_Extra_Fijo_Exento") = rwNominaGuardada(x)("fTExtraFijoExento").ToString
                    fila.Item("Tiempo_Extra_Ocasional") = rwNominaGuardada(x)("fTExtraOcasional").ToString
                    fila.Item("Desc_Sem_Obligatorio") = rwNominaGuardada(x)("fDescSemObligatorio").ToString
                    fila.Item("Vacaciones_proporcionales") = rwNominaGuardada(x)("fVacacionesProporcionales").ToString
                    fila.Item("Aguinaldo_gravado") = rwNominaGuardada(x)("fAguinaldoGravado").ToString
                    fila.Item("Aguinaldo_exento") = rwNominaGuardada(x)("fAguinaldoExento").ToString
                    fila.Item("Total_Aguinaldo") = Math.Round(Double.Parse(rwNominaGuardada(x)("fAguinaldoGravado").ToString) + Double.Parse(rwNominaGuardada(x)("fAguinaldoExento").ToString), 2)
                    fila.Item("Prima_vac_gravado") = rwNominaGuardada(x)("fPrimaVacacionalGravado").ToString
                    fila.Item("Prima_vac_exento") = rwNominaGuardada(x)("fPrimaVacacionalExento").ToString

                    fila.Item("Total_Prima_vac") = Math.Round(Double.Parse(rwNominaGuardada(x)("fPrimaVacacionalGravado").ToString) + Double.Parse(rwNominaGuardada(x)("fPrimaVacacionalExento").ToString), 2)
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
                    fila.Item("Operadora") = rwNominaGuardada(x)("fXurtep").ToString
                    fila.Item("Prestamo_Personal_A") = rwNominaGuardada(x)("fPrestamoPerA").ToString
                    fila.Item("Adeudo_Infonavit_A") = rwNominaGuardada(x)("fAdeudoInfonavitA").ToString
                    fila.Item("Diferencia_Infonavit_A") = rwNominaGuardada(x)("fDiferenciaInfonavitA").ToString
                    fila.Item("Asimilados") = rwNominaGuardada(x)("fAsimilados").ToString
                    fila.Item("Retenciones_Operadora") = rwNominaGuardada(x)("fRetencionOperadora").ToString
                    fila.Item("%_Comisión") = rwNominaGuardada(x)("fPorComision").ToString
                    fila.Item("Comisión_Operadora") = rwNominaGuardada(x)("fComisionXurtep").ToString
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

                ' MessageBox.Show("Datos cargados", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)


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


                        ' MessageBox.Show("Datos cargados", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
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

    Private Sub cmdComision_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdComision.Click
        'Enviar datos a excel
        Dim SQL As String, Alter As Boolean = False

        Dim promotor As String = ""
        Dim filaExcel As Integer = 5
        Dim dialogo As New SaveFileDialog()
        Dim contadorfacturas As Integer
        Dim Operadora As Double
        Dim ISR As Double
        Dim Infonavit As Double
        Dim Pension As Double
        Dim costo As Double
        Dim comision As Double
        Dim retenciones As Double
        Dim Asimilados As Double
        Dim comisionasimilados As Double
        Dim sueldoTMM As Double

        Alter = True
        Try

            SQL = "select departamentos.cNombre,sum(fsalariobase) as salariobase, sum(fOperadora) as Operadora,"
            SQL &= " sum(fRetencionOperadora) as retencion, sum(fComisionOperadora) as Comoperadora,"
            SQL &= " sum(fInfonavit) as infonavit,sum(fInfonavitBanterior) as infonavitanterior, sum(fAjusteInfonavit) as ajusteinfonavit,"
            SQL &= " sum(fPensionAlimenticia) as pensionalimenticia, sum(fIsr) as ISR,CostoSocial"
            SQL &= " from (nomina inner join departamentos on nomina.fkiIdDepartamento=departamentos.iIdDepartamento)"
            SQL &= " inner join (select fkiIdDepartamento,sum(fTotalCostoSocial) as CostoSocial"
            SQL &= " from nomina"
            SQL &= " where fkiIdPeriodo =" & cboperiodo.SelectedValue & " And iEstatusEmpleado =" & cboserie.SelectedIndex & " And iTiponomina = 0"
            SQL &= " group by fkiIdDepartamento"
            SQL &= " ) as CS on departamentos.iIdDepartamento=CS.fkiIdDepartamento"
            SQL &= " where fkiIdPeriodo =" & cboperiodo.SelectedValue & " And iEstatusEmpleado =" & cboserie.SelectedIndex
            SQL &= " group by departamentos.cNombre,CostoSocial"
            SQL &= " order by departamentos.cNombre"

            Dim rwFilas As DataRow() = nConsulta(SQL)

            If rwFilas.Length > 0 Then
                Dim libro As New ClosedXML.Excel.XLWorkbook
                Dim hoja As IXLWorksheet = libro.Worksheets.Add("Nomina")
                'Dim hoja2 As IXLWorksheet = libro.Worksheets.Add("Resumen pago")

                hoja.Column("B").Width = 20
                hoja.Column("C").Width = 15
                hoja.Column("D").Width = 15
                hoja.Column("E").Width = 15
                hoja.Column("F").Width = 15
                hoja.Column("G").Width = 15
                hoja.Column("H").Width = 15
                hoja.Column("I").Width = 15
                hoja.Column("J").Width = 15
                hoja.Column("K").Width = 15
                hoja.Column("L").Width = 15
                hoja.Column("M").Width = 15
                hoja.Column("N").Width = 15
                hoja.Column("O").Width = 15


                hoja.Cell(1, 2).Value = "Comision Nomina"
                hoja.Range(1, 2, 1, 2).Style.Font.SetBold(True)
                hoja.Cell(2, 2).Value = "Fecha:" & Date.Now.ToShortDateString & " " & Date.Now.ToShortTimeString
                hoja.Cell(3, 2).Value = "PERIODO: " & cboperiodo.Text
                hoja.Range(3, 2, 3, 2).Style.Font.SetBold(True)

                'hoja.Cell(3, 2).Value = ":"
                'hoja.Cell(3, 3).Value = ""

                hoja.Range(4, 2, 4, 15).Style.Font.FontSize = 10
                hoja.Range(4, 2, 4, 15).Style.Font.SetBold(True)
                hoja.Range(4, 2, 4, 15).Style.Alignment.WrapText = True
                hoja.Range(4, 2, 4, 15).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                hoja.Range(4, 1, 4, 15).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                'hoja.Range(4, 1, 4, 18).Style.Fill.BackgroundColor = XLColor.BleuDeFrance
                hoja.Range(4, 2, 4, 15).Style.Fill.BackgroundColor = XLColor.FromHtml("#538DD5")
                hoja.Range(4, 2, 4, 15).Style.Font.FontColor = XLColor.FromHtml("#FFFFFF")

                hoja.Range(5, 2, 1000, 26).Style.NumberFormat.NumberFormatId = 4

                'Format = ("$ #,###,##0.00")
                'hoja.Cell(4, 1).Value = "Num"

                hoja.Cell(4, 2).Value = "Barco"
                hoja.Cell(4, 3).Value = "Dispersión"
                hoja.Cell(4, 4).Value = "Costo Social"
                hoja.Cell(4, 5).Value = "Retenciones"
                hoja.Cell(4, 6).Value = "Comisión"
                hoja.Cell(4, 7).Value = "Subtotal Routes"
                hoja.Cell(4, 8).Value = "IVA"
                hoja.Cell(4, 9).Value = "TOTAL"
                hoja.Cell(4, 10).Value = ""

                hoja.Cell(4, 11).Value = "Dispersión"
                hoja.Cell(4, 12).Value = "Comisión"
                hoja.Cell(4, 13).Value = "Subtotal Biryusa"
                hoja.Cell(4, 14).Value = "IVA"
                hoja.Cell(4, 15).Value = "TOTAL"


                filaExcel = 5
                contadorfacturas = 1

                For x As Integer = 0 To rwFilas.Length - 1

                    Operadora = Double.Parse(rwFilas(x)("Operadora"))
                    ISR = Double.Parse(rwFilas(x)("ISR"))
                    Infonavit = Double.Parse(rwFilas(x)("Infonavit"))
                    Pension = Double.Parse(rwFilas(x)("pensionalimenticia"))
                    costo = Double.Parse(rwFilas(x)("CostoSocial"))
                    comision = Double.Parse(rwFilas(x)("Comoperadora"))
                    retenciones = ISR + Infonavit + Pension
                    sueldoTMM = Double.Parse(rwFilas(x)("salariobase"))
                    Asimilados = sueldoTMM - Infonavit - Pension - Operadora
                    comisionasimilados = Asimilados * 0.02



                    'Barco
                    hoja.Cell(filaExcel + x, 2).Value = rwFilas(x)("cNombre")
                    'Dispersion
                    hoja.Cell(filaExcel + x, 3).Value = Operadora

                    'Costo
                    hoja.Cell(filaExcel + x, 4).Value = costo
                    'Retenciones
                    hoja.Cell(filaExcel + x, 5).Value = retenciones
                    'Comision
                    hoja.Cell(filaExcel + x, 6).Value = comision
                    'Subtotal
                    hoja.Cell(filaExcel + x, 7).Value = Operadora + costo + retenciones + comision
                    'IVA
                    hoja.Cell(filaExcel + x, 8).Value = Math.Round((Operadora + costo + retenciones + comision) * 0.16, 2)
                    'TOTAL
                    hoja.Cell(filaExcel + x, 9).FormulaA1 = "=SUM(G" & filaExcel + x & ":H" & filaExcel + x & ")"
                    'nada
                    hoja.Cell(filaExcel + x, 10).Value = ""
                    'Dispersion Asimilados
                    hoja.Cell(filaExcel + x, 11).Value = Asimilados
                    'Comision
                    hoja.Cell(filaExcel + x, 12).Value = comisionasimilados
                    'Subtotal
                    hoja.Cell(filaExcel + x, 13).Value = Asimilados + comisionasimilados
                    'IVA
                    hoja.Cell(filaExcel + x, 14).Value = Math.Round((Asimilados + comisionasimilados) * 0.16, 2)
                    'TOTAL
                    hoja.Cell(filaExcel + x, 15).FormulaA1 = "=SUM(M" & filaExcel + x & ":N" & filaExcel + x & ")"

                Next


                hoja.Cell(filaExcel + rwFilas.Length, 2).FormulaA1 = "=SUM(B" & filaExcel & ":B" & filaExcel + rwFilas.Length - 1 & ")"
                hoja.Cell(filaExcel + rwFilas.Length, 3).FormulaA1 = "=SUM(C" & filaExcel & ":C" & filaExcel + rwFilas.Length - 1 & ")"
                hoja.Cell(filaExcel + rwFilas.Length, 4).FormulaA1 = "=SUM(D" & filaExcel & ":D" & filaExcel + rwFilas.Length - 1 & ")"
                hoja.Cell(filaExcel + rwFilas.Length, 5).FormulaA1 = "=SUM(E" & filaExcel & ":E" & filaExcel + rwFilas.Length - 1 & ")"
                hoja.Cell(filaExcel + rwFilas.Length, 6).FormulaA1 = "=SUM(F" & filaExcel & ":F" & filaExcel + rwFilas.Length - 1 & ")"
                hoja.Cell(filaExcel + rwFilas.Length, 7).FormulaA1 = "=SUM(G" & filaExcel & ":G" & filaExcel + rwFilas.Length - 1 & ")"
                hoja.Cell(filaExcel + rwFilas.Length, 8).FormulaA1 = "=SUM(H" & filaExcel & ":G" & filaExcel + rwFilas.Length - 1 & ")"
                hoja.Cell(filaExcel + rwFilas.Length, 9).FormulaA1 = "=SUM(I" & filaExcel & ":I" & filaExcel + rwFilas.Length - 1 & ")"
                hoja.Cell(filaExcel + rwFilas.Length, 10).FormulaA1 = "=SUM(J" & filaExcel & ":J" & filaExcel + rwFilas.Length - 1 & ")"
                hoja.Cell(filaExcel + rwFilas.Length, 11).FormulaA1 = "=SUM(K" & filaExcel & ":K" & filaExcel + rwFilas.Length - 1 & ")"
                hoja.Cell(filaExcel + rwFilas.Length, 12).FormulaA1 = "=SUM(L" & filaExcel & ":L" & filaExcel + rwFilas.Length - 1 & ")"
                hoja.Cell(filaExcel + rwFilas.Length, 13).FormulaA1 = "=SUM(M" & filaExcel & ":M" & filaExcel + rwFilas.Length - 1 & ")"
                hoja.Cell(filaExcel + rwFilas.Length, 14).FormulaA1 = "=SUM(N" & filaExcel & ":N" & filaExcel + rwFilas.Length - 1 & ")"
                hoja.Cell(filaExcel + rwFilas.Length, 15).FormulaA1 = "=SUM(O" & filaExcel & ":O" & filaExcel + rwFilas.Length - 1 & ")"


                hoja.Range(filaExcel + rwFilas.Length, 2, filaExcel + dtgDatos.Rows.Count, 11).Style.Font.SetBold(True)


                '##### HOJA NUMERO 2 RESUMEN PAGO


                dialogo.DefaultExt = "*.xlsx"
                dialogo.FileName = "Resumen Comision"
                dialogo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
                dialogo.ShowDialog()
                libro.SaveAs(dialogo.FileName)
                'libro.SaveAs("c:\temp\control.xlsx")
                'libro.SaveAs(dialogo.FileName)
                'apExcel.Quit()
                libro = Nothing

                MessageBox.Show("Archivo generado", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            Else
                MessageBox.Show("No hay datos a mostrar", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub cmdReporteInfonavit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReporteInfonavit.Click

    End Sub

    Private Sub cmdResumenInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdResumenInfo.Click
        Dim SQL As String
        Dim filaExcel As Integer = 5
        Dim contador As Integer
        Dim dialogo As New SaveFileDialog()

        Dim Forma As New frmConcentradoInfonavit

        If Forma.ShowDialog = Windows.Forms.DialogResult.OK Then
            SQL = "select iBimestre,iAnio,Calculoinfonavit.fkiIdEmpleadoC,cNombreLargo,Calculoinfonavit.cTipoFactor,Calculoinfonavit.fFactor,Monto,retenido from (calculoinfonavit "
            SQL &= " inner join empleadosC on calculoinfonavit.fkiIdEmpleadoC=empleadosC.iIdEmpleadoC)"
            SQL &= " inner join (select fkiIdEmpleadoC, sum (cantidad) as retenido from (DetalleDescInfonavit "
            SQL &= " inner join empleadosC on DetalleDescInfonavit.fkiIdEmpleadoC=empleadosC.iIdEmpleadoC)"
            SQL &= " where(Numbimestre =" & Forma.gBimestre & " And anio =" & Forma.gAnio & ")"
            SQL &= " group by fkiIdEmpleadoC) as detalle on empleadosC.iIdEmpleadoC=detalle.fkiIdEmpleadoC"
            SQL &= " where(iBimestre = " & Forma.gBimestre & " And iAnio = " & Forma.gAnio & ")"
            SQL &= " order by cnombreLargo"

            Dim rwFilas As DataRow() = nConsulta(SQL)

            If rwFilas.Length > 0 Then
                Dim libro As New ClosedXML.Excel.XLWorkbook
                Dim hoja As IXLWorksheet = libro.Worksheets.Add("Nomina")
                'Dim hoja2 As IXLWorksheet = libro.Worksheets.Add("Resumen pago")

                hoja.Column("B").Width = 15
                hoja.Column("C").Width = 15
                hoja.Column("D").Width = 40
                hoja.Column("E").Width = 15
                hoja.Column("F").Width = 15
                hoja.Column("G").Width = 15
                hoja.Column("H").Width = 15



                hoja.Cell(1, 2).Value = "Concentrado Infonavit"
                hoja.Range(1, 2, 1, 2).Style.Font.SetBold(True)
                hoja.Cell(2, 2).Value = "Fecha:" & Date.Now.ToShortDateString & " " & Date.Now.ToShortTimeString
                hoja.Cell(3, 2).Value = "PERIODO: " & cboperiodo.Text
                hoja.Range(3, 2, 3, 2).Style.Font.SetBold(True)

                'hoja.Cell(3, 2).Value = ":"
                'hoja.Cell(3, 3).Value = ""

                hoja.Range(4, 2, 4, 15).Style.Font.FontSize = 10
                hoja.Range(4, 2, 4, 15).Style.Font.SetBold(True)
                hoja.Range(4, 2, 4, 15).Style.Alignment.WrapText = True
                hoja.Range(4, 2, 4, 15).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                hoja.Range(4, 1, 4, 15).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                'hoja.Range(4, 1, 4, 18).Style.Fill.BackgroundColor = XLColor.BleuDeFrance
                hoja.Range(4, 2, 4, 15).Style.Fill.BackgroundColor = XLColor.FromHtml("#538DD5")
                hoja.Range(4, 2, 4, 15).Style.Font.FontColor = XLColor.FromHtml("#FFFFFF")

                hoja.Range(5, 7, 1000, 8).Style.NumberFormat.NumberFormatId = 4

                'Format = ("$ #,###,##0.00")
                'hoja.Cell(4, 1).Value = "Num"

                hoja.Cell(4, 2).Value = "Año"
                hoja.Cell(4, 3).Value = "Bimestre"
                hoja.Cell(4, 4).Value = "Nombre"
                hoja.Cell(4, 5).Value = "Tipo Factor"
                hoja.Cell(4, 6).Value = "Factor"
                hoja.Cell(4, 7).Value = "Monto Bimestre"
                hoja.Cell(4, 8).Value = "Retenido"



                filaExcel = 5
                contador = 1

                For x As Integer = 0 To rwFilas.Length - 1






                    'Año
                    hoja.Cell(filaExcel + x, 2).Value = rwFilas(x)("iAnio")
                    'bimestre
                    hoja.Cell(filaExcel + x, 3).Value = rwFilas(x)("iBimestre")
                    'nombre
                    hoja.Cell(filaExcel + x, 4).Value = rwFilas(x)("cNombreLargo")
                    'Tipo Factor
                    hoja.Cell(filaExcel + x, 5).Value = rwFilas(x)("cTipoFactor")
                    'Factor
                    hoja.Cell(filaExcel + x, 6).Value = rwFilas(x)("fFactor")
                    'Monto bimestre
                    hoja.Cell(filaExcel + x, 7).Value = rwFilas(x)("Monto")
                    'Retenido
                    hoja.Cell(filaExcel + x, 8).Value = rwFilas(x)("retenido")


                Next




                '##### HOJA NUMERO 2 RESUMEN PAGO


                dialogo.DefaultExt = "*.xlsx"
                dialogo.FileName = "Resumen Infonavit Bimestre " & Forma.gBimestre & " Año " & Forma.gAnio
                dialogo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
                dialogo.ShowDialog()
                libro.SaveAs(dialogo.FileName)
                'libro.SaveAs("c:\temp\control.xlsx")
                'libro.SaveAs(dialogo.FileName)
                'apExcel.Quit()
                libro = Nothing

                MessageBox.Show("Archivo generado", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            Else
                MessageBox.Show("No hay datos a mostrar", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If


        End If
    End Sub

    Private Sub NoCalcularInofnavitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NoCalcularInofnavitToolStripMenuItem.Click
        Try
            Dim iFila As DataGridViewRow = Me.dtgDatos.CurrentRow()
            iFila.Tag = "1"
            iFila.Cells(1).Style.BackColor = Color.Yellow
        Catch ex As Exception

        End Try
    End Sub


    Private Sub ActivarCalculoInfonavitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ActivarCalculoInfonavitToolStripMenuItem.Click
        Try
            Dim iFila As DataGridViewRow = Me.dtgDatos.CurrentRow()
            iFila.Tag = ""
            iFila.Cells(1).Style.BackColor = Color.White
        Catch ex As Exception

        End Try
    End Sub


    Private Sub pnlCatalogo_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pnlCatalogo.Paint

    End Sub
    Private Sub ToolStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked

    End Sub

    Private Sub EditarEmpleadoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditarEmpleadoToolStripMenuItem.Click

    End Sub


    Private Sub cmdReporteMensual_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReporteMensual.Click
        Try
            Dim sql As String
            Dim filaExcel As Integer = 0
            Dim dialogo As New SaveFileDialog()
            Dim periodo, fechadepago As String
            Dim mes, anio As String
            Dim fechapagoletra() As String

            'Dim sueldobase, tef, teo, desc, vacpro, taginaldo, tpvac, tpercepciones, comAsim, prestamo, comXurtep, comComple, costoSocial As Double
            'Dim IMSS, SAR, INFONAVIT, ISN As Double

            'Dim sueldobaseD, tefD, teoD, descD, vacproD, taginaldoD, tpvacD, tpercepcionesD, comAsimD, prestamoD, comXurtepD, comCompleD, costoSocialD As Double
            'Dim IMSSD, SARD, INFONAVITD, ISND As Double
            'Dim amarrados, arboleda, azteca, cedros, miramar, verde, cruz, montserrat, blanca, ciari, janitzio, luis, ignacio, gabriel, diego, jose, grande, creciente, colorada, subsea88 As Integer
            'Dim passavera, margot As Integer

            If dtgDatos.Rows.Count > 0 Then

                'Clonar el XLSX
                Dim ruta As String
                ruta = My.Application.Info.DirectoryPath() & "\Archivos\reportemensual.xlsx"

                Dim book As New ClosedXML.Excel.XLWorkbook(ruta)
                Dim libro As New ClosedXML.Excel.XLWorkbook

                book.Worksheet(1).CopyTo(libro, "REPORTE")
                Dim hoja As IXLWorksheet = libro.Worksheets(0)


                recorrerFilasColumnas(hoja, 4, 100, 50, "clear")
                hoja.Range("A3", "K3").Style.NumberFormat.Format = "@"

                '<<<<<<<<<<<ECO III>>>>>>>>>>>
                Dim rwPeriodo0 As DataRow() = nConsulta("Select * from periodos where iIdPeriodo=" & cboperiodo.SelectedValue)
                If rwPeriodo0 Is Nothing = False Then
                    periodo = MonthString(rwPeriodo0(0).Item("iMes")).ToUpper & " " & (rwPeriodo0(0).Item("iEjercicio"))


                End If
                Dim rwNominaGuardada As DataRow()
                sql = "EXEC getSumatoriasNomina "
                sql &= cboperiodo.SelectedValue & " ,"
                sql &= cboTipoNomina.SelectedIndex & " ,"
                sql &= cboserie.SelectedIndex
                rwNominaGuardada = nConsulta(sql)

                hoja.Cell(3, 1).Value = "XURTEP MARINOS  ECO III. - Veracruz"
                If rwNominaGuardada Is Nothing = False Then
                    hoja.Cell(3, 2).Value = CInt(rwNominaGuardada(0)("iDiasTrabajados").ToString) * 12 ' Total horas hombre
                    hoja.Cell(3, 3).Value = CInt(rwNominaGuardada(0)("Empleados").ToString) ' Contar cantidad de empleados
                    hoja.Cell(3, 6).Value = CInt(rwNominaGuardada(0)("EdadPromedio").ToString) 'Promedio edad general
                End If

                sql = "EXEC getSumatoriasNomina2 "
                sql &= cboperiodo.SelectedValue & " ,"
                sql &= cboTipoNomina.SelectedIndex & " ,"
                sql &= cboserie.SelectedIndex & " ,"
                sql &= "0" 'WOMAN
                rwNominaGuardada = nConsulta(sql)
                If rwNominaGuardada Is Nothing = False Then
                    If CInt(rwNominaGuardada(0)("xGenero").ToString) > 0 Then
                        hoja.Cell(3, 4).Value = CInt(rwNominaGuardada(0)("xGenero").ToString) 'Numero de mujeres
                        hoja.Cell(3, 10).Value = CInt(rwNominaGuardada(0)("EdadPromedio").ToString)  'Edad promedio  mujeres
                    Else
                        hoja.Cell(3, 4).Value = 0 'Numero de mujeres
                        hoja.Cell(3, 10).Value = 0  'Edad promedio  mujeres
                    End If

                End If

                sql = "EXEC getSumatoriasNomina2 "
                sql &= cboperiodo.SelectedValue & " ,"
                sql &= cboTipoNomina.SelectedIndex & " ,"
                sql &= cboserie.SelectedIndex & " ,"
                sql &= "1" 'MAN
                rwNominaGuardada = nConsulta(sql)
                If rwNominaGuardada Is Nothing = False Then
                    If CInt(rwNominaGuardada(0)("xGenero").ToString) > 0 Then
                        hoja.Cell(3, 5).Value = CInt(rwNominaGuardada(0)("xGenero").ToString) 'Numero de hombres
                        hoja.Cell(3, 8).Value = CInt(rwNominaGuardada(0)("EdadPromedio").ToString)  'Edad promedio 
                    Else
                        hoja.Cell(3, 5).Value = 0 'Numero de hombres
                        hoja.Cell(3, 8).Value = 0 'Edad promedio 
                    End If

                End If

                ''Promedio de años imss

                Dim sumatoriaaños As Integer = 0
                Dim sumatoriaH As Integer = 0
                Dim sumatoriaM As Integer = 0
                Dim count As Integer = 0
                Dim countH As Integer = 0
                Dim countM As Integer = 0
                For x As Integer = 0 To dtgDatos.Rows.Count - 1
                    'sql = "EXEC getFechaIMSSAntiguedad 0,"
                    'sql &= dtgDatos.Rows(x).Cells(2).Value & " ,"
                    'sql &= getDatosNomina(0, dtgDatos.Rows(x).Cells(3).Value, dtgDatos.Rows(x).Cells(18).Value, "idnomina") & " ,"
                    'sql &= "'" & Date.Parse(getDatosNomina(2, dtgDatos.Rows(x).Cells(2).Value, dtgDatos.Rows(x).Cells(18).Value, "fechaantiguedad")) & "' ,"
                    'sql &= "'" & getDatosNomina(2, dtgDatos.Rows(x).Cells(2).Value, dtgDatos.Rows(x).Cells(18).Value, "clave") & "' ,"
                    'sql &= getDatosNomina(3, dtgDatos.Rows(x).Cells(2).Value, dtgDatos.Rows(x).Cells(18).Value, "isexo")
                    'rwNominaGuardada = nConsulta(sql)

                    Dim fechaantiguedad As Date = Date.Parse(getDatosNomina(2, dtgDatos.Rows(x).Cells(2).Value, dtgDatos.Rows(x).Cells(18).Value, "fechaantiguedad"))
                    Dim fechahoy As Date = Date.Now
                    sumatoriaaños = sumatoriaaños + (fechahoy.Year - fechaantiguedad.Year)
                    count = count + 1

                    If getDatosNomina(3, dtgDatos.Rows(x).Cells(2).Value, dtgDatos.Rows(x).Cells(18).Value, "isexo") = 0 Then
                        fechaantiguedad = Date.Parse(getDatosNomina(2, dtgDatos.Rows(x).Cells(2).Value, dtgDatos.Rows(x).Cells(18).Value, "fechaantiguedad"))
                        sumatoriaM = sumatoriaaños + (fechahoy.Year - fechaantiguedad.Year)
                        countM = countM + 1
                    Else
                        fechaantiguedad = Date.Parse(getDatosNomina(2, dtgDatos.Rows(x).Cells(2).Value, dtgDatos.Rows(x).Cells(18).Value, "fechaantiguedad"))
                        sumatoriaH = sumatoriaaños + (fechahoy.Year - fechaantiguedad.Year)
                        countH = countH + 1

                    End If

                Next


                hoja.Cell(3, 7).Value = CInt(sumatoriaaños / count) 'PROMEDIO TOTAL
                If countM <> 0 Then
                    hoja.Cell(3, 9).Value = IIf(countM = 0, 0, CInt(sumatoriaM / countM)) 'PROMEDIO TOTAL MUJER
                End If
                If countH <> 0 Then
                    hoja.Cell(3, 11).Value = IIf(countH = 0, 0, CInt(sumatoriaH / countH)) 'PROMEDIO TOTAL HOMBRE
                End If


                'Guardar Excel
                'Titulo
                Dim moment As Date = Date.Now()
                Dim month As Integer = moment.Month
                Dim year As Integer = moment.Year


                dialogo.FileName = "Reporte Estadistico ECO III " & periodo
                dialogo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx"

                If dialogo.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

                    libro.SaveAs(dialogo.FileName)
                    libro = Nothing

                    MessageBox.Show("Archivo generado correctamente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("No se guardo el archivo", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                End If
            End If
        Catch ex As Exception

            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Try
    End Sub


    Function getDatosNomina(ByRef tipodato As String, ByRef trabjador As String, ByRef dias As String, ByRef tipe As String) As String

        Dim valorretorno As String
        Dim sql As String

        If tipodato = 0 Then


            sql = "select * from Nomina inner join EmpleadosC on fkiIdEmpleadoC=iIdEmpleadoC"
            sql &= " where Nomina.fkiIdEmpresa = 1 And fkiIdPeriodo = " & cboperiodo.SelectedValue
            sql &= " and Nomina.iEstatus=1 and iEstatusEmpleado=" & cboserie.SelectedIndex
            sql &= " and iTipoNomina=1"
            sql &= " and EmpleadosC.cCodigoEmpleado=" & trabjador
            sql &= " and Nomina.iDiasTrabajados=" & dias
            sql &= " order by " & "Nomina.Buque, cNombreLargo"

            Dim rwNominaGuardada As DataRow() = nConsulta(sql)

            'If rwNominaGuardadaFinal Is Nothing = False Then
            If rwNominaGuardada Is Nothing = False Then

                Select Case tipe
                    Case "idnomina"
                        valorretorno = rwNominaGuardada(0)("iIdNomina").ToString
                    Case "sueldoO"
                        valorretorno = rwNominaGuardada(0)("fSalarioBase").ToString
                    Case "prestamoA"
                        valorretorno = rwNominaGuardada(0)("fPrestamoPerA").ToString
                    Case "Asimilado"
                        valorretorno = rwNominaGuardada(0)("fAsimilados").ToString
                    Case "descuentoInfoA"
                        valorretorno = rwNominaGuardada(0)("fAdeudoInfonavitA").ToString

                End Select


            End If
        ElseIf tipodato = 2 Then
            sql &= " SELECT * FROM IngresoBajaAlta AS imss"
            sql &= " JOIN empleadosC as emp"
            sql &= " ON (IMSS.fkiIdEmpleado=emp.iIdEmpleadoC)"
            sql &= " WHERE imss.fkiIdEmpleado=" & trabjador
            sql &= "  and Clave='A'"
            sql &= " ORDER BY fechabajaimss DESC"
            Dim rwNominaGuardada As DataRow() = nConsulta(sql)

            'If rwNominaGuardadaFinal Is Nothing = False Then
            If rwNominaGuardada Is Nothing = False Then
                Select Case tipe
                    Case "fechaantiguedad"
                        valorretorno = rwNominaGuardada(0)("fecha").ToString
                    Case "clave"
                        valorretorno = rwNominaGuardada(0)("Clave").ToString
                End Select

            End If
        ElseIf tipodato = 3 Then
            sql &= "SELECT * FROM empleadosC "
            sql &= "WHERE iIdEmpleadoC = " & trabjador
            Dim rwNominaGuardada As DataRow() = nConsulta(sql)


            If rwNominaGuardada Is Nothing = False Then

                Select Case tipe
                    Case "isexo"
                        valorretorno = rwNominaGuardada(0)("iSexo").ToString
                End Select
            End If

        End If


        Return (valorretorno)
    End Function

    Private Sub btnAsimilados_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAsimilados.Click
        ExisteEnLista("asim")
    End Sub


    Function generarLayoutAsimilados(ByVal dtgD As DataGridView, ByVal path As String)
        Try

            Dim filaExcel As Integer = 0
            Dim dialogo As New SaveFileDialog()
            Dim periodo As String

            If dtgDatos.Rows.Count > 0 Then

                Dim ruta As String
                ruta = My.Application.Info.DirectoryPath() & "\Archivos\asimilados.xlsm"

                Dim book As New ClosedXML.Excel.XLWorkbook(ruta)
                Dim libro As New ClosedXML.Excel.XLWorkbook
                book.Worksheet(1).CopyTo(libro, "ASIMILADOS")
                Dim hoja As IXLWorksheet = libro.Worksheets(0)

                filaExcel = 2

                Dim app, apm, nom, nombrecompleto, clabe, banco, cuenta As String

                'limpiar
                recorrerFilasColumnas(hoja, filaExcel, dtgD.Rows.Count + 50, 50, "clear")

                hoja.Range(2, 1, dtgD.Rows.Count, 11).Style.Font.FontSize = 10
                hoja.Range(2, 1, dtgD.Rows.Count, 11).Style.Font.SetBold(True)
                hoja.Range(2, 1, dtgD.Rows.Count, 11).Style.Alignment.WrapText = True
                hoja.Range(2, 1, dtgD.Rows.Count, 11).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                hoja.Range(2, 1, dtgD.Rows.Count, 11).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                hoja.Range(2, 1, dtgD.Rows.Count, 11).Style.NumberFormat.Format = "@"

                If cboTipoNomina.SelectedIndex = 1 Then
                    TipoNomina = "0"
                    llenargridD("0")
                Else
                    TipoNomina = "1"

                End If

                For x As Integer = 0 To dtgD.Rows.Count - 1
                    If (dtgD.Rows(x).Cells(3).Value Is Nothing = False) Then

                        Dim empleado As DataRow() = nConsulta("Select * from empleadosC where cCodigoEmpleado=" & dtgD.Rows(x).Cells(3).Value)
                        If empleado Is Nothing = False Then
                            nombrecompleto = empleado(0).Item("cNombre") & " " & empleado(0).Item("cApellidoP") & " " & empleado(0).Item("cApellidoM")
                            app = empleado(0).Item("cApellidoP")
                            apm = empleado(0).Item("cApellidoM")
                            nom = empleado(0).Item("cNombre")
                            clabe = empleado(0).Item("Clabe")
                            cuenta = empleado(0).Item("NumCuenta")
                            Dim bank As DataRow() = nConsulta("select * from bancos where iIdBanco =" & empleado(0).Item("fkiIdBanco"))
                            If bank Is Nothing = False Then
                                banco = IIf(bank(0).Item("cBANCO") = "BBVA BANCOMER", "BANCOMER", bank(0).Item("cBANCO"))
                            End If
                        End If

                        hoja.Cell(filaExcel + x, 1).Value = app 'Paterno
                        hoja.Cell(filaExcel + x, 2).Value = apm 'Materno
                        hoja.Cell(filaExcel + x, 3).Value = nom 'Nombre
                        If dtgD.Rows(x).Cells(50).Value < 1 Then
                            hoja.Cell(filaExcel + x, 4).Value = 0.0
                        Else
                            hoja.Cell(filaExcel + x, 4).Value = CDbl(dtgD.Rows(x).Cells(50).Value * 2) ' + CDbl(getDatosNomina(0, dtgD.Rows(x).Cells(3).Value, dtgD.Rows(x).Cells(18).Value, "Asimilado")) ' asimilados
                        End If

                        hoja.Cell(filaExcel + x, 5).Value = dtgD.Rows(x).Cells(8).Value
                        hoja.Cell(filaExcel + x, 6).Value = dtgD.Rows(x).Cells(18).Value ' Dias Trabjados
                        hoja.Cell(filaExcel + x, 7).Value = banco ' Banco
                        hoja.Cell(filaExcel + x, 8).Value = clabe ' Clabe
                        hoja.Cell(filaExcel + x, 9).Value = cuenta ' Cuenta
                        hoja.Cell(filaExcel + x, 10).Value = dtgD.Rows(x).Cells(7).Value 'Curp   
                        hoja.Cell(filaExcel + x, 11).Value = dtgD.Rows(x).Cells(6).Value ' rfc
                    End If

                Next x

                'dialogo.FileName = "ASIMILADOS SIMPLE "
                'dialogo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
                ' ''  dialogo.ShowDialog()

                'If dialogo.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                '    ' OK button pressed
                '    libro.SaveAs(dialogo.FileName)
                '    libro = Nothing
                '    MessageBox.Show("Archivo generado correctamente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                'Else
                '    MessageBox.Show("No se guardo el archivo", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

                'End If
                libro.SaveAs(path)
                libro = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    
    

    Private Sub NoCalcularCostoSocialToolStripMenuItem_Click_1(sender As System.Object, e As System.EventArgs) Handles NoCalcularCostoSocialToolStripMenuItem.Click
        Try
            Dim iFila As DataGridViewRow = Me.dtgDatos.CurrentRow()
            iFila.Cells(2).Tag = "1"
            iFila.Cells(2).Style.BackColor = Color.Green
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ActivarCalculoCostoSocialToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ActivarCalculoCostoSocialToolStripMenuItem.Click
        Dim iFila As DataGridViewRow = Me.dtgDatos.CurrentRow()
        iFila.Cells(2).Tag = ""
        iFila.Cells(2).Style.BackColor = Color.White
    End Sub

    Private Sub tsbbuscar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbbuscar.Click

        Try
            Dim dialogo As New SaveFileDialog()
            Dim sql As String
            Dim Forma As New frmBuscar
            Dim temp As Integer = 0
            Dim encontro As Boolean = False
            If Forma.ShowDialog = Windows.Forms.DialogResult.OK Then

                For Each fila As DataGridViewRow In dtgDatos.Rows

                    fila.DefaultCellStyle.BackColor = Color.White

                    If fila.Cells.Item(4).Value.ToString().Contains(Forma.txtbuscar.Text.ToUpper) Then
                        fila.DefaultCellStyle.BackColor = Color.Yellow
                        encontro = True
                        temp = temp + 1

                    End If
                Next

            End If

            If encontro = False Then
                MsgBox("No se encontro nada")
            Else
                MsgBox("Se encontrarón " & temp & " Registro")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub cmdAcumuladoMary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAcumuladoMary.Click
        Try
        Dim SQL As String
        Dim filaExcel As Integer = 0
        Dim filatmp As Integer = 0
        Dim dialogo As New SaveFileDialog()
        Dim periodo, iejercicio, iMes As String

        Dim ruta As String

        Dim rwPeriodo0 As DataRow() = nConsulta("Select * from periodos where iIdPeriodo=" & cboperiodo.SelectedValue)
        If rwPeriodo0 Is Nothing = False Then
            Dim Fechafin As Date = rwPeriodo0(0).Item("dFechaFin")
            periodo = "1 " & MonthString(rwPeriodo0(0).Item("iMes")).ToUpper & " AL " & Fechafin.Day & " " & MonthString(rwPeriodo0(0).Item("iMes")).ToUpper & " " & rwPeriodo0(0).Item("iEjercicio")
            iejercicio = (rwPeriodo0(0).Item("iEjercicio"))
            iMes = MonthString(rwPeriodo0(0).Item("iMes")).ToUpper
        End If

        ruta = My.Application.Info.DirectoryPath() & "\Archivos\acumuladosmarinos19.xlsx"

        Dim book As New ClosedXML.Excel.XLWorkbook(ruta)
        Dim libro As New ClosedXML.Excel.XLWorkbook

        book.Worksheet(1).CopyTo(libro, iMes)

        Dim hoja As IXLWorksheet = libro.Worksheets(0)


        filaExcel = 2
        Dim nombrebuque As String
        Dim inicio As Integer = 0
        Dim contadorexcelbuqueinicial As Integer = 0
        Dim contadorexcelbuquefinal As Integer = 0
        Dim total As Integer = dtgDatos.Rows.Count - 1

        Dim fecha As String

        '<<<<<<<<<<<<<<<<<<<<<<<<<<Marinos>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            recorrerFilasColumnas(hoja, 2, dtgDatos.Rows.Count + 10, 40, "clear")

          
        SQL = "EXEC acumuladosMarinos " & cboperiodo.SelectedValue
        Dim rwFilas As DataRow() = nConsulta(SQL)

        If rwFilas.Length > 0 Then

                For x As Integer = 0 To rwFilas.Length - 1
                    hoja.Range(1, 5, filaExcel + x, 33).Style.NumberFormat.NumberFormatId = 4
                    hoja.Range(1, 1, filaExcel + x, 4).Style.NumberFormat.Format = "@"


                    hoja.Cell(filaExcel + x, 1).Value = rwFilas(x)("cCodigoEmpleado")
                    hoja.Cell(filaExcel + x, 2).Value = rwFilas(x)("cNombreLargo")
                    hoja.Cell(filaExcel + x, 3).Value = rwFilas(x)("cRFC")
                    hoja.Cell(filaExcel + x, 4).Value = rwFilas(x)("ccurp")
                    hoja.Cell(filaExcel + x, 5).Value = rwFilas(x)("fSalarioBase")
                    hoja.Cell(filaExcel + x, 6).Value = rwFilas(x)("fSueldoBruto")
                    hoja.Cell(filaExcel + x, 7).Value = rwFilas(x)("fTExtraFijoGravado")
                    hoja.Cell(filaExcel + x, 8).Value = rwFilas(x)("fTExtraFijoExento")
                    hoja.Cell(filaExcel + x, 9).Value = rwFilas(x)("fTExtraOcasional")
                    hoja.Cell(filaExcel + x, 10).Value = rwFilas(x)("fDescSemObligatorio")
                    hoja.Cell(filaExcel + x, 11).Value = rwFilas(x)("fVacacionesProporcionales")
                    hoja.Cell(filaExcel + x, 12).Value = rwFilas(x)("fAguinaldoGravado")
                    hoja.Cell(filaExcel + x, 13).Value = rwFilas(x)("fAguinaldoExento")
                    hoja.Cell(filaExcel + x, 14).Value = rwFilas(x)("fPrimaVacacionalGravado")
                    hoja.Cell(filaExcel + x, 15).Value = rwFilas(x)("fPrimaVacacionalExento")
                    hoja.Cell(filaExcel + x, 16).Value = rwFilas(x)("fTotalPercepciones")
                    hoja.Cell(filaExcel + x, 17).Value = rwFilas(x)("fTotalPercepcionesISR")
                    hoja.Cell(filaExcel + x, 18).Value = rwFilas(x)("fIncapacidad")
                    hoja.Cell(filaExcel + x, 19).Value = rwFilas(x)("fIsr")
                    hoja.Cell(filaExcel + x, 20).Value = rwFilas(x)("fImss")
                    hoja.Cell(filaExcel + x, 21).Value = rwFilas(x)("fInfonavit")
                    hoja.Cell(filaExcel + x, 22).Value = rwFilas(x)("fInfonavitBanterior")
                    hoja.Cell(filaExcel + x, 23).Value = rwFilas(x)("fAjusteInfonavit")
                    hoja.Cell(filaExcel + x, 24).Value = rwFilas(x)("fPensionAlimenticia")
                    hoja.Cell(filaExcel + x, 25).Value = rwFilas(x)("fPrestamo")
                    hoja.Cell(filaExcel + x, 26).Value = rwFilas(x)("fFonacot")
                    hoja.Cell(filaExcel + x, 27).Value = rwFilas(x)("fSubsidioGenerado")
                    hoja.Cell(filaExcel + x, 28).Value = rwFilas(x)("fSubsidioAplicado")
                    hoja.Cell(filaExcel + x, 29).Value = rwFilas(x)("fXurtep")
                    hoja.Cell(filaExcel + x, 30).Value = rwFilas(x)("fPrestamoPerA")
                    hoja.Cell(filaExcel + x, 31).Value = rwFilas(x)("fAdeudoInfonavitA")
                    hoja.Cell(filaExcel + x, 32).Value = rwFilas(x)("fDiferenciaInfonavitA")
                    hoja.Cell(filaExcel + x, 33).FormulaA1 = "=E" & filaExcel + x & "-(U" & filaExcel + x & "+V" & filaExcel + x & "+W" & filaExcel + x & "+X" & filaExcel + x & "+Y" & filaExcel + x & "+Z" & filaExcel + x & ")-AC" & filaExcel + x & "-AD" & filaExcel + x & "-AE" & filaExcel + x & "-AF" & filaExcel + x & ""

                Next
        End If

        dialogo.DefaultExt = "*.xlsx"
            dialogo.FileName = "Reporte Acumulados  Marinos"
        dialogo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
        dialogo.ShowDialog()
        libro.SaveAs(dialogo.FileName)
        'libro.SaveAs("c:\temp\control.xlsx")
        'libro.SaveAs(dialogo.FileName)
        'apExcel.Quit()
        libro = Nothing

            MessageBox.Show("Archivo generado", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString())
        End Try

    End Sub

   


    Private Sub SoloRegistroACalcularToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SoloRegistroACalcularToolStripMenuItem.Click
        Try
            Dim iFila As DataGridViewRow = Me.dtgDatos.CurrentRow()
            iFila.Cells(3).Tag = "1"
            iFila.Cells(3).Style.BackColor = Color.Brown
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DesactivarSoloRegistroACalcularToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DesactivarSoloRegistroACalcularToolStripMenuItem.Click
        Try
            Dim iFila As DataGridViewRow = Me.dtgDatos.CurrentRow()
            iFila.Cells(3).Tag = ""
            iFila.Cells(3).Style.BackColor = Color.White
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Costo0ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles Costo0ToolStripMenuItem.Click
        Try
            Dim iFila As DataGridViewRow = Me.dtgDatos.CurrentRow()
            iFila.Cells(4).Tag = "1"
            iFila.Cells(4).Style.BackColor = Color.Purple
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DesactivarCostoCeroToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DesactivarCostoCeroToolStripMenuItem.Click
        Try
            Dim iFila As DataGridViewRow = Me.dtgDatos.CurrentRow()
            iFila.Cells(4).Tag = ""
            iFila.Cells(4).Style.BackColor = Color.White
        Catch ex As Exception

        End Try
    End Sub

    Function DiasMes(NumPeriodo As Integer) As Integer
        Try
            Dim sql As String

            Dim FechaInicioPeriodo1 As Date
            Dim dias As Integer
            sql = "select * from periodos where iIdPeriodo= " & NumPeriodo
            Dim rwPeriodo As DataRow() = nConsulta(Sql)
            dias = 0
            If rwPeriodo Is Nothing = False Then
                FechaInicioPeriodo1 = Date.Parse(rwPeriodo(0)("dFechaInicio"))

                dias = DateTime.DaysInMonth(Year(FechaInicioPeriodo1), Month(FechaInicioPeriodo1))





            End If
            Return dias

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 0
        End Try
    End Function


    Private Sub RegistroTotalDiasToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles RegistroTotalDiasToolStripMenuItem.Click
        Try
            Dim iFila As DataGridViewRow = Me.dtgDatos.CurrentRow()
            iFila.Cells(5).Tag = "1"
            iFila.Cells(5).Style.BackColor = Color.Purple
            'chkCalSoloMarcados.Checked = True

        Catch ex As Exception

        End Try
    End Sub

    Private Sub DesactivarRegistroTotalDiasToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DesactivarRegistroTotalDiasToolStripMenuItem.Click
        Try
            Dim iFila As DataGridViewRow = Me.dtgDatos.CurrentRow()
            iFila.Cells(5).Tag = ""
            iFila.Cells(5).Style.BackColor = Color.White
        Catch ex As Exception

        End Try
    End Sub

    Private Sub calcularcostosocial(posicion1 As Integer, dias As Integer)
        Dim Sql = "select * from puestos inner join costosocial on puestos.iidPuesto= costosocial.fkiIdPuesto where puestos.cnombre='" & dtgDatos.Rows(posicion1).Cells(11).FormattedValue & "' and anio=" & aniocostosocial
        Dim rwCostoSocial As DataRow() = nConsulta(Sql)
        If rwCostoSocial Is Nothing = False Then
            If dtgDatos.Rows(Posicion1).Cells(10).Value >= 55 Then

                If dtgDatos.Rows(Posicion1).Cells(5).Tag = "" Then
                    'verificar los dias del mes
                    dtgDatos.Rows(Posicion1).Cells(55).Value = Math.Round(Double.Parse(rwCostoSocial(0)("imsstopado")) / 30 * dias, 2)
                    dtgDatos.Rows(Posicion1).Cells(56).Value = Math.Round(Double.Parse(rwCostoSocial(0)("RCVtopado")) / 30 * dias, 2)
                    dtgDatos.Rows(Posicion1).Cells(57).Value = Math.Round(Double.Parse(rwCostoSocial(0)("infonavittopado")) / 30 * dias, 2)
                    dtgDatos.Rows(Posicion1).Cells(58).Value = Math.Round(Double.Parse(dtgDatos.Rows(Posicion1).Cells(33).Value) * 0.03 + (Double.Parse(dtgDatos.Rows(Posicion1).Cells(33).Value) * 0.03 * 0.33), 2)
                    dtgDatos.Rows(Posicion1).Cells(59).Value = Math.Round(Double.Parse(dtgDatos.Rows(Posicion1).Cells(55).Value) + Double.Parse(dtgDatos.Rows(Posicion1).Cells(56).Value) + Double.Parse(dtgDatos.Rows(Posicion1).Cells(57).Value) + Double.Parse(dtgDatos.Rows(Posicion1).Cells(58).Value), 2)
                Else
                    dtgDatos.Rows(Posicion1).Cells(55).Value = Math.Round(Double.Parse(rwCostoSocial(0)("imsstopado")) / 30 * dtgDatos.Rows(Posicion1).Cells(18).Value, 2)
                    dtgDatos.Rows(Posicion1).Cells(56).Value = Math.Round(Double.Parse(rwCostoSocial(0)("RCVtopado")) / 30 * dtgDatos.Rows(Posicion1).Cells(18).Value, 2)
                    dtgDatos.Rows(Posicion1).Cells(57).Value = Math.Round(Double.Parse(rwCostoSocial(0)("infonavittopado")) / 30 * dtgDatos.Rows(Posicion1).Cells(18).Value, 2)
                    dtgDatos.Rows(Posicion1).Cells(58).Value = Math.Round(Double.Parse(dtgDatos.Rows(Posicion1).Cells(33).Value) * 0.03 + (Double.Parse(dtgDatos.Rows(Posicion1).Cells(33).Value) * 0.03 * 0.33), 2)
                    dtgDatos.Rows(Posicion1).Cells(59).Value = Math.Round(Double.Parse(dtgDatos.Rows(Posicion1).Cells(55).Value) + Double.Parse(dtgDatos.Rows(Posicion1).Cells(56).Value) + Double.Parse(dtgDatos.Rows(Posicion1).Cells(57).Value) + Double.Parse(dtgDatos.Rows(Posicion1).Cells(58).Value), 2)
                End If
            Else
                If dtgDatos.Rows(Posicion1).Cells(5).Tag = "" Then
                    dtgDatos.Rows(Posicion1).Cells(55).Value = Math.Round(Double.Parse(rwCostoSocial(0)("imss")) / 30 * dias, 2)
                    dtgDatos.Rows(Posicion1).Cells(56).Value = Math.Round(Double.Parse(rwCostoSocial(0)("RCV")) / 30 * dias, 2)
                    dtgDatos.Rows(Posicion1).Cells(57).Value = Math.Round(Double.Parse(rwCostoSocial(0)("Infonavit")) / 30 * dias, 2)
                    dtgDatos.Rows(Posicion1).Cells(58).Value = Math.Round(Double.Parse(dtgDatos.Rows(Posicion1).Cells(33).Value) * 0.03 + (Double.Parse(dtgDatos.Rows(Posicion1).Cells(33).Value) * 0.03 * 0.33), 2)
                    dtgDatos.Rows(Posicion1).Cells(59).Value = Math.Round(Double.Parse(dtgDatos.Rows(Posicion1).Cells(55).Value) + Double.Parse(dtgDatos.Rows(Posicion1).Cells(56).Value) + Double.Parse(dtgDatos.Rows(Posicion1).Cells(57).Value) + Double.Parse(dtgDatos.Rows(Posicion1).Cells(58).Value), 2)
                Else
                    dtgDatos.Rows(Posicion1).Cells(55).Value = Math.Round(Double.Parse(rwCostoSocial(0)("imss")) / 30 * dtgDatos.Rows(Posicion1).Cells(18).Value, 2)
                    dtgDatos.Rows(Posicion1).Cells(56).Value = Math.Round(Double.Parse(rwCostoSocial(0)("RCV")) / 30 * dtgDatos.Rows(Posicion1).Cells(18).Value, 2)
                    dtgDatos.Rows(Posicion1).Cells(57).Value = Math.Round(Double.Parse(rwCostoSocial(0)("Infonavit")) / 30 * dtgDatos.Rows(Posicion1).Cells(18).Value, 2)
                    dtgDatos.Rows(Posicion1).Cells(58).Value = Math.Round(Double.Parse(dtgDatos.Rows(Posicion1).Cells(33).Value) * 0.03 + (Double.Parse(dtgDatos.Rows(Posicion1).Cells(33).Value) * 0.03 * 0.33), 2)
                    dtgDatos.Rows(Posicion1).Cells(59).Value = Math.Round(Double.Parse(dtgDatos.Rows(Posicion1).Cells(55).Value) + Double.Parse(dtgDatos.Rows(Posicion1).Cells(56).Value) + Double.Parse(dtgDatos.Rows(Posicion1).Cells(57).Value) + Double.Parse(dtgDatos.Rows(Posicion1).Cells(58).Value), 2)
                End If
            End If
            dtgDatos.Rows(Posicion1).Cells(59).Style.BackColor = Color.Chocolate
            dtgDatos.Rows(Posicion1).Cells(59).Tag = "1"
        End If
    End Sub
End Class