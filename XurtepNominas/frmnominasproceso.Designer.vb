﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmnominasproceso
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmnominasproceso))
        Me.cmdComision = New System.Windows.Forms.Button()
        Me.cmdReporteInfonavit = New System.Windows.Forms.Button()
        Me.cmdInfonavit = New System.Windows.Forms.Button()
        Me.layoutTimbrado = New System.Windows.Forms.Button()
        Me.pnlProgreso = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.pgbProgreso = New System.Windows.Forms.ProgressBar()
        Me.pnlCatalogo = New System.Windows.Forms.Panel()
        Me.cmdResumenInfo = New System.Windows.Forms.Button()
        Me.cmdSubirDatos = New System.Windows.Forms.Button()
        Me.btnReporte = New System.Windows.Forms.Button()
        Me.cmdexcel = New System.Windows.Forms.Button()
        Me.cmdrecibosA = New System.Windows.Forms.Button()
        Me.cboTipoNomina = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cboserie = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.chkgrupo = New System.Windows.Forms.CheckBox()
        Me.chkinter = New System.Windows.Forms.CheckBox()
        Me.cbobancos = New System.Windows.Forms.ComboBox()
        Me.chkSindicato = New System.Windows.Forms.CheckBox()
        Me.chkAll = New System.Windows.Forms.CheckBox()
        Me.cmdreiniciar = New System.Windows.Forms.Button()
        Me.cmdincidencias = New System.Windows.Forms.Button()
        Me.cmdlayouts = New System.Windows.Forms.Button()
        Me.cmdguardarfinal = New System.Windows.Forms.Button()
        Me.cmdguardarnomina = New System.Windows.Forms.Button()
        Me.cmdcalcular = New System.Windows.Forms.Button()
        Me.dtgDatos = New System.Windows.Forms.DataGridView()
        Me.cmdverdatos = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboperiodo = New System.Windows.Forms.ComboBox()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.tsbEmpleados = New System.Windows.Forms.ToolStripButton()
        Me.tsbPeriodos = New System.Windows.Forms.ToolStripButton()
        Me.tsbpuestos = New System.Windows.Forms.ToolStripButton()
        Me.tsbdeptos = New System.Windows.Forms.ToolStripButton()
        Me.tsbImportar = New System.Windows.Forms.ToolStripButton()
        Me.tsbLayout = New System.Windows.Forms.ToolStripButton()
        Me.tsbIEmpleados = New System.Windows.Forms.ToolStripButton()
        Me.cMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.EliminarDeLaListaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AgregarTrabajadoresToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EditarEmpleadoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NoCalcularInofnavitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ActivarCalculoInfonavitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.pnlProgreso.SuspendLayout()
        Me.pnlCatalogo.SuspendLayout()
        CType(Me.dtgDatos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip1.SuspendLayout()
        Me.cMenu.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdComision
        '
        Me.cmdComision.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdComision.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdComision.Location = New System.Drawing.Point(344, 520)
        Me.cmdComision.Name = "cmdComision"
        Me.cmdComision.Size = New System.Drawing.Size(99, 27)
        Me.cmdComision.TabIndex = 42
        Me.cmdComision.Text = "Comisión"
        Me.cmdComision.UseVisualStyleBackColor = True
        '
        'cmdReporteInfonavit
        '
        Me.cmdReporteInfonavit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdReporteInfonavit.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReporteInfonavit.Location = New System.Drawing.Point(239, 520)
        Me.cmdReporteInfonavit.Name = "cmdReporteInfonavit"
        Me.cmdReporteInfonavit.Size = New System.Drawing.Size(99, 27)
        Me.cmdReporteInfonavit.TabIndex = 41
        Me.cmdReporteInfonavit.Text = "Saldo Infonavit"
        Me.cmdReporteInfonavit.UseVisualStyleBackColor = True
        '
        'cmdInfonavit
        '
        Me.cmdInfonavit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdInfonavit.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdInfonavit.Location = New System.Drawing.Point(121, 520)
        Me.cmdInfonavit.Name = "cmdInfonavit"
        Me.cmdInfonavit.Size = New System.Drawing.Size(112, 27)
        Me.cmdInfonavit.TabIndex = 40
        Me.cmdInfonavit.Text = "Reporte Infonavit"
        Me.cmdInfonavit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdInfonavit.UseVisualStyleBackColor = True
        '
        'layoutTimbrado
        '
        Me.layoutTimbrado.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.layoutTimbrado.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.layoutTimbrado.Location = New System.Drawing.Point(3, 520)
        Me.layoutTimbrado.Name = "layoutTimbrado"
        Me.layoutTimbrado.Size = New System.Drawing.Size(112, 27)
        Me.layoutTimbrado.TabIndex = 39
        Me.layoutTimbrado.Text = "Layout Timbrado"
        Me.layoutTimbrado.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.layoutTimbrado.UseVisualStyleBackColor = True
        '
        'pnlProgreso
        '
        Me.pnlProgreso.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.pnlProgreso.Controls.Add(Me.Label2)
        Me.pnlProgreso.Controls.Add(Me.pgbProgreso)
        Me.pnlProgreso.Location = New System.Drawing.Point(454, 240)
        Me.pnlProgreso.Name = "pnlProgreso"
        Me.pnlProgreso.Size = New System.Drawing.Size(449, 84)
        Me.pnlProgreso.TabIndex = 38
        Me.pnlProgreso.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(154, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(135, 19)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Procesando recibos"
        '
        'pgbProgreso
        '
        Me.pgbProgreso.Location = New System.Drawing.Point(17, 12)
        Me.pgbProgreso.Name = "pgbProgreso"
        Me.pgbProgreso.Size = New System.Drawing.Size(413, 30)
        Me.pgbProgreso.TabIndex = 0
        '
        'pnlCatalogo
        '
        Me.pnlCatalogo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlCatalogo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlCatalogo.Controls.Add(Me.cmdResumenInfo)
        Me.pnlCatalogo.Controls.Add(Me.cmdSubirDatos)
        Me.pnlCatalogo.Controls.Add(Me.btnReporte)
        Me.pnlCatalogo.Controls.Add(Me.cmdexcel)
        Me.pnlCatalogo.Controls.Add(Me.cmdrecibosA)
        Me.pnlCatalogo.Controls.Add(Me.cboTipoNomina)
        Me.pnlCatalogo.Controls.Add(Me.Label4)
        Me.pnlCatalogo.Controls.Add(Me.cboserie)
        Me.pnlCatalogo.Controls.Add(Me.Label3)
        Me.pnlCatalogo.Controls.Add(Me.chkgrupo)
        Me.pnlCatalogo.Controls.Add(Me.chkinter)
        Me.pnlCatalogo.Controls.Add(Me.cbobancos)
        Me.pnlCatalogo.Controls.Add(Me.chkSindicato)
        Me.pnlCatalogo.Controls.Add(Me.chkAll)
        Me.pnlCatalogo.Controls.Add(Me.cmdreiniciar)
        Me.pnlCatalogo.Controls.Add(Me.cmdincidencias)
        Me.pnlCatalogo.Controls.Add(Me.cmdlayouts)
        Me.pnlCatalogo.Controls.Add(Me.cmdguardarfinal)
        Me.pnlCatalogo.Controls.Add(Me.cmdguardarnomina)
        Me.pnlCatalogo.Controls.Add(Me.cmdcalcular)
        Me.pnlCatalogo.Controls.Add(Me.dtgDatos)
        Me.pnlCatalogo.Controls.Add(Me.cmdverdatos)
        Me.pnlCatalogo.Controls.Add(Me.Label1)
        Me.pnlCatalogo.Controls.Add(Me.cboperiodo)
        Me.pnlCatalogo.Location = New System.Drawing.Point(0, 57)
        Me.pnlCatalogo.Name = "pnlCatalogo"
        Me.pnlCatalogo.Size = New System.Drawing.Size(1357, 451)
        Me.pnlCatalogo.TabIndex = 37
        '
        'cmdResumenInfo
        '
        Me.cmdResumenInfo.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdResumenInfo.Location = New System.Drawing.Point(1171, 33)
        Me.cmdResumenInfo.Name = "cmdResumenInfo"
        Me.cmdResumenInfo.Size = New System.Drawing.Size(147, 28)
        Me.cmdResumenInfo.TabIndex = 30
        Me.cmdResumenInfo.Text = "Concentrado Infonavit"
        Me.cmdResumenInfo.UseVisualStyleBackColor = True
        '
        'cmdSubirDatos
        '
        Me.cmdSubirDatos.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSubirDatos.Location = New System.Drawing.Point(1051, 33)
        Me.cmdSubirDatos.Name = "cmdSubirDatos"
        Me.cmdSubirDatos.Size = New System.Drawing.Size(103, 28)
        Me.cmdSubirDatos.TabIndex = 29
        Me.cmdSubirDatos.Text = "Subir datos"
        Me.cmdSubirDatos.UseVisualStyleBackColor = True
        '
        'btnReporte
        '
        Me.btnReporte.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReporte.Location = New System.Drawing.Point(942, 33)
        Me.btnReporte.Name = "btnReporte"
        Me.btnReporte.Size = New System.Drawing.Size(103, 28)
        Me.btnReporte.TabIndex = 28
        Me.btnReporte.Text = "Reporte a Excel"
        Me.btnReporte.UseVisualStyleBackColor = True
        '
        'cmdexcel
        '
        Me.cmdexcel.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdexcel.Location = New System.Drawing.Point(1003, 3)
        Me.cmdexcel.Name = "cmdexcel"
        Me.cmdexcel.Size = New System.Drawing.Size(93, 27)
        Me.cmdexcel.TabIndex = 27
        Me.cmdexcel.Text = "Enviar a Excel"
        Me.cmdexcel.UseVisualStyleBackColor = True
        '
        'cmdrecibosA
        '
        Me.cmdrecibosA.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdrecibosA.Location = New System.Drawing.Point(877, 3)
        Me.cmdrecibosA.Name = "cmdrecibosA"
        Me.cmdrecibosA.Size = New System.Drawing.Size(122, 27)
        Me.cmdrecibosA.TabIndex = 26
        Me.cmdrecibosA.Text = "Asimilado Simple"
        Me.cmdrecibosA.UseVisualStyleBackColor = True
        '
        'cboTipoNomina
        '
        Me.cboTipoNomina.FormattingEnabled = True
        Me.cboTipoNomina.Items.AddRange(New Object() {"Abordo", "Descanso"})
        Me.cboTipoNomina.Location = New System.Drawing.Point(437, 2)
        Me.cboTipoNomina.Name = "cboTipoNomina"
        Me.cboTipoNomina.Size = New System.Drawing.Size(134, 27)
        Me.cboTipoNomina.TabIndex = 25
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(395, 7)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(41, 19)
        Me.Label4.TabIndex = 24
        Me.Label4.Text = "Tipo:"
        '
        'cboserie
        '
        Me.cboserie.FormattingEnabled = True
        Me.cboserie.Items.AddRange(New Object() {"A", "B", "C", "D", "E"})
        Me.cboserie.Location = New System.Drawing.Point(330, 3)
        Me.cboserie.Name = "cboserie"
        Me.cboserie.Size = New System.Drawing.Size(59, 27)
        Me.cboserie.TabIndex = 21
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(289, 7)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(45, 19)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Serie:"
        '
        'chkgrupo
        '
        Me.chkgrupo.AutoSize = True
        Me.chkgrupo.BackColor = System.Drawing.Color.Transparent
        Me.chkgrupo.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkgrupo.Location = New System.Drawing.Point(871, 36)
        Me.chkgrupo.Name = "chkgrupo"
        Me.chkgrupo.Size = New System.Drawing.Size(65, 22)
        Me.chkgrupo.TabIndex = 19
        Me.chkgrupo.Text = "Grupo"
        Me.chkgrupo.UseVisualStyleBackColor = False
        '
        'chkinter
        '
        Me.chkinter.AutoSize = True
        Me.chkinter.BackColor = System.Drawing.Color.Transparent
        Me.chkinter.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkinter.Location = New System.Drawing.Point(433, 37)
        Me.chkinter.Name = "chkinter"
        Me.chkinter.Size = New System.Drawing.Size(110, 22)
        Me.chkinter.TabIndex = 18
        Me.chkinter.Text = "Interbancario"
        Me.chkinter.UseVisualStyleBackColor = False
        '
        'cbobancos
        '
        Me.cbobancos.FormattingEnabled = True
        Me.cbobancos.Location = New System.Drawing.Point(541, 33)
        Me.cbobancos.Name = "cbobancos"
        Me.cbobancos.Size = New System.Drawing.Size(252, 27)
        Me.cbobancos.TabIndex = 17
        '
        'chkSindicato
        '
        Me.chkSindicato.AutoSize = True
        Me.chkSindicato.BackColor = System.Drawing.Color.Transparent
        Me.chkSindicato.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSindicato.Location = New System.Drawing.Point(343, 37)
        Me.chkSindicato.Name = "chkSindicato"
        Me.chkSindicato.Size = New System.Drawing.Size(84, 22)
        Me.chkSindicato.TabIndex = 16
        Me.chkSindicato.Text = "Sindicato"
        Me.chkSindicato.UseVisualStyleBackColor = False
        '
        'chkAll
        '
        Me.chkAll.AutoSize = True
        Me.chkAll.BackColor = System.Drawing.Color.Transparent
        Me.chkAll.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAll.Location = New System.Drawing.Point(59, 36)
        Me.chkAll.Name = "chkAll"
        Me.chkAll.Size = New System.Drawing.Size(107, 22)
        Me.chkAll.TabIndex = 15
        Me.chkAll.Text = "Marcar todos"
        Me.chkAll.UseVisualStyleBackColor = False
        '
        'cmdreiniciar
        '
        Me.cmdreiniciar.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdreiniciar.Location = New System.Drawing.Point(1228, 3)
        Me.cmdreiniciar.Name = "cmdreiniciar"
        Me.cmdreiniciar.Size = New System.Drawing.Size(111, 27)
        Me.cmdreiniciar.TabIndex = 14
        Me.cmdreiniciar.Text = "Reiniciar Nomina"
        Me.cmdreiniciar.UseVisualStyleBackColor = True
        '
        'cmdincidencias
        '
        Me.cmdincidencias.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdincidencias.Location = New System.Drawing.Point(1111, 3)
        Me.cmdincidencias.Name = "cmdincidencias"
        Me.cmdincidencias.Size = New System.Drawing.Size(111, 27)
        Me.cmdincidencias.TabIndex = 13
        Me.cmdincidencias.Text = "Excel Incidencias"
        Me.cmdincidencias.UseVisualStyleBackColor = True
        '
        'cmdlayouts
        '
        Me.cmdlayouts.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdlayouts.Location = New System.Drawing.Point(799, 34)
        Me.cmdlayouts.Name = "cmdlayouts"
        Me.cmdlayouts.Size = New System.Drawing.Size(66, 27)
        Me.cmdlayouts.TabIndex = 11
        Me.cmdlayouts.Text = "Layout"
        Me.cmdlayouts.UseVisualStyleBackColor = True
        '
        'cmdguardarfinal
        '
        Me.cmdguardarfinal.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdguardarfinal.Location = New System.Drawing.Point(782, 3)
        Me.cmdguardarfinal.Name = "cmdguardarfinal"
        Me.cmdguardarfinal.Size = New System.Drawing.Size(92, 27)
        Me.cmdguardarfinal.TabIndex = 9
        Me.cmdguardarfinal.Text = "Guardar Final"
        Me.cmdguardarfinal.UseVisualStyleBackColor = True
        '
        'cmdguardarnomina
        '
        Me.cmdguardarnomina.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdguardarnomina.Location = New System.Drawing.Point(715, 3)
        Me.cmdguardarnomina.Name = "cmdguardarnomina"
        Me.cmdguardarnomina.Size = New System.Drawing.Size(63, 27)
        Me.cmdguardarnomina.TabIndex = 8
        Me.cmdguardarnomina.Text = "Guardar"
        Me.cmdguardarnomina.UseVisualStyleBackColor = True
        '
        'cmdcalcular
        '
        Me.cmdcalcular.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdcalcular.Location = New System.Drawing.Point(651, 3)
        Me.cmdcalcular.Name = "cmdcalcular"
        Me.cmdcalcular.Size = New System.Drawing.Size(63, 27)
        Me.cmdcalcular.TabIndex = 7
        Me.cmdcalcular.Text = "Calcular"
        Me.cmdcalcular.UseVisualStyleBackColor = True
        '
        'dtgDatos
        '
        Me.dtgDatos.AllowUserToAddRows = False
        Me.dtgDatos.AllowUserToDeleteRows = False
        Me.dtgDatos.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgDatos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dtgDatos.Location = New System.Drawing.Point(1, 65)
        Me.dtgDatos.Name = "dtgDatos"
        Me.dtgDatos.Size = New System.Drawing.Size(1349, 379)
        Me.dtgDatos.TabIndex = 6
        '
        'cmdverdatos
        '
        Me.cmdverdatos.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdverdatos.Location = New System.Drawing.Point(575, 3)
        Me.cmdverdatos.Name = "cmdverdatos"
        Me.cmdverdatos.Size = New System.Drawing.Size(71, 27)
        Me.cmdverdatos.TabIndex = 5
        Me.cmdverdatos.Text = "Ver datos"
        Me.cmdverdatos.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 19)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Periodo:"
        '
        'cboperiodo
        '
        Me.cboperiodo.FormattingEnabled = True
        Me.cboperiodo.Location = New System.Drawing.Point(72, 3)
        Me.cboperiodo.Name = "cboperiodo"
        Me.cboperiodo.Size = New System.Drawing.Size(191, 27)
        Me.cboperiodo.TabIndex = 3
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbEmpleados, Me.tsbPeriodos, Me.tsbpuestos, Me.tsbdeptos, Me.tsbImportar, Me.tsbLayout, Me.tsbIEmpleados})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1357, 54)
        Me.ToolStrip1.TabIndex = 36
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'tsbEmpleados
        '
        Me.tsbEmpleados.Image = CType(resources.GetObject("tsbEmpleados.Image"), System.Drawing.Image)
        Me.tsbEmpleados.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbEmpleados.Name = "tsbEmpleados"
        Me.tsbEmpleados.Size = New System.Drawing.Size(118, 51)
        Me.tsbEmpleados.Text = "Importar Empleados"
        Me.tsbEmpleados.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'tsbPeriodos
        '
        Me.tsbPeriodos.Image = CType(resources.GetObject("tsbPeriodos.Image"), System.Drawing.Image)
        Me.tsbPeriodos.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbPeriodos.Name = "tsbPeriodos"
        Me.tsbPeriodos.Size = New System.Drawing.Size(106, 51)
        Me.tsbPeriodos.Text = "Importar Períodos"
        Me.tsbPeriodos.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'tsbpuestos
        '
        Me.tsbpuestos.Image = CType(resources.GetObject("tsbpuestos.Image"), System.Drawing.Image)
        Me.tsbpuestos.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbpuestos.Name = "tsbpuestos"
        Me.tsbpuestos.Size = New System.Drawing.Size(101, 51)
        Me.tsbpuestos.Text = "Importar Puestos"
        Me.tsbpuestos.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'tsbdeptos
        '
        Me.tsbdeptos.Image = CType(resources.GetObject("tsbdeptos.Image"), System.Drawing.Image)
        Me.tsbdeptos.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbdeptos.Name = "tsbdeptos"
        Me.tsbdeptos.Size = New System.Drawing.Size(96, 51)
        Me.tsbdeptos.Text = "Importar deptos"
        Me.tsbdeptos.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'tsbImportar
        '
        Me.tsbImportar.Image = CType(resources.GetObject("tsbImportar.Image"), System.Drawing.Image)
        Me.tsbImportar.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbImportar.Name = "tsbImportar"
        Me.tsbImportar.Size = New System.Drawing.Size(70, 51)
        Me.tsbImportar.Text = "Incidencias"
        Me.tsbImportar.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.tsbImportar.ToolTipText = "Importar incidencias"
        '
        'tsbLayout
        '
        Me.tsbLayout.AutoSize = False
        Me.tsbLayout.Image = CType(resources.GetObject("tsbLayout.Image"), System.Drawing.Image)
        Me.tsbLayout.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbLayout.Name = "tsbLayout"
        Me.tsbLayout.Size = New System.Drawing.Size(90, 51)
        Me.tsbLayout.Text = "Layouts"
        Me.tsbLayout.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.tsbLayout.Visible = False
        '
        'tsbIEmpleados
        '
        Me.tsbIEmpleados.Image = CType(resources.GetObject("tsbIEmpleados.Image"), System.Drawing.Image)
        Me.tsbIEmpleados.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbIEmpleados.Name = "tsbIEmpleados"
        Me.tsbIEmpleados.Size = New System.Drawing.Size(69, 51)
        Me.tsbIEmpleados.Text = "Empleados"
        Me.tsbIEmpleados.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'cMenu
        '
        Me.cMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EliminarDeLaListaToolStripMenuItem, Me.AgregarTrabajadoresToolStripMenuItem, Me.EditarEmpleadoToolStripMenuItem, Me.NoCalcularInofnavitToolStripMenuItem, Me.ActivarCalculoInfonavitToolStripMenuItem})
        Me.cMenu.Name = "cMenu"
        Me.cMenu.Size = New System.Drawing.Size(205, 114)
        '
        'EliminarDeLaListaToolStripMenuItem
        '
        Me.EliminarDeLaListaToolStripMenuItem.Name = "EliminarDeLaListaToolStripMenuItem"
        Me.EliminarDeLaListaToolStripMenuItem.Size = New System.Drawing.Size(204, 22)
        Me.EliminarDeLaListaToolStripMenuItem.Text = "Eliminar de la Lista"
        '
        'AgregarTrabajadoresToolStripMenuItem
        '
        Me.AgregarTrabajadoresToolStripMenuItem.Name = "AgregarTrabajadoresToolStripMenuItem"
        Me.AgregarTrabajadoresToolStripMenuItem.Size = New System.Drawing.Size(204, 22)
        Me.AgregarTrabajadoresToolStripMenuItem.Text = "Agregar Trabajadores"
        '
        'EditarEmpleadoToolStripMenuItem
        '
        Me.EditarEmpleadoToolStripMenuItem.Name = "EditarEmpleadoToolStripMenuItem"
        Me.EditarEmpleadoToolStripMenuItem.Size = New System.Drawing.Size(204, 22)
        Me.EditarEmpleadoToolStripMenuItem.Text = "Editar Empleado"
        '
        'NoCalcularInofnavitToolStripMenuItem
        '
        Me.NoCalcularInofnavitToolStripMenuItem.Name = "NoCalcularInofnavitToolStripMenuItem"
        Me.NoCalcularInofnavitToolStripMenuItem.Size = New System.Drawing.Size(204, 22)
        Me.NoCalcularInofnavitToolStripMenuItem.Text = "No Calcular inofnavit"
        '
        'ActivarCalculoInfonavitToolStripMenuItem
        '
        Me.ActivarCalculoInfonavitToolStripMenuItem.Name = "ActivarCalculoInfonavitToolStripMenuItem"
        Me.ActivarCalculoInfonavitToolStripMenuItem.Size = New System.Drawing.Size(204, 22)
        Me.ActivarCalculoInfonavitToolStripMenuItem.Text = "Activar Calculo Infonavit"
        '
        'frmnominasproceso
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(1357, 553)
        Me.Controls.Add(Me.cmdComision)
        Me.Controls.Add(Me.cmdReporteInfonavit)
        Me.Controls.Add(Me.cmdInfonavit)
        Me.Controls.Add(Me.layoutTimbrado)
        Me.Controls.Add(Me.pnlProgreso)
        Me.Controls.Add(Me.pnlCatalogo)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "frmnominasproceso"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Nominas Proceso"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlProgreso.ResumeLayout(False)
        Me.pnlProgreso.PerformLayout()
        Me.pnlCatalogo.ResumeLayout(False)
        Me.pnlCatalogo.PerformLayout()
        CType(Me.dtgDatos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.cMenu.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdComision As System.Windows.Forms.Button
    Friend WithEvents cmdReporteInfonavit As System.Windows.Forms.Button
    Friend WithEvents cmdInfonavit As System.Windows.Forms.Button
    Friend WithEvents layoutTimbrado As System.Windows.Forms.Button
    Friend WithEvents pnlProgreso As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents pgbProgreso As System.Windows.Forms.ProgressBar
    Friend WithEvents pnlCatalogo As System.Windows.Forms.Panel
    Friend WithEvents cmdResumenInfo As System.Windows.Forms.Button
    Friend WithEvents cmdSubirDatos As System.Windows.Forms.Button
    Friend WithEvents btnReporte As System.Windows.Forms.Button
    Friend WithEvents cmdexcel As System.Windows.Forms.Button
    Friend WithEvents cmdrecibosA As System.Windows.Forms.Button
    Friend WithEvents cboTipoNomina As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboserie As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents chkgrupo As System.Windows.Forms.CheckBox
    Friend WithEvents chkinter As System.Windows.Forms.CheckBox
    Friend WithEvents cbobancos As System.Windows.Forms.ComboBox
    Friend WithEvents chkSindicato As System.Windows.Forms.CheckBox
    Friend WithEvents chkAll As System.Windows.Forms.CheckBox
    Friend WithEvents cmdreiniciar As System.Windows.Forms.Button
    Friend WithEvents cmdincidencias As System.Windows.Forms.Button
    Friend WithEvents cmdlayouts As System.Windows.Forms.Button
    Friend WithEvents cmdguardarfinal As System.Windows.Forms.Button
    Friend WithEvents cmdguardarnomina As System.Windows.Forms.Button
    Friend WithEvents cmdcalcular As System.Windows.Forms.Button
    Friend WithEvents dtgDatos As System.Windows.Forms.DataGridView
    Friend WithEvents cmdverdatos As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboperiodo As System.Windows.Forms.ComboBox
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbEmpleados As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbPeriodos As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbpuestos As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbdeptos As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbImportar As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbLayout As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbIEmpleados As System.Windows.Forms.ToolStripButton
    Friend WithEvents cMenu As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents EliminarDeLaListaToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AgregarTrabajadoresToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EditarEmpleadoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NoCalcularInofnavitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ActivarCalculoInfonavitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
