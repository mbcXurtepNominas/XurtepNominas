﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConcentradoInfonavit
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
        Me.cbobimestre = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdAceptar = New System.Windows.Forms.Button()
        Me.cmdCancelar = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboAnio = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'cbobimestre
        '
        Me.cbobimestre.FormattingEnabled = True
        Me.cbobimestre.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6"})
        Me.cbobimestre.Location = New System.Drawing.Point(106, 12)
        Me.cbobimestre.Name = "cbobimestre"
        Me.cbobimestre.Size = New System.Drawing.Size(59, 26)
        Me.cbobimestre.TabIndex = 22
        Me.cbobimestre.UseWaitCursor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(101, 18)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "Num bimestre:"
        Me.Label1.UseWaitCursor = True
        '
        'cmdAceptar
        '
        Me.cmdAceptar.Location = New System.Drawing.Point(49, 58)
        Me.cmdAceptar.Name = "cmdAceptar"
        Me.cmdAceptar.Size = New System.Drawing.Size(84, 31)
        Me.cmdAceptar.TabIndex = 24
        Me.cmdAceptar.Text = "Aceptar"
        Me.cmdAceptar.UseVisualStyleBackColor = True
        '
        'cmdCancelar
        '
        Me.cmdCancelar.Location = New System.Drawing.Point(167, 58)
        Me.cmdCancelar.Name = "cmdCancelar"
        Me.cmdCancelar.Size = New System.Drawing.Size(84, 31)
        Me.cmdCancelar.TabIndex = 25
        Me.cmdCancelar.Text = "Cancelar"
        Me.cmdCancelar.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(173, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(37, 18)
        Me.Label2.TabIndex = 27
        Me.Label2.Text = "Año:"
        Me.Label2.UseWaitCursor = True
        '
        'cboAnio
        '
        Me.cboAnio.FormattingEnabled = True
        Me.cboAnio.Items.AddRange(New Object() {"2018", "2019", "2020", "2021", "2022", "2023"})
        Me.cboAnio.Location = New System.Drawing.Point(213, 12)
        Me.cboAnio.Name = "cboAnio"
        Me.cboAnio.Size = New System.Drawing.Size(74, 26)
        Me.cboAnio.TabIndex = 26
        Me.cboAnio.UseWaitCursor = True
        '
        'frmConcentradoInfonavit
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(300, 102)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboAnio)
        Me.Controls.Add(Me.cmdCancelar)
        Me.Controls.Add(Me.cmdAceptar)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbobimestre)
        Me.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmConcentradoInfonavit"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Concentrado Infonavit"
        Me.UseWaitCursor = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cbobimestre As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdAceptar As System.Windows.Forms.Button
    Friend WithEvents cmdCancelar As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboAnio As System.Windows.Forms.ComboBox
End Class
