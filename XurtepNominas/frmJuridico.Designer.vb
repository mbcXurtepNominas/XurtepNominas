<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmJuridico
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmJuridico))
        Me.cmdingreso = New System.Windows.Forms.Button()
        Me.cmdPlanta = New System.Windows.Forms.Button()
        Me.btnAsimilados = New System.Windows.Forms.Button()
        Me.cmdDeterminado = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdingreso
        '
        Me.cmdingreso.Image = CType(resources.GetObject("cmdingreso.Image"), System.Drawing.Image)
        Me.cmdingreso.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdingreso.Location = New System.Drawing.Point(100, 160)
        Me.cmdingreso.Name = "cmdingreso"
        Me.cmdingreso.Size = New System.Drawing.Size(87, 72)
        Me.cmdingreso.TabIndex = 38
        Me.cmdingreso.Text = "S. Ingreso"
        Me.cmdingreso.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdingreso.UseVisualStyleBackColor = True
        '
        'cmdPlanta
        '
        Me.cmdPlanta.Image = CType(resources.GetObject("cmdPlanta.Image"), System.Drawing.Image)
        Me.cmdPlanta.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdPlanta.Location = New System.Drawing.Point(153, 5)
        Me.cmdPlanta.Name = "cmdPlanta"
        Me.cmdPlanta.Size = New System.Drawing.Size(140, 72)
        Me.cmdPlanta.TabIndex = 41
        Me.cmdPlanta.Text = "Planta Proceso"
        Me.cmdPlanta.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdPlanta.UseVisualStyleBackColor = True
        '
        'btnAsimilados
        '
        Me.btnAsimilados.Image = CType(resources.GetObject("btnAsimilados.Image"), System.Drawing.Image)
        Me.btnAsimilados.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnAsimilados.Location = New System.Drawing.Point(7, 160)
        Me.btnAsimilados.Name = "btnAsimilados"
        Me.btnAsimilados.Size = New System.Drawing.Size(87, 72)
        Me.btnAsimilados.TabIndex = 45
        Me.btnAsimilados.Text = " Asimilados"
        Me.btnAsimilados.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnAsimilados.UseVisualStyleBackColor = True
        Me.btnAsimilados.Visible = False
        '
        'cmdDeterminado
        '
        Me.cmdDeterminado.Image = CType(resources.GetObject("cmdDeterminado.Image"), System.Drawing.Image)
        Me.cmdDeterminado.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdDeterminado.Location = New System.Drawing.Point(7, 5)
        Me.cmdDeterminado.Name = "cmdDeterminado"
        Me.cmdDeterminado.Size = New System.Drawing.Size(140, 72)
        Me.cmdDeterminado.TabIndex = 46
        Me.cmdDeterminado.Text = "Determinado"
        Me.cmdDeterminado.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdDeterminado.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.cmdDeterminado)
        Me.Panel1.Controls.Add(Me.btnAsimilados)
        Me.Panel1.Controls.Add(Me.cmdPlanta)
        Me.Panel1.Controls.Add(Me.cmdingreso)
        Me.Panel1.Location = New System.Drawing.Point(3, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(478, 250)
        Me.Panel1.TabIndex = 99
        '
        'frmJuridico
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(493, 279)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmJuridico"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Juridico"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdingreso As System.Windows.Forms.Button
    Friend WithEvents cmdPlanta As System.Windows.Forms.Button
    Friend WithEvents btnAsimilados As System.Windows.Forms.Button
    Friend WithEvents cmdDeterminado As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
End Class
