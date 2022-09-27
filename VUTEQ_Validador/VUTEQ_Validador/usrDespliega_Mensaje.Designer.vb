<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class usrDespliega_Mensaje
    Inherits System.Windows.Forms.UserControl

    'UserControl reemplaza a Dispose para limpiar la lista de componentes.
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
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'txtMensaje
        '
        Me.txtMensaje.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMensaje.Font = New System.Drawing.Font("Tahoma", 20.0!)
        Me.txtMensaje.Location = New System.Drawing.Point(0, 0)
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.Size = New System.Drawing.Size(896, 111)
        Me.txtMensaje.TabIndex = 2
        Me.txtMensaje.Text = "Area de Mensajes"
        Me.txtMensaje.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtMensaje.Visible = False
        '
        'Timer1
        '
        '
        'usrDespliega_Mensaje
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.txtMensaje)
        Me.Name = "usrDespliega_Mensaje"
        Me.Size = New System.Drawing.Size(896, 111)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtMensaje As System.Windows.Forms.TextBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer

End Class
