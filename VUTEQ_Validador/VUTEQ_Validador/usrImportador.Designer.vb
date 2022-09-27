<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class usrImportador
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
        Me.tmrIntervalos = New System.Windows.Forms.Timer(Me.components)
        Me.UsrDespliega_Mensaje1 = New VUTEQ_Validador.usrDespliega_Mensaje()
        Me.SuspendLayout()
        '
        'tmrIntervalos
        '
        Me.tmrIntervalos.Interval = 2000
        '
        'UsrDespliega_Mensaje1
        '
        Me.UsrDespliega_Mensaje1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.UsrDespliega_Mensaje1.Location = New System.Drawing.Point(0, 0)
        Me.UsrDespliega_Mensaje1.Mensaje = "Area de Mensajes"
        Me.UsrDespliega_Mensaje1.Name = "UsrDespliega_Mensaje1"
        Me.UsrDespliega_Mensaje1.Size = New System.Drawing.Size(684, 95)
        Me.UsrDespliega_Mensaje1.TabIndex = 0
        '
        'usrImportador
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.UsrDespliega_Mensaje1)
        Me.Name = "usrImportador"
        Me.Size = New System.Drawing.Size(684, 95)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tmrIntervalos As System.Windows.Forms.Timer
    Friend WithEvents UsrDespliega_Mensaje1 As VUTEQ_Validador.usrDespliega_Mensaje

End Class
