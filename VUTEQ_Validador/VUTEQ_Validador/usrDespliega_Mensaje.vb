Public Class usrDespliega_Mensaje
    Public Property Mensaje As String
        Get
            Return txtMensaje.Text
        End Get
        Set(ByVal value As String)
            txtMensaje.Text = value
        End Set
    End Property

    Public Sub MostrarMensaje(ByVal mensaje As String, ByVal TipoError As String, ByVal tiempo As Integer)
        Me.BringToFront()

        txtMensaje.Visible = True
        Timer1.Interval = tiempo

        If TipoError = "OK" Then
            txtMensaje.BackColor = Color.LimeGreen
            txtMensaje.BorderStyle = Windows.Forms.BorderStyle.Fixed3D
            txtMensaje.ForeColor = Color.Black
            txtMensaje.Text = mensaje
        ElseIf TipoError = "ERROR" Then
            txtMensaje.BackColor = Color.Red
            txtMensaje.BorderStyle = Windows.Forms.BorderStyle.Fixed3D
            txtMensaje.ForeColor = Color.Yellow
            txtMensaje.Text = mensaje
        ElseIf TipoError = "WRNG" Then
            txtMensaje.BackColor = Color.Yellow
            txtMensaje.BorderStyle = Windows.Forms.BorderStyle.Fixed3D
            txtMensaje.ForeColor = Color.Black
            txtMensaje.Text = mensaje
        End If
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        txtMensaje.Visible = False
        txtMensaje.Text = ""
        Timer1.Stop()
        Me.SendToBack()
    End Sub
End Class
