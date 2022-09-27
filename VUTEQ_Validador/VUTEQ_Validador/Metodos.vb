Imports System.Data.SqlClient
Imports System.IO
Imports System.Net.Mail

'Imports EASendMail


Module Metodos

    ''' <summary>
    ''' Metodo para realizar consultas, insert, updates y deletes con SQL.
    ''' </summary>
    ''' <param name="comando"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ObtenerDatos(ByVal comando As String) As DataTable
        Try
            Using cnConexion As New SqlConnection(My.Settings.glbcnSQL)
                Dim Dt As DataTable
                Dim Da As New SqlDataAdapter
                Dim Cmd As New SqlCommand
                If Not cnConexion.State = ConnectionState.Open Then
                    cnConexion.Open()
                End If
                With Cmd
                    .CommandType = CommandType.Text
                    .CommandText = comando
                    .Connection = cnConexion
                End With
                Da.SelectCommand = Cmd
                Dt = New DataTable
                Da.Fill(Dt)
                If cnConexion.State = ConnectionState.Open Then
                    cnConexion.Close()
                End If
                Return Dt
            End Using
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function ObtenerDatosProduccion(ByVal comando As String) As DataTable
        Try
            Using cnProduccion As New SqlConnection(My.Settings.glbcnProd)
                Dim Dt As DataTable
                Dim Da As New SqlDataAdapter
                Dim Cmd As New SqlCommand
                If Not cnProduccion.State = ConnectionState.Open Then
                    cnProduccion.Open()
                End If
                With Cmd
                    .CommandType = CommandType.Text
                    .CommandText = comando
                    .Connection = cnProduccion
                End With
                Da.SelectCommand = Cmd
                Dt = New DataTable
                Da.Fill(Dt)
                If cnProduccion.State = ConnectionState.Open Then
                    cnProduccion.Close()
                End If
                Return Dt
            End Using
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Sub SendEmail(ByVal cuerpo As String)
        'Definicion e instanciacion de la clase MailMessage
        Dim MMessage As MailMessage = New MailMessage()
        'Rellenamos los parametros usuales para el envio de un email
        For Each cadena As String In My.Settings.glbNotificaciones.Split(";")
            MMessage.To.Add(cadena)
        Next
        MMessage.From = New MailAddress("enviamail@aidctech.com.mx", "Validador", System.Text.Encoding.UTF8)
        MMessage.Subject = "Notificaciones del Validador"
        'MMessage.Attachments.Add(New Attachment(ruta))
        MMessage.Body = cuerpo
        MMessage.BodyEncoding = System.Text.Encoding.UTF8
        MMessage.IsBodyHtml = False 'Formato texto plano
        'Definimos nuestras credenciales para el unvio de emails a traves de Gmail
        Dim SClient As New SmtpClient()
        SClient.Credentials = New System.Net.NetworkCredential("enviamail@aidctech.com.mx", "3nvi@M@il22")
        SClient.Host = "mail.aidctech.com.mx"
        SClient.DeliveryMethod = SmtpDeliveryMethod.Network
        'SClient.EnableSsl = True
        'Capturamos los errores en el envio
        Try
            SClient.Send(MMessage)
            Console.ReadLine()
            EscribirLog("Mensaje Enviado")
        Catch ex As System.Net.Mail.SmtpException
            If Not ex.InnerException Is Nothing Then
                EscribirLog(ex.InnerException.Message)
            Else
                EscribirLog(ex.Message)
            End If
            Console.ReadLine()
        End Try
    End Sub

    'Public Sub SendEmail(ByVal cuerpo As String, ByVal Asunto As String)
    '    'Definicion e instanciacion de la clase MailMessage
    '    Dim MMessage As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage()
    '    'Rellenamos los parametros usuales para el envio de un email
    '    For Each variable As String In My.Settings.glbNotificaciones.Split(";")
    '        If Not String.IsNullOrWhiteSpace(variable) Then
    '            MMessage.To.Add(variable)
    '        End If
    '    Next

    '    MMessage.From = New Net.Mail.MailAddress("enviamail@aidctech.com.mx", "Validador", System.Text.Encoding.UTF8)
    '    MMessage.Subject = Asunto
    '    MMessage.SubjectEncoding = System.Text.Encoding.UTF8
    '    MMessage.Body = cuerpo
    '    MMessage.BodyEncoding = System.Text.Encoding.UTF8
    '    MMessage.IsBodyHtml = False 'Formato texto plano
    '    'Definimos nuestras credenciales para el unvio de emails a traves de Gmail
    '    'Dim SClient As New SmtpMail("TryIt")
    '    'SClient.From = "saul.ibarra@aidctech.com.mx"
    '    'SClient.To = "saul_iz93@hotmail.com"
    '    'SClient.Subject = Asunto
    '    'SClient.TextBody = cuerpo
    '    'Dim oServer As New SmtpServer("smtp.gmail.com")
    '    'oServer.User = "sistema.encuestas2021@gmail.com"
    '    'oServer.Password = "2021encuestas"
    '    'oServer.Port = 587
    '    'oServer.ConnectType = SmtpConnectType.ConnectTryTLS

    '    'Dim oSmtp As New SmtpClient()
    '    'oSmtp.SendMail(oServer, SClient)

    '    SClient.Credentials = New System.Net.NetworkCredential("", "2021encuestas")
    '    SClient.Host = "smtp.gmail.com" 'Servidor SMTP de Gmail
    '    SClient.Port = 587 'Puerto del SMTP de Gmail
    '    SClient.EnableSsl = True 'Habilita el SSL, necesio en Gmail
    '    SClient.con()
    '    'Capturamos los errores en el envio
    '    Try
    '        'SClient.Send(MMessage)
    '    Catch ex As System.Net.Mail.SmtpException
    '        If Not ex.InnerException Is Nothing Then
    '            Throw New ArgumentException(ex.InnerException.Message)
    '        Else
    '            Throw New ArgumentException(ex.Message)
    '        End If
    '    End Try
    'End Sub

    Public Sub EscribirLog(ByVal DescError As String)
        On Error GoTo Errorxx
        Dim mError As String = ""
        Dim objStreamWriter As StreamWriter

        objStreamWriter = New StreamWriter(My.Settings.glbLogFile.Trim, True)  ' Append true
        mError = Str(Err.Number) & "-" & Trim(Err.Description) & "; " & _
            Format(Now, "yyyy-MM-dd HH:mm:ss") & "; " & DescError.Trim & "." & vbCr
        objStreamWriter.WriteLine(mError)
        objStreamWriter.Close()

ErroR_ResumE:
        mError = ""
        Exit Sub
Errorxx:
        mError = Str(Err.Number) & "-" & Trim(Err.Description) & " # " & _
            Format(Now, "YYYY-mm-dd hh:mm:ss") & " Error al escribir el error en log # "
        EscribirLog(mError)
        GoTo ErroR_ResumE
    End Sub

    Public Sub ValidaParte(ByVal parte As String, ByVal secuencia As String)
        Try
            If Not String.IsNullOrWhiteSpace(parte) Then
                'Valido que el numero de parte exista.
                Dim dtParte As DataTable = ObtenerDatos("SELECT * FROM Part_Master WHERE (tbl_PartNumber = '" & parte & "')")
                If dtParte.Rows.Count = 0 Then
                    Throw New ArgumentException("El numero de parte " & parte & " no se encontro en la base de datos, viene en la secuencia " & secuencia & ".")
                End If

                'Valido el que el nivel de ingenieria sea valido y este vigente.
                Dim dtNivel As DataTable = ObtenerDatos("SELECT * FROM NivelesIngenieria " &
                " WHERE (tbl_PartNumber = '" & parte & "') AND (tbl_Status = 'A') AND (tbl_StartDate <= GETDATE() OR tbl_StartDate IS NULL) AND (tbl_StopDate >= GETDATE() OR tbl_StopDate IS NULL)")
                If dtNivel.Rows.Count = 0 Then
                    Throw New ArgumentException("El numero de parte " & parte & " no tiene una vigencia valida o no esta activo.")
                End If


            End If
        Catch ex As Exception
            EscribirLog(ex.Message)
        End Try
    End Sub
End Module
