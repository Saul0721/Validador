Imports System.Text.RegularExpressions

Public Class usrImportador
    Dim st, fecha, hora, csn, vin, PVI, año_modelo, panel, cable, harness, consola, insulator, grill, lampLH, LampRH, VsrLH, VsrRH, tnrLH, tnrRH, Support, CvrRet, BtSpacer, magnet,
        RetnrHld, BrketHl, AbsvrLH1, AbsvrRH2, AbsvrLH3, AbsvrRH4, AbsvrLH5, AbsvrLH6, AbsvrLH7, AbsvrLH8, AbsvrLH9, AbsvrLH0, fin As String
    Dim idcrfm, status, orden, fan, radiator, hose, condenser, fullSec As String

    Dim comando As String = ""
    Private Sub usrImportador_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Public Sub IniciarTimer()
        Try
            Me.tmrIntervalos.Interval = My.Settings.glbTiempo
            Me.tmrIntervalos.Start()
            Me.UsrDespliega_Mensaje1.MostrarMensaje("El Validador ha iniciado", "OK", 3500)
            EscribirLog(" El Validador ha iniciado correctamente.")
        Catch ex As Exception
            EscribirLog(ex.Message)
        End Try
    End Sub

    Public Sub DetenerImportador()
        Me.tmrIntervalos.Stop()
        Me.UsrDespliega_Mensaje1.MostrarMensaje("El Validador se ha detenido.", "ERROR", 5000)
        EscribirLog(" El Validador se ha detenido.")
    End Sub

    Private Sub tmrIntervalos_Tick(sender As Object, e As EventArgs) Handles tmrIntervalos.Tick
        Try
            'Ir a buscar las secuencias que aun no se validan en la base de datos.
            Dim dtSecuencias As DataTable = ObtenerDatos(" SELECT * " & _
                                            " FROM Sequence_Hdr " & _
                                            " WHERE (sh_SequenceID = 0) AND (sh_Validated = 0) AND (sh_Hold = 0) ORDER BY sh_AutoInc")

            If dtSecuencias.Rows.Count > 0 Then
                'Detengo el Timer hasta que acabe de procesarlas y mando notificacion.
                DetenerImportador()
                Me.UsrDespliega_Mensaje1.MostrarMensaje("Secuencias encontradas, validando....", "OK", 5000)
                EscribirLog("Secuencias encontradas, validando...")

                'Recorrerlas y leer el campo de FullCSN para separar valores.
                For Each row As DataRow In dtSecuencias.Rows
                    'Threading.Thread.Sleep(500)
                    Dim secuencia As String = row.Item("sh_FullSeq").ToString.ToUpper

                    'Por cada secuencia, aplicare la expresion regular para separar cada END de cada secuencia y asi partirlo en un arreglo.
                    Dim arreglo() As String = Regex.Split(secuencia, "END\b")

                    'Recorro cada parte del arreglo para partir los datos de la secuencia.
                    For index = 0 To arreglo.Length - 1
                        If Not String.IsNullOrWhiteSpace(arreglo(index)) Then
                            fullSec = arreglo(index)

                            'Podria usar el LENGHT de la secuencia para saber que tipo de comodity es.
                            If fullSec.Length > 289 Then
                            Else
                            End If

                            'FIN Validacion
                            'Aqui modificar el tipo de secuencia que se estara insertando.
                            'Formato para consolas CRFM.
                            'idcrfm = Mid(arreglo(index), 1, 1)
                            'status = Mid(arreglo(index), 2, 2)
                            'orden = Mid(arreglo(index), 4, 6)
                            'csn = Mid(Mid(arreglo(index), 10, 11), 4, 7).Trim
                            'vin = Mid(arreglo(index), 21, 17)
                            'PVI = Mid(arreglo(index), 38, 9)
                            'año_modelo = Mid(arreglo(index), 47, 5)
                            'fan = Mid(arreglo(index), 52, 8)
                            'radiator = Mid(arreglo(index), 60, 8)
                            'hose = Mid(arreglo(index), 68, 8)
                            'condenser = Mid(arreglo(index), 76, 8)
                            'fecha = Mid(arreglo(index), 84, 8)
                            'hora = Mid(arreglo(index), 92, 5)

                            st = Mid(arreglo(index), 1, 4)
                            fecha = Mid(arreglo(index), 5, 8)
                            hora = Mid(arreglo(index), 13, 5)
                            'csn = Mid(Mid(arreglo(index), 10, 11), 4, 7).Trim
                            csn = Mid(arreglo(index), 18, 11)
                            csn = Mid(csn, 4, 7)
                            vin = Mid(arreglo(index), 29, 17)
                            PVI = Mid(arreglo(index), 46, 9)
                            año_modelo = Mid(arreglo(index), 55, 5)
                            panel = Mid(arreglo(index), 63, 8)
                            cable = Mid(arreglo(index), 71, 8)
                            harness = Mid(arreglo(index), 79, 8)
                            consola = Mid(arreglo(index), 87, 8)
                            insulator = Mid(arreglo(index), 95, 8)
                            grill = Mid(arreglo(index), 103, 8)
                            lampLH = Mid(arreglo(index), 111, 8)
                            LampRH = Mid(arreglo(index), 119, 8)
                            VsrLH = Mid(arreglo(index), 127, 8)
                            VsrRH = Mid(arreglo(index), 135, 8)
                            tnrLH = Mid(arreglo(index), 143, 8)
                            tnrRH = Mid(arreglo(index), 151, 8)
                            Support = Mid(arreglo(index), 159, 8)
                            CvrRet = Mid(arreglo(index), 167, 8)
                            BtSpacer = Mid(arreglo(index), 175, 8)
                            magnet = Mid(arreglo(index), 183, 8)
                            RetnrHld = Mid(arreglo(index), 191, 8)
                            BrketHl = Mid(arreglo(index), 199, 8)
                            AbsvrLH1 = Mid(arreglo(index), 207, 8)
                            AbsvrRH2 = Mid(arreglo(index), 215, 8)
                            AbsvrLH3 = Mid(arreglo(index), 223, 8)
                            AbsvrRH4 = Mid(arreglo(index), 231, 8)
                            AbsvrLH5 = Mid(arreglo(index), 239, 8)
                            AbsvrLH6 = Mid(arreglo(index), 247, 8)
                            AbsvrLH7 = Mid(arreglo(index), 255, 8)
                            AbsvrLH8 = Mid(arreglo(index), 263, 8)
                            AbsvrLH9 = Mid(arreglo(index), 271, 8)
                            AbsvrLH0 = Mid(arreglo(index), 279, 8)
                            'fin = Mid(secuencia, 287, 3) END punto que parte cada secuencia, en este punto esta parte ya no existe por que la expresion regular la partio.

                            Dim dtExiste As DataTable = ObtenerDatos("SELECT * FROM Sequence_Hdr WHERE (sh_SequenceID = " & csn & ")")
                            If dtExiste.Rows.Count > 0 Then
                                SendEmail("Se detecto secuencia repetida " & csn & ", esta no fue ingresada para evitar duplicados.")
                            End If
                            EscribirLog(" Se encontraron secuencias recibidas, Secuencia: " & csn & ", validando...")

                            ObtenerDatos("INSERT INTO [tbl_Log] ([tbl_CSN] ,[tbl_Type] ,[tbl_Error] ,[tbl_DateTime]) " &
                                         " VALUES ('" & csn & "', 'OK', 'Secuencia insertada correctamente de forma automatica.', getDate())")

                            ObtenerDatos("UPDATE Sequence_Hdr " & _
                                         "SET sh_SequenceID = " & csn & ", sh_Validated = 1, sh_ValidatedDT = getDate() " & _
                                         "WHERE sh_FullSeq = '" & secuencia & "'")

                            'Aqui validar si no es toldos, insertarle los campos que cada Comodity aceptara.
                            ObtenerDatos("INSERT INTO [Sequence_Det] ([sd_SeqID],[sd_PartNum],[sd_ClientPart],[sd_Vin],[sd_Panel],[sd_Cable],[sd_Harness],[sd_Console],[sd_Insultr],[sd_Grill],[sd_LampLH],[sd_LampRH],[sd_SunshadeLH],[sd_SunshadeRH],[sd_RtnrLH],[sd_RtnrRH],[sd_Support],[sd_CvrRetBt],[sd_Spacer],[sd_Magnet],[sd_RtnrHld], [sd_BrketHl],[sd_AbsorverLH1],[sd_AbsorverRH1],[sd_AbsorverLH2],[sd_AbsorverRH2],[sd_AbsorverLH3],[sd_AbsorverLH4],[sd_AbsorverLH5],[sd_AbsorverLH6],[sd_AbsorverLH7],[sd_AbsorverLH8], [sd_Fan], [sd_Radiator], [sd_Hose], [sd_Condenser], [sd_PVI],[sd_anoModelo]) " &
                            " VALUES (" & csn & ", 'PARTNUM', 'ClientPart', '" & vin & "', '" & panel & "', '" & cable & "', '" & harness & "', '" & consola & "', '" & insulator & "', '" & grill & "', '" & lampLH & "', '" & LampRH & "', '" & VsrLH & "', '" & VsrRH & "', '" & tnrLH & "', '" & tnrRH & "', '" & Support & "', '" & CvrRet & "', '" & BtSpacer & "', '" & magnet & "', '" & RetnrHld & "', '" & BrketHl & "', '" & AbsvrLH1 & "', '" & AbsvrRH2 & "', '" & AbsvrLH3 & "', '" & AbsvrRH4 & "', '" & AbsvrLH5 & "', '" & AbsvrLH6 & "', '" & AbsvrLH7 & "', '" & AbsvrLH8 & "', '" & AbsvrLH9 & "', '" & AbsvrLH0 & "', '" & fan & "', '" & radiator & "', '" & hose & "', '" & condenser & "', '" & PVI & "', '" & año_modelo & "')")

                            System.Threading.Thread.Sleep(2500)
                            'Replico la informacion, para pasar de una base de datos a otra.
                            Dim dtSecHeader As DataTable = ObtenerDatos("SELECT * FROM Sequence_Hdr WHERE sh_FullSeq = '" & secuencia & "'")
                            For Each datos As DataRow In dtSecHeader.Rows
                                Dim parte1 As String = Mid(datos.Item("sh_FullSeq").ToString, 1, 145)
                                Dim parte2 As String = Mid(datos.Item("sh_FullSeq").ToString, 146, 290)

                                comando = datos.Item("sh_SequenceID") & ",'" & parte1 & "'" & _
                                "'" & parte2 & "', '" & DateTime.Parse(datos.Item("sh_Received")).ToString("yyyy-MM-dd HH:mm:ss") & "', '" & datos.Item("sh_type").ToString & "', '" & datos.Item("sh_Cust").ToString & "', '" & datos.Item("sh_Validated").ToString & "', '" & DateTime.Parse(datos.Item("sh_ValidatedDT")).ToString("yyyy-MM-dd HH:mm:ss") & "'"

                                EscribirLog("COMANDO: " & comando)
                                ObtenerDatosProduccion("INSERT INTO Sequence_Hdr(sh_SequenceID, sh_FullSeq, sh_Received, sh_type, sh_Cust, sh_Validated, sh_ValidatedDT) " & _
                                " VALUES(" & comando & ")")
                            Next
                            ObtenerDatosProduccion("INSERT INTO [Sequence_Det] ([sd_SeqID],[sd_PartNum],[sd_ClientPart],[sd_Vin],[sd_Panel],[sd_Cable],[sd_Harness],[sd_Console],[sd_Insultr],[sd_Grill],[sd_LampLH],[sd_LampRH],[sd_SunshadeLH],[sd_SunshadeRH],[sd_RtnrLH],[sd_RtnrRH],[sd_Support],[sd_CvrRetBt],[sd_Spacer],[sd_Magnet],[sd_RtnrHld], [sd_BrketHl],[sd_AbsorverLH1],[sd_AbsorverRH1],[sd_AbsorverLH2],[sd_AbsorverRH2],[sd_AbsorverLH3],[sd_AbsorverLH4],[sd_AbsorverLH5],[sd_AbsorverLH6],[sd_AbsorverLH7],[sd_AbsorverLH8], [sd_Fan], [sd_Radiator], [sd_Hose], [sd_Condenser], [sd_PVI],[sd_anoModelo]) " &
                            " VALUES (" & csn & ", 'PARTNUM', 'ClientPart', '" & vin & "', '" & panel & "', '" & cable & "', '" & harness & "', '" & consola & "', '" & insulator & "', '" & grill & "', '" & lampLH & "', '" & LampRH & "', '" & VsrLH & "', '" & VsrRH & "', '" & tnrLH & "', '" & tnrRH & "', '" & Support & "', '" & CvrRet & "', '" & BtSpacer & "', '" & magnet & "', '" & RetnrHld & "', '" & BrketHl & "', '" & AbsvrLH1 & "', '" & AbsvrRH2 & "', '" & AbsvrLH3 & "', '" & AbsvrRH4 & "', '" & AbsvrLH5 & "', '" & AbsvrLH6 & "', '" & AbsvrLH7 & "', '" & AbsvrLH8 & "', '" & AbsvrLH9 & "', '" & AbsvrLH0 & "', '" & fan & "', '" & radiator & "', '" & hose & "', '" & condenser & "', '" & PVI & "', '" & año_modelo & "')")
                            'Fin Replica de informacion

                            UsrDespliega_Mensaje1.MostrarMensaje("Secuencia " & csn & " Insertada Correctamente.", "OK", 3000)
                            EscribirLog("Secuencia " & csn & " Insertada Correctamente.")

                            'Refresco el grid de las secuencias
                            Form1.dgvSecuencias.AutoGenerateColumns = False
                            Form1.dgvSecuencias.DataSource = Nothing
                            Form1.dgvSecuencias.DataSource = ObtenerDatos("SELECT * FROM Sequence_Hdr ORDER BY sh_AutoInc")
                            Form1.dgvSecuencias.ClearSelection()

                            Threading.Thread.Sleep(1500)

                        End If
                    Next


                Next

                'Cuando termine, inicio nuevamente el timer.
                Me.IniciarTimer()
            Else
                UsrDespliega_Mensaje1.MostrarMensaje("Esperando Secuencias sin validar.", "WRNG", 3000)
            End If
        Catch ex As Exception
            Me.UsrDespliega_Mensaje1.Visible = True
            Me.UsrDespliega_Mensaje1.MostrarMensaje(ex.Message, "ERROR", 5000)
            EscribirLog(ex.Message & comando)

            'Enviar correo para notificar Error.
            SendEmail(ex.Message)
            'enviando el correo.

            ObtenerDatos("INSERT INTO [tbl_Log] ([tbl_CSN] ,[tbl_Type] ,[tbl_Error] ,[tbl_DateTime]) VALUES ('" & csn & "', 'ERROR', '" & ex.Message & "', getDate())")

            ObtenerDatos("IF (NOT EXISTS(SELECT * FROM Sequence_Hdr WHERE sh_FullSeq = '" & fullSec & "')) " & _
      " BEGIN " & _
      "     INSERT INTO [Sequence_Hdr] ([sh_SequenceID] ,[sh_FullSeq] ,[sh_Received] ,[sh_type] ,[sh_Cust], [sh_Hold], [sh_HoldDT])" & _
      "     VALUES (" & csn & ", '" & fullSec & "', getDate(), 'Automatico', 'GM', 1, getDate())" & _
      " END " & _
      " ELSE " & _
      " BEGIN " & _
      "     UPDATE Sequence_Hdr SET sh_Hold= 'True', sh_HoldDT = getDate() WHERE (sh_FullSeq ='" & fullSec & "')" & _
      " END ")
        Finally
            'Refresco el grid de las secuencias
            Form1.dgvSecuencias.AutoGenerateColumns = False
            Form1.dgvSecuencias.DataSource = Nothing
            Form1.dgvSecuencias.DataSource = ObtenerDatos("SELECT * FROM Sequence_Hdr ORDER BY sh_AutoInc")
            Form1.dgvSecuencias.ClearSelection()

            'Refresco el Log
            Form1.dgvLog.AutoGenerateColumns = False
            Form1.dgvLog.DataSource = Nothing
            Form1.dgvLog.DataSource = ObtenerDatos("SELECT TOP (100) * FROM tbl_Log ORDER BY tbl_DateTime ASC")
            Form1.dgvLog.ClearSelection()
        End Try
    End Sub

    Public Sub ValidarParte(ByVal parte As String)
        If Not String.IsNullOrWhiteSpace(parte) Then
            Dim dtParte As DataTable = ObtenerDatos("SELECT * FROM NivelesIngenieria " & _
   " WHERE (tbl_PartNumber = '" & parte & "') AND (tbl_StartDate <= GETDATE()) AND (tbl_StopDate >= GETDATE())")

            If dtParte.Rows.Count = 0 Then
                Throw New ArgumentException("No existe el numero de parte " & parte & " en el control de niveles de ingenieria.")
                Exit Sub
            End If

            For Each row As DataRow In dtParte.Rows
                If row.Item("tbl_Status") <> "A" Then
                    Throw New ArgumentException("El numero de parte " & parte & " que llego en la secuencia " & csn & " no tiene el nivel de ingenieria actualizado.")
                End If
            Next
        End If
    End Sub
End Class
