Imports System.IO
Imports System.Text.RegularExpressions

Public Class Form1
    Dim st, fecha, hora, csn, vin, PVI, año_modelo, panel, cable, harness, consola, insulator, grill, lampLH, LampRH, VsrLH, VsrRH, tnrLH, tnrRH, Support, CvrRet, BtSpacer, magnet,
        RetnrHld, BrketHl, AbsvrLH1, AbsvrRH2, AbsvrLH3, AbsvrRH4, AbsvrLH5, AbsvrLH6, AbsvrLH7, AbsvrLH8, AbsvrLH9, AbsvrLH0, fin As String

    Dim idcrfm, status, orden, fan, radiator, hose, condenser, fullSec As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            'Refresco el grid de las secuencias
            Me.WindowState = FormWindowState.Maximized
            dgvSecuencias.AutoGenerateColumns = False
            dgvSecuencias.DataSource = Nothing
            dgvSecuencias.DataSource = ObtenerDatos("SELECT * FROM Sequence_Hdr ORDER BY sh_AutoInc")
            dgvSecuencias.ClearSelection()

            'Refresco el Log
            dgvLog.AutoGenerateColumns = False
            dgvLog.DataSource = Nothing
            dgvLog.DataSource = ObtenerDatos("SELECT TOP (100) * FROM tbl_Log ORDER BY tbl_ID")
            dgvLog.ClearSelection()


            ''Secuencias Manuales
            OpenFileDialog1.Filter = "Csv Files (*.csv)|*.csv|Text File (*.txt)| *.txt"
            ''

            Me.UsrDespliega_Mensaje1.Visible = False
            Me.UsrImportador1.IniciarTimer()
        Catch ex As Exception
            EscribirLog(ex.Message)
        End Try
    End Sub

    Private Sub btnSeleccionar_Click(sender As Object, e As EventArgs) Handles btnSeleccionar.Click
        OpenFileDialog1.ShowDialog()
        lblRuta.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub btnCargar_Click(sender As Object, e As EventArgs) Handles btnCargar.Click
        Dim texto As Stream = OpenFileDialog1.OpenFile()
        Using lector As New StreamReader(texto)
            RichTextBox1.Text = lector.ReadToEnd
        End Using
    End Sub

    Private Sub btnProcesar_Click(sender As Object, e As EventArgs) Handles btnProcesar.Click
        Try
            Dim secuencia As String = RichTextBox1.Text.ToUpper
            Dim arreglo() As String = Regex.Split(secuencia, "END\b")
            RichTextBox1.Text = ""

            For index = 0 To arreglo.Length - 1
                If Not String.IsNullOrWhiteSpace(arreglo(index)) Then
                    RichTextBox1.Text += "-------------------- Nuevo --------------------" & Environment.NewLine
                    arreglo(index) = arreglo(index).TrimStart
                    RichTextBox1.Text += arreglo(index) & Environment.NewLine
                    fullSec = arreglo(index)

                    'Formato para consolas CRFM.
                    idcrfm = Mid(arreglo(index), 1, 1)
                    status = Mid(arreglo(index), 2, 2)
                    orden = Mid(arreglo(index), 4, 6)
                    csn = Mid(Mid(arreglo(index), 10, 11), 4, 7).Trim
                    vin = Mid(arreglo(index), 21, 17)
                    PVI = Mid(arreglo(index), 38, 9)
                    año_modelo = Mid(arreglo(index), 47, 5)
                    fan = Mid(arreglo(index), 52, 8)
                    radiator = Mid(arreglo(index), 60, 8)
                    hose = Mid(arreglo(index), 68, 8)
                    condenser = Mid(arreglo(index), 76, 8)
                    fecha = Mid(arreglo(index), 84, 8)
                    hora = Mid(arreglo(index), 92, 5)
                    Me.EscribirDetallesCRFM()
                    'Termina formato para consolas CRFM.

                    'Habra 3 tipos de formatos de secuencias.
                    'st = Mid(arreglo(index), 1, 4)
                    'fecha = Mid(arreglo(index), 5, 8)
                    'hora = Mid(arreglo(index), 13, 5)
                    'csn = Mid(Mid(arreglo(index), 18, 11), 4, 9).Trim
                    'vin = Mid(arreglo(index), 29, 17)
                    'PVI = Mid(arreglo(index), 46, 9)
                    'año_modelo = Mid(arreglo(index), 55, 5)
                    'panel = Mid(arreglo(index), 63, 8)
                    'cable = Mid(arreglo(index), 71, 8)
                    'harness = Mid(arreglo(index), 79, 8)
                    'consola = Mid(arreglo(index), 87, 8)
                    'insulator = Mid(arreglo(index), 95, 8)
                    'grill = Mid(arreglo(index), 103, 8)

                    'lampLH = Mid(arreglo(index), 111, 8)
                    'LampRH = Mid(arreglo(index), 119, 8)
                    'VsrLH = Mid(arreglo(index), 127, 8)
                    'VsrRH = Mid(arreglo(index), 135, 8)
                    'tnrLH = Mid(arreglo(index), 143, 8)
                    'tnrRH = Mid(arreglo(index), 151, 8)
                    'Support = Mid(arreglo(index), 159, 8)
                    'CvrRet = Mid(arreglo(index), 167, 8)
                    'BtSpacer = Mid(arreglo(index), 175, 8)
                    'magnet = Mid(arreglo(index), 183, 8)
                    'RetnrHld = Mid(arreglo(index), 191, 8)
                    'BrketHl = Mid(arreglo(index), 199, 8)

                    'AbsvrLH1 = Mid(arreglo(index), 207, 8)
                    'AbsvrRH2 = Mid(arreglo(index), 215, 8)
                    'AbsvrLH3 = Mid(arreglo(index), 223, 8)
                    'AbsvrRH4 = Mid(arreglo(index), 231, 8)

                    'AbsvrLH5 = Mid(arreglo(index), 239, 8)
                    'AbsvrLH6 = Mid(arreglo(index), 247, 8)
                    'AbsvrLH7 = Mid(arreglo(index), 255, 8)
                    'AbsvrLH8 = Mid(arreglo(index), 263, 8)
                    'AbsvrLH9 = Mid(arreglo(index), 271, 8)
                    'AbsvrLH0 = Mid(arreglo(index), 279, 8)
                    'fin = Mid(secuencia, 287, 3)
                    'Me.EscribirDetalles()

                    Dim dtExiste As DataTable = ObtenerDatos("SELECT * FROM Sequence_Hdr WHERE (sh_SequenceID = " & csn & ") AND (sh_Validated = 'True')")
                    If dtExiste.Rows.Count > 0 Then
                        Throw New ArgumentException("Secuencia Repetida " & csn & " ya ha sido agregada anteriormente.")
                    End If

                    'Enviar correo para notificar Error.
                    ObtenerDatos("INSERT INTO [tbl_Log] ([tbl_CSN] ,[tbl_Type] ,[tbl_Error] ,[tbl_DateTime]) " &
                                 " VALUES ('" & csn & "', 'OK', 'Secuencia insertada correctamente.', getDate())")

                    ObtenerDatos("IF (NOT EXISTS(SELECT * FROM Sequence_Hdr WHERE (sh_FullSeq = '" & fullSec & "'))) " & _
      " BEGIN " & _
      "     INSERT INTO [Sequence_Hdr] ([sh_SequenceID] ,[sh_FullSeq] ,[sh_Received] ,[sh_type] ,[sh_Cust], [sh_Validated], [sh_ValidatedDT])" & _
      "     VALUES (" & csn & ", '" & arreglo(index) & "', getDate(), 'Manual', 'GM', 'True', getDate())" & _
      " END " & _
      " ELSE " & _
      " BEGIN " & _
      "     UPDATE Sequence_Hdr SET sh_SequenceID = '" & csn & "', sh_Validated = 'True', sh_ValidatedDT = getDate(), sh_Hold= 'False' WHERE (sh_FullSeq ='" & fullSec & "')" & _
      " END ")

                    ObtenerDatos("INSERT INTO [Sequence_Det] ([sd_SeqID],[sd_PartNum],[sd_ClientPart],[sd_Vin],[sd_Panel],[sd_Cable],[sd_Harness],[sd_Console],[sd_Insultr],[sd_Grill],[sd_LampLH],[sd_LampRH],[sd_SunshadeLH],[sd_SunshadeRH],[sd_RtnrLH],[sd_RtnrRH],[sd_Support],[sd_CvrRetBt],[sd_Spacer],[sd_Magnet],[sd_RtnrHld], [sd_BrketHl],[sd_AbsorverLH1],[sd_AbsorverRH1],[sd_AbsorverLH2],[sd_AbsorverRH2],[sd_AbsorverLH3],[sd_AbsorverLH4],[sd_AbsorverLH5],[sd_AbsorverLH6],[sd_AbsorverLH7],[sd_AbsorverLH8], [sd_PVI]) " &
                    " VALUES (" & csn & ", 'PARTNUM', 'ClientPart', '" & vin & "', '" & panel & "', '" & cable & "', '" & harness & "', '" & consola & "', '" & insulator & "', '" & grill & "', '" & lampLH & "', '" & LampRH & "', '" & VsrLH & "', '" & VsrRH & "', '" & tnrLH & "', '" & tnrRH & "', '" & Support & "', '" & CvrRet & "', '" & BtSpacer & "', '" & magnet & "', '" & RetnrHld & "', '" & BrketHl & "', '" & AbsvrLH1 & "', '" & AbsvrRH2 & "', '" & AbsvrLH3 & "', '" & AbsvrRH4 & "', '" & AbsvrLH5 & "', '" & AbsvrLH6 & "', '" & AbsvrLH7 & "', '" & AbsvrLH8 & "', '" & AbsvrLH9 & "', '" & AbsvrLH0 & "', '" & PVI & "')")

                    UsrDespliega_Mensaje1.MostrarMensaje("Secuencia " & csn & " Insertada Correctamente.", "OK", 3000)
                    Threading.Thread.Sleep(500)
                End If
            Next
        Catch ex As Exception
            Me.UsrDespliega_Mensaje1.Visible = True
            Me.UsrDespliega_Mensaje1.MostrarMensaje(ex.Message, "ERROR", 5000)
            EscribirLog(ex.Message)

            'Enviar correo para notificar Error.
            SendEmail(ex.Message)
            'enviando el correo.

            ObtenerDatos("INSERT INTO [tbl_Log] ([tbl_CSN] ,[tbl_Type] ,[tbl_Error] ,[tbl_DateTime]) VALUES ('" & csn & "', 'ERROR', '" & ex.Message & "', getDate())")

            ObtenerDatos("IF (NOT EXISTS(SELECT * FROM Sequence_Hdr WHERE sh_SequenceID = '" & csn & "')) " & _
      " BEGIN " & _
      "     INSERT INTO [Sequence_Hdr] ([sh_SequenceID] ,[sh_FullSeq] ,[sh_Received] ,[sh_type] ,[sh_Cust], [sh_Hold], [sh_HoldDT])" & _
      "     VALUES (" & csn & ", '" & fullSec & "', getDate(), 'Manual', 'GM', 1, getDate())" & _
      " END ")
        Finally
            'Refresco el grid de las secuencias
            dgvSecuencias.AutoGenerateColumns = False
            dgvSecuencias.DataSource = Nothing
            dgvSecuencias.DataSource = ObtenerDatos("SELECT * FROM Sequence_Hdr ORDER BY sh_AutoInc")
            dgvSecuencias.ClearSelection()

            'Refresco el Log
            dgvLog.AutoGenerateColumns = False
            dgvLog.DataSource = Nothing
            dgvLog.DataSource = ObtenerDatos("SELECT TOP (100) * FROM tbl_Log ORDER BY tbl_ID")
            dgvLog.ClearSelection()
        End Try
    End Sub

    Private Sub cmbEventos_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbEventos.SelectedIndexChanged
        Try
            dgvLog.AutoGenerateColumns = False
            dgvLog.DataSource = Nothing
            If cmbEventos.SelectedIndex = 0 Then
                dgvLog.DataSource = ObtenerDatos("SELECT TOP (100) * FROM tbl_Log ORDER BY tbl_ID")
            ElseIf cmbEventos.SelectedIndex = 1 Then
                dgvLog.DataSource = ObtenerDatos("SELECT TOP (100) * FROM tbl_Log WHERE tbl_Type = 'OK' ORDER BY tbl_ID")
            ElseIf cmbEventos.SelectedIndex = 2 Then
                dgvLog.DataSource = ObtenerDatos("SELECT TOP (100) * FROM tbl_Log WHERE tbl_Type = 'ERROR' ORDER BY tbl_ID")
            End If
            dgvLog.ClearSelection()
        Catch ex As Exception
            EscribirLog(ex.Message)
        End Try
    End Sub


    Private Sub dgvSecuencias_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles dgvSecuencias.CellFormatting
        If e.ColumnIndex = 14 Then
            If e.Value Then
                dgvSecuencias.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Yellow
            End If
        End If
    End Sub

    Private Sub dgvLog_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles dgvLog.CellFormatting
        If e.ColumnIndex = 2 Then
            If e.Value = "OK" Then
                e.CellStyle.BackColor = Color.LimeGreen
                e.CellStyle.ForeColor = Color.Black
            ElseIf e.Value = "ERROR" Then
                e.CellStyle.BackColor = Color.Red
                e.CellStyle.ForeColor = Color.Black
            Else
                e.CellStyle.BackColor = Color.Yellow
                e.CellStyle.ForeColor = Color.Black
            End If

        End If

    End Sub

    Public Sub ValidarParte(ByVal parte As String)
        If Not String.IsNullOrWhiteSpace(parte) Then
            Dim dtParte As DataTable = ObtenerDatos("SELECT * FROM NivelesIngenieria " & _
   " WHERE tbl_PartNumber = '" & parte & "' AND (tbl_StartDate <= GETDATE()) AND (tbl_StopDate >= GETDATE())")

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

    Public Sub EscribirDetalles()
        RichTextBox1.Text += "ST: " & st & Environment.NewLine
        RichTextBox1.Text += "FECHA: " & fecha & Environment.NewLine
        RichTextBox1.Text += "HORA: " & hora & Environment.NewLine
        RichTextBox1.Text += "CCN: " & csn & Environment.NewLine
        RichTextBox1.Text += "VIN: " & vin & Environment.NewLine

        RichTextBox1.Text += "PVI: " & PVI & Environment.NewLine
        RichTextBox1.Text += "AÑO_MODELO: " & año_modelo & Environment.NewLine
        RichTextBox1.Text += "PANEL: " & panel & Environment.NewLine
        RichTextBox1.Text += "CABLE: " & cable & Environment.NewLine
        RichTextBox1.Text += "HARNESS: " & harness & Environment.NewLine

        RichTextBox1.Text += "CONSOLA: " & consola & Environment.NewLine
        RichTextBox1.Text += "GRILL: " & grill & Environment.NewLine
        RichTextBox1.Text += "LAMP LH: " & lampLH & Environment.NewLine
        RichTextBox1.Text += "LAMP RH: " & LampRH & Environment.NewLine
        RichTextBox1.Text += "VSR LH: " & VsrLH & Environment.NewLine
        RichTextBox1.Text += "VSR RH: " & VsrRH & Environment.NewLine

        RichTextBox1.Text += "TNR LH: " & tnrLH & Environment.NewLine
        RichTextBox1.Text += "TNR RH: " & tnrRH & Environment.NewLine
        RichTextBox1.Text += "SUPPORT: " & Support & Environment.NewLine
        RichTextBox1.Text += "CVR RET: " & CvrRet & Environment.NewLine
        RichTextBox1.Text += "SPACER: " & BtSpacer & Environment.NewLine
        RichTextBox1.Text += "MAGNET: " & magnet & Environment.NewLine

        RichTextBox1.Text += "RETN HLD: " & RetnrHld & Environment.NewLine
        RichTextBox1.Text += "BRAKET HLD: " & BrketHl & Environment.NewLine
        RichTextBox1.Text += "ABSORVER LH: " & AbsvrLH1 & Environment.NewLine
        RichTextBox1.Text += "ABSORVER RH: " & AbsvrRH2 & Environment.NewLine
        RichTextBox1.Text += "ABSORVER LH: " & AbsvrLH3 & Environment.NewLine
        RichTextBox1.Text += "ABSORVER RH: " & AbsvrRH4 & Environment.NewLine

        RichTextBox1.Text += "ABSORVER LH: " & AbsvrLH5 & Environment.NewLine
        RichTextBox1.Text += "ABSORVER LH: " & AbsvrLH6 & Environment.NewLine
        RichTextBox1.Text += "ABSORVER LH: " & AbsvrLH7 & Environment.NewLine
        RichTextBox1.Text += "ABSORVER LH: " & AbsvrLH8 & Environment.NewLine
        RichTextBox1.Text += "ABSORVER LH: " & AbsvrLH9 & Environment.NewLine
        RichTextBox1.Text += "ABSORVER LH: " & AbsvrLH0 & Environment.NewLine
        RichTextBox1.Text += "END: " & fin & Environment.NewLine
    End Sub
    Public Sub EscribirDetallesCRFM()
        RichTextBox1.Text += "ID: " & idcrfm & Environment.NewLine
        RichTextBox1.Text += "CONS: " & status & Environment.NewLine
        RichTextBox1.Text += "ORDEN: " & orden & Environment.NewLine
        RichTextBox1.Text += "CCN: " & csn & Environment.NewLine
        RichTextBox1.Text += "VIN: " & vin & Environment.NewLine
        RichTextBox1.Text += "PVI: " & PVI & Environment.NewLine
        RichTextBox1.Text += "AÑO_MODELO: " & año_modelo & Environment.NewLine
        RichTextBox1.Text += "FAN: " & fan & Environment.NewLine
        RichTextBox1.Text += "RADIATOR: " & radiator & Environment.NewLine
        RichTextBox1.Text += "HOSE: " & hose & Environment.NewLine
        RichTextBox1.Text += "CONDENSER: " & condenser & Environment.NewLine
        RichTextBox1.Text += "FECHA: " & fecha & Environment.NewLine
        RichTextBox1.Text += "HORA: " & hora & Environment.NewLine
    End Sub
End Class
