Imports System.IO
Imports System.Net.Mail
Imports System.Data.OleDb
Imports System.Net
Imports System.Windows
Imports System.Net.Mime
Imports System.Drawing.Printing
Imports System.Threading
Imports System.Security.Principal
Imports Org.BouncyCastle.OpenSsl
Imports Org.BouncyCastle.Crypto
Imports Org.BouncyCastle.Security
Imports System.Text

Module MODULOGENERAL
    Public RESPUESTA As String

    Public Sub EnviarMsgWA(ByVal templateId As String, ByVal toPhoneNumber As String, ByVal additionalParams As Dictionary(Of String, String))
        Dim baseUrl As String = My.Settings.UrlBaseWSWA
        Dim queryParams As New List(Of String)
        queryParams.Add($"templateId={Uri.EscapeDataString(templateId)}")
        queryParams.Add($"toPhoneNumber={Uri.EscapeDataString(toPhoneNumber)}")

        For Each kvp In additionalParams
            queryParams.Add($"{Uri.EscapeDataString(kvp.Key)}={Uri.EscapeDataString(kvp.Value)}")
        Next

        Dim queryString As String = String.Join("&", queryParams)
        Dim fullUrl As String = $"{baseUrl}?{queryString}"

        Using client As New WebClient()
            Try
                Dim responseData As String = client.DownloadString(fullUrl)
            Catch ex As WebException
                If ex.Response IsNot Nothing Then
                    Using responseStream = ex.Response.GetResponseStream()
                        Using reader As New System.IO.StreamReader(responseStream)
                            Dim errorResponse = reader.ReadToEnd()
                        End Using
                    End Using
                End If
            End Try
        End Using

    End Sub
    Public Sub CARGAX(ByRef LISTA As List(Of String), ByRef CB As ComboBox, ByVal VALOR As String)
        Dim X As Integer
        For X = 0 To LISTA.Count - 1
            If LISTA(X) = VALOR Then
                CB.SelectedIndex = X
                Exit Sub
            End If
        Next
    End Sub
    Public Function ENVIASMS(ByVal TEL As String, ByVal MSG As String)
        Try
            Dim SQL As New SqlClient.SqlCommand("SELECT DBO.MENSAGESMS('" + TEL + "','" + MSG + "')", frmPrincipal.CONX)
            SQL.ExecuteNonQuery()
            SQL.Dispose()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function LLENACOMBOBOX(ByRef CB As ComboBox, ByRef LISTA As List(Of String), ByVal QUERY As String, ByVal CADENA As String) As Boolean
        CB.Items.Clear()
        LISTA.Clear()
        Dim CONX As New SqlClient.SqlConnection(CADENA)
        CONX.Open()
        Dim SQL As New SqlClient.SqlCommand(QUERY, CONX)
        Dim LEC As SqlClient.SqlDataReader
        LEC = SQL.ExecuteReader
        CB.Items.Clear()
        LISTA.Clear()
        While LEC.Read
            LISTA.Add(LEC(0))
            CB.Items.Add(LEC(1))
        End While
        LEC.Close()
        SQL.Dispose()
        CONX.Close()
        Try
            CB.SelectedIndex = 0
            Return True
        Catch ex As Exception

        End Try
        Return False
    End Function
    Public Function LLENA2LISTAS(ByRef LISTA As List(Of String), ByRef LISTA2 As List(Of String), ByVal QUERY As String, ByVal CADENA As String) As Boolean
        LISTA.Clear()
        LISTA2.Clear()
        Dim CONX As New SqlClient.SqlConnection(CADENA)
        CONX.Open()
        Dim SQL As New SqlClient.SqlCommand(QUERY, CONX)
        Dim LEC As SqlClient.SqlDataReader
        LEC = SQL.ExecuteReader
        LISTA.Clear()
        While LEC.Read
            LISTA.Add(LEC(0))
            LISTA2.Add(LEC(1))
        End While
        LEC.Close()
        SQL.Dispose()
        CONX.Close()
        Try
            Return True
        Catch ex As Exception

        End Try
        Return False
    End Function
    Private Function SetDefPrinter(ByVal sNombreImpresora As String) As Boolean
        'Parámetro especifica nombre de impresora para poner por defecto.
        'La pongo por defecto y la quito.

        Dim WshNetwork As Object
        Dim pd As New PrintDocument

        WshNetwork = Microsoft.VisualBasic.CreateObject("WScript.Network")

        Try
            WshNetwork.SetDefaultPrinter(sNombreImpresora)
            pd.PrinterSettings.PrinterName = sNombreImpresora
            If pd.PrinterSettings.IsValid Then
                Return True
            Else
                WshNetwork.SetDefaultPrinter(sNombreImpresora)
                Return False
            End If
        Catch exptd As Exception
            'WshNetwork.SetDefaultPrinter(sNombreImpresora)
            Return False
        Finally
            WshNetwork = Nothing
            pd = Nothing
        End Try
    End Function
    Public RESPUESTAMG As String
    Public Function BDEjecutarSql(ByVal QUERY As String, ByVal CADENA As String) As String

        RESPUESTAMG = "OK"
        Try
            Dim CONX As New SqlClient.SqlConnection(CADENA)
            CONX.Open()
            Dim SQL As New SqlClient.SqlCommand(QUERY, CONX)
            SQL.ExecuteNonQuery()
            SQL.Dispose()
            CONX.Close()
        Catch ex As Exception
            RESPUESTAMG = ex.Message.ToString
        End Try

        Return RESPUESTAMG
    End Function
    Public Function BDEjecutarSql(ByVal QUERY As String, ByVal CADENA As String, ByVal INI As DateTime, ByVal FIN As DateTime) As String
        Dim REG As String
        REG = "OK"
        Try
            Dim CONX As New SqlClient.SqlConnection(CADENA)
            CONX.Open()
            Dim SQL As New SqlClient.SqlCommand(QUERY, CONX)
            SQL.Parameters.Add("@INI", SqlDbType.DateTime).Value = INI
            SQL.Parameters.Add("@FIN", SqlDbType.DateTime).Value = FIN
            SQL.ExecuteNonQuery()
            SQL.Dispose()
            CONX.Close()
        Catch ex As Exception
            REG = ex.Message.ToString
        End Try

        Return REG
    End Function

    Public Function CadenaVacia(ByVal texto As String) As Boolean
        If String.IsNullOrEmpty(Trim(texto)) Then
            Return True
        End If
        Return False
    End Function
    Public Function BDExtraeUnDato(ByVal QUERY As String, ByVal CADENA As String, Optional ByVal DEFAULTNULL As String = "") As String
        Dim REG As String
        REG = ""
        Dim CONX As New SqlClient.SqlConnection(CADENA)
        CONX.Open()
        Dim SQL As New SqlClient.SqlCommand(QUERY, CONX)
        Dim LEC As SqlClient.SqlDataReader
        LEC = SQL.ExecuteReader
        If LEC.Read Then
            If LEC(0) Is DBNull.Value Then
                REG = DEFAULTNULL
            Else
                REG = LEC(0)
            End If
        End If
        LEC.Close()
        SQL.Dispose()
        CONX.Close()
        Return REG
    End Function
    Public Function EsAdministrador() As Boolean
        Thread.GetDomain.SetPrincipalPolicy(PrincipalPolicy.WindowsPrincipal)
        Dim currentPrincipal As WindowsPrincipal = TryCast(Thread.CurrentPrincipal, WindowsPrincipal)
        Return currentPrincipal.IsInRole(WindowsBuiltInRole.Administrator)
    End Function

    Public Function LLENACOMBOBOX2LISTAS(ByRef CB As ComboBox, ByRef LISTA As List(Of String), ByRef LISTA2 As List(Of String), ByVal QUERY As String, ByVal CADENA As String) As Boolean
        CB.Items.Clear()
        LISTA.Clear()
        LISTA2.Clear()
        Dim CONX As New SqlClient.SqlConnection(CADENA)
        CONX.Open()
        Dim SQL As New SqlClient.SqlCommand(QUERY, CONX)
        Dim LEC As SqlClient.SqlDataReader
        LEC = SQL.ExecuteReader
        CB.Items.Clear()
        LISTA.Clear()
        While LEC.Read
            LISTA.Add(LEC(0))
            LISTA2.Add(LEC(1))
            CB.Items.Add(LEC(2))
        End While
        LEC.Close()
        SQL.Dispose()
        CONX.Close()
        Try
            CB.SelectedIndex = 0
            Return True
        Catch ex As Exception

        End Try
        Return False
    End Function
    Public Function LLENACOMBOBOX3LISTAS(ByRef CB As ComboBox, ByRef LISTA As List(Of String), ByRef LISTA2 As List(Of String), ByRef LISTA3 As List(Of String), ByVal QUERY As String, ByVal CADENA As String) As Boolean
        CB.Items.Clear()
        LISTA.Clear()
        LISTA2.Clear()
        LISTA3.Clear()
        Dim CONX As New SqlClient.SqlConnection(CADENA)
        CONX.Open()
        Dim SQL As New SqlClient.SqlCommand(QUERY, CONX)
        Dim LEC As SqlClient.SqlDataReader
        LEC = SQL.ExecuteReader
        CB.Items.Clear()
        LISTA.Clear()
        While LEC.Read
            LISTA.Add(LEC(0))
            LISTA2.Add(LEC(1))
            LISTA3.Add(LEC(2))
            CB.Items.Add(LEC(3))
        End While
        LEC.Close()
        SQL.Dispose()
        CONX.Close()
        Try
            CB.SelectedIndex = 0
            Return True
        Catch ex As Exception

        End Try
        Return False
    End Function
    Public Function LLENACOMBOBOX2(ByRef CB As ComboBox, ByRef LISTA As List(Of String), ByVal QUERY As String, ByVal CADENA As String, ByVal TODOSLOS As String, ByVal CLAPRIM As String) As Boolean
        CB.Items.Clear()
        LISTA.Clear()
        Dim CONX As New SqlClient.SqlConnection(CADENA)
        CONX.Open()
        Dim SQL As New SqlClient.SqlCommand(QUERY, CONX)
        Dim LEC As SqlClient.SqlDataReader
        LEC = SQL.ExecuteReader
        CB.Items.Clear()
        LISTA.Clear()
        CB.Items.Add(TODOSLOS)
        LISTA.Add(CLAPRIM)
        While LEC.Read
            LISTA.Add(LEC(0))
            CB.Items.Add(LEC(1))
        End While
        LEC.Close()
        SQL.Dispose()
        CONX.Close()
        Try
            CB.SelectedIndex = 0
            Return True
        Catch ex As Exception

        End Try
        Return False
    End Function
    Public Function LLENATABLAIF(ByVal QUERY As String, ByVal CADENA As String, ByVal INI As DateTime, ByVal FIN As DateTime) As DataTable
        Dim CONX As New SqlClient.SqlConnection(CADENA)
        CONX.Open()
        Dim DA As New SqlClient.SqlDataAdapter(QUERY, CONX)
        DA.SelectCommand.Parameters.Add("@INI", SqlDbType.DateTime).Value = INI
        DA.SelectCommand.Parameters.Add("@FIN", SqlDbType.DateTime).Value = FIN
        DA.SelectCommand.CommandTimeout = 300
        Dim DT As New DataTable
        DA.Fill(DT)
        Return DT
    End Function
    Public Function InfoHoraEquipo() As String
        Return Format(Now, "hh:mm:ss tt")
    End Function
    Public Function InfoFechayHoraEquipo() As String
        Return Format(Now, "DD/MM/YYYY hh:mm:ss tt")
    End Function
    Public Function InfoFechayHoraEquipo(ByVal Formato As String) As String
        Return Format(Now, Formato)
    End Function
    Public Function InfoFechaEquipo() As String
        Return Format(Now, "DD/MM/YYYY hh:mm:ss tt")
    End Function
    Public Function InfoFechaEquipo(ByVal Formato As String) As String
        Return Format(Now, formato)
    End Function
    Public Function LLENATABLA(ByVal QUERY As String, ByVal CADENA As String) As DataTable
        Dim CONX As New SqlClient.SqlConnection(CADENA)
        CONX.Open()
        Dim DA As New SqlClient.SqlDataAdapter(QUERY, CONX)
        DA.SelectCommand.CommandTimeout = 300
        Dim DT As New DataTable
        DA.Fill(DT)
        Return DT
    End Function

    'Public Function IMPRIMIRREPORTE(ByVal REP As CrystalDecisions.CrystalReports.Engine.ReportDocument, ByVal DT As DataTable, ByVal NCOPIAS As Integer, ByVal NombreImpresora As String) As Boolean
    '    If NombreImpresora = "" Then
    '    Else
    '        For i As Integer = 0 To PrinterSettings.InstalledPrinters.Count - 1
    '            Dim a As New PrinterSettings()
    '            a.PrinterName = PrinterSettings.InstalledPrinters(i).ToString()
    '            If a.PrinterName.ToUpper = NombreImpresora.ToUpper Then
    '                NombreImpresora = PrinterSettings.InstalledPrinters(i).ToString()
    '                REP.PrintOptions.PrinterName = NombreImpresora
    '            End If

    '        Next
    '    End If
    '    REP.SetDataSource(DT)
    '    REP.PrintToPrinter(NCOPIAS, False, 0, 0)
    '    REP.Dispose()
    '    Return False
    'End Function
    'Public Function MOSTRARREPORTE(ByVal REP As CrystalDecisions.CrystalReports.Engine.ReportDocument, ByVal NombreVentana As String, ByVal DT As DataTable, ByVal NombreImpresora As String) As Boolean
    '    Dim FRM As New System.Windows.Forms.Form
    '    Dim CRV As New CrystalDecisions.Windows.Forms.CrystalReportViewer
    '    If NombreImpresora = "" Then
    '    Else
    '        For i As Integer = 0 To PrinterSettings.InstalledPrinters.Count - 1
    '            Dim a As New PrinterSettings()
    '            a.PrinterName = PrinterSettings.InstalledPrinters(i).ToString()
    '            If a.PrinterName.ToUpper = NombreImpresora.ToUpper Then
    '                NombreImpresora = PrinterSettings.InstalledPrinters(i).ToString()
    '                REP.PrintOptions.PrinterName = NombreImpresora
    '            End If

    '        Next
    '    End If
    '    REP.SetDataSource(DT)
    '    CRV.ReportSource = REP
    '    CRV.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    '    CRV.Dock = System.Windows.Forms.DockStyle.Fill
    '    FRM.Controls.Add(CRV)
    '    FRM.Text = NombreVentana
    '    FRM.WindowState = FormWindowState.Maximized
    '    FRM.ShowDialog()
    '    REP.Dispose()
    '    CRV.Dispose()
    '    FRM.Dispose()
    '    Return False
    'End Function
    'Public Function GUARDARREPORTE(ByVal REP As CrystalDecisions.CrystalReports.Engine.ReportDocument, ByVal DT As DataTable, ByVal TIPOEXPORTAR As CrystalDecisions.Shared.ExportFormatType, ByVal ProgramaDefault As String, ByVal ExtensionArchivo As String, ByVal MensajePregunta As String, ByVal NombreArchivo As String, ByVal NombreImpresora As String) As Boolean
    '    Dim xyz As Short
    '    xyz = MessageBox.Show(MensajePregunta, "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
    '    If xyz <> 6 Then
    '        Exit Function
    '    End If
    '    Dim SFD As New System.Windows.Forms.SaveFileDialog
    '    SFD.FileName = NombreArchivo
    '    SFD.Filter = "Archivos de " + ProgramaDefault + " (*." + ExtensionArchivo + ")|*." + ExtensionArchivo + "|" + Chr(34) + "All files (*.*)|*.*;"
    '    If SFD.ShowDialog = DialogResult.OK Then
    '        Try
    '            If System.IO.File.Exists(SFD.FileName) = True Then
    '                System.IO.File.Delete(SFD.FileName)
    '            End If
    '        Catch ex As Exception
    '            MessageBox.Show("La informacion No se puede Guardar, Archivo en Uso", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '            Return False
    '            Exit Function
    '        End Try
    '        If NombreImpresora = "" Then
    '        Else
    '            For i As Integer = 0 To PrinterSettings.InstalledPrinters.Count - 1
    '                Dim a As New PrinterSettings()
    '                a.PrinterName = PrinterSettings.InstalledPrinters(i).ToString()
    '                If a.PrinterName.ToUpper = NombreImpresora.ToUpper Then
    '                    NombreImpresora = PrinterSettings.InstalledPrinters(i).ToString()
    '                    REP.PrintOptions.PrinterName = NombreImpresora
    '                End If

    '            Next
    '        End If
    '        REP.SetDataSource(DT)
    '        REP.ExportToDisk(TIPOEXPORTAR, SFD.FileName)
    '        MessageBox.Show("El archivo ha sido Guardado Exitosamente", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
    '    End If
    '    Return False
    'End Function
    'Public Function GUARDARREPORTEDIRECTO(ByVal REP As CrystalDecisions.CrystalReports.Engine.ReportDocument, ByVal DT As DataTable, ByVal TIPOEXPORTAR As CrystalDecisions.Shared.ExportFormatType, ByVal RutaNombreArchivo As String, ByVal NombreImpresora As String) As Boolean
    '    RESPUESTA = ""
    '    Try
    '        If System.IO.File.Exists(RutaNombreArchivo) = True Then
    '            System.IO.File.Delete(RutaNombreArchivo)
    '        End If
    '    Catch ex As Exception
    '        MessageBox.Show("La informacion No se puede Guardar, Archivo en Uso: " + RutaNombreArchivo, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '        RESPUESTA = ex.Message
    '        Return False
    '        Exit Function
    '    End Try
    '    If NombreImpresora = "" Then
    '    Else
    '        For i As Integer = 0 To PrinterSettings.InstalledPrinters.Count - 1
    '            Dim a As New PrinterSettings()
    '            a.PrinterName = PrinterSettings.InstalledPrinters(i).ToString()
    '            If a.PrinterName.ToUpper = NombreImpresora.ToUpper Then
    '                NombreImpresora = PrinterSettings.InstalledPrinters(i).ToString()
    '                REP.PrintOptions.PrinterName = NombreImpresora
    '            End If

    '        Next
    '    End If
    '    REP.SetDataSource(DT)
    '    Try
    '        REP.ExportToDisk(TIPOEXPORTAR, RutaNombreArchivo)
    '        Return True
    '    Catch ex As Exception
    '        RESPUESTA = ex.Message
    '        Return False
    '    End Try

    '    'MessageBox.Show("El archivo ha sido Guardado Exitosamente", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
    '    Return False
    'End Function
    Public Sub ENVIARMAIL(ByVal LDESTINATARIOS As List(Of String), ByVal ArchivoAdjunto As String, ByVal SUBJECT As String, ByVal BODY As String)
        Dim MSG As New MailMessage
        Dim SMTP As New SmtpClient("smtp.gmail.com", 587)
        'Dim disposition As ContentDisposition
        'disposition = ATT.ContentDisposition
        MSG.From = New MailAddress("correosadjuntos@gmail.com", "correosadjuntos@gmail.com", System.Text.Encoding.UTF8)
        Dim X As Integer
        For X = 0 To LDESTINATARIOS.Count - 1
            MSG.To.Add(LDESTINATARIOS(X))
        Next

        SMTP.Credentials = New System.Net.NetworkCredential("correosadjuntos@gmail.com", "abretesesamo")
        SMTP.Host = "smtp.gmail.com"
        SMTP.EnableSsl = True
        SMTP.Timeout = 300

        If ArchivoAdjunto = "" Then
        Else
            If System.IO.File.Exists(ArchivoAdjunto) = True Then
                Dim ATT As New Attachment(ArchivoAdjunto)
                Try
                    MSG.Attachments.Add(ATT)
                Catch ex As SmtpException
                    Exit Sub
                End Try
            End If
        End If
        MSG.Subject = SUBJECT
        MSG.SubjectEncoding = System.Text.Encoding.UTF8
        MSG.Body = BODY
        Try
            SMTP.Send(MSG)
            MessageBox.Show("Mail Enviado Con Exito ", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        Catch ex As SmtpException
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try
    End Sub
    Public Function SUMARTIEMPOAFECHA(ByVal FECHA As DateTime, ByVal HORA As DateTime) As DateTime
        Dim DT As DateTime
        DT = FECHA.Date
        DT = DT.AddHours(HORA.Hour)
        DT = DT.AddMinutes(HORA.Minute)
        DT = DT.AddSeconds(HORA.Second)
        Return DT
    End Function
    Public Function Image2Bytes(ByVal img As Image) As Byte()
        Dim sTemp As String = Path.GetTempFileName()
        Dim fs As New FileStream(sTemp, FileMode.OpenOrCreate, FileAccess.ReadWrite)
        img.Save(fs, System.Drawing.Imaging.ImageFormat.Png)
        fs.Position = 0
        '
        Dim imgLength As Integer = CInt(fs.Length)
        Dim bytes(0 To imgLength - 1) As Byte
        fs.Read(bytes, 0, imgLength)
        fs.Close()
        Return bytes
    End Function

    Public Function Bytes2Image(ByVal bytes() As Byte) As Image
        If bytes Is Nothing Then Return Nothing
        '
        Dim ms As New MemoryStream(bytes)
        Dim bm As Bitmap = Nothing
        Try
            bm = New Bitmap(ms)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try
        Return bm
    End Function
    '' aqui esta el metodo que creo en el modulo general...
    Public Function str2nothing(ByVal TEXTO As String) As String
        Try
            If Trim(TEXTO) = "" Or String.IsNullOrEmpty(TEXTO) Then
                Return Nothing
            Else
                Return Trim(TEXTO)
            End If
        Catch ex As Exception
            Return Nothing
        End Try

    End Function
    Public Function Truncar(ByVal Numero As Double, Optional ByVal Decimales As Byte = 0) As Double
        Dim lngPotencia As Long
        lngPotencia = 10 ^ Decimales
        Numero = Int(Numero * lngPotencia)
        Truncar = Numero / lngPotencia
    End Function
    Public Function Truncar2(ByVal Numero As Double, Optional ByVal Decimales As Byte = 0) As String
        Dim lngPotencia As Long
        Dim Truncar As Double
        lngPotencia = 10 ^ Decimales
        Numero = Int(Numero * lngPotencia)
        Truncar = Numero / lngPotencia
        Dim FORMATO As String
        FORMATO = "###########0."
        For X = 0 To Decimales - 1
            FORMATO = FORMATO.Insert(13 + X, "0")
        Next

        Dim REG As String
        'REG = Format(Truncar, "###########0.0000")
        REG = Format(Truncar, FORMATO)

        Return REG
    End Function
End Module
