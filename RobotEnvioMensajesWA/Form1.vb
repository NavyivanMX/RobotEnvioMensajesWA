Imports Newtonsoft
Imports Newtonsoft.Json

Public Class frmPrincipal
    Dim CONT As Integer
    Dim ContResp As Integer
    Public IP, Sistema, Version As String
    Public CadenaConexion As String
    Public CONX As New SqlClient.SqlConnection
    Dim BD, USER, PWDBD As String
    Dim LLOGS As New List(Of String)

    Private Sub frmPrincipal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Sistema = "Timbrado"
        Version = "BETA"
        IP = "LOCALHOST"
        IP = "navyserver.ddns.net"
        LBD.Clear()
        LBD.Add("MensajeoWa")
        'CBBD.Items.Clear()
        ''IP = " 127.0.0.1"
        IP = "structureserver3.ddns.net"
        BD = "MensajeoWa"
        USER = "MarosIlocal"
        PWDBD = "Loco*1234"


        'IP = "navyserver.ddns.net"
        'BD = "MensajeoWa"
        'USER = "dbaadmin"
        'PWDBD = "Xoporte1234."


        CadenaConexion = "Data Source=" + IP + ",1433;Network Library=DBMSSOCN;Initial Catalog=" + BD + ";User ID=" + USER + ";Password=" + PWDBD + ""
        'CadenaConexion = "Data Source=" + IP + ",1433;Network Library=DBMSSOCN;Initial Catalog=FACTURACIONELECTRONICA;User ID=" + USER + ";Password=" + PWDBD + ""
        Label1.Text = ""
        CONX.ConnectionString = CadenaConexion
        SB.Items(0).Text = "Los Mochis, Sin. Hoy es " + FormatDateTime(Now, DateFormat.LongDate) + " "

        Dim FL As New System.IO.FileInfo(Application.ExecutablePath)
        LBLEJE.Text = "Ejecutando desde: " + FL.DirectoryName + " " + IP
        CONX.Open()
        CHECACONX()
        CONT = 6


        If EsAdministrador() Then
            Label2.Text = "Es Administrador"
        Else
            Label2.Text = "NO Es Administrador"
            'Me.Close()
        End If
        Timer1.Start()
    End Sub
    Dim TEMPLATEID, PHONENUMBER, JSONPARAMS, MENSAJEID As String
    Private Sub EnviarMensajes()
        If Not CHECACONX() Then
            Exit Sub
        End If
        Dim DT As New DataTable

        DT = LLENATABLA("SELECT TOP 50  TEMPLATEID,PHONENUMBER,JSONPARAMS,ID  FROM MENSAJES  with (nolock) WHERE SENDED=0", CadenaConexion)
        Dim X As Integer
        For X = 0 To DT.Rows.Count - 1
            TEMPLATEID = DT.Rows(X).Item(0).ToString
            PHONENUMBER = DT.Rows(X).Item(1).ToString
            JSONPARAMS = DT.Rows(X).Item(2).ToString
            MENSAJEID = DT.Rows(X).Item(3).ToString

            WriteLog("Enviando mensaje to " + PHONENUMBER + Format(Now, "hh:mm:ss tt"))
            ACTUALIZALOG()
            Dim PARAMS As Dictionary(Of String, String) = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(JSONPARAMS)
            Dim resp As New WAMessageResponse
            resp = EnviarMsgWA(TEMPLATEID, PHONENUMBER, PARAMS)
            If resp.Sid <> "" Then
                WriteLog("Enviando con éxito " + PHONENUMBER + " " + Format(Now, "hh:mm:ss tt"))
                ACTUALIZALOG()
                BDEjecutarSql("update Mensajes SET Sended=1,SendId='" + resp.Sid + "',DateSent=GETDATE(),Status='" + resp.Status + "',ErrorCode='" + resp.ErrorCode.ToString + "',ErrorMessage='" + resp.ErrorMessage + "' WHERE Id='" + MENSAJEID + "'", CadenaConexion)
            End If
        Next

    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        SB.Items(1).Text = "Envio de Mensajes WA   " + Format(Now, "hh:mm:ss tt")
        CONT = CONT - 1

        If CONT <= 0 Then

            Timer1.Stop()
            For Each BaseDatos As String In LBD
                Try
                    CONX.Close()
                    CadenaConexion = "Data Source=" + IP + ",1433;Network Library=DBMSSOCN;Initial Catalog=" + BaseDatos + ";User ID=" + USER + ";Password=" + PWDBD + ""
                    CONX.ConnectionString = CadenaConexion

                    CONX.Open()
                    WriteLog("Revisando en " + BaseDatos)
                    If HayMensajesSinEnviar() Then
                        WriteLog("Hay mensajes para enviar " + Format(Now, "hh:mm:ss tt"))
                        ACTUALIZALOG()
                        EnviarMensajes()
                    Else
                        WriteLog("No hay mensajes para enviar " + Format(Now, "hh:mm:ss tt"))
                        ACTUALIZALOG()

                    End If

                Catch ex As Exception
                    WriteLog(ex.Message + "  " + Format(Now, "hh:mm:ss tt"))
                    ACTUALIZALOG()
                End Try
            Next

            CONT = 30 '0
            ACTUALIZALOG()

            Timer1.Start()
        End If
    End Sub

    Dim LBD As New List(Of String)


    Public Function CHECACONX() As Boolean
        If Me.CONX.State = ConnectionState.Closed Or Me.CONX.State = ConnectionState.Broken Then
            Try
                Me.CONX.Open()
            Catch ex As Exception
                MessageBox.Show("La Conexión NO esta realizada, La Informacion No se ha Procesado, Intente en un momento por Favor", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return False
            End Try
        End If
        Return True
    End Function

    Public Sub WriteLog(ByVal message As String)
        If Not message Is Nothing Then
            LLOGS.Add(message)
        End If

    End Sub
    Private Sub ACTUALIZALOG()
        Dim X As Integer
        For X = 0 To LLOGS.Count - 1
            Me.LBLOGO.Items.Add(LLOGS(X))
        Next
        LBLOGO.SelectedIndex = LBLOGO.Items.Count - 1
        LBLOGO.ClearSelected()
        LLOGS.Clear()
    End Sub

    Private Function HayMensajesSinEnviar() As Boolean
        'Return True
        ' Return False
        If Not CHECACONX() Then
            Return False
        End If
        Dim CUANTOS As Integer
        Dim SQL As New SqlClient.SqlCommand("SELECT COUNT (Id) FROM Mensajes with (nolock)  WHERE Sended=0 and DateCreated>=dateadd(day,-3,getdate()) ", CONX)
        Dim LEC As SqlClient.SqlDataReader
        LEC = SQL.ExecuteReader
        If LEC.Read Then
            CUANTOS = LEC(0)
        End If
        LEC.Close()
        SQL.Dispose()
        If CUANTOS >= 1 Then
            Return True
        Else
            Return False
        End If
    End Function
End Class
