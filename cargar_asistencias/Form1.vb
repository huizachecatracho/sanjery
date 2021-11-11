'******************************************************************************************************
' Aplicacion    : Carga de registros para la asistencia de apilac                                     *
' Fecha         : Diciembre 11 de 2018                                                                *
' Autor         : Manuel Salvador Ramirez Camacho                                                     *
' Lugar         : APILAC                                                                              *
' Observaciones : Esta aplicacion Ejecuta un Procedimiento Almacenado en el servidor 10.34.67.10      *
'                 EXEC pa_ExportAccessEvent 'AAAA-MM-DD'                                              *
'                 y el resultado lo carga a una base de datos en el servicor 10.34.68.34              *
'                 en la tabla LoggedEventDay34                                                        *
'******************************************************************************************************
Public Class Form1
    Public bdlenl34 As New lenel34 'Instancia de la clase para la BD del servidor 10.34.68.34
    Public bdlenel As New bdlenl10 'Instancia de la clase para la BD del servidor 10.34.67.10
    Private Sub MonthCalendar1_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Button1.Visible = False
        ProgressBar1.Visible = True
        Dim stardate As Date = MonthCalendar1.SelectionStart ' Dia Inicial
        Dim enddate As Date = MonthCalendar1.SelectionEnd    ' Dia Final
        Dim Dias As Double = DateDiff(DateInterval.Day, stardate, enddate) ' Numero de dias 
        Dim Avance As Double = 100 ' Avance para la barra de Progreso

        If Dias <> 0 Then
            Avance = 100 / Dias
        End If

        Do While stardate <= enddate 'Ciclo para los dias 
            Carga(stardate) ' Funcion para cargar el dia 
            stardate = stardate.AddDays(1)

            If ProgressBar1.Value < 100 Then
                ProgressBar1.Value += Avance ' Avance de la barra de progreso
            End If
        Loop

        Button1.Visible = True
        ProgressBar1.Visible = False
        MsgBox("La asistencia ha sido Cargada")
    End Sub

    Public Sub Carga(Dia As Date)
        'conectar a lenel y a la bd del 34
        bdlenel.conectar()
        bdlenl34.conectar()

        Dim Elimina_dia As String
        Elimina_dia = "Delete LoggedEventDay34 where Year(Fecha) = " & Year(Dia) & " and Month(Fecha) = " & Month(Dia) & "And Day(Fecha) = " & (Dia.Day)
        bdlenl34.ejecuta_sentencia(Elimina_dia) ' Se elimina el dia de acuerdo al ciclo 

        Dim Consulta As String
        Dim Inserta As String
        Consulta = "EXEC pa_ExportAccessEvent '" & Format(Dia, "yyyy-MM-dd") & "'" ' Se Arma la Instruccion para ejecutar el Procedimiento almacenado
        bdlenel.consultar(Consulta) ' Se ejecuta el procedimiento almacenado
        Do While bdlenel.reader.Read ' Ciclo de lectura y se arma la instruccion para insertar
            Inserta = "insert into LoggedEventDay34 values('"
            Inserta = Inserta & bdlenel.reader("Descripcion").ToString & "','"
            Inserta = Inserta & bdlenel.reader("Apellidos").ToString & "','"
            Inserta = Inserta & bdlenel.reader("Nombres").ToString & "','"
            Inserta = Inserta & bdlenel.reader("DEVICE").ToString & "','"
            Inserta = Inserta & bdlenel.reader("BADGEID").ToString & "','"
            Inserta = Inserta & Format(bdlenel.reader("Fecha"), "dd/MM/yyyy HH:mm:ss").ToString & "','"
            Inserta = Inserta & bdlenel.reader("n_empleado").ToString & "','"
            Inserta = Inserta & bdlenel.reader("IMSS").ToString & "','"
            Inserta = Inserta & bdlenel.reader("PLACA").ToString & "')"
            bdlenl34.ejecuta_sentencia(Inserta) ' Se ejecuta la instruccion para insertar en la bd del servidor 10.34.68.21
            ' Debug.Print(Inserta.ToString)
        Loop

        Dim Actualiza_datos As String
        Actualiza_datos = "update LoggedEventDay34 set n_empleado = '00' + n_empleado where isnumeric(n_empleado)<> 0 and LEN(n_empleado) = 3 "
        bdlenl34.ejecuta_sentencia(Actualiza_datos) ' Se actualiza el numero del empleado 

        'Cerramos las Bases de datos Lenel y 34
        bdlenel.cerrar()
        bdlenl34.cerrar()
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' ProgressBar1.Visible = False
    End Sub
End Class
