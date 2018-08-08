Imports System.Data
Imports System.Configuration
Imports System.Net.Mail
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.TimeSpan
Imports System.Globalization.CultureInfo
Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop
Imports System.Threading.Thread
Imports System.Globalization
Imports System.IO
Imports Microsoft.Office.Interop.Excel

Public Class _Default
    Inherits System.Web.UI.Page
    Dim ds As DataSet
    Public sConta As Integer
    Public sSQL As String
    Public myCommand As OleDbCommand
    Public dsr As OleDbDataReader
    Public sTotal As Double
    Public ra As Integer
    Public widestData As Integer = 0
    Enum xlsOption
        xlsSaveAs
        xlsOpen
    End Enum
    'Conexão com o MYSQL
    Const ConnStr As String = "Driver={MySQL ODBC 5.1 Driver};" + "Server=172.16.0.29;Database=glpi;uid=root;pwd=sdinhaf12;option=3"


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not Page.IsPostBack Then

            Dim sData = Now()
            Dim Ano As String = Trim(Replace(Mid(sData, 7, 4), "/", ""))
            Dim Mes As String = Trim(Replace(Mid(sData, 4, 2), "/", ""))
            Dim Dia As String = Trim(Replace(Mid(sData, 1, 2), "/", ""))
            If Len(Dia) = 1 Then Dia = "0" + Dia
            If Len(Mes) = 1 Then Mes = "0" + Mes

            TxtDataini.Text = Dia + "/" + Mes + "/" + Ano
            TxtDatafini.Text = Dia + "/" + Mes + "/" + Ano

            widestData = 0

        End If
    End Sub

    Private Sub CarregarGrid()

        Dim myConnection As MySqlConnection
        Dim myDataAdapter As MySqlDataAdapter
        Dim myDataSet As DataSet
        Dim strSQL As String

        Dim Ano As String = Trim(Replace(Mid(TxtDataini.Text, 7, 4), "/", ""))
        Dim Mes As String = Trim(Replace(Mid(TxtDataini.Text, 4, 2), "/", ""))
        Dim Dia As String = Trim(Replace(Mid(TxtDataini.Text, 1, 2), "/", ""))

        Dim sInicio As String = Ano + "-" + Mes + "-" + Dia

        Ano = Trim(Replace(Mid(TxtDatafini.Text, 7, 4), "/", ""))
        Mes = Trim(Replace(Mid(TxtDatafini.Text, 4, 2), "/", ""))
        Dia = Trim(Replace(Mid(TxtDatafini.Text, 1, 2), "/", ""))

        Dim sFinal As String = Ano + "-" + Mes + "-" + Dia

        myConnection = New MySqlConnection("server=172.16.0.29; user id=root; password=sdinhaf12; database=glpi; pooling=false;")
        If sInicio = sFinal Then


            strSQL = "SELECT a.id as ID,a.name as DESCRICAO_PROBLEMA,a.status as Status,  a.date as Data, a.priority  as Prioridade, " &
                     "concat(b.firstname,' ',b.realname) as Requerente, concat(c.firstname,' ',c.realname) as Atribuido, a.solvedate " &
                     "As DataResolucao,a.global_validation As Validacao,a.date_mod  As UltimaAtualizacao,'' as GERENCIA, 0 as Ociosidade, 0 as Espera " &
                     "FROM glpi_tickets a " &
                     "left outer join glpi_users b on b.id = a.users_id_recipient " &
                     "left outer join glpi_users c on c.id = a.users_id_lastupdater where date >= '" & sInicio & "' and (a.name <> '')"

        Else



            strSQL = "Select a.id As ID,a.name As DESCRICAO_PROBLEMA,a.status As Status,  a.Date As Data, a.priority  As Prioridade, " &
                     "concat(b.firstname,' ',b.realname) as Requerente,  concat(c.firstname,' ',c.realname) as Atribuido, a.solvedate " &
                     "As DataResolucao,a.global_validation As Validacao,a.date_mod  As UltimaAtualizacao,'' as GERENCIA, 0 as Ociosidade, 0 as Espera " &
                     "FROM glpi_tickets a " &
                     "left outer join glpi_users b on b.id = a.users_id_recipient " &
                     "left outer join glpi_users c on c.id = a.users_id_lastupdater where date >= '" & sInicio & "' and date <= '" & sFinal & "' and (a.name <> '')"

        End If

        myDataAdapter = New MySqlDataAdapter(strSQL, myConnection)

        myDataSet = New DataSet()
        myDataAdapter.Fill(myDataSet, "test")
        gdItens.DataSource = myDataSet

        'DataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader)

        gdItens.DataBind()



    End Sub
    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click

        If (TxtDataini.Text <> "") And (TxtDatafini.Text <> "") Then
            CarregarGrid()
            BtnSalvar.Visible = True
        Else
            Response.Write("<script language='javascript'>window.alert('PREENCHA AS DATAS!!!');</script>")
        End If


    End Sub

    Protected Sub gdItens_DataBound(ByVal sender As Object, ByVal e As EventArgs) Handles gdItens.DataBound

        Dim sDiasUteis As Integer = 0
        Dim sContador As Integer = 0
        'Preparar as definições para a geração da planilha
        For ContadorLinhas As Integer = 0 To Me.gdItens.Rows.Count - 1

            sContador = sContador + 1
            'Testar se há algum erro
            'If (Me.gdItens.Rows(ContadorLinhas).Cells(0).Text = "4976") Then

            '    Dim Leo As Integer = 1

            'End If

            'STATUS
            If (Me.gdItens.Rows(ContadorLinhas).Cells(2).Text <> "") Then

                If (Me.gdItens.Rows(ContadorLinhas).Cells(2).Text = "6") Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(2).Text = "Fechado"

                End If

                If (Me.gdItens.Rows(ContadorLinhas).Cells(2).Text = "2") Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(2).Text = "Processando"

                End If

                'If (Me.gdItens.Rows(ContadorLinhas).Cells(3).Text = "solved") Then

                '    Me.gdItens.Rows(ContadorLinhas).Cells(3).Text = "Solucionado"

                'End If

                'If (Me.gdItens.Rows(ContadorLinhas).Cells(3).Text = "assign") Then

                '    Me.gdItens.Rows(ContadorLinhas).Cells(3).Text = "Processando (atribuido)"

                'End If

                'If (Me.gdItens.Rows(ContadorLinhas).Cells(3).Text = "pending") Then

                '    Me.gdItens.Rows(ContadorLinhas).Cells(3).Text = "Pendente"

                'End If

                'If (Me.gdItens.Rows(ContadorLinhas).Cells(3).Text = "waiting") Then

                '    Me.gdItens.Rows(ContadorLinhas).Cells(3).Text = "Pendente"

                'End If

                'If (Me.gdItens.Rows(ContadorLinhas).Cells(3).Text = "plan") Then

                '    Me.gdItens.Rows(ContadorLinhas).Cells(3).Text = "Processando (atribuido)"

                'End If



            End If

            'PRIORIDADE
            If (Me.gdItens.Rows(ContadorLinhas).Cells(4).Text <> "") Then

                If (Me.gdItens.Rows(ContadorLinhas).Cells(4).Text = "4") Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(4).Text = "Alta"

                End If

                If (Me.gdItens.Rows(ContadorLinhas).Cells(4).Text = "3") Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(4).Text = "Média"

                End If

            End If

            'REQUERENTE
            If (Me.gdItens.Rows(ContadorLinhas).Cells(5).Text <> "") Then

                Dim sCodigo As String = Me.gdItens.Rows(ContadorLinhas).Cells(0).Text
                Dim myConnectionL As MySqlConnection
                Dim DsX As MySqlDataReader
                Dim myCommand As MySqlCommand

                myConnectionL = New MySqlConnection("server=172.16.0.29; user id=root; password=sdinhaf12; database=glpi; pooling=false;")
                myConnectionL.Open()

                sSQL = "SELECT b.name FROM glpi_tickets_users a " &
                        "left outer join glpi_users b On b.id = a.users_id " &
                        "where a.tickets_id =  '" & sCodigo & "'LIMIT 1"

                myCommand = New MySqlCommand(sSQL, myConnectionL)
                myCommand.CommandTimeout = 0
                DsX = myCommand.ExecuteReader

                If DsX.Read() Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(5).Text = DsX(0)

                End If

                DsX.Close()
                myConnectionL.Close()


            Else

                Dim sCodigo As String = Me.gdItens.Rows(ContadorLinhas).Cells(0).Text
                Dim myConnectionL As MySqlConnection
                Dim DsX As MySqlDataReader
                Dim myCommand As MySqlCommand

                myConnectionL = New MySqlConnection("server=172.16.0.29; user id=root; password=sdinhaf12; database=glpi; pooling=false;")
                myConnectionL.Open()

                sSQL = "SELECT b.name FROM glpi_tickets_users a " &
                        "left outer join glpi_users b On b.id = a.users_id " &
                        "where a.tickets_id =  '" & sCodigo & "'LIMIT 1"

                myCommand = New MySqlCommand(sSQL, myConnectionL)
                myCommand.CommandTimeout = 0
                DsX = myCommand.ExecuteReader

                If DsX.Read() Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(5).Text = DsX(0)

                End If

                DsX.Close()
                myConnectionL.Close()

            End If


            'VALIDAÇÃO
            If (Me.gdItens.Rows(ContadorLinhas).Cells(8).Text <> "") Then

                If (Me.gdItens.Rows(ContadorLinhas).Cells(8).Text = "none") Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(8).Text = "Não esta sujeito a aprovação"

                End If

                If (Me.gdItens.Rows(ContadorLinhas).Cells(8).Text = "accepted") Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(8).Text = "Aceito"

                End If

                If (Me.gdItens.Rows(ContadorLinhas).Cells(8).Text = "waiting") Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(8).Text = "Esperando por uma validação"

                End If

            End If

            'DESENV OU INFRA

            'ATRIBUIDO A VALIDAÇÃO
            If (Me.gdItens.Rows(ContadorLinhas).Cells(6).Text <> "") Then

                If (Me.gdItens.Rows(ContadorLinhas).Cells(6).Text = "Paulo C&#233;zar Cavalcante CONCEI&#199;&#195;O") Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(10).Text = "SUPORTE"

                End If


                If (Me.gdItens.Rows(ContadorLinhas).Cells(6).Text = "Manoel Leonardo Metelis Florindo") Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(10).Text = "APLICATIVOS"

                End If


                If (Me.gdItens.Rows(ContadorLinhas).Cells(6).Text = "Andre Nascimento Sidou") Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(10).Text = "SUPORTE"

                End If



                If (Trim(Me.gdItens.Rows(ContadorLinhas).Cells(6).Text) = "Almir Junior") Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(10).Text = "SUPORTE"

                End If

                If (Trim(Me.gdItens.Rows(ContadorLinhas).Cells(6).Text) = "Bruno da Silva Paulo") Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(10).Text = "SUPORTE"

                End If

                If (Trim(Me.gdItens.Rows(ContadorLinhas).Cells(6).Text) = "George Schnnyder Araujo de Souza") Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(10).Text = "SUPORTE"

                End If

                If (Trim(Me.gdItens.Rows(ContadorLinhas).Cells(6).Text) = "Luiz Renato Batista da Silva") Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(10).Text = "SUPORTE"

                End If

                If (Trim(Me.gdItens.Rows(ContadorLinhas).Cells(6).Text) = "Thomas Nunes") Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(10).Text = "SUPORTE"

                End If

                If (Trim(Me.gdItens.Rows(ContadorLinhas).Cells(6).Text) = "Julie Patricia Pinheiro Moreira") Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(10).Text = "SUPORTE"

                End If

                If (Trim(Me.gdItens.Rows(ContadorLinhas).Cells(6).Text) = "Priscila Macedo Pessoa") Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(10).Text = "SUPORTE"

                End If

                If (Trim(Me.gdItens.Rows(ContadorLinhas).Cells(6).Text) = "Eron Miranda dos Santos") Then

                    Me.gdItens.Rows(ContadorLinhas).Cells(10).Text = "SUPORTE"

                End If



            End If

            'OCIOSIDADE
            If (Me.gdItens.Rows(ContadorLinhas).Cells(9).Text <> "") And (Me.gdItens.Rows(ContadorLinhas).Cells(9).Text <> "&nbsp;") Then

                Dim sDataInicio = Mid(Me.gdItens.Rows(ContadorLinhas).Cells(3).Text, 1, 10)
                Dim sHoje = Mid(Me.gdItens.Rows(ContadorLinhas).Cells(9).Text, 1, 10)
                If (Me.gdItens.Rows(ContadorLinhas).Cells(9).Text <> "") And (Me.gdItens.Rows(ContadorLinhas).Cells(9).Text <> "&nbsp;") Then

                    If sHoje = "" Or sHoje = "&nbsp;" Then
                        sHoje = Mid(Now, 1, 10)
                    End If

                    sDiasUteis = CalculaDiasUteis(sDataInicio, sHoje)

                    Me.gdItens.Rows(ContadorLinhas).Cells(11).Text = CStr(sDiasUteis)

                End If

            End If



            'ESPERA
            If (Me.gdItens.Rows(ContadorLinhas).Cells(3).Text <> "") And (Me.gdItens.Rows(ContadorLinhas).Cells(3).Text <> "&nbsp;") Then

                Dim sDataInicio = Mid(Me.gdItens.Rows(ContadorLinhas).Cells(3).Text, 1, 10)
                Dim sResolvido = Mid(Me.gdItens.Rows(ContadorLinhas).Cells(7).Text, 1, 10)

                If sResolvido = "" Or sResolvido = "&nbsp;" Then
                    sResolvido = Mid(Now, 1, 10)
                End If

                sDiasUteis = CalculaDiasUteis(sDataInicio, sResolvido)

                Me.gdItens.Rows(ContadorLinhas).Cells(12).Text = CStr(sDiasUteis)

            End If

        Next

        lblAviso.Text = "Registros gerados: " + CStr(sContador)
        lblAviso.Visible = True


    End Sub

    Public Function CalculaDiasUteis(ByVal DataIni, ByVal Datafinal) As Integer
        Dim cont As Integer

        cont = 0
        Dim Dt As Date = CDate(DataIni)

        While Dt <= CDate(Datafinal)

            If Weekday(Dt) <> 6 And Weekday(Dt) <> 7 Then
                cont = cont + 1
            End If

            Dt = Dt.AddDays(1)

        End While

        CalculaDiasUteis = cont


    End Function

    Protected Sub gdItens_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gdItens.RowDataBound

        Dim colWidth As Integer
        colWidth = 250

        For i As Integer = 0 To gdItens.Columns.Count - 1
            If (i = 7) Then
                gdItens.Columns(i).ItemStyle.Width = colWidth
                gdItens.Columns(i).ItemStyle.Wrap = False
            End If
        Next

    End Sub

    Protected Sub BtnSalvar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles BtnSalvar.Click

        Dim sData As String = CStr(Now)
        Dim Ano As String = Trim(Replace(Mid(sData, 7, 4), "/", ""))
        Dim Mes As String = Trim(Replace(Mid(sData, 4, 2), "/", ""))
        Dim Dia As String = Trim(Replace(Mid(sData, 1, 2), "/", ""))
        Dim sArquivo = "KPI_" + DateTime.Now.ToString("ddMMyy") + "_" + DateTime.Now.ToString("hhmmss") + ".xls"

        'Exporta a Grid para o Excel
        exportarExcel(sArquivo)

    End Sub
    Sub exportarExcel(ByVal sNomeArquivo As String)

        Dim tw As New StringWriter()
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        Dim frm As HtmlForm = New HtmlForm()

        Response.ContentType = "application/vnd.ms-excel"
        Response.AddHeader("content-disposition", "attachment;filename=" + sNomeArquivo)
        Response.Charset = ""
        EnableViewState = False

        Controls.Add(frm)
        frm.Controls.Add(gdItens)
        frm.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

    End Sub

    Protected Sub gdItens_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gdItens.PageIndexChanging

        gdItens.PageIndex = e.NewPageIndex
        CarregarGrid()

    End Sub

 
    Protected Sub gdItens_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gdItens.RowCreated

        'AJUSTAR DINAMICAMENTE O TAMANHO DAS COLUNAS DENTRO DO DATAGRIDVIEW
        e.Row.Cells(1).Width = New Unit(100, UnitType.Mm)
        e.Row.Cells(1).Wrap = False

        e.Row.Cells(2).HorizontalAlign = HorizontalAlign.Center
        e.Row.Cells(2).Wrap = False

        e.Row.Cells(3).Width = New Unit(100, UnitType.Mm)
        e.Row.Cells(3).HorizontalAlign = HorizontalAlign.Center
        e.Row.Cells(3).Wrap = False

        e.Row.Cells(4).Width = New Unit(100, UnitType.Mm)
        e.Row.Cells(4).Wrap = False

        e.Row.Cells(5).HorizontalAlign = HorizontalAlign.Center
        e.Row.Cells(5).Wrap = False

        e.Row.Cells(6).Width = New Unit(76, UnitType.Mm)
        e.Row.Cells(6).Wrap = False

        e.Row.Cells(7).Width = New Unit(76, UnitType.Mm)
        e.Row.Cells(7).Wrap = False

        e.Row.Cells(8).Width = New Unit(76, UnitType.Mm)
        e.Row.Cells(8).Wrap = False
        e.Row.Cells(8).Visible = False

        e.Row.Cells(9).Width = New Unit(76, UnitType.Mm)
        e.Row.Cells(9).Wrap = False

        e.Row.Cells(10).Width = New Unit(76, UnitType.Mm)
        e.Row.Cells(10).Wrap = False

        e.Row.Cells(11).Width = New Unit(76, UnitType.Mm)
        e.Row.Cells(11).Wrap = False

        e.Row.Cells(12).HorizontalAlign = HorizontalAlign.Center
        e.Row.Cells(12).Wrap = False

        'e.Row.Cells(13).HorizontalAlign = HorizontalAlign.Center
        'e.Row.Cells(13).Wrap = False

        'e.Row.Cells(14).HorizontalAlign = HorizontalAlign.Center
        'e.Row.Cells(14).Wrap = False


    End Sub
End Class