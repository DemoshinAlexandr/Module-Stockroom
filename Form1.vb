'подключение библиотек
Imports System.Math
Imports Microsoft.Office
Imports DevExpress.XtraSplashScreen
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid
Imports System.Data.SqlClient

'главная форма проекта
Partial Public Class Form1

	'подключение, настройка, инициализация компоненотов DevExpress
    Inherits DevExpress.XtraBars.Ribbon.RibbonForm

    Shared Sub New()
        DevExpress.UserSkins.BonusSkins.Register()
        DevExpress.Skins.SkinManager.EnableFormSkins()
    End Sub
	
    Public Sub New()
        InitializeComponent()
    End Sub

    Public ListIndex As Integer
    Public PayInSlip As String
    Public skladSelected As Boolean
    Public ArrayTag() As String
    Private downHitInfo As GridHitInfo = Nothing

    'загрузка формы
    Public Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            gModuleName = "Склад"
            Me.Visible = False
            frmLogin.ShowDialog() 'показ формы проверки пользователя
            If gIsPasswordTrue Then
                StartForm.ShowDialog() 'показ формы выбора склада
                If skladSelected = False Then
                    Application.Exit()
                    Exit Sub
                End If
            Else
                Application.Exit()
                Exit Sub
            End If
            Me.WindowState = FormWindowState.Maximized
            Me.Text = "Модуль ""Склад"" - " & gWHS_NAME.ToString 'название
            SQLEx("select convert (date,getdate())") 'получение серверного времени
            While reader.Read
                gDate = (Convert.ToString(reader.Item(0)))
                gDate = gDate.Substring(0, 10)
            End While
            ConClose()
            opt = 1
            RadioButtonSum.Checked = True
            ЗаполнениеКомбобокcа(ComboBox1, "SELECT WareType, WareTypeName FROM Groups")
            XtraTabControl1.TabPages.Item(0).Show()
            XtraTabControl1.TabPages.Item(6).PageVisible = False
            XtraTabControl1_Click(sender, e)
            BigButton.Left = XtraTabControl1.Left
            BigButton.Width = XtraTabControl1.Width - 5
            Me.Show()

        Catch Ex As Exception
            MsgBox(Ex.Message, , "Ошибка при загрузке главной формы")
        End Try

    End Sub


	'функция для заполнения грида с заказами
    Public Sub отрисовка_заказов(Optional ByVal WareCode As Integer = 0)
        Try
            Dim where As String
            If WareCode <> 0 Then
                where = " where (o.deleted  = 0)  AND W.WareCode = " & WareCode.ToString & " "
            Else
                where = " where (o.deleted  = 0)"
            End If
            dt = New DataTable
            dt.Columns.Add("Orders_code", GetType(String))
            dt.Columns.Add("Организация-поставщик", GetType(String))
            dt.Columns.Add("Дата заявки", GetType(String))
            dt.Columns.Add("Дата оплаты", GetType(String))
            dt.Columns.Add("НДС(%)", GetType(String))
            dt.Columns.Add("НДС?", GetType(String))
            dt.Columns.Add("Номер счета", GetType(String))
            dt.Columns.Add("Дата счета", GetType(String))
            dt.Columns.Add("Номер вход. док-та", GetType(String))
            dt.Columns.Add("Дата док-та", GetType(String))
            dt.Columns.Add("Комментарий", GetType(String))

            SQLEx("SELECT DISTINCT  O.Orders_code, " &
              "      	C.cli_name As [Организация-поставщик], " &
              "     	O.orders_date As [Дата заявки],  " &
              "    		O.order_note As [Платежные док-ты],  " &
              "     	O.orders_date_pay as [Дата оплаты],  " &
              "     	ISNULL (O.NDS, 0) As [НДС(%)],  " &
              "     	ISNULL(O.NDS_inc, 'нет') As [НДС?],  " &
              "     	O.InvoiceNumber As [Номер счета], " &
              "     	O.InvoiceDate As [Дата счета], " &
              "     	D.doc_text As [Номер вход. док-та],  " &
              "     	ISNULL(D.DocDT, D.data_make) As [Дата док-та],  " &
              "     	ISNULL(D.DocNote, '') As Комментарий" &
              "		 FROM   Orders AS O " &
              "      INNER JOIN (SELECT OB.orders_code " &
             "                        FROM   OrdersBody AS OB " &
             "                        INNER JOIN WareBalance AS WB ON OB.ware_code = WB.WareCode " &
             "                        WHERE  ( WB.smena_no = " & gWHS & " ) " &
             "                        GROUP  BY OB.orders_code) AS qWB ON O.orders_code = qWB.orders_code " &
             "       LEFT OUTER JOIN Documents AS D ON O.doc_code = D.doc_code " &
             "       LEFT OUTER JOIN Clients AS C ON O.cli_code = C.cli_code " &
             "       INNER JOIN OrdersBody AS OB ON O.orders_code = OB.orders_code " &
             "       INNER JOIN Wares AS W ON OB.ware_code = W.WareCode " & where)

            While reader.Read
                dt.Rows.Add()
                dt.Rows(dt.Rows.Count() - 1).Item(0) = reader.Item("Orders_code")
                dt.Rows(dt.Rows.Count() - 1).Item(1) = reader.Item("Организация-поставщик")
                dt.Rows(dt.Rows.Count() - 1).Item(2) = reader.Item("Дата заявки")
                dt.Rows(dt.Rows.Count() - 1).Item(3) = reader.Item("Дата оплаты")
                dt.Rows(dt.Rows.Count() - 1).Item(4) = reader.Item("НДС(%)")
                dt.Rows(dt.Rows.Count() - 1).Item(5) = reader.Item("НДС?")
                dt.Rows(dt.Rows.Count() - 1).Item(6) = reader.Item("Номер счета")
                dt.Rows(dt.Rows.Count() - 1).Item(7) = reader.Item("Дата счета")
                dt.Rows(dt.Rows.Count() - 1).Item(8) = reader.Item("Номер вход. док-та")
                dt.Rows(dt.Rows.Count() - 1).Item(9) = reader.Item("Дата док-та")
                dt.Rows(dt.Rows.Count() - 1).Item(10) = reader.Item("Комментарий")
            End While
            ConClose()

        Catch Ex As Exception
            ConClose()
            MsgBox(Ex.Message)
        End Try
    End Sub


    Public Sub отрисовка_прихода()
        Try
            dt = New DataTable
            SQLEx("SELECT doc_code, doc_type, PartnerCode, Название, Поставщик, Дата, [НДС (%)], [В том числе НДС?], Комментарий FROM ft_Sklad_DocList_In (" & gWHS & ") ORDER BY Дата")
            dt.Load(reader)
            ConClose()
        Catch Ex As Exception
            ConClose()
            MsgBox(Ex.Message)
        End Try
    End Sub

    Public Sub отрисовка_остатков(ByVal t As Integer)
        Try
            Dim SQL As String = ""
            ConClose()
            If opt = 1 Then
                dt = New DataTable
                dt.Columns.Add("Warecode", GetType(String))
                dt.Columns.Add("WareType", GetType(String))
                dt.Columns.Add("Тип", GetType(String))
                dt.Columns.Add("Наименование", GetType(String))
                dt.Columns.Add("Ед", GetType(String))
                dt.Columns.Add("Ост.", GetType(Double))
                dt.Columns.Add("Вес(кг)", GetType(Double))
                dt.Columns.Add("В отправке", GetType(Double))
                dt.Columns.Add("Цена", GetType(Double))
                dt.Columns.Add("Исх. треб-я", GetType(String))
                dt.Columns.Add("Приход по заявке", GetType(String))
                dt.Columns.Add("Короткое наим.", GetType(String))
                SQL = "SELECT Warecode, WareType, CASE WHEN WareType = 2 THEN 'Ст.' WHEN WareType = 1 THEN 'Мат.' END as Тип, Warename As Наименование, OKEI_Name As Ед, Num As [Ост.], Weight As [Вес(кг)], " &
                      " NumSend As [В отправке], Price As Цена, OutDoc As [Исх. треб-я], InDoc As [Приход по заявке], ' ' as [Короткое наим.] FROM ft_Sklad_Balance ( " & gWHS & ", " & t & " ) ORDER BY WareType, WareName"
                SQLEx(SQL)
                While reader.Read
                    dt.Rows.Add()
                    dt.Rows(dt.Rows.Count() - 1).Item(0) = reader.Item("Warecode")
                    dt.Rows(dt.Rows.Count() - 1).Item(1) = reader.Item("WareType")
                    dt.Rows(dt.Rows.Count() - 1).Item(2) = reader.Item("Тип")
                    dt.Rows(dt.Rows.Count() - 1).Item(3) = reader.Item("Наименование")
                    dt.Rows(dt.Rows.Count() - 1).Item(4) = reader.Item("Ед")
                    dt.Rows(dt.Rows.Count() - 1).Item(5) = reader.Item("Ост.")
                    dt.Rows(dt.Rows.Count() - 1).Item(6) = reader.Item("Вес(кг)")
                    dt.Rows(dt.Rows.Count() - 1).Item(7) = reader.Item("В отправке")
                    dt.Rows(dt.Rows.Count() - 1).Item(8) = Round(reader.Item("Цена"), 2)
                    dt.Rows(dt.Rows.Count() - 1).Item(9) = reader.Item("Исх. треб-я")
                    dt.Rows(dt.Rows.Count() - 1).Item(10) = reader.Item("Приход по заявке")
                    dt.Rows(dt.Rows.Count() - 1).Item(11) = reader.Item("Короткое наим.")
                End While
                ConClose()
            ElseIf opt = 2 Then
                dt = New DataTable
                SQL = "SELECT Warecode, WareType, CASE WHEN WareType = 2 THEN 'Ст.' WHEN WareType = 1 THEN 'Мат.' END as Тип, Warename As Наименование, OKEI_Name As Ед, Num As [Остаток(Все ячейки)], NumSend as [Остаток в ячейке], RackName As [Стеллаж], CellName As [Ячейка], " &
                      " MaxNum As [Вместительность], WareShortName as [Короткое наименование] FROM ft_Sklad_Balance_Cells ( " & gWHS & ", " & t & " ) ORDER BY WareType, WareName, ColSort "
                SQLEx(SQL)
                dt.Load(reader)
                ConClose()
            ElseIf opt = 3 Then
                dt = New DataTable
                SQL = "SELECT Warecode, WareType, CASE WHEN WareType = 2 THEN 'Ст.' WHEN WareType = 1 THEN 'Мат.' END as Тип, Warename As Наименование, OKEI_Name As Ед, WhsNum As [Остаток на складе], Num as [Незакрепленный остаток], WareShortName as [Короткое наим.] " &
                      "FROM ft_Sklad_Balance_Unfixed ( " & gWHS & ", " & t & " ) ORDER BY WareType, WareName "
                SQLEx(SQL)
                dt.Load(reader)
                ConClose()
            End If
        Catch Ex As Exception
            ConClose()
            MsgBox(Ex.Message)
        End Try
    End Sub

    'ОБНОВЛЕНИЕ СПИСКА РАСХОДНЫХ ДОКУМЕНТОВ
    Public Sub отрисовка_расходов()
        Try
            dt = New DataTable
            dt.Columns.Add("doc_type", GetType(String))
            dt.Columns.Add("Код смены", GetType(String))
            dt.Columns.Add("doc_code", GetType(String))
            dt.Columns.Add("№", GetType(String))
            dt.Columns.Add("Дата", GetType(Date))
            dt.Columns.Add("Получатель", GetType(String))
            dt.Columns.Add("Заказ", GetType(String))
            dt.Columns.Add("Комментарий", GetType(String))

            SQLEx("SELECT * FROM ft_Sklad_DocList_Out ( " & gWHS & ", " & IIF_S(CBool(gWareCode), CStr(gWareCode), "NULL") & " ) ORDER BY Номер")

            While reader.Read()
                dt.Rows.Add()
                dt.Rows(dt.Rows.Count() - 1).Item(0) = reader.Item("doc_type")
                dt.Rows(dt.Rows.Count() - 1).Item(1) = reader.Item("MasterCode")
                dt.Rows(dt.Rows.Count() - 1).Item(2) = reader.Item("doc_code")
                dt.Rows(dt.Rows.Count() - 1).Item(3) = reader.Item("номер")
                If Not reader.Item("Дата") = "" Then
                    dt.Rows(dt.Rows.Count() - 1).Item(4) = reader.Item("Дата")
                Else
                    dt.Rows(dt.Rows.Count() - 1).Item(4) = DBNull.Value
                End If
                dt.Rows(dt.Rows.Count() - 1).Item(5) = reader.Item("V_user")
                dt.Rows(dt.Rows.Count() - 1).Item(6) = reader.Item("ContractNo")
                dt.Rows(dt.Rows.Count() - 1).Item(7) = reader.Item("sNote")
            End While
            ConClose()
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Public Sub отрисовка_поставщиков()
        Try
            dt = New DataTable
            SQLEx("SELECT cli_name as [Наименование поставщика], cli_code FROM clients WHERE data_del IS NULL ORDER BY cli_name")
            dt.Load(reader)
            ConClose()
        Catch Ex As Exception
            ConClose()
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Public Sub отрисовка_накладных(ByVal a As Boolean, Optional ByVal treb As Boolean = False)
        Try
            Dim where As String = ""
            If a = False Then
                where = "WHERE  D.data_reg BETWEEN ' " & gDate & " ' AND ' " & gDate & gDay & "' "
            Else
                where = "WHERE  D.data_reg BETWEEN ' " & DTPicker1.Value.ToString.Substring(0, 10) & " ' AND ' " & DTPicker1.Value.ToString.Substring(0, 10) & gDay & "' "
            End If
            dt = New DataTable
            If treb = False Then
                SQLEx("SELECT D.doc_code as doc_code, " &
        "       ISNULL(( CASE D.doc_type " &
        "                 WHEN 1 THEN 'Приходная накл.' " &
        "        WHEN 3 THEN 'Расходная накл.' " &
        "        WHEN 100 THEN 'Возврат поставщику' " &
        "           WHEN 101 THEN 'Возврат со смены' " &
        "           WHEN 22 THEN ( CASE " &
        "                          WHEN D.MasterCode = " & gWHS & " THEN 'Пе редача на склад (Расход)' " &
        "                         WHEN D.cli_code = " & gWHS & " THEN 'Приход со склада (Приход)' " &
        "                                 END ) " &
        "                END ), '')               AS Документ, " &
        "       ISNULL(D.doc_text, '')            AS Название, " &
        "       CONVERT(VARCHAR, D.data_reg, 104) AS Дата, " &
        "       CASE " &
        "         WHEN D.doc_type = 1 THEN C.cli_name " &
        "         WHEN D.doc_type = 100 THEN C.cli_name " &
        "         WHEN D.doc_type = 3 THEN VW.[Смена] " &
        "         WHEN D.doc_type = 101 THEN VW.[Смена] " &
        "         WHEN D.doc_type = 22 THEN ( CASE " &
        "                                       WHEN D.MasterCode = " & gWHS & " THEN S_Client.StructureName " &
        "                                       WHEN D.cli_code = " & gWHS & " THEN S_Master.StructureName " &
        "                                     END ) " &
        "       END                               AS Контрагент, " &
        "       ISNULL(D.DocNote, '')     AS Комментарий, " &
        "       ISNULL(D.PayInSlip, '')   AS [Приходный ордер], " &
              " ISNULL(REVERSE(STUFF(REVERSE((SELECT DISTINCT ContractNo + ',' FROM Contracts AS C JOIN NaklContract AS NC ON NC.ContractID = C.ContractID JOIN Requirement AS R ON R.nakl_no = NC.nakl_no JOIN DocumentDocument AS DD ON DD.ParentID = R.doc_code WHERE DD.ChildID = D.doc_code FOR XML PATH (''))),1,1,'')), '') AS Заказ " &
        " FROM   Documents AS D " &
        "       LEFT OUTER JOIN vwViewSmenaNames AS VW ON D.MasterCode = VW.StructureID " &
        "       LEFT OUTER JOIN Structure AS S_Client ON D.cli_code = S_Client.StructureID " &
        "       LEFT OUTER JOIN Structure AS S_Master ON D.MasterCode = S_Master.StructureID " &
        "       LEFT OUTER JOIN Clients AS C ON D.cli_code = C.cli_code " &
        where &
        "   AND D.doc_type IN ( 1, 3, 101, 100, 22 )  " &
        "   AND CASE " &
        "         WHEN D.doc_type = 1 " &
        "              AND D.MasterCode = " & gWHS & " THEN 1 " &
        "         WHEN D.doc_type = 100 " &
        "              AND D.MasterCode = " & gWHS & " THEN 1 " &
        "         WHEN D.doc_type = 3 " &
        "              AND D.cli_code = " & gWHS & " THEN 1 " &
        "         WHEN D.doc_type = 101 " &
        "              AND D.cli_code = " & gWHS & " THEN 1 " &
        "         WHEN D.doc_type = 22 " &
        "              AND ( D.cli_code = " & gWHS & " " &
        "                     OR D.MasterCode = 640 ) THEN 1 " &
        "         ELSE 0 " &
        "       END = 1 " &
        " ORDER  BY D.data_reg ")
            Else
                SQLEx("SELECT RName AS [Наименование], DTMod AS [Дата], RID FROM [ОАОКЭТЗ_FILES]..Reports WHERE ParentFolder = 16 AND DTMod BETWEEN '" & DTPicker1.Value.ToString.Substring(0, 10) & "' AND '" & DTPicker1.Value.ToString.Substring(0, 10) & gDay & " '")
            End If
            dt.Load(reader)
            ConClose()

        Catch Ex As Exception
            ConClose()
            MsgBox(Ex.Message)
        End Try
    End Sub
	

    Public Sub Отрисовка_возврата()
        Try
            Dim SQL As String
            dt = New DataTable

            SQL = ""
            SQL = SQL & "SELECT D.doc_text  AS Название, "
            SQL = SQL & "       C.cli_name  AS Поставщик, "
            SQL = SQL & "       ISNULL(D.DocDT, D.data_make) AS Дата, "
            SQL = SQL & "       D.NDS       AS [НДС (%)], "
            SQL = SQL & "       D.NDS_inc   AS [В том числе НДС?], "
            SQL = SQL & "       ISNULL(D.DocNote, '') Комментарий, "
            SQL = SQL & "       D.doc_code as [doc_code], "
            SQL = SQL & "       D.cli_code "
            SQL = SQL & "FROM   Documents AS D "
            SQL = SQL & "       LEFT OUTER JOIN Clients AS C ON D.cli_code = C.cli_code "
            SQL = SQL & "WHERE  D.doc_type = 100 "
            SQL = SQL & "   AND D.data_reg IS NULL "
            SQL = SQL & "   AND D.MasterCode = " & gWHS

            SQLEx(SQL)
            dt.Load(reader)
            ConClose()
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Sub пересчет()
        Try
            Dim V_code As String
            Dim i As Integer = 0
            GridViewRashodClick.Columns.Clear()
            dt = New DataTable
            V_code = GridViewRashod.GetFocusedDataRow.Item("doc_code").ToString
            ArrayTag = New String(CInt(CN("Select Count(*) FROM ft_Sklad_DocList_Out_DocContent ( " & gWHS & ", " & V_code & " ) "))) {}

            SQLEx("Select * FROM ft_Sklad_DocList_Out_DocContent ( " & gWHS & ", " & V_code & " ) ORDER BY Вид, Наименование")

            dt.Columns.Add("Тип", GetType(String))
            dt.Columns.Add("Наименование", GetType(String))
            dt.Columns.Add("Ед", GetType(String))
            dt.Columns.Add("Треб.", GetType(String))
            dt.Columns.Add("К выдаче", GetType(String))
            dt.Columns.Add("В кг", GetType(String))
            dt.Columns.Add("Остатки", GetType(Double))
            dt.Columns.Add("Аналоги", GetType(String))
            dt.Columns.Add("Короткое наим.", GetType(String))
            dt.Columns.Add("Warecode", GetType(String))
            dt.Columns.Add("provodka_code", GetType(String))

            While reader.Read
                ArrayTag(i) = ""
                dt.Rows.Add()
                dt.Rows(dt.Rows.Count() - 1).Item(0) = reader.Item("Вид")
                dt.Rows(dt.Rows.Count() - 1).Item(1) = reader.Item("Наименование")
                dt.Rows(dt.Rows.Count() - 1).Item("Короткое наим.") = reader.Item("Короткое наим.")
                dt.Rows(dt.Rows.Count() - 1).Item("Треб.") = reader.Item("Требуется")
                dt.Rows(dt.Rows.Count() - 1).Item(2) = reader.Item("Ед")
                'На всех документах, кроме ЛЗК - авто-запись NumFact
                If GridViewRashod.GetFocusedDataRow.Item("doc_type").ToString <> "21" Then ArrayTag(i) = reader.Item("NumFact").ToString
                If reader.Item("отпущено") Is DBNull.Value Then dt.Rows(dt.Rows.Count() - 1).Item(4) = "" Else dt.Rows(dt.Rows.Count() - 1).Item(4) = Round(CDec(reader.Item("отпущено")), 3)
                If reader.Item("Num") Is DBNull.Value Then dt.Rows(dt.Rows.Count() - 1).Item(6) = 0 Else dt.Rows(dt.Rows.Count() - 1).Item(6) = Round(CDec(reader.Item("Num")), 3)
                dt.Rows(dt.Rows.Count() - 1).Item("Аналоги") = reader.Item("Аналоги")
                dt.Rows(dt.Rows.Count() - 1).Item("Warecode") = reader.Item("Warecode")
                dt.Rows(dt.Rows.Count() - 1).Item("provodka_code") = reader.Item("provodka_code")
                i = i + 1
            End While
            ConClose()

        Catch Ex As Exception
            ConClose()
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    'Расход
    Public Function Sklad_Reports_Day_OUT_GetSQL() As String
        Dim SQL As String

        SQL = ""
        SQL = SQL & "SELECT Case When W.WareType = 1 Then 'Материалы' When W.Waretype = 2 Then 'Стандартные изделия' End As Тип, "
        SQL = SQL & "  W.WareName as Наименование, "
        SQL = SQL & " ISNULL(W.WareShortName, '') AS [Короткое наим.], "
        SQL = SQL & "       O.OKEI_Name as Ед, "
        SQL = SQL & "       P.NumFact as [Кол-во], "
        SQL = SQL & "       CASE "
        SQL = SQL & "         WHEN D.doc_type = 3 THEN VW.[Смена] "
        SQL = SQL & "         WHEN D.doc_type = 22 THEN S_Client.StructureName "
        SQL = SQL & "         ELSE C.cli_name "
        SQL = SQL & "       END AS Поставщик "

        SQL = SQL & " FROM   Provodka AS P "
        SQL = SQL & "       LEFT OUTER JOIN Documents AS D ON P.doc_code = D.doc_code "
        SQL = SQL & "       LEFT OUTER JOIN Wares AS W ON P.ware_code = W.WareCode "
        SQL = SQL & "       LEFT OUTER JOIN OKEI AS O ON  O.OKEI_Code = CASE WHEN P.OKEI_Code is null THEN W.OKEI_Code ELSE P.OKEI_Code END "
        SQL = SQL & "       LEFT OUTER JOIN vwViewSmenaNames AS VW ON D.MasterCode = VW.StructureID "
        SQL = SQL & "       LEFT OUTER JOIN Structure AS S_Client ON D.cli_code = S_Client.StructureID "
        SQL = SQL & "       LEFT OUTER JOIN Clients AS C ON D.cli_code = C.cli_code "

        SQL = SQL & " WHERE  "
        SQL = SQL & "   D.doc_type IN ( 3, 100, 22 ) "

        SQL = SQL & "  AND D.data_reg BETWEEN ' " & gDate & " ' AND ' " & gDate & gDay & "' "

        SQL = SQL & "   AND CASE "
        SQL = SQL & "         WHEN D.doc_type = 3 "
        SQL = SQL & "              AND D.cli_code = " & gWHS & " THEN 1 "
        SQL = SQL & "         WHEN D.doc_type = 100 "
        SQL = SQL & "              AND D.MasterCode = " & gWHS & " THEN 1 "
        SQL = SQL & "         WHEN D.doc_type = 22 "
        SQL = SQL & "              AND D.MasterCode = " & gWHS & " THEN 1 "
        SQL = SQL & "         ELSE 0 "
        SQL = SQL & "       END = 1 "

        SQL = SQL & " ORDER  BY W.WareName "

        Sklad_Reports_Day_OUT_GetSQL = SQL
    End Function

	
    'Приход
    Public Function Sklad_Reports_Day_IN_GetSQL() As String
        Dim SQL As String

        SQL = ""
        SQL = SQL & "SELECT Case When W.WareType = 1 Then 'Материалы' When W.Waretype = 2 Then 'Стандартные изделия' End As Тип, "
        SQL = SQL & " W.WareName As Наименование, "
        SQL = SQL & "ISNULL(W.WareShortName, '') AS [Короткое наим.] , "
        SQL = SQL & "       O.OKEI_Name as Ед, "
        SQL = SQL & "       P.NumFact as [Кол-во], "
        SQL = SQL & "       CASE "
        SQL = SQL & "         WHEN D.doc_type = 101 THEN VW.[Смена] "
        SQL = SQL & "         WHEN D.doc_type = 22 THEN S_Master.StructureName "
        SQL = SQL & "         ELSE C.[cli_name] "
        SQL = SQL & "       END as Поставщик "

        SQL = SQL & "FROM   Provodka AS P "
        SQL = SQL & "       LEFT OUTER JOIN Documents AS D ON P.doc_code = D.doc_code "
        SQL = SQL & "       LEFT OUTER JOIN Wares AS W ON P.ware_code = W.WareCode "
        SQL = SQL & "       LEFT OUTER JOIN OKEI AS O ON O.OKEI_Code = "

        SQL = SQL & "                         CASE"
        SQL = SQL & "                              WHEN p.OKEI_Code Is Null"
        SQL = SQL & "                              then W.OKEI_Code"
        SQL = SQL & "                              Else p.OKEI_Code"
        SQL = SQL & "                          End"

        SQL = SQL & "       LEFT OUTER JOIN vwViewSmenaNames AS VW ON D.MasterCode = VW.StructureID "
        SQL = SQL & "       LEFT OUTER JOIN Structure AS S_Master ON D.MasterCode = S_Master.StructureID "
        SQL = SQL & "       LEFT OUTER JOIN Clients AS C ON D.cli_code = C.cli_code "

        SQL = SQL & " WHERE  "
        SQL = SQL & "   D.doc_type IN ( 1, 101, 22 ) "
        SQL = SQL & " AND D.data_reg BETWEEN ' " & gDate & " ' AND ' " & gDate & gDay & "' "
        SQL = SQL & "   AND CASE "
        SQL = SQL & "         WHEN D.doc_type = 1 "
        SQL = SQL & "              AND D.MasterCode = " & gWHS & " THEN 1 "
        SQL = SQL & "         WHEN D.doc_type = 101 "
        SQL = SQL & "              AND D.cli_code = " & gWHS & " THEN 1 "
        SQL = SQL & "         WHEN D.doc_type = 22 "
        SQL = SQL & "              AND D.cli_code = " & gWHS & " THEN 1 "
        SQL = SQL & "         ELSE 0 "
        SQL = SQL & "       END = 1 "

        SQL = SQL & "ORDER  BY W.WareName "

        Sklad_Reports_Day_IN_GetSQL = SQL
    End Function


    Sub CellWare_load(ByVal WHS As Integer)

        Try
            Dim SQL As String = ""
            Dim XML As String = ""

            With FrmCellWare

                Select Case XtraTabControl1.SelectedTabPage.Text

                    Case "Расход"
                        XML = Sklad_WareOut_MakeXML()
                        SQL = "EXEC pr_Sklad_CellWare '" & XML & "','' , " & WHS & ", 'Список ячеек'"

                    Case "Возврат поставщику"
                        XML = Sklad_WareOut2_MakeXML()
                        SQL = "EXEC pr_Sklad_CellWare '" & XML & "','' , " & WHS & ", 'Список ячеек'"

                    Case "Приход"
                        XML = Sklad_WareOut3_MakeXML()
                        SQL = "EXEC pr_Sklad_CellWare '" & XML & "','' , " & WHS & ",  'Список ячеек" & IIF_S(WHS = WHS, " - Приход'", "'")

                    Case "Приход по заявке"
                        XML = Sklad_WareOut4_MakeXML()
                        SQL = "EXEC pr_Sklad_CellWare '" & XML & "','' , " & WHS & ", 'Список ячеек - Приход'"
                End Select


                dt = New DataTable
                SQLEx(SQL)

                dt.Columns.Add("Наименование", GetType(String))
                dt.Columns.Add("Ед", GetType(String))
                dt.Columns.Add("План", GetType(String))
                dt.Columns.Add("Факт", GetType(String))
                dt.Columns.Add("Склад", GetType(String))
                dt.Columns.Add("Стеллаж", GetType(String))
                dt.Columns.Add("Ячейка", GetType(String))
                dt.Columns.Add("Кол в ячейке", GetType(String))
                dt.Columns.Add("Вместительность", GetType(String))
                dt.Columns.Add("WareCode", GetType(String))
                dt.Columns.Add("CWID", GetType(String))

                While reader.Read
                    dt.Rows.Add()
                    dt.Rows(dt.Rows.Count() - 1).Item(0) = reader.Item("Warename")
                    dt.Rows(dt.Rows.Count() - 1).Item(1) = reader.Item("Okei_name")
                    dt.Rows(dt.Rows.Count() - 1).Item(2) = reader.Item("SumNum")
                    dt.Rows(dt.Rows.Count() - 1).Item(3) = reader.Item("Num")
                    dt.Rows(dt.Rows.Count() - 1).Item(4) = reader.Item("S1Name")
                    dt.Rows(dt.Rows.Count() - 1).Item(5) = reader.Item("S2Name")
                    dt.Rows(dt.Rows.Count() - 1).Item(6) = reader.Item("S3Name")
                    dt.Rows(dt.Rows.Count() - 1).Item(7) = reader.Item("NumWhs")
                    dt.Rows(dt.Rows.Count() - 1).Item(8) = reader.Item("MaxNum")
                    dt.Rows(dt.Rows.Count() - 1).Item(9) = reader.Item("WareCode")
                    dt.Rows(dt.Rows.Count() - 1).Item(10) = reader.Item("CWID")
                End While

                FrmCellWare.GridCtrlCellWare.DataSource = dt
                ConClose()

                FrmCellWare.GridViewCellWare.Columns.Item("WareCode").Visible = False
                FrmCellWare.GridViewCellWare.Columns.Item("CWID").Visible = False
            End With
        Catch Ex As Exception
            ConClose()
            MsgBox(Ex.Message)
        End Try

    End Sub

	
    'ПОСТРОЕНИЕ XML СТРОКИ ДЛЯ ОТПРАВКИ В ПРОЦЕДУРУ СОХРАНЕНИЯ НА СЕРВЕР
    Function Sklad_WareSend_MakeXML() As String

        Dim i As Integer
        Dim xml As String
        xml = "<root>"

        For i = 0 To GridViewRashodClick.RowCount - 1
            xml = xml & "<row  provodka_code  = " & Chr(34) & GridViewRashodClick.GetDataRow(i).Item("provodka_code").ToString & Chr(34)
            xml = xml & "      WareCode       = " & Chr(34) & GridViewRashodClick.GetDataRow(i).Item("Warecode").ToString & Chr(34)
            xml = xml & "      NumPlan        = " & Chr(34) & Sng2DB(CSng(GridViewRashodClick.GetDataRow(i).Item("Треб.")), 3) & Chr(34)
            xml = xml & "      NumFact        = " & Chr(34) & Sng2DB(CSng(ArrayTag(i)), 3) & Chr(34) & " />"
        Next

        xml = xml & "</root>"
        Sklad_WareSend_MakeXML = xml
    End Function

	
    'ПОСТРОЕНИЕ XML СТРОКИ ДЛЯ ОТПРАВКИ В ПРОЦЕДУРУ СОХРАНЕНИЯ НА СЕРВЕР
    Function Sklad_WareOut_MakeXML() As String

        Dim i As Integer
        Dim XML As String
        XML = "<root>"

        For i = 0 To GridViewRashodClick.RowCount - 1
            If Len(GridViewRashodClick.GetDataRow(i).Item(4).ToString) > 0 And GridViewRashodClick.GetDataRow(i).Item(4).ToString <> "0" Then
                XML = XML & "<row "

                XML = XML & "WareCode="
                XML = XML & gQ
                XML = XML & GridViewRashodClick.GetDataRow(i).Item("warecode").ToString
                XML = XML & gQ

                XML = XML & " Num="
                XML = XML & gQ
                XML = XML & Replace(ArrayTag(i), ",", ".")
                XML = XML & gQ

                XML = XML & " />"
            End If
        Next

        XML = XML & "</root>"
        Sklad_WareOut_MakeXML = XML
    End Function

	
    Function Sklad_WareOut4_MakeXML() As String

        Dim i As Integer
        Dim XML As String
        XML = "<root>"

        For i = 0 To GridViewPoZaiyvkeClick.RowCount - 1

            If Len(GridViewPoZaiyvkeClick.GetDataRow(i).Item(5)) > 0 And Len(GridViewPoZaiyvkeClick.GetDataRow(i).Item(6)) <> 0 And GridViewPoZaiyvkeClick.GetDataRow(i).Item(5).ToString <> "" Then
                XML = XML & "<row "
                XML = XML & "WareCode="
                XML = XML & gQ
                XML = XML & GridViewPoZaiyvkeClick.GetDataRow(i).Item("Warecode").ToString
                XML = XML & gQ

                XML = XML & " Num="
                XML = XML & gQ
                XML = XML & Replace(GridViewPoZaiyvkeClick.GetDataRow(i).Item("Кол-во").ToString, ",", ".")
                XML = XML & gQ

                XML = XML & " NewOKEI="
                XML = XML & gQ
                XML = XML & CN("SELECT OKEI_Code FROM OKEI WHERE OKEI_name='" & GridViewPoZaiyvkeClick.GetDataRow(i).Item("Ед. изм.").ToString & "' ").ToString()
                XML = XML & gQ

                XML = XML & " />"

            End If
        Next

        XML = XML & "</root>"
        Sklad_WareOut4_MakeXML = XML
    End Function
	

    Function Sklad_WareOut3_MakeXML() As String

        Dim XML As String
        XML = "<root>"

        For i = 0 To GridViewPrihodClick.RowCount - 1
            If Len(GridViewPrihodClick.GetDataRow(i).Item("Кол-во").ToString) > 0 And GridViewPrihodClick.GetDataRow(i).Item("Кол-во").ToString <> "0" Then
                XML = XML & "<row "
                XML = XML & "WareCode="
                XML = XML & gQ
                XML = XML & GridViewPrihodClick.GetDataRow(i).Item("Ware_Code").ToString
                XML = XML & gQ
                XML = XML & " Num="
                XML = XML & gQ
                XML = XML & Replace$(GridViewPrihodClick.GetDataRow(i).Item("Кол-во").ToString, ",", ".")
                XML = XML & gQ
                XML = XML & " NewOKEI="
                XML = XML & gQ
                XML = XML & CN("SELECT OKEI_Code FROM Wares WHERE WareCode='" & GridViewPrihodClick.GetDataRow(i).Item("Ware_Code").ToString & "' ").ToString
                XML = XML & gQ
                XML = XML & " />"
            End If
        Next

        XML = XML & "</root>"
        Sklad_WareOut3_MakeXML = XML
    End Function


    Function Sklad_CellWareOut_MakeXML() As String

        Dim XML As String
        XML = "<root>"

        For i = 0 To FrmCellWare.GridViewCellWare.RowCount - 1
            If Len(FrmCellWare.GridViewCellWare.GetDataRow(i).Item(3).ToString) > 0 And FrmCellWare.GridViewCellWare.GetDataRow(i).Item(3).ToString <> "0" Then
                XML = XML & "<row "

                XML = XML & "WareCode="
                XML = XML & gQ
                XML = XML & FrmCellWare.GridViewCellWare.GetDataRow(i).Item("WareCode").ToString
                XML = XML & gQ

                XML = XML & " Num="
                XML = XML & gQ
                XML = XML & (Replace$(FrmCellWare.GridViewCellWare.GetDataRow(i).Item("Факт").ToString, ",", "."))
                XML = XML & gQ

                XML = XML & " />"
            End If
        Next

        XML = XML & "</root>"
        Sklad_CellWareOut_MakeXML = XML
    End Function


    Function Sklad_WareOut2_MakeXML() As String

        Dim xml As String
        xml = "<root>"

        For i = 0 To GridViewVozvratClick.RowCount - 1
            If Len(GridViewVozvratClick.GetDataRow(i).Item("Цена").ToString) > 0 And GridViewVozvratClick.GetDataRow(i).Item("Цена").ToString <> "0" Then
                xml = xml & "<row "

                xml = xml & "WareCode="
                xml = xml & gQ
                xml = xml & GridViewVozvratClick.GetDataRow(i).Item("WareCode").ToString
                xml = xml & gQ

                xml = xml & " Num="
                xml = xml & gQ

                xml = xml & Replace(GridViewVozvratClick.GetDataRow(i).Item("Кол-во").ToString, ",", ".")
                xml = xml & gQ

                xml = xml & " />"

            Else
                MsgBox("ВНИМАНИЕ - Не допускается приход\возврат ресурса по нулевой цене." & vbCr & vbCr & "Ресурс: " & GridViewVozvratClick.GetDataRow(i).Item("Наименование").ToString, vbInformation, "Неверный ввод")
            End If
        Next

        xml = xml & "</root>"
        Sklad_WareOut2_MakeXML = xml
    End Function
	

    Function Sklad_CellWarePR_MakeXML() As String

        Dim XML As String
        XML = "<root>"

        For i = 0 To FrmCellWare.GridViewCellWare.RowCount - 1
            If Len(FrmCellWare.GridViewCellWare.GetDataRow(i).Item(3).ToString) > 0 And FrmCellWare.GridViewCellWare.GetDataRow(i).Item(3).ToString <> "0" Then
                XML = XML & "<row "

                XML = XML & "WareCode="
                XML = XML & gQ
                XML = XML & FrmCellWare.GridViewCellWare.GetDataRow(i).Item("WareCode").ToString
                XML = XML & gQ

                XML = XML & " Num="
                XML = XML & gQ
                XML = XML & (Replace$(FrmCellWare.GridViewCellWare.GetDataRow(i).Item("Факт").ToString, ",", "."))
                XML = XML & gQ

                XML = XML & " CWID="
                XML = XML & gQ
                XML = XML & (Replace$(FrmCellWare.GridViewCellWare.GetDataRow(i).Item("CWID").ToString, ",", "."))
                XML = XML & gQ

                XML = XML & " NewOKEI="
                XML = XML & gQ

                XML = XML & CN("SELECT OKEI_Code FROM OKEI WHERE OKEI_name='" & FrmCellWare.GridViewCellWare.GetDataRow(i).Item("Ед").ToString & "' ").ToString

                XML = XML & gQ
                XML = XML & " />"
            End If
        Next

        XML = XML & "</root>"
        Sklad_CellWarePR_MakeXML = XML
    End Function

    Public Sub mew_in(ByVal PayInSlip As String)
        Try

            'ПРОВЕДЕНИЕ ДОКУМЕНТА
            SQLExNQ("UPDATE Documents SET data_reg = GETDATE(), UserID = " & gUserId & ", PayInSlip = '" & PayInSlip & "' WHERE doc_code=" & GridViewPrihod.GetFocusedDataRow.Item("doc_code").ToString)
            If CInt(CN("select Count(*) from Documents Where doc_code = " & GridViewPrihod.GetFocusedDataRow.Item("doc_code").ToString)) = 0 Then Exit Sub

            'ПЕЧАТЬ ПРИХОДНОЙ НАКЛАДНОЙ НА СКЛАДА ОТ ПОСТАВЩИКА (Типовая межотраслевая форма № М-4)
            If CBool(StrComp(GridViewPrihod.GetFocusedDataRow.Item("doc_type").ToString, "22")) = False Then
                If MsgBox("Печатать?", vbYesNo + vbQuestion, "Печать требования-накладной (по форме М-11)") = vbYes Then
                    Dim path As String = pathMyDocs & "\frM11.fr3"
                    IO.File.WriteAllBytes(path, My.Resources.frM11)
                    Label14.Visible = False : M11_ToFR(FRX, GridViewPrihod.GetFocusedDataRow.Item("doc_code").ToString, 1, , , , , gWHS) 'ПЕЧАТЬ ЧЕРЕЗ FastREport
                End If
            Else
                If MsgBox("Печатать?", vbYesNo + vbQuestion, "Печать приходной накладной (по форме М-4)") = vbYes Then
                    Dim path As String = pathMyDocs & "\frM4.fr3"
                    IO.File.WriteAllBytes(path, My.Resources.frM4)
                    Label14.Visible = False : M4_ToFR(FRX, CInt(GridViewPrihod.GetFocusedDataRow.Item("doc_code")), 1, , gWHS) 'ПЕЧАТЬ ЧЕРЕЗ FastREport
                End If
            End If

            Exit Sub

        Catch
            MsgBox("ВНИМАНИЕ - Произошла ошибка. " & vbCr & "Пожалуйста повторите попытку..." & vbCr & vbCr & "Описание ошибки: " & gErr, vbCritical, "Ошибка сохранения")
        End Try
    End Sub

	
    'ПЕРЕМЕЩЕНИЕ РЕСУРСОВ МЕЖДУ СКЛАДАМИ - ПРИХОД НА ПРИНИМАЮЩЕМ СКЛАДЕ
    Function Sklad_WareSend_Receive(ByVal DOCCODE As String) As Boolean

        Sklad_WareSend_Receive = True

        SQLEx("SELECT WB.WareCode FROM Provodka AS P INNER JOIN Documents AS D ON P.doc_code = D.doc_code INNER JOIN WareBalance AS WB ON P.ware_code = WB.WareCode AND D.MasterCode = WB.smena_no AND ( ISNULL(P.NumFact, 0) > WB.Num OR P.NumFact IS NULL ) WHERE D.doc_code = " & DOCCODE)

        While reader.Read
            For i = 0 To GridViewPrihodClick.RowCount - 1
                If GridViewPrihodClick.GetDataRow(i).Item("ware_code").ToString = reader.Item("Warecode").ToString Then
                    'нужно выделить красным цветом ячейку
                    Form1.GridViewPrihodClick.ActiveEditor.BackColor = Color.Red
                End If
            Next
        End While

        MsgBox("ВНИМАНИЕ - На отдающем складе нет требуемого количества либо оно указано неверно.", vbInformation, "Ошибка сохранения")
        Sklad_WareSend_Receive = False

    End Function


    ' Приход по заявке
    Public Sub order_fix(ByVal PayInSlip As String)

        Try

            Dim AllZero As Boolean = False
            Dim var_code As Integer
            Dim sTMP As String = ""
            Dim i As Integer
            Dim var_cli As Integer
            Dim sDocText As String
            Dim isVIP As Boolean
            Dim var_doc_code As String = ""
            Dim OldPrice As Double
            Dim OldNum As Double
            Dim OldNumAll As Double
            Dim NewNum As Double
            Dim NewNumAll As Double
            Dim Average As Double

            'индексы колонок строк документа
            Const iColNum = 4
            Const iColPrice = 5
            Const iColCost = 6
            Const iColDeliv = 7

            Dim cCost As Double
            Dim cDelivery As Double
            Dim cNum As Double
            Dim a As Integer
            Dim b As Integer
            Dim cWeight As Double

            For i = 0 To GridViewPoZaiyvkeClick.RowCount - 1
                If CInt(GridViewPoZaiyvkeClick.GetDataRow(i).Item("Кол-во")) <> 0 Then AllZero = True
            Next

            If AllZero = False Then
                MsgBox("ВНИМАНИЕ - Ни в одной позиции документа не указано количество", vbInformation, "Ошибка сохранения")
                Exit Sub
            End If

            var_code = CInt(GridViewPoZaiyvke.GetFocusedDataRow.Item("Orders_code"))
            '  If gCellWare Then
            Dim MsB As Integer
            MsB = MsgBox("Ресурсы будут записаны в ячейки по умолчанию." & vbCr & "Хотите просмотреть/изменить список?", vbYesNo, "Информация")
            If MsB = vbYes Then
                sTMP = FrmCellWare.Show2(True, gWHS)
            ElseIf MsB = vbNo Then
                sTMP = FrmCellWare.Show2(False, gWHS)
            End If

            If Len(sTMP) = 0 Then Exit Sub

            For i = 0 To GridViewPoZaiyvkeClick.RowCount - 1
                If Len(Trim(GridViewPoZaiyvkeClick.GetDataRow(i).Item(iColCost).ToString)) > 0 Then
                    'есть необходимость проверять в каких единицах идет подсчет

                    dt.Columns.Add("NumFact", GetType(String))
                    dt.Columns.Add("№ п/п", GetType(String))
                    dt.Columns.Add("Название", GetType(String))
                    dt.Columns.Add("Ед. изм.", GetType(String))
                    dt.Columns.Add("Кол-во", GetType(String))
                    dt.Columns.Add("Цена", GetType(String))
                    dt.Columns.Add("Стоимость", GetType(String))
                    dt.Columns.Add("+Доставка", GetType(String))
                    dt.Columns.Add("Тип", GetType(String))
                    dt.Columns.Add("Примечание", GetType(String))
                    dt.Columns.Add("Короткое наим.", GetType(String))
                    dt.Columns.Add("Аналоги", GetType(String))
                    dt.Columns.Add("Warecode", GetType(String))

                    'a - хранит OKEI_Code с прихода
                    a = CInt(CN("SELECT OKEI_Code FROM OKEI WHERE OKEI_Name ='" & GridViewPoZaiyvkeClick.GetDataRow(i).Item("Ед. изм.").ToString & "' "))
                    'b - хранит OKEI_Code в заявке
                    b = CInt(CN("SELECT CASE WHEN OrdersBody.OKEI_Code Is Null THEN Wares.OKEI_Code  Else OrdersBody.OKEI_Code END As OKEI_Code From OrdersBody,Wares WHERE orders_code =" & var_code & " AND ware_code=wareCode AND wareCode = " & GridViewPoZaiyvkeClick.GetDataRow(i).Item("Warecode").ToString)) ' 

                    'получаем коэффициент
                    cWeight = CDbl(CN("SELECT TOP 1 ISNULL(Weight, 0) As Weight FROM Wares WHERE WareName = '" & GridViewPoZaiyvkeClick.GetDataRow(i).Item("Название").ToString & " ' "))

                    ' в (i, iColNum) хранится в тех единицах, в которых получено, поэтому преобразовывать требуется к тем единицам, в которых закупалос
                    If (a = b) Then
                        SQLExNQ("UPDATE OrdersBody SET NumFact = ROUND(ISNULL(NumFact, 0) + REPLACE('" & GridViewPoZaiyvkeClick.GetDataRow(i).Item("NumFact").ToString & "', ',', '.'), 3) WHERE orders_code=" & var_code & " AND ware_code=" & GridViewPoZaiyvkeClick.GetDataRow(i).Item("Warecode").ToString)
                        SQLExNQ("UPDATE OrdersBody SET Price = REPLACE('" & GridViewPoZaiyvkeClick.GetDataRow(i).Item("Цена").ToString & "', ',', '.') WHERE orders_code=" & var_code & " AND ware_code=" & GridViewPoZaiyvkeClick.GetDataRow(i).Item("Warecode").ToString)
                    Else
                        If (a = 2) Then 'приход в кг, а заявка в учетных единицах, значит надо списывать в уч.ед
                            SQLExNQ(" UPDATE OrdersBody SET NumFact = ROUND(ISNULL(NumFact, 0) + REPLACE('" & Round((CDbl(GridViewPoZaiyvkeClick.GetDataRow(i).Item(iColNum)) / cWeight), 3) & "', ',', '.'), 3) Where orders_code = " & var_code & " And ware_code = " & GridViewPoZaiyvkeClick.GetDataRow(i).Item("Warecode").ToString)
                            SQLExNQ(" UPDATE OrdersBody SET Price = REPLACE('" & GridViewPoZaiyvkeClick.GetDataRow(i).Item(4).ToString & "', ',', '.') WHERE orders_code=" & var_code & " AND ware_code=" & GridViewPoZaiyvkeClick.GetDataRow(i).Item("Warecode").ToString)
                        Else 'b=2 приход в учетных единицах, а заявка в кг, значит надо списывать в кг
                            SQLExNQ(" UPDATE OrdersBody SET NumFact = ROUND(ISNULL(NumFact, 0) + REPLACE('" & Round((CDbl(GridViewPoZaiyvkeClick.GetDataRow(i).Item(iColNum)) * cWeight), 3) & "', ',', '.'), 3) WHERE orders_code=" & var_code & " AND ware_code=" & GridViewPoZaiyvkeClick.GetDataRow(i).Item("Warecode").ToString)
                            SQLExNQ(" UPDATE OrdersBody SET Price = REPLACE('" & GridViewPoZaiyvkeClick.GetDataRow(i).Item(4).ToString & "', ',', '.') WHERE orders_code=" & var_code & " AND ware_code=" & GridViewPoZaiyvkeClick.GetDataRow(i).Item("Warecode").ToString)
                        End If
                    End If
                End If
            Next

            'Запись в таблицу "orders"
            var_cli = CInt(CN("SELECT cli_code FROM orders WHERE orders_code = " & var_code))

            isVIP = CBool(CN("SELECT ISNULL(IsVIP, 0) FROM orders WHERE orders_code = " & var_code))

            'Запись в таблицу "Documents"
            sDocText = GridViewPoZaiyvke.GetFocusedDataRow.Item("Номер вход. док-та").ToString

            'Documents
            If Len(mInDocOrder_SavedKey) > 0 Then
                    var_doc_code = GETP(mInDocOrder_SavedKey, "doc_code")

                    SQLExNQ("UPDATE Documents SET data_reg  = GETDATE(), " & _
                               "                     PayInSlip = ' " & PayInSlip & " ' " & _
                               "                     Delivery  = ROUND(REPLACE('" & NeS(Trim$(txtDelivery2.text)) & "', ',', '.'), 2), " & _
                               "                     doc_text  = '" & sDocText & "', " & _
                               "                     ordersCode = " & var_code & " , " & _
                               "                     UserID = " & gUserId & ", DocDT = '" & Form1.GridViewPoZaiyvke.GetFocusedDataRow.Item("Дата документа").ToString & "', DocNote = '" & Form1.GridViewPoZaiyvke.GetFocusedDataRow.Item("Комментарий").ToString & "' " &
                               "WHERE doc_code=" & var_doc_code)

            Else
                SQLExNQ("INSERT INTO Documents (cli_code, MasterCode, data_make, data_reg, NDS, NDS_inc, doc_text, doc_type, Delivery, UserID, IsVIP, DocDT, DocNote, DocPay, PayInSlip, ordersCode) " &
                        "VALUES (" & var_cli & ", " & gWHS & ", GETDATE(), GETDATE(), " & GridViewPoZaiyvke.GetFocusedDataRow.Item("НДС(%)").ToString & ", '" & GridViewPoZaiyvke.GetFocusedDataRow.Item(5).ToString & "', " &
                        IIF_S(Len(sDocText) > 0, "'" & sDocText & "'", "Null") & ", 1, ROUND(REPLACE('" & NeS(Trim$(txtDelivery2.Text)) & "', ',', '.'), 2), " & gUserId & ", " & IIF_S(isVIP, "1", "NULL") &
                        ", '" & GridViewPoZaiyvke.GetFocusedDataRow.Item("Дата док-та").ToString & "', '" & GridViewPoZaiyvke.GetFocusedDataRow.Item("Комментарий").ToString &
                        "', (SELECT TOP 1 Order_note FROM Orders WHERE orders_code = " & var_code & "),' " & PayInSlip & " ', " & var_code & " )")
                'Код нового документа
                var_doc_code = CN("SELECT MAX(doc_code) FROM Documents").ToString
            End If

            'удаление связки-документ
            SQLExNQ("UPDATE Orders SET doc_code = NULL WHERE orders_code = " & var_code)

            'Запись проводок на складе
            'чтобы проверял по стоимости, а не по цене
            For i = 0 To GridViewPoZaiyvkeClick.RowCount - 1
                If CBool(StrComp(Nz(Trim(GridViewPoZaiyvkeClick.GetDataRow(i).Item(iColNum).ToString).ToString).ToString, "0")) And Len(Trim(GridViewPoZaiyvkeClick.GetDataRow(i).Item(iColCost).ToString).ToString) > 0 Then
                    cCost = CDbl(GP(GridViewPoZaiyvkeClick.Tag.ToString, "[" & i & "_Cost]"))
                    cDelivery = CDbl(GP(GridViewPoZaiyvkeClick.Tag.ToString, "[" & i & "_Delivery]"))
                    cNum = CDbl(Trim(GridViewPoZaiyvkeClick.GetDataRow(i).Item(iColNum).ToString))
                    'в таблицу Provodka добавлено поле OKEI_Code, которое хранит код в тех единицах в которых фактически пришло
                    Dim OkeiCode As Integer
                    OkeiCode = CInt(CN("SELECT OKEI_Code FROM OKEI WHERE OKEI_Name = '" & GridViewPoZaiyvkeClick.GetDataRow(i).Item(3).ToString & "' "))
                    'Проводки

                    Dim SQL As String

                    SQL = ""
                    SQL = SQL & "INSERT INTO provodka ( OKEI_Code, ware_code, ware_type, doc_code, NUM, NumFact" & Switch(Len(Trim(GridViewPoZaiyvkeClick.GetDataRow(i).Item(iColPrice).ToString)) > 0, ", Price, Delivery, PriceMinusDelivery").ToString & ") "
                    SQL = SQL & "VALUES ( " & OkeiCode & " ," & GridViewPoZaiyvkeClick.GetDataRow(i).Item("Warecode").ToString & ", "
                    If CBool(StrComp(GridViewPoZaiyvkeClick.GetDataRow(i).Item(8).ToString, "М")) = False Then
                        SQL = SQL & "1, "
                    ElseIf CBool(StrComp(GridViewPoZaiyvkeClick.GetDataRow(i).Item(8).ToString, "Ст")) = False Then
                        SQL = SQL & "2, "
                    Else
                        SQL = SQL & "5, "
                    End If
                    SQL = SQL & var_doc_code & " , REPLACE('" & Round(cNum, 3) & "', ',', '.'), "
                    SQL = SQL & " REPLACE('" & Round(cNum, 3) & "', ',', '.') "
                    SQL = SQL & Switch(Len(Trim(GridViewPoZaiyvkeClick.GetDataRow(i).Item(iColPrice).ToString)) > 0, ", " & RCr(cDelivery) & ", " & RCr(cDelivery) & ", " & RCr(cCost)).ToString & ") "
                    SQLExNQ(SQL)                                                                                                                    ' Rce (cDelivery - cCost)

                    With GridViewPoZaiyvkeClick

                        SQLEx("SELECT ISNULL(W.Price, 0) AS Price, ISNULL(WB.NUM, 0) AS Num FROM WareBalance WB INNER JOIN Wares W ON W.WareCode=WB.WareCode " &
                              "WHERE WB.WareCode = " & GridViewPoZaiyvkeClick.GetDataRow(i).Item("Warecode").ToString & " And smena_no = " & gWHS)
                        While reader.Read
                            'Старое кол-во на складе
                            OldNum = CDbl(reader.Item("Num"))
                            'Старая цена
                            OldPrice = CDbl(reader.Item("Price"))
                        End While

                        ConClose()

                        'Старое кол-во на всех складах
                        OldNumAll = CDbl(CN("SELECT ISNULL((SELECT SUM(WB.NUM) AS Num FROM WareBalance WB JOIN Structure S ON S.StructureID = WB.smena_no WHERE WB.WareCode = " & GridViewPoZaiyvkeClick.GetDataRow(i).Item("Warecode").ToString & " AND S.StructureType = 5), 0)"))

                        'если пришло не в учетных
                        'a - хранит OKEI_Code с прихода
                        a = CInt(CN("SELECT OKEI_Code FROM OKEI WHERE OKEI_Name ='" & GridViewPoZaiyvkeClick.GetDataRow(i).Item(3).ToString & "' "))
                        'b - хранит OKEI_Code учетных единиц из справочника
                        b = CInt(CN("SELECT Wares.OKEI_Code From OrdersBody,Wares WHERE orders_code =" & var_code & " AND ware_code=wareCode AND wareCode = " & GridViewPoZaiyvkeClick.GetDataRow(i).Item("Warecode").ToString))
                        'получаем коэффициент
                        cWeight = CDbl(CN("SELECT TOP 1 ISNULL(Weight, 0) As Weight FROM Wares WHERE WareName = '" & GridViewPoZaiyvkeClick.GetDataRow(i).Item(2).ToString & " ' "))

                        ' в Form1.GridViewPoZaiyvkeClick.GetDataRow(i).Item(iColNum) должно храниться в учетных единицах
                        If (a <> b) Then
                            Dim NewCost As Double
                            'Запоминаем стоимость
                            NewCost = CDbl(GridViewPoZaiyvkeClick.GetDataRow(i).Item(iColDeliv))
                            If (a = 2) Then 'приход в кг, надо в расчетных
                                GridViewPoZaiyvkeClick.GetDataRow(i).Item(iColNum) = Round((CDbl(GridViewPoZaiyvkeClick.GetDataRow(i).Item(iColNum)) / cWeight), 3)
                            Else 'b=2 приход в учетных единицах, а заявка в кг, значит надо списывать в кг
                                GridViewPoZaiyvkeClick.GetDataRow(i).Item(iColNum) = Round((CDbl(GridViewPoZaiyvkeClick.GetDataRow(i).Item(iColNum)) * cWeight), 3)
                            End If
                            'цена в уч.единцицах измерения будет отличаться от цены в тех единицах, в которых пришло, поэтому надо пересчитать
                            GridViewPoZaiyvkeClick.GetDataRow(i).Item(iColPrice) = Round(CDbl(GridViewPoZaiyvkeClick.GetDataRow(i).Item(iColDeliv)) / CDbl(GridViewPoZaiyvkeClick.GetDataRow(i).Item(iColNum)), 2)
                        End If

                        'Новое количество на складе
                        'изменить! проверить в каких единицах пришло, и если не в учетных, то перевести в учетные
                        NewNum = CDbl(.GetDataRow(i).Item(iColNum)) + OldNum

                        'Обновление кол-ва на складе данного ресурса
                        SQLExNQ("UPDATE WareBalance SET Num = " & Dbl2DB(NewNum, 3) & " " &
                                   "WHERE  WareCode = " & GridViewPoZaiyvkeClick.GetDataRow(i).Item("Warecode").ToString & " AND smena_no = " & gWHS)

                        'Новое кол-во на всех складах
                        NewNumAll = OldNumAll + CDbl(.GetDataRow(i).Item(iColNum))

                        'расчет новой цены (усреднение)
                        If OldNumAll = 0 Then

                            'A/C
                            Average = CDbl(.GetDataRow(i).Item(iColPrice))
                        Else

                            '(A+B)/C
                            Average = (0 + OldNumAll * OldPrice) / NewNumAll  ' CDbl(.GetDataRow(i).Item(iColDeliv))
                        End If

                        'обновление цены на складе
                        SQLExNQ("UPDATE Wares SET Price = " & Dbl2DB(Average, 2) & " WHERE WareCode = " & GridViewPoZaiyvkeClick.GetDataRow(i).Item("Warecode").ToString)

                    End With
                End If
            Next

            'СБРОС НДС В ЗАЯВКЕ
            SQLExNQ("UPDATE Orders SET NDS = NULL, NDS_Inc = NULL WHERE orders_code = " & var_code)

            'Запись в ячейки
            'снова собираем строку для передачи в процедуру. т.к. на экране уже отображены только учетные единицы
            sTMP = FrmCellWare.Show2(False, gWHS)
            ' If gCellWare Then
            SQLExNQ("EXEC pr_Sklad_CellWare '<root></root>', '" & sTMP & "', '', 'Запись в ячейки' ")
            '  End If

            'ПЕЧАТЬ ПРИХОДНОЙ НАКЛАДНОЙ НА СКЛАДА ОТ ПОСТАВЩИКА (Типовая межотраслевая форма № М-4)
            If MsgBox("Печатать?", vbYesNo + vbQuestion, "Печать приходной накладной (по форме М-4)") = vbYes Then
                Dim path As String = pathMyDocs & "\frM4.fr3"
                IO.File.WriteAllBytes(path, My.Resources.frM4)
                M4_ToFR(FRX, CInt(var_doc_code), 1, , gWHS) 'ПЕЧАТЬ ЧЕРЕЗ FastREport
            End If

            'ОБНОВЛЕНИЕ
            отрисовка_заказов()
            GridCtrlPoZaiyvke.DataSource = dt
            GridCtrlPoZaiyvkeClick.Visible = False
            BigButton.Enabled = False

            Exit Sub

        Catch
            ConClose()
            MsgBox("Ошибка сохранения" & vbCrLf & "Документ не будет проведен.." & vbCr & vbCr & "Описание ошибки: """ & Err.Description & " (" & Err.Number & ")" & """", vbCritical, "Ошибка сохранения")
        End Try

    End Sub


    'ПРИХОД - ИЗМЕНИТЬ КОЛИЧЕСТВО
    Public Sub InNumChange(ByVal GV As DevExpress.XtraGrid.Views.Grid.GridView, ByVal fSkladTransfer As Boolean)
        Dim ANSW As Object

        Try

            'Запрос количества
            ANSW = CheckNum(InputBox("Количество:", "[ " & GV.GetFocusedDataRow.Item("Тип").ToString & " ] - " & GV.GetFocusedDataRow.Item("Наименование").ToString, GV.GetFocusedDataRow.Item("Кол-во").ToString, )) ' Form1.ActiveControl.ColumnHeaders(1).Width + Form1.LEFT + Form1.ActiveControl.LEFT + Form1.ActiveControl.Container.LEFT, Form1.TOP + Form1.ActiveControl.TOP + Form1.ActiveControl.Container.TOP + lvwItem.TOP + 2400))
            If ANSW Is Nothing Or ANSW Is DBNull.Value Then Exit Sub
            If CDbl(ANSW) < 0 Then Exit Sub

            'ПЕРЕДАЧА РЕСУРСОВ МЕЖДУ СКЛАДАМИ
            If fSkladTransfer And CDbl(ANSW) > CDbl(Replace(GV.GetFocusedDataRow.Item("Кол-во").ToString, ".", ",")) Then
                MsgBox("ВНИМАНИЕ - Количество 'К выдаче' должно быть равным либо меньше требуемого по документу.", vbInformation, "Неверный ввод")
                Exit Sub
            End If

            'Редактирование целевой таблицы
            SQLExNQ("UPDATE provodka SET " & IIF_S(fSkladTransfer, "NumFact", "Num") & " = " & Sng2DB(CSng(ANSW), 3) & " WHERE provodka_code=" & GV.GetFocusedDataRow.Item("provodka_code").ToString)

            If XtraTabControl1.SelectedTabPage.Text = "Приход" And fSkladTransfer = False Then
                Delivery()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    'УСТАНОВКА ФЛАГА ЮЗЕРА НА ДОКУМЕНТЕ
    Function UserFlags_SET(ByVal SFIELD As String, ByVal lVALUE As Integer, ByVal fClearFlag As Boolean) As Boolean
        Dim sTMP As Object = Nothing

        'ВЫХОД ЮЗЕРА ИЗ ПРОГРАММЫ
        If Len(SFIELD) = 0 Then
            SQLExNQ("UPDATE Users SET doc_code=NULL,orders_code=NULL WHERE UserID=" & gUserId)
            UserFlags_SET = False
            Exit Function
        End If

        SQLEx("SELECT UserFullName FROM Users WHERE " & SFIELD & " = " & lVALUE & " AND UserID <> " & gUserId)
        While reader.Read
            sTMP = reader.Item(0)
        End While
        ConClose()

        If Not sTMP Is Nothing Then

            UserFlags_SET = True
            MsgBox("ВНИМАНИЕ - Документ уже открыт пользователем: " & sTMP.ToString, vbInformation, "Документ занят другим пользователем")

        Else
            SQLExNQ("UPDATE Users SET " & SFIELD.ToString & " = " & IIf(fClearFlag, "NULL", lVALUE).ToString & " WHERE UserID=" & gUserId)
            UserFlags_SET = False
        End If
    End Function

	
    Sub cboWare_Load()
        Dim XML As String = ""

        With frmNewCellWare

            .cboWare.Items.Clear()
            .cboRack.Items.Clear()
            .cboCell.Items.Clear()
            .txtNumWhs.Text = ""
            .txtNum.Text = ""

            Select Case XtraTabControl1.SelectedTabPage.Text

                Case "Расход"
                    XML = Sklad_WareOut_MakeXML()
                Case "Возврат поставщику"
                    '       XML = Sklad_WareOut2_MakeXML(Form1.lvwRight(2))
                Case "Приход"
                    XML = Sklad_WareOut3_MakeXML()
                Case "Приход по заявке"
                    XML = Sklad_WareOut4_MakeXML()

            End Select

            SQLEx("EXEC pr_Sklad_CellWare '" & XML & "', '', " & gWHS & ", 'Список Wares'")

            While reader.Read()
                .cboWare.Items.Add(reader.Item("WareName"))
            End While

        End With
    End Sub


    Public Function ReturnToClient(ByVal DOCCODE As Integer) As Boolean
        ' проверяем заполнение
        ' надо проверить одновременно заполнение и соответствие наличия на складе

        Dim SQL As String, ans As Object
        Dim var_Ed As String = ""
        Dim sTMP As String

        'надо проставить факт выдачи
        SQLExNQ("UPDATE provodka SET provodka.NumFact = provodka.NUM WHERE provodka.doc_code=" & DOCCODE)

        SQL = ""
        SQL = SQL & "SELECT O.OKEI_Name       AS [Ед измерения], "
        SQL = SQL & "       P.NumFact         AS отпущено, "
        SQL = SQL & "       CASE"
        SQL = SQL & "           WHEN O.OKEI_Code = W.OKEI_Code"
        SQL = SQL & "           THEN  P.NumFact"
        SQL = SQL & "           ELSE  P.NumFact / W.Weight"
        SQL = SQL & "       END   AS [в_уч_ед],"
        SQL = SQL & "       P.provodka_code, "
        SQL = SQL & "       W.WareCode, "
        SQL = SQL & "       W.WareName        AS Название, "
        SQL = SQL & "       ISNULL(WB.Num, 0) AS ОСТАТОК "

        SQL = SQL & "FROM   Provodka AS P "
        SQL = SQL & "       LEFT OUTER JOIN WareBalance AS WB "
        SQL = SQL & "                       RIGHT OUTER JOIN Wares AS W ON WB.WareCode = W.WareCode ON P.ware_code = W.WareCode "
        SQL = SQL & "       LEFT OUTER JOIN OKEI AS O ON O.OKEI_Code = CASE WHEN P.OKEI_Code is null THEN W.OKEI_Code ELSE P.OKEI_Code END  "

        SQL = SQL & "WHERE  P.doc_code = " & DOCCODE & " "
        SQL = SQL & "   AND ISNULL(WB.smena_no, " & gWHS & ") = " & gWHS

        SQLEx(SQL)

        While reader.Read
            ' проверяем на  заполнение
            If reader.Item("отпущено") Is Nothing Then
                ans = MsgBox("Не указано сколько отпущено", vbExclamation, reader.Item("Название").ToString).ToString

                con.Close()
                ReturnToClient = True
                Exit Function
            End If

            '  проверяем на наличие остатков
            If CDbl(Format(CDbl(reader.Item("Остаток")) - CDbl(reader.Item("в_уч_ед")), "#0.000")) < 0 Then
                ans = MsgBox("На складе не хватает нужного количества " & var_Ed, vbExclamation, reader.Item("Название").ToString).ToString

                ReturnToClient = True
                con.Close()
                Exit Function
            End If
        End While
        con.Close()

        'ЗАПРОС НА ОТОБРАЖЕНИЕ ФОРМЫ ЯЧЕЕК СКЛАДОВ
        sTMP = FrmCellWare.Show2(MsgBox("Ресурсы будут списаны с ячеек по умолчанию." & vbCr & "Хотите просмотреть\изменить список?", vbQuestion + vbYesNo, "Информация") = vbYes, gWHS)
        If Len(sTMP) = 0 Then Exit Function

        SQL = ""
        SQL = SQL & "UPDATE WB "
        SQL = SQL & "SET    Num = ROUND(ISNULL(WB.Num, 0) - "
        SQL = SQL & "                                   CASE "
        SQL = SQL & "                                       when p.OKEI_Code = W.OKEI_Code "
        SQL = SQL & "                                       then P.NumFact "
        SQL = SQL & "                                       else P.NumFact/W.Weight "
        SQL = SQL & "                                    End "
        SQL = SQL & "                                       , 3), "
        SQL = SQL & "       Price = CASE "
        SQL = SQL & "                 WHEN ROUND(ISNULL(WB.Num, 0) - "
        SQL = SQL & "                                   CASE "
        SQL = SQL & "                                       when p.OKEI_Code = W.OKEI_Code "
        SQL = SQL & "                                       then P.NumFact "
        SQL = SQL & "                                       else P.NumFact/W.Weight "
        SQL = SQL & "                                    End "

        SQL = SQL & "                                                               , 3) > 0 THEN ( CASE "
        SQL = SQL & "                                                                           WHEN ( ISNULL(WB.Price, 0) * ISNULL(WB.Num, 0) - P.Price ) > 0 THEN ROUND(( ISNULL(WB.Price, 0) * ISNULL(WB.Num, 0) - P.Price ) / ( ISNULL(WB.Num, 0) - "
        SQL = SQL & "                                                                                                                                                                                                                    CASE "
        SQL = SQL & "                                                                                                                                                                                                                             when p.OKEI_Code = W.OKEI_Code "
        SQL = SQL & "                                                                                                                                                                                                                             then P.NumFact "
        SQL = SQL & "                                                                                                                                                                                                                             else P.NumFact/W.Weight "
        SQL = SQL & "                                                                                                                                                                                                                    End "
        SQL = SQL & "                                                                                                                                                                                                                                   ), 2) "
        SQL = SQL & "                                                                           ELSE 0 "
        SQL = SQL & "                                                                         END ) "
        SQL = SQL & "                 ELSE ISNULL(WB.Price, 0) "
        SQL = SQL & "               END "
        SQL = SQL & "FROM   Provodka AS P "
        SQL = SQL & "       LEFT OUTER JOIN WareBalance AS WB ON P.ware_code = WB.WareCode "
        SQL = SQL & "       LEFT JOIN Wares AS W ON WB.WareCode = W.WareCode "
        SQL = SQL & "WHERE  ( P.doc_code = " & DOCCODE & " ) "
        SQL = SQL & "   AND ( ISNULL(WB.smena_no, " & gWHS & ") = " & gWHS & " ) "

        SQLExNQ(SQL)

        SQL = ""
        SQL = SQL & "UPDATE P "
        SQL = SQL & "SET    P.Price              = ROUND(ISNULL(WB.Price, 0) * P.NumFact, 2), "
        SQL = SQL & "       P.PriceMinusDelivery = ROUND(ISNULL(WB.Price, 0) * P.NumFact, 2) "

        SQL = SQL & "FROM   provodka P "
        SQL = SQL & "       LEFT JOIN WareBalance WB "
        SQL = SQL & "       ON     P.ware_code = WB.WareCode "

        SQL = SQL & "WHERE  P.doc_code         = " & DOCCODE & " AND ISNULL(WB.smena_no, " & gWHS & ") = " & gWHS
        SQLExNQ(SQL)

        ' 3 - закрыть документ
        SQL = "UPDATE Documents SET data_reg = GETDATE(), UserID = " & gUserId & " WHERE doc_code=" & DOCCODE
        SQLExNQ(SQL)
        ' 4 - пересчитать требования

        'Списание с ячеек
        SQLExNQ("EXEC pr_Sklad_CellWare '<root></root>', '" & sTMP & "', '', 'Списание с ячеек' ")

        'ОБНОВЛЕНИЕ
        FillSkladInLVW_Return(CBool(IIf(GridViewVozvrat.GetFocusedDataRow.Item("В том числе НДС?").ToString = "Да", True, False)))

        ' надо пересчитать остатки
        отрисовка_остатков("Материалы")
        GridCtrlOstatki.DataSource = dt
        GridViewOstatki.BestFitColumns()

        Text = "Материалы"
        Exit Function
    End Function


    'СОЗДАНИЕ ДОКУМЕНТА ПЕРЕДАЧИ РЕСУРСОВ НА ДРУГОЙ СКЛАД
    Sub Sklad_WareSend_NewDoc(ByVal WHS As Integer)
        Try
            Dim NewDocText As String
            '1) Создание документа
            NewDocText = CN("EXEC pr_All_NewDoc @doc_type = 22, @MasterCode = " & gWHS & ", @cli_code = " & WHS).ToString
            SQLExNQ("UPDATE Documents SET data_make = NULL WHERE doc_type=22 AND doc_text='" & NewDocText & "'")

            '2) Обновление картинки
            ContxMnuStrRashod.Close()
            SplashScreenManager.ShowForm(Me, GetType(SplashScreen), True, True, False)
            отрисовка_расходов()
            GridCtrlRashod.DataSource = dt
            SplashScreenManager.CloseForm()

            '3) Фокус на созданный документ
            For i = 0 To GridViewRashod.RowCount - 1
                If CBool(StrComp(GridViewRashod.GetDataRow(i).Item("№").ToString, NewDocText)) = False Then
                    GridViewRashod.FocusedRowHandle = i
                    Exit For
                End If
            Next
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub


    'Показать\спрятать подвал
    Public Sub showSumAndDelivery(ByVal Show As Boolean, ByVal Where As Integer, Optional ByVal fReturnMode As Boolean = False)
        Try
            With Me
                If Not fReturnMode Then
                    Select Case Show
                        Case True
                            .txtDelivery.Visible = True
                            .lblDelivery.Visible = True
                            .cmdDelivery.Visible = True

                        Case False
                            .txtDelivery.Visible = False
                            .lblDelivery.Visible = False
                            .cmdDelivery.Visible = False
                            .lblSum.Visible = False

                            If Where = 1 Then .lblSum.Visible = False Else .lblSum2.Visible = False
                    End Select

                Else
                    Select Case Show
                        Case True
                            .txtDeliveryReturn.Visible = True
                            .lblDeliveryReturn.Visible = True
                            .cmdDeliveryReturn.Visible = True
                              .lblSumReturn.Visible = True

                        Case False
                            .txtDeliveryReturn.Visible = False
                            .lblDeliveryReturn.Visible = False
                            .cmdDeliveryReturn.Visible = False
                            .lblSumReturn.Visible = False
                    End Select
                End If
            End With
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    'Пишем в Lable "lblSum" сумму по документу
    Public Sub WriteSum(Optional ByVal fReturnMode As Boolean = False)
        Try
            With Me
                If .GridViewPrihod.RowCount = 0 Or (.GridViewPrihodClick.RowCount = 0) Then showSumAndDelivery(False, 1) : Exit Sub
                showSumAndDelivery(Not fReturnMode, 1)

                SQLEx("SELECT Sum(PriceMinusDelivery) " & _
                         "FROM provodka " & _
                         "WHERE doc_code = " & (.GridViewPrihod.GetFocusedDataRow.Item("doc_code").ToString))

                'Заполняем сумму
                If reader.HasRows = False Then Exit Sub

                .lblSum.Visible = True

                While reader.Read
                    If IsDBNull(reader.Item(0)) Then .lblSum.Text = "" : .lblSum.Visible = False : ConClose() : Exit Sub
                    'Итогo
                    Dim d As Double
                    If Trim(.txtDelivery.Text) = "" Then
                        d = 0
                    Else
                        d = CDbl(Trim(.txtDelivery.Text))
                    End If

                    TotalAndNDS = CDbl(CDbl(reader.Item(0)) + CDbl(IIF_S(.GridViewPrihod.GetFocusedDataRow.Item("В том числе НДС?").ToString = "Да", CStr(d / (1 + CInt(NeS(.GridViewPrihod.GetFocusedDataRow.Item("НДС (%)").ToString)) / 100)), Trim$(NeS(d.ToString)))))
                    .lblSum.Text = "ИТОГО: " & FormatCurrency(TotalAndNDS)

                    'ндс
                    If CInt(Nz(.GridViewPrihod.GetFocusedDataRow("НДС (%)"))) <> 0 Then 'Если есть ндс - пишем его

                        .lblSum.Text = .lblSum.Text & _
                                          " (НДС(" & .GridViewPrihod.GetFocusedDataRow.Item("НДС (%)").ToString & "%): " & _
                                          Format(TotalAndNDS * CInt(.GridViewPrihod.GetFocusedDataRow.Item("НДС (%)")) / 100, "Currency") & ")" & vbCr & _
                                          "ВСЕГО: " & Format(TotalAndNDS + (TotalAndNDS * CInt(.GridViewPrihod.GetFocusedDataRow.Item("НДС (%)"))) / 100, "Currency")
                    End If
                End While
                ConClose()
            End With
        Catch Ex As Exception
            MsgBox(Ex.Message)
            ConClose()
        End Try
    End Sub


    'Пишем в Lable "lblSum2" сумму по документу вкладки "Приход по заявке"
    Public Sub WriteSum2()
        Dim Sum As Double = 0, i As Integer

        With Me
            For i = 0 To .GridViewPoZaiyvkeClick.RowCount - 1
                If .GridViewPoZaiyvkeClick.GetDataRow(i).Item("Стоимость").ToString <> "" Then
                    Sum = Sum + CDbl(.GridViewPoZaiyvkeClick.GetDataRow(i).Item("Стоимость"))
                End If
            Next i

            If Sum <= 0 Then
                .lblSum2.Text = ""
                Exit Sub
            End If

            If .txtDelivery2.Text = "" Then .txtDelivery2.Text = "0"

            Sum = Sum + CDbl(IIf(.GridViewPoZaiyvke.GetFocusedDataRow.Item("НДС?").ToString = "Да", CDbl(.txtDelivery2.Text) / (1 + CInt(NeS(.GridViewPoZaiyvke.GetFocusedDataRow.Item("НДС(%)").ToString)) / 100), CDbl(Trim(.txtDelivery2.Text))))
            .lblSum2.Text = "ИТОГО: " & FormatCurrency(Sum)

            .lblSum2.Text = .lblSum2.Text & " (НДС(" & .GridViewPoZaiyvke.GetFocusedDataRow.Item("НДС(%)").ToString & "%): " & _
                                  Format(Sum * CInt(.GridViewPoZaiyvke.GetFocusedDataRow.Item("НДС(%)")) / 100, "Currency") & ")" & vbCr & _
                                  "ВСЕГО: " & Format(Sum + (Sum * CInt(.GridViewPoZaiyvke.GetFocusedDataRow.Item("НДС(%)"))) / 100, "Currency")

            .txtDelivery2.Visible = True
            .lblDelivery2.Visible = True
            .cmdDelivery2.Visible = True
            .lblSum2.Visible = True
            .txtDelivery2.Text = "0"

        End With
    End Sub


    Sub obnovka()
        Try
            GridCtrlPoZaiyvkeClick.Visible = False
            GridCtrlRashodClick.Visible = False
            GridCtrlNaklClick.Visible = False

            SQLEx("select convert (date,getdate())")
            While reader.Read
                gDate = (Convert.ToString(reader.Item(0)))
                gDate = gDate.Substring(0, 10)
            End While
            ConClose()

            отрисовка_заказов()
            GridCtrlPoZaiyvke.DataSource = dt

            отрисовка_прихода()
            GridCtrlPrihod.DataSource = dt
            ComboBox1.Text = "Материалы"
            GridCtrlOstatki.DataSource = dt
            отрисовка_расходов()
            GridCtrlRashod.DataSource = dt
            отрисовка_поставщиков()
            GridCtrlPostavshiki.DataSource = dt
            отрисовка_накладных(False)
            GridCtrlNakl.DataSource = dt
            Me.Text = "Модуль Склад - " & gWHS_NAME
            Me.Visible = True

            GridViewPoZaiyvke.Columns.Item(0).Visible = False
            GridCtrlPrihodClick.Visible = False
            GridViewPrihod.Columns.Item("doc_code").Visible = False
            GridViewPrihod.Columns.Item("doc_type").Visible = False
            GridViewPrihod.Columns.Item(2).Visible = False
            GridViewOstatki.Columns.Item(1).Visible = False
            GridViewNakl.Columns.Item(0).Visible = False

            Label14.Visible = False
            lblDelivery.Visible = False
            cmdDelivery.Visible = False
            txtDelivery.Visible = False
            DTPicker1.Enabled = False

        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try

    End Sub

	
    'изменение значения в комбобоксе выбора типа остатков
    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        Try
            отрисовка_остатков(DirectCast(ComboBox1.SelectedItem, KeyValuePair(Of Integer, String)).Key)
            GridCtrlOstatki.DataSource = dt
            GridViewOstatki.Columns.Item("Warecode").Visible = False
            GridViewOstatki.Columns.Item("WareType").Visible = False
            GridViewOstatki.BestFitColumns()
            Dim i As Integer
            For i = 0 To GridViewOstatki.Columns.Count - 4
                GridViewOstatki.Columns.Item(i).BestFit()
            Next
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    'радио кнопка - по всем ячейкам    
    Private Sub RadioButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonSum.Click
        Try
            opt = 1
            отрисовка_остатков(DirectCast(ComboBox1.SelectedItem, KeyValuePair(Of Integer, String)).Key)
            GridViewOstatki.Columns.Clear()
            GridCtrlOstatki.DataSource = dt
            GridViewOstatki.Columns.Item("Warecode").Visible = False
            GridViewOstatki.Columns.Item("WareType").Visible = False
            GridViewOstatki.BestFitColumns()
            Dim i As Integer
            For i = 0 To GridViewOstatki.Columns.Count - 4
                GridViewOstatki.Columns.Item(i).BestFit()
            Next
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    'радио кнопка - по ячейкам
    Private Sub RadioButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonYach.Click
        Try
            opt = 2
            отрисовка_остатков(DirectCast(ComboBox1.SelectedItem, KeyValuePair(Of Integer, String)).Key)
            GridViewOstatki.Columns.Clear()
            GridCtrlOstatki.DataSource = dt
            GridViewOstatki.Columns.Item("Warecode").Visible = False
            GridViewOstatki.Columns.Item("WareType").Visible = False
            GridViewOstatki.BestFitColumns()
            Dim i As Integer
            For i = 0 To GridViewOstatki.Columns.Count - 4
                GridViewOstatki.Columns.Item(i).BestFit()
            Next
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    'радио кнопка - незакрепленный
    Private Sub RadioButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonNezakr.Click
        Try
            opt = 3
            отрисовка_остатков(DirectCast(ComboBox1.SelectedItem, KeyValuePair(Of Integer, String)).Key)
            GridViewOstatki.Columns.Clear()
            GridCtrlOstatki.DataSource = dt
            GridViewOstatki.Columns.Item("Warecode").Visible = False
            GridViewOstatki.Columns.Item("WareType").Visible = False
            GridViewOstatki.BestFitColumns()
            Dim i As Integer
            For i = 0 To GridViewOstatki.Columns.Count - 4
                GridViewOstatki.Columns.Item(i).BestFit()
            Next
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize, Me.DoubleClick
        Try
            If WindowState = FormWindowState.Maximized Then
                cScrHeight = CInt(System.Windows.Forms.SystemInformation.PrimaryMonitorSize.Height)
                cScrWidth = CInt(System.Windows.Forms.SystemInformation.PrimaryMonitorSize.Width)
            Else
                cScrHeight = 100
                cScrWidth = 100
            End If

            fraReport.Width = CInt(cScrWidth / 5 * 2.8)
            XtraTabControl1.Width = CInt(cScrWidth / 5 * 2.8)
            Panel3.Width = CInt(cScrWidth / 5 * 2.2) - 5
            GridCtrlOstatki.Width = CInt(cScrWidth / 5 * 2.2) - 5
            ComboBox1.Width = CInt(cScrWidth / 5 * 2.2) - 30
            Panel3.Left = fraReport.Width
            PanelOstatki.Width = CInt(cScrWidth / 5 * 2.2) - 5
            PanelOstatki.Left = fraReport.Width
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub XtraTabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles XtraTabControl1.Click
        Try
            Select Case XtraTabControl1.SelectedTabPage.Text

                Case "Приход"

                    If BtnReportDay.Text = "Закрыть" Then

                        GridCtrlPrihodClick.Visible = False
                        GridCtrlRashodClick.Visible = False
                        GridViewPrihod.Columns.Clear()
                        dt = New DataTable
                        SQLEx(Sklad_Reports_Day_IN_GetSQL())
                        dt.Load(reader)
                        ConClose()
                        GridCtrlPrihod.DataSource = dt
                        BigButton.Text = "Провести приходный документ"
                        BigButton.Enabled = False
                        DTPicker1.Enabled = False

                    Else

                        отрисовка_прихода()
                        GridCtrlPrihod.DataSource = dt
                        GridViewPrihod.Columns.Item("doc_code").Visible = False
                        GridViewPrihod.Columns.Item("doc_type").Visible = False
                        GridViewPrihod.Columns.Item(2).Visible = False
                        GridCtrlPoZaiyvkeClick.Visible = False
                        GridCtrlRashodClick.Visible = False
                        GridCtrlNaklClick.Visible = False
                        GridCtrlVozvratClick.Visible = False
                        SmallButton.Visible = False
                        BigButton.Enabled = False
                        BigButton.Text = "Провести приходный документ"
                        GridCtrlPrihodClick.Visible = False
                        Label14.Visible = False
                        lblDelivery.Visible = False
                        cmdDelivery.Visible = False
                        txtDelivery.Visible = False
                        DTPicker1.Enabled = False
                        BigButton.Left = XtraTabControl1.Left
                        BigButton.Width = XtraTabControl1.Width - 5
                        DTPicker1.Value = Now
                        lblSum.Visible = False
                    End If

                Case "Приход по заявке"
                    отрисовка_заказов(gWareCode)
                    GridCtrlPoZaiyvke.DataSource = dt
                    GridViewPoZaiyvke.Columns.Item(0).Visible = False
                    GridCtrlPoZaiyvkeClick.Visible = False
                    GridCtrlRashodClick.Visible = False
                    GridCtrlNaklClick.Visible = False
                    GridCtrlVozvratClick.Visible = False
                    SmallButton.Visible = True
                    SmallButton.Enabled = False
                    SmallButton.Text = "Данные о документе"
                    BigButton.Enabled = False
                    BigButton.Text = "Провести приходный документ"
                    GridCtrlPrihodClick.Visible = False
                    Label14.Visible = False
                    lblDelivery.Visible = False
                    cmdDelivery.Visible = False
                    txtDelivery.Visible = False
                    DTPicker1.Enabled = False
                    BigButton.Left = XtraTabControl1.Left + SmallButton.Width + 10
                    BigButton.Width = XtraTabControl1.Width - SmallButton.Width - 15
                    DTPicker1.Value = Now

                    GridViewPoZaiyvke.BestFitColumns()
                    GridViewPoZaiyvke.Columns("Дата заявки").AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                    GridViewPoZaiyvke.Columns("Дата оплаты").AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                    GridViewPoZaiyvke.Columns("Дата счета").AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                    GridViewPoZaiyvke.Columns("НДС?").AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                Case "Расход"


                    If BtnReportDay.Text = "Закрыть" Then

                        GridViewRashod.Columns.Clear()
                        dt = New DataTable
                        SQLEx(Sklad_Reports_Day_OUT_GetSQL())
                        dt.Load(reader)
                        ConClose()
                        GridCtrlRashod.DataSource = dt
                        GridViewRashod.BestFitColumns()
                        GridCtrlRashod.Height = 700
                        BigButton.Text = "Провести приходный документ"
                        BigButton.Enabled = False
                        DTPicker1.Enabled = False

                    Else

                        If gWareCode = 0 Then
                            SplashScreenManager.ShowForm(Me, GetType(SplashScreen), True, True, False)
                        End If

                        GridViewRashod.Columns.Clear()
                        отрисовка_расходов()
                        GridCtrlRashod.DataSource = dt
                        ConClose()
                        GridViewRashod.Columns.Item("Код смены").Visible = False
                        GridCtrlPoZaiyvkeClick.Visible = False
                        GridCtrlRashodClick.Visible = False
                        GridCtrlNaklClick.Visible = False
                        GridCtrlVozvratClick.Visible = False
                        GridViewRashod.Columns.Item("doc_code").Visible = False
                        GridViewRashod.Columns.Item("doc_type").Visible = False
                        SmallButton.Visible = True
                        SmallButton.Enabled = False
                        SmallButton.Text = "Печать этикетки"
                        BigButton.Visible = True
                        BigButton.Text = "Провести приходный документ"
                        GridCtrlPrihodClick.Visible = False
                        Label14.Visible = False
                        lblDelivery.Visible = False
                        cmdDelivery.Visible = False
                        txtDelivery.Visible = False
                        DTPicker1.Enabled = False
                        BigButton.Left = XtraTabControl1.Left + SmallButton.Width + 10
                        BigButton.Width = XtraTabControl1.Width - SmallButton.Width - 15
                        DTPicker1.Value = Now
                        GridViewRashod.BestFitColumns()
                        GridCtrlRashod.Height = 215

                        If gWareCode = 0 Then
                            SplashScreenManager.CloseForm()
                        End If
                    End If

                Case "Список поставщиков"

                    отрисовка_поставщиков()
                    GridCtrlPostavshiki.DataSource = dt
                    GridCtrlPoZaiyvkeClick.Visible = False
                    GridCtrlRashodClick.Visible = False
                    GridCtrlNaklClick.Visible = False
                    GridCtrlVozvratClick.Visible = False
                    GridViewPostavshiki.Columns.Item("cli_code").Visible = False
                    SmallButton.Visible = False
                    BigButton.Enabled = True
                    BigButton.Text = "Новый поставщик"
                    GridCtrlPrihodClick.Visible = False
                    Label14.Visible = False
                    lblDelivery.Visible = False
                    cmdDelivery.Visible = False
                    txtDelivery.Visible = False
                    DTPicker1.Enabled = False
                    BigButton.Left = XtraTabControl1.Left
                    BigButton.Width = XtraTabControl1.Width - 5
                    DTPicker1.Value = Now

                Case "Накладные"

                    отрисовка_накладных(False, False)
                    GridCtrlNakl.DataSource = dt
                    GridViewNakl.Columns.Item(0).Visible = False
                    GridCtrlPoZaiyvkeClick.Visible = False
                    GridCtrlRashodClick.Visible = False
                    GridCtrlNaklClick.Visible = False
                    GridCtrlVozvratClick.Visible = False
                    SmallButton.Visible = False
                    BigButton.Enabled = False
                    BigButton.Text = "Распечатать документ"
                    GridCtrlPrihodClick.Visible = False
                    Label14.Visible = False
                    lblDelivery.Visible = False
                    cmdDelivery.Visible = False
                    txtDelivery.Visible = False
                    DTPicker1.Enabled = True
                    BigButton.Left = XtraTabControl1.Left
                    BigButton.Width = XtraTabControl1.Width - 5
                    DTPicker1.Value = Now

                Case "Требования-накладные"

                    отрисовка_накладных(False, True)
                    GridCtrlTrebNakl.DataSource = dt
                    GridViewTrebNakl.Columns.Item(2).Visible = False
                    GridCtrlPoZaiyvkeClick.Visible = False
                    GridCtrlRashodClick.Visible = False
                    GridCtrlNaklClick.Visible = False
                    GridCtrlVozvratClick.Visible = False
                    SmallButton.Visible = False
                    BigButton.Enabled = False
                    BigButton.Text = "Распечатать документ"
                    GridCtrlPrihodClick.Visible = False
                    Label14.Visible = False
                    lblDelivery.Visible = False
                    cmdDelivery.Visible = False
                    txtDelivery.Visible = False
                    DTPicker1.Enabled = True
                    BigButton.Left = XtraTabControl1.Left
                    BigButton.Width = XtraTabControl1.Width - 5
                    DTPicker1.Value = Now

                Case "Возврат поставщику"

                    Отрисовка_возврата()
                    GridCtrlVozvrat.DataSource = dt
                    GridCtrlPoZaiyvkeClick.Visible = False
                    GridCtrlRashodClick.Visible = False
                    GridCtrlNaklClick.Visible = False
                    SmallButton.Visible = False
                    BigButton.Enabled = False
                    BigButton.Text = "Провести возвратный документ"
                    GridCtrlPrihodClick.Visible = False
                    GridViewVozvrat.Columns.Item("doc_code").Visible = False
                    GridViewVozvrat.Columns.Item("cli_code").Visible = False
                    Label14.Visible = False
                    lblDelivery.Visible = False
                    cmdDelivery.Visible = False
                    txtDelivery.Visible = False
                    DTPicker1.Enabled = False
                    BigButton.Left = XtraTabControl1.Left
                    BigButton.Width = XtraTabControl1.Width - 5
                    DTPicker1.Value = Now

            End Select
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BigButton.Click  ' БОЛЬШАЯ КНОПКА
        Try
            Select Case XtraTabControl1.SelectedTabPage.Text

                Case "Приход"
                    gDoccode = GridViewPrihod.GetFocusedDataRow.Item("doc_code").ToString
                    PayInSlip = InputBox("Введите номер приходного ордера", "Приходной ордер")
                    fReturnFromSmena = IIF_B(CBool(StrComp(GridViewPrihod.GetFocusedDataRow.Item("doc_type").ToString, "1")) = False, False, True)

                    If PayInSlip = "" Then
                        If CBool(MsgBox("Нельзя осуществить приход. Не указан номер приходного ордера", , "Ошибка")) Then
                        End If
                    Else
                        If SaveMatStand(fReturnFromSmena) Then mew_in(PayInSlip)
                    End If

                    отрисовка_прихода()
                    GridCtrlPrihod.DataSource = dt
                    GridCtrlPrihodClick.Visible = False
                    BigButton.Enabled = False

                Case "Приход по заявке"
                    If Len(GridViewPoZaiyvke.GetFocusedDataRow.Item("Номер вход. док-та").ToString) = 0 Then
                        MsgBox("Нельзя осуществить приход. Не указан номер документа", , "Ошибка")
                        Exit Sub
                    Else
                        PayInSlip = InputBox("Введите номер приходного ордера", "Приходный ордер")
                        If PayInSlip = "" Then
                            MsgBox("Нельзя осуществить приход. Не указан номер приходного ордера", , "Ошибка")
                            Exit Sub
                        Else
                            order_fix(PayInSlip)
                        End If
                    End If
                    txtDelivery2.Visible = False
                    lblDelivery2.Visible = False
                    cmdDelivery2.Visible = False
                    lblSum2.Visible = False

                Case "Расход"
                    close_treb()

                Case "Список поставщиков"
                    new_cli()

                Case "Накладные"
                    If GridViewNakl.GetFocusedDataRow.Item("Документ").ToString = "Приходная накл." Or GridViewNakl.GetFocusedDataRow.Item("Документ").ToString = "Возврат со смены" Then
                        Dim path As String = pathMyDocs & "\frM4.fr3"
                        IO.File.WriteAllBytes(path, My.Resources.frM4)
                        M4_ToFR(FRX, CInt(GridViewNakl.GetFocusedDataRow.Item("doc_code")), 1, , gWHS)
                    Else
                        Dim path As String = pathMyDocs & "\frM11.fr3"
                        IO.File.WriteAllBytes(path, My.Resources.frM11)
                        M11_ToFR(FRX, GridViewNakl.GetFocusedDataRow.Item("doc_code").ToString, 1, , , , , gWHS)
                    End If

                Case "Возврат поставщику"
                    If ReturnToClient(CInt(GridViewVozvrat.GetFocusedDataRow.Item("doc_code"))) Then Exit Select

                    Отрисовка_возврата()
                    GridCtrlVozvrat.DataSource = dt

            End Select
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Public Function SaveMatStand(Optional ByVal fReturnFromSmena As Boolean = False) As Boolean
        Try
            Dim iWareCode As String
            Dim cWareBalanceNum As Double
            Dim cCost As Double
            Dim Average As Double
            Dim AverageNDS As Double
            Dim NewNum As Double
            Dim OldPrice As Double
            Dim OldPriceNDS As Double
            Dim OldNum As Double
            Dim OldNumAll As Double
            Dim NewNumAll As Double
            Dim cNum As Double
            Dim ANSW As Object
            Dim sTMP As String
            Dim fBlank As Boolean
            Dim fDocType As Integer
            Dim iWareCodeForBlank As Integer
            Dim iBlanks_ContractID As Integer 'Заготовки - фикт. заказ
            Dim iBlanks_NaklNo As Integer 'Заготовки - фикт. С\Н
            Dim iBlanks_Act_no As Integer 'Заготовки - единый Act_no для данного ELID
            Dim cCostLVWColIndex As String : cCostLVWColIndex = (IIf(Not fReturnFromSmena, "Стоимость", "Стоимость").ToString)

            If GridViewPrihod.GetFocusedDataRow.Item("doc_type").ToString = "22" Then fDocType = 22

            'Перемещение ресурсов между складами
            If fDocType = 22 Then
                SaveMatStand = Sklad_WareSend_Receive(GridViewPrihod.GetFocusedDataRow.Item("doc_code").ToString)
                If Not SaveMatStand Then ConClose() : Exit Function
            End If

            'Проверка на наличие указанного количества ресурсов в WareBalance
            If fReturnFromSmena Then

                For i = 0 To GridViewPrihodClick.RowCount - 1

                    iWareCode = GridViewPrihodClick.GetDataRow(i).Item("ware_code").ToString
                    fBlank = CBool(StrComp(GridViewPrihodClick.GetDataRow(i).Item("ware_type").ToString, "5")) = False
                    If fBlank = False Then fBlank = CBool(StrComp(GridViewPrihodClick.GetDataRow(i).Item("ware_type").ToString, "6")) = False

                    If fBlank Then
                        iWareCodeForBlank = CInt(CN("SELECT ware_code FROM provodka WHERE provodka_code = " & GridViewPrihodClick.GetDataRow(i).Item("provodka_code").ToString & " "))
                        SQLEx("SELECT ROUND(num, 3) AS Num FROM Balance WHERE Act_no =(" & iWareCodeForBlank & " ) AND smena_no = " & GridViewPrihod.GetFocusedDataRow.Item("PartnerCode").ToString)
                    Else
                        SQLEx("SELECT ROUND(num, 3) AS Num FROM WareBalance WHERE WareCode = " & iWareCode & " AND smena_no = " & GridViewPrihod.GetFocusedDataRow.Item("PartnerCode").ToString)
                    End If

                    If reader.Read Then cWareBalanceNum = CInt(reader.Item("Num")) Else cWareBalanceNum = 0
                    ConClose()

                    If cWareBalanceNum < CDbl(GridViewPrihodClick.GetDataRow(i).Item("Кол-во")) Then
                        ANSW = MsgBox("Указанное количество ресурса [ " & GridViewPrihodClick.GetDataRow(i).Item("Наименование").ToString & " ] за сменой не числится: " & cWareBalanceNum & " < " & GridViewPrihodClick.GetDataRow(i).Item("Кол-во").ToString & vbCr & _
                                      "Продолжить приходование? Данная позиция пойдет в количестве = [ " & cWareBalanceNum & " ] ?", vbExclamation + vbYesNo, "Ошибка прихода")

                        If CInt(ANSW) = vbNo Then Exit Function
                    End If
                Next

            End If

            'ПРИСВАИВАНИЕ КОЛИЧЕСТВА МАТЕРИАЛА ВЗЯТОГО ИЗ ЯЧЕЙКИ/ПОЛОЖЕННОГО В ЯЧЕЙКУ
            sTMP = FrmCellWare.Show2(MsgBox("Ресурсы будут записаны в ячейки по умолчанию." & vbCr & "Хотите просмотреть\изменить список?", vbQuestion + vbYesNo, "Информация") = vbYes, gWHS)
            If Len(sTMP) = 0 Then Exit Function

            'ЗАГОТОВКИ - ОПРЕДЕЛЕНИЕ ФИКТИВНЫХ ContractID, nakl_no
            If gBlanks Then

                'Contracts
                iBlanks_ContractID = CInt(CN("SELECT ISNULL((SELECT ContractID FROM Contracts WHERE ContractType = 1 AND doc_name = 'Сырье' AND ContractCode = 'Сырье' AND doc_begin IS NULL), 0)"))
                If iBlanks_ContractID = 0 Then
                    SQLExNQ("INSERT Contracts(doc_name, ContractCode, doc_begin, doc_end, doc_del, ArchDT, IsInner) VALUES ('Сырье', 'Сырье', NULL, GETDATE(), GETDATE(), GETDATE(), 1)")
                    iBlanks_ContractID = CInt(CN("SELECT SCOPE_IDENTITY()"))
                End If

                'Nakl
                iBlanks_NaklNo = CInt(CN("SELECT ISNULL((SELECT nakl_no FROM Nakl WHERE nakl_type = 4), 0)"))
                If iBlanks_NaklNo = 0 Then
                    SQLExNQ("INSERT Nakl (nakl_type, beg_data, end_data, Num, Fact, IsGroup, UserID) VALUES ( 4, GETDATE(), GETDATE(), 1, NULL, 1, " & gUserId & " )")
                    iBlanks_NaklNo = CInt(CN("SELECT SCOPE_IDENTITY()"))
                End If
            End If


            'Идем по таблице ресурсов и совершаем приход
            'ЗАПИСЬ В ЦИКЛЕ ПО СТРОКАМ ДОКУМЕНТА
            For i = 0 To GridViewPrihodClick.RowCount - 1

                iWareCode = GridViewPrihodClick.GetDataRow(i).Item("ware_code").ToString
                fBlank = CBool(StrComp(GridViewPrihodClick.GetDataRow(i).Item("ware_type").ToString, "5")) = False
                If fBlank = False Then fBlank = CBool(StrComp(GridViewPrihodClick.GetDataRow(i).Item("ware_type").ToString, "6")) = False

                'ЗАГОТОВКИ - ОПРЕДЕЛЕНИЕ\СОЗДАНИЕ Act_no ДЛЯ ДАННОГО ELID
                If fBlank Then
                    iBlanks_Act_no = CInt(CN("SELECT ISNULL((SELECT NCP.fact_no FROM NaklContract NCP JOIN Nakl N ON N.nakl_no = NCP.nakl_no JOIN NaklContract NCC ON NCC.ElementID = NCP.ElementID WHERE N.nakl_type = 4 AND NCC.fact_no = " & iWareCodeForBlank & " ), 0)"))
                    If iBlanks_Act_no = 0 Then

                        'Act
                        SQLExNQ("INSERT Act (Act_type) VALUES (5)")
                        iBlanks_Act_no = CInt(CN("SELECT SCOPE_IDENTITY()"))

                        'NC
                        SQLExNQ("INSERT NaklContract (nakl_no, ContractID, Num, ElementID, Cost, FactCost, fact_no) " & _
                                   "SELECT " & iBlanks_NaklNo & ", " & iBlanks_ContractID & ", 1, NC.ElementID, 0, 0, " & iBlanks_Act_no & " FROM NaklContract NC WHERE NC.fact_no = " & iWareCodeForBlank)

                        'IO
                        SQLExNQ("INSERT In_out(fact_no, fact_date, fact_col, nakl_no, ELID, COST, IsEnd, MasterStep, SMENANO, NaklContractID, UserID) " & _
                                   "SELECT NC.fact_no, GETDATE(), 1, NC.nakl_no, NC.ElementID, 0, 1, 0, " & gWHS & ", NC.NaklContractID, " & gUserId & " FROM NaklContract NC WHERE NC.NaklContractID = SCOPE_IDENTITY()")
                    End If
                End If


                'Вытаскиваем КолВо передаваемого на смене, с которой передавали 
                If fBlank Then
                    SQLEx("SELECT ROUND(Num, 3) AS Num FROM Balance WHERE Act_no=" & iWareCodeForBlank & "  AND smena_no=" & GridViewPrihod.GetFocusedDataRow.Item("PartnerCode").ToString)
                Else
                    SQLEx("SELECT ROUND(Num, 3) AS Num FROM WareBalance WHERE WareCode=" & iWareCode & " AND smena_no=" & GridViewPrihod.GetFocusedDataRow.Item("PartnerCode").ToString)
                End If
                While reader.Read
                    If reader.HasRows = True Then cWareBalanceNum = CDbl(reader.Item(0)) Else cWareBalanceNum = 0
                End While
                ConClose()

                If cWareBalanceNum > 0 Or Not fReturnFromSmena Then

                    If fReturnFromSmena Then
                        If cWareBalanceNum >= CDbl(GridViewPrihodClick.GetDataRow(i).Item("Кол-во")) Then 'Если ресурса хватает с избытком - уменьшаем остаток на смене
                            If fBlank Then
                                SQLExNQ("UPDATE Balance SET Num = ROUND(Num - REPLACE('" & GridViewPrihodClick.GetDataRow(i).Item("Кол-во").ToString & "', ',', '.'), 3) " & _
                                           "WHERE  Act_no= " & iWareCodeForBlank & "  AND smena_no=" & GridViewPrihod.GetFocusedDataRow.Item("PartnerCode").ToString)
                            Else
                                SQLExNQ("UPDATE WareBalance SET Num = ROUND(Num - REPLACE('" & GridViewPrihodClick.GetDataRow(i).Item("Кол-во").ToString & "', ',', '.'), 3) " & _
                                           "WHERE  WareCode=" & iWareCode & " AND smena_no=" & GridViewPrihod.GetFocusedDataRow.Item("PartnerCode").ToString)
                            End If
                        Else 'Если ресурса не хватает - берем сколько хватает
                            GridViewPrihodClick.GetDataRow(i).Item("Кол-во") = cWareBalanceNum

                            If fBlank Then
                                SQLExNQ("DELETE FROM Balance WHERE Act_no=" & iWareCodeForBlank & "  AND smena_no=" & GridViewPrihod.GetFocusedDataRow.Item("PartnerCode").ToString)
                            Else
                                SQLExNQ("DELETE FROM WareBalance WHERE WareCode=" & iWareCode & " AND smena_no=" & GridViewPrihod.GetFocusedDataRow.Item("PartnerCode").ToString)
                            End If
                        End If
                    End If


                    'ПРИХОД

                    OldNumAll = CDbl(CN("SELECT ISNULL((SELECT SUM(WB.NUM) AS Num FROM WareBalance WB JOIN Structure S ON S.StructureID = WB.smena_no WHERE WB.WareCode = " & iWareCode & " AND S.StructureType = 5), 0)"))

                    If fBlank Then
                        SQLEx("SELECT ISNULL(B.Num, 0) AS Num, ISNULL(E.PriceWHS, 0) AS Price " & _
                                "FROM Balance B JOIN Elements E ON E.ElementID = B.ELID WHERE B.smena_no = " & gWHS & " AND B.Act_no= " & iBlanks_Act_no)
                    Else
                        SQLEx("SELECT ISNULL(WB.Num, 0) AS Num, ISNULL(W.Price, 0) AS Price " & _
                                 "FROM WareBalance WB JOIN Wares W ON W.WareCode = WB.WareCode WHERE WB.smena_no = " & gWHS & " AND WB.WareCode= " & iWareCode)
                    End If

                    While reader.Read
                        If reader.HasRows = True Then '(A+B)/C
                            OldNum = CDbl(reader.Item(0))
                            OldPrice = CDbl(reader.Item(1))
                        Else
                            OldNum = 0
                            OldPrice = 0
                        End If
                    End While
                    ConClose()

                    cCost = CDbl(GridViewPrihodClick.GetDataRow(i).Item("С доставкой"))
                    cNum = CDbl(GridViewPrihodClick.GetDataRow(i).Item("Кол-во"))


                    'ОБНОВЛЕНИЕ СТОИМОСТИ ПРОВОДКИ
                    'в Provodka должно лежать количество в пришедших единицах
                    SQLExNQ("UPDATE Provodka SET NumFact             = " & Dbl2DB(cNum, 3) & ", " & _
                                                  " Price               = " & RCr(cCost) & ", " & _
                                                  " PriceMinusDelivery  = " & RCr(cCost) & " " & _
                             "WHERE  provodka_code     = " & GridViewPrihodClick.GetDataRow(i).Item("provodka_code").ToString)

                    NewNum = CDbl(Replace(GridViewPrihodClick.GetDataRow(i).Item("кол-во").ToString, ".", ",")) + OldNum

                    'Новое кол-во на складах
                    NewNumAll = CDbl(Replace(GridViewPrihodClick.GetDataRow(i).Item("кол-во").ToString, ".", ",")) + OldNumAll 'C

                    'РАСЧЕТ Price
                    If OldNumAll = 0 Then
                        Average = CDbl(GridViewPrihodClick.GetDataRow(i).Item(cCostLVWColIndex)) / CDbl(CDbl(Replace(GridViewPrihodClick.GetDataRow(i).Item("Кол-во").ToString, ".", ",")))  'A/C
                    Else
                        Average = (cCost + OldNumAll * OldPrice) / NewNumAll  '(A+B)/C
                    End If

                    'ЗАГОТОВКА - ЗАПИСЬ В NE
                    If fBlank Then
                        SQLExNQ("INSERT INTO NaklElement (fact_col, fact_no, fact_no_in, DTNaklElement, SmenaNo) " & _
                                   "SELECT " & Dbl2DB(NewNum, 3) & ", " & iBlanks_Act_no & ", " & iWareCode & ", GETDATE(), " & gWHS)
                    End If

                    'зАпись новой цены в справочник
                    If fBlank Then
                        SQLExNQ("UPDATE E SET E.PriceWHS = " & Dbl2DB(Average, 2) & " FROM Elements E JOIN NaklContract NC ON NC.ElementID = E.ElementID WHERE NC.fact_no = " & iWareCodeForBlank)
                    Else
                        SQLExNQ("UPDATE Wares SET Price = " & Dbl2DB(Average, 2) & " WHERE WareCode = " & iWareCode)
                    End If

                    'Запись в наличие склада
                    If fBlank Then

                        SQLExNQ(" UPDATE  Balance SET Num        = " & Dbl2DB(NewNum, 3) & _
                                " WHERE   Act_no              = " & iBlanks_Act_no & _
                                "     AND smena_no               = " & gWHS)
                        'если это первый ресурс на складе - INSERT
                        If CInt(CN("Select Count(*) from WareBalance where WareCode = " & iWareCode & " and smena_no = " & gWHS)) = 0 Then
                            SQLExNQ("INSERT INTO Balance (smena_no, Act_no, Num, ELID) " & _
                                     "SELECT " & gWHS & ", " & iBlanks_Act_no & ", " & Dbl2DB(NewNum, 3) & ", NC.ElementID FROM NaklContract NC WHERE NC.fact_no = " & iBlanks_Act_no)

                        End If
                        SQLExNQ(" UPDATE  WareBalance SET Num       = " & Dbl2DB(NewNum, 3) & _
                                    " WHERE   WareCode                  = " & iWareCode & _
                                     "     AND smena_no                  = " & gWHS)
                        'если это первый ресурс на складе - INSERT
                        If CInt(CN("Select Count(*) from WareBalance where WareCode = " & iWareCode & " and smena_no = " & gWHS)) = 0 Then
                            SQLExNQ("INSERT INTO WareBalance (smena_no, WareCode, Num) " & _
                                "VALUES (" & gWHS & ", " & iWareCodeForBlank & ", " & Dbl2DB(NewNum, 3) & ")")
                        End If
                    Else
                        SQLExNQ(" UPDATE  WareBalance SET Num       = " & Dbl2DB(NewNum, 3) & _
                                   " WHERE   WareCode                  = " & iWareCode & _
                                   "     AND smena_no                  = " & gWHS)
                        'если это первый ресурс на складе - INSERT
                        If CInt(CN("Select Count(*) from WareBalance where WareCode = " & iWareCode & " and smena_no = " & gWHS)) = 0 Then

                            SQLExNQ("INSERT INTO WareBalance (smena_no, WareCode, Num) " & _
                                       "VALUES (" & gWHS & ", " & iWareCode & ", " & Dbl2DB(NewNum, 3) & ")")

                        End If
                    End If
                End If
            Next 

            'Запись в ячейки
            SQLExNQ("EXEC pr_Sklad_CellWare '<root></root>', '" & sTMP & "', '', 'Запись в ячейки' ")

            If fDocType = 22 Then

                'Списание с ячеек
                sTMP = FrmCellWare.Show2(False, CInt(GridViewPrihod.GetFocusedDataRow.Item("PartnerCode")))
                SQLExNQ("EXEC pr_Sklad_CellWare '<root></root>', '" & sTMP & "', '', 'Списание с ячеек' ")
            End If

            SaveMatStand = True
            Exit Function
        Catch Ex As Exception
            MsgBox("Ошибка сохранения" & vbCrLf & "Документ не будет проведен.." & vbCr & vbCr & "Описание ошибки: """ & gErr, vbCritical, "Ошибка сохранения")
        End Try

    End Function

	
    Public Sub new_cli()
        Try
            frmClient.ShowDialog()
            отрисовка_поставщиков()
            GridCtrlPostavshiki.DataSource = dt

            Dim cli_code As String = CN("select max(cli_code) from clients").ToString

            For i = 0 To GridViewPostavshiki.RowCount - 1
                If GridViewPostavshiki.GetDataRow(i).Item("cli_code") = cli_code Then
                    GridViewPostavshiki.FocusedRowHandle = i
                    Exit For
                End If
            Next

        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub SmallButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SmallButton.Click
        Try
            Select Case XtraTabControl1.SelectedTabPage.Text
                Case "Приход по заявке"
                    ind = 2
                    InEdit = True
                    frmNewDocument.ShowDialog()
                Case "Расход"
                    Dim path As String = pathMyDocs & "\frM113.fr3"
                    IO.File.WriteAllBytes(path, My.Resources.frM113)
                    M113_ToFR(FRX, GridViewRashod.GetFocusedDataRow.Item("doc_code").ToString, 1, , gWHS)
            End Select
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    ' клик по заявкам
    Private Sub GridViewPoZaiyvke_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridViewPoZaiyvke.Click
        rowPoZaiyvkeClick = 2 ' два если был произведен клик по названию столбцы
    End Sub

	
    Private Sub GridViewPoZaiyvke_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridViewPoZaiyvke.DoubleClick
        Try
            If rowPoZaiyvkeClick = 1 Then
                If GridViewPoZaiyvke.GetFocusedDataRow Is Nothing Then
                    Exit Sub
                Else
                    Dim i As Integer = GridViewPoZaiyvke.FocusedRowHandle
                    SmallButton.PerformClick()
                    GridViewPoZaiyvke.FocusedRowHandle = i
                End If
            End If
            rowPoZaiyvkeClick = 0
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub


    Public Sub GridViewPoZaiyvke_RowClick(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.RowClickEventArgs) Handles GridViewPoZaiyvke.RowClick
        Try
            Dim Str As String = ""
            Dim i As Integer
            Dim j As Integer
            Dim разница As Double
            Dim MousePos As Object = MousePosition
            rowPoZaiyvkeClick = 1 'один если был клик по строке

            If GridViewPoZaiyvke.GetFocusedDataRow Is Nothing Then Exit Sub

            frmList.e2 = e
            frmList.sender2 = sender

            SQLEx("SELECT ISNULL(REPLACE ((SELECT CAST(ROW_NUMBER() OVER (ORDER BY W.WareType, W.WareName) AS VARCHAR) + ') ' + CASE " &
                   "                          WHEN W.WareType = 1 THEN 'Мат. - ' " &
                   "                          WHEN W.WareType = 2 THEN 'Ст. - ' " &
                   "                        END + W.WareName + '§' " &
                   "                 FROM   WareBalance WB " &
                   "                 LEFT JOIN CellWare CW ON CW.warecode = WB.warecode " &
                   "                                                 AND CW.smenaNo = WB.smena_no " &
                   "                 JOIN Wares W ON W.WareCode = WB.WareCode " &
                   "                 RIGHT JOIN OrdersBody OB ON OB.ware_code = W.WareCode " &
                   "                 WHERE  WB.smena_no = " & gWHS & " " &
                   "                    AND Ob.orders_code = " & GridViewPoZaiyvke.GetFocusedDataRow.Item("Orders_code").ToString &
                   "                 GROUP  BY WB.WareCode, " &
                   "                           W.WareType, " &
                   "                           W.WareName, " &
                   "                           WB.smena_no, " &
                   "                           WB.Num " &
                   "                 HAVING ROUND(WB.Num, 3) <> ROUND(SUM(CW.Num), 3) " &
                   "                         OR ROUND(SUM(CW.Num), 3) IS NULL ORDER BY W.WareType, W.WareName " &
                   "                 FOR XML PATH('')), '§', CHAR(13)), '') ")

            While reader.Read()
                Str = Convert.ToString(reader.Item(0))
            End While
            ConClose()

            If Len(Str) > 0 Then MsgBox("ВНИМАНИЕ - Данным ресурсам не присвоена ячейка хранения на складе или распределено не все имеющееся количество " & vbCr & vbCr & Str, vbInformation, "Информация")

            dt = New DataTable
            SQLEx(" SELECT W.Warecode as Warecode," &
            "        W.WareName  AS Название," &
            "        CASE" &
            "          WHEN OB.OKEI_Code is NULL THEN  O.OKEI_Name" &
            "          Else Ox.OKEI_Name" &
            "        END         As ед_изм," &
            "        ROUND(OB.Num, 3) Num," &
            "        ROUND(OB.NumFact, 3) NumFact," &
            "        OB.price  AS цена ," &
            "        CASE" &
            "          WHEN W.WareType = 1 THEN 'М'" &
            "          Else 'Ст'" &
            "        END         AS Тип," &
            "        OB.note as Примечание," &
            "        W.WareShortName  AS [Короткое наим.]," &
            "        CASE WHEN AN.ParentID IS NULL THEN '-'" &
            "        Else '+'" &
            "        END                 Аналоги" &
             " FROM   Wares AS W" &
             "        INNER JOIN WareBalance ON W.WareCode = WareBalance.WareCode" &
             "        RIGHT OUTER JOIN OrdersBody AS OB ON W.WareCode = OB.ware_code" &
             "        LEFT OUTER JOIN OKEI AS O ON W.OKEI_Code = O.OKEI_Code" &
             "       LEFT OUTER JOIN OKEI AS Ox ON OB.OKEI_Code = Ox.OKEI_Code" &
             "        LEFT OUTER JOIN (" &
             "        SELECT DISTINCT ParentID" &
             "        FROM  WareClone AS WC" &
             "        ) AS AN ON AN.ParentID = W.WareCode" &
             " Where OB.orders_code = " & GridViewPoZaiyvke.GetFocusedDataRow.Item("Orders_code").ToString &
             "    AND WareBalance.smena_no = " & gWHS)
            ' " ORDER  BY Название")

            i = 1
            j = 0

            dt.Columns.Add("NumFact", GetType(String))
            dt.Columns.Add("№ п/п", GetType(String))
            dt.Columns.Add("Название", GetType(String))
            dt.Columns.Add("Ед. изм.", GetType(String))
            dt.Columns.Add("Кол-во", GetType(String))
            dt.Columns.Add("Цена", GetType(String))
            dt.Columns.Add("Стоимость", GetType(String))
            dt.Columns.Add("+Доставка", GetType(String))
            dt.Columns.Add("Тип", GetType(String))
            dt.Columns.Add("Примечание", GetType(String))
            dt.Columns.Add("Короткое наим.", GetType(String))
            dt.Columns.Add("Аналоги", GetType(String))
            dt.Columns.Add("Warecode", GetType(String))
            GridViewPoZaiyvkeClick.Tag = ""

            While reader.Read
                If Not TypeOf (reader.Item("NumFact")) Is DBNull Then разница = Round(CDbl(reader.Item("Num")) - CDbl(reader.Item("NumFact")), 3) Else разница = CDbl(reader.Item("Num"))
                If разница < 0 Then разница = 0
                Dim стоим As Double
                If Not TypeOf (reader.Item("цена")) Is DBNull Then стоим = CDbl(reader.Item("цена")) * разница

                If стоим > 0 Then

                    GridViewPoZaiyvkeClick.Tag = SETP(GridViewPoZaiyvkeClick.Tag.ToString, "[" & j & "_Cost]", dec_sep(стоим.ToString))
                    GridViewPoZaiyvkeClick.Tag = SETP(GridViewPoZaiyvkeClick.Tag.ToString, "[" & j & "_InputCost]", dec_sep(стоим.ToString))
                    GridViewPoZaiyvkeClick.Tag = SETP(GridViewPoZaiyvkeClick.Tag.ToString, "[" & j & "_Delivery]", dec_sep(стоим.ToString))
                End If

                j = j + 1

                dt.Rows.Add()
                dt.Rows(dt.Rows.Count() - 1).Item(0) = reader.Item("NumFact")
                dt.Rows(dt.Rows.Count() - 1).Item(1) = i
                dt.Rows(dt.Rows.Count() - 1).Item(2) = reader.Item("Название")
                dt.Rows(dt.Rows.Count() - 1).Item(3) = reader.Item("Ед_изм")
                dt.Rows(dt.Rows.Count() - 1).Item(4) = разница
                dt.Rows(dt.Rows.Count() - 1).Item(5) = reader.Item("Цена")
                dt.Rows(dt.Rows.Count() - 1).Item(6) = ""
                dt.Rows(dt.Rows.Count() - 1).Item(7) = ""
                dt.Rows(dt.Rows.Count() - 1).Item(8) = reader.Item("Тип")
                dt.Rows(dt.Rows.Count() - 1).Item(9) = reader.Item("Примечание")
                dt.Rows(dt.Rows.Count() - 1).Item(10) = reader.Item("Короткое наим.")
                dt.Rows(dt.Rows.Count() - 1).Item(11) = reader.Item("Аналоги")
                dt.Rows(dt.Rows.Count() - 1).Item(12) = reader.Item("Warecode")
                i = i + 1

            End While

            ConClose()
            GridCtrlPoZaiyvkeClick.DataSource = dt
            GridCtrlPoZaiyvkeClick.Visible = True
            GridViewPoZaiyvkeClick.Columns.Item("Warecode").Visible = False
            GridViewPoZaiyvkeClick.Columns.Item("NumFact").Visible = False
            GridViewPoZaiyvkeClick.BestFitColumns()
            BigButton.Enabled = True
            SmallButton.Enabled = True

            If e.Button = Windows.Forms.MouseButtons.Right Then
                ContxMnuStrPoZaiyvke.Show(MousePos)
            End If

            txtDelivery2.Visible = False
            lblDelivery2.Visible = False
            cmdDelivery2.Visible = False
            lblSum2.Visible = False

        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Public Sub GridViewPrihod_RowClick(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.RowClickEventArgs) Handles GridViewPrihod.RowClick
        Try
            Dim sTMP As String = ""
            Dim Strl As String = ""
            Dim Strlp As Integer
            Dim MousePos As Object = MousePosition

            DeliverySum = 0
            lblSum.Visible = False

            If GridViewPrihod.GetFocusedDataRow Is Nothing Or BtnReportDay.Text = "Закрыть" Then Exit Sub

            If GridViewPrihod.GetFocusedDataRow.Item("doc_type").ToString = "22" Then

                SQLEx("SELECT ISNULL(REPLACE(STUFF((SELECT '§' + CASE W.WareType " &
                      "                               WHEN 1 THEN '[Мат.]' " &
                      "                                     ELSE '[Станд.]' " &
                      "                                   END + ' - ' + W.WareName " &
                      "                      FROM   Provodka P " &
                      "                             JOIN Documents D ON D.doc_code = P.doc_code " &
                      "                             LEFT JOIN WareBalance WB ON WB.smena_no = D.cli_code " &
                      "                                                         AND WB.WareCode = P.ware_code " &
                      "                             JOIN Wares W ON W.WareCode = P.ware_code " &
                      "                      WHERE  D.doc_code = " & GridViewPrihod.GetFocusedDataRow.Item("Doc_code").ToString & " " &
                      "                         AND WB.ID IS NULL " &
                      "                      ORDER  BY W.WareType, " &
                      "                                W.WareName " &
                      "                      FOR XML PATH('')), 1, 1, ''), '§', CHAR(13)), '')  ")
                While reader.Read()
                    sTMP = reader.Item(0).ToString
                End While
                ConClose()

                If Len(sTMP) > 0 Then
                    If MsgBox("ВНИМАНИЕ - Указанные ниже ресурсы не закреплены за складом-получателем." & vbCr & "Произвести их закрепление?" & vbCr & vbCr & sTMP, vbQuestion + vbYesNo, "Ресурсы не закреплены за складом-получателем") = vbYes Then

                        SQLExNQ("INSERT WareBalance " &
                         "       (WareCode, " &
                         "        smena_no, " &
                         "        Num, " &
                         "        Price) " &
                         " SELECT W.WareCode, " &
                         "       D.cli_code, " &
                         "       0, " &
                         "       0 " &
                         " FROM   Provodka P " &
                         "       JOIN Documents D ON D.doc_code = P.doc_code " &
                         "       LEFT JOIN WareBalance WB ON WB.smena_no = D.cli_code " &
                         "                                   AND WB.WareCode = P.ware_code " &
                         "       JOIN Wares W ON W.WareCode = P.ware_code " &
                         " WHERE  D.doc_code = " & GridViewPrihod.GetFocusedDataRow.Item("Doc_code").ToString & " " &
                         "   AND WB.ID IS NULL ")

                    End If
                End If
            End If

            SQLEx("SELECT ISNULL(REPLACE ((SELECT CAST(ROW_NUMBER() OVER (ORDER BY W.WareType, W.WareName) AS VARCHAR) + ') ' + CASE " &
             "                          WHEN W.WareType = 1 THEN 'Мат. - ' " &
             "                          WHEN W.WareType = 2 THEN 'Ст. - ' " &
             "                        END + W.WareName + '§' " &
             "                 FROM   WareBalance WB " &
             "                        LEFT JOIN CellWare CW ON CW.warecode = WB.warecode " &
             "                                               AND CW.smenaNo = WB.smena_no " &
             "                        JOIN Provodka P ON P.ware_code = WB.WareCode " &
             "                        JOIN Wares W ON W.WareCode = WB.WareCode " &
             "                 WHERE  WB.smena_no = " & gWHS & " " &
             "                    AND P.doc_code = " & GridViewPrihod.GetFocusedDataRow.Item("Doc_code").ToString & " " &
             "                 GROUP  BY WB.WareCode, " &
             "                           W.WareType, " &
             "                           W.WareName, " &
             "                           WB.smena_no, " &
             "                           WB.Num " &
             "                 HAVING ROUND(WB.Num, 3) <> ROUND(SUM(CW.Num), 3) " &
             "                         OR ROUND(SUM(CW.Num), 3) IS NULL ORDER BY W.WareType, W.WareName " &
             "                 FOR XML PATH('')), '§', CHAR(13)), '') ")

            While reader.Read()
                Strl = reader.Item(0).ToString
            End While
            ConClose()

            If Len(Strl) > 0 Then MsgBox("ВНИМАНИЕ - данным ресурсам не присвоена ячейка храниния на складе или распределено не все имеющееся количество: " & vbCr & vbCr & Strl, vbInformation, "Информация")

            SQLEx("SELECT COUNT(*) FROM Documents WHERE doc_Code = " & GridViewPrihod.GetFocusedDataRow.Item("Doc_code").ToString & " AND data_reg IS NOT NULL")
            While reader.Read()
                Strlp = CInt(reader.Item(0))
            End While
            ConClose()

            If Strlp > 0 Then
                MsgBox("ВНИМАНИЕ - Документ уже проведен", vbExclamation, "Документ уже проведен")
            End If


            SQLEx("SELECT CASE WHEN Ware_type = 2 THEN 'Ст.' WHEN Ware_type = 1 THEN 'Мат.' END as Тип, Наименование, " &
                  " Кол As [кол-во], " &
                  " PriceMinusDelivery / Кол As Цена, " &
                  " PriceMinusDelivery as [Стоимость],  " &
                  " Стоимость as [С доставкой], Ед, [Короткое наим.] as [Короткое наим.]  FROM ft_Sklad_DocList_In_DocContent_Roman (" & GridViewPrihod.GetFocusedDataRow.Item("Doc_code").ToString & ") ORDER BY ware_type, Наименование")
            While reader.Read()
                If reader.Item(5) Is DBNull.Value Then
                    DeliverySum = DeliverySum
                Else
                    DeliverySum = DeliverySum + CDbl(reader.Item(5))
                End If
            End While
            ConClose()

            dt = New DataTable

            dt.Columns.Add("ware_code", GetType(String))
            dt.Columns.Add("ware_type", GetType(String))
            dt.Columns.Add("provodka_code", GetType(String))
            dt.Columns.Add("PriceMinusDelivery", GetType(String))
            dt.Columns.Add("Тип", GetType(String))
            dt.Columns.Add("Наименование", GetType(String))
            dt.Columns.Add("Короткое наим.", GetType(String))
            dt.Columns.Add("Кол-во", GetType(String))
            If GridViewPrihod.GetFocusedDataRow.Item("doc_type").ToString = "22" Then
                dt.Columns.Add("К выдаче", GetType(String))
                dt.Columns.Add("Ед", GetType(String))
            Else
                dt.Columns.Add("Ед", GetType(String))
                dt.Columns.Add("Стоимость", GetType(String))
                dt.Columns.Add("Цена", GetType(String))
                dt.Columns.Add("С доставкой", GetType(String))
            End If

            SQLEx("SELECT ware_code, ware_type, provodka_code, PriceMinusDelivery, NumFact,  CASE WHEN Ware_type = 2 THEN 'Ст.' WHEN Ware_type = 1 THEN 'Мат.' END as Тип, Наименование, " &
                  " [Короткое наим.] as [Короткое наим.], Кол As [Кол-во], Ед, " &
                  " Round (PriceMinusDelivery, 5) as [Стоимость],  " &
                  " Round (PriceMinusDelivery / Кол, 5) As Цена, " &
                  " Стоимость as [С доставкой] FROM ft_Sklad_DocList_In_DocContent_Roman (" & GridViewPrihod.GetFocusedDataRow.Item("Doc_code").ToString & ") ORDER BY ware_type, Наименование")

            While reader.Read
                dt.Rows.Add()
                If reader.Item("ware_type") = 5 Or reader.Item("ware_type") = 6 Then
                    dt.Rows(dt.Rows.Count() - 1).Item(0) = CN("SELECT W.WareCode FROM In_out IO JOIN Wares W ON W.ElementID  = IO.ELID " & _
                                          " WHERE IO.fact_no = " & reader.Item("ware_code").ToString)
                Else
                    dt.Rows(dt.Rows.Count() - 1).Item(0) = reader.Item("ware_code")
                End If
                dt.Rows(dt.Rows.Count() - 1).Item(1) = reader.Item("ware_type")
                dt.Rows(dt.Rows.Count() - 1).Item(2) = reader.Item("provodka_code")
                dt.Rows(dt.Rows.Count() - 1).Item(3) = reader.Item("PriceMinusDelivery")
                dt.Rows(dt.Rows.Count() - 1).Item(4) = reader.Item("Тип")
                dt.Rows(dt.Rows.Count() - 1).Item(5) = reader.Item("Наименование")
                dt.Rows(dt.Rows.Count() - 1).Item(6) = reader.Item("Короткое наим.")
                dt.Rows(dt.Rows.Count() - 1).Item(7) = reader.Item("Кол-во")
                If GridViewPrihod.GetFocusedDataRow.Item("doc_type").ToString = "22" Then
                    dt.Rows(dt.Rows.Count() - 1).Item(8) = reader.Item("NumFact")
                    dt.Rows(dt.Rows.Count() - 1).Item(9) = reader.Item("Ед")
                Else
                    dt.Rows(dt.Rows.Count() - 1).Item(8) = reader.Item("Ед")
                    dt.Rows(dt.Rows.Count() - 1).Item(9) = System.Math.Round(CDbl(reader.Item("Стоимость")), 2)
                    dt.Rows(dt.Rows.Count() - 1).Item(10) = reader.Item("Цена")
                    dt.Rows(dt.Rows.Count() - 1).Item(11) = System.Math.Round(CDbl(reader.Item("С доставкой")), 2)
                End If
            End While

            ConClose()

            Label14.Text = "Наполнение документа : " & GridViewPrihod.GetFocusedDataRow.Item("Название").ToString

            GridViewPrihodClick.Columns.Clear()
            GridCtrlPrihodClick.DataSource = dt
            GridCtrlPrihodClick.Visible = True
            GridViewPrihodClick.Columns.Item("ware_code").Visible = False
            GridViewPrihodClick.Columns.Item("ware_type").Visible = False
            GridViewPrihodClick.Columns.Item("provodka_code").Visible = False
            GridViewPrihodClick.Columns.Item("PriceMinusDelivery").Visible = False
            GridViewPrihodClick.BestFitColumns()
            Label14.Visible = True
            lblDelivery.Visible = True
            cmdDelivery.Visible = True
            txtDelivery.Visible = True
            BigButton.Enabled = True
            txtDelivery.Text = "0"

            If GridViewPrihod.GetFocusedDataRow.Item("doc_type").ToString <> "22" Then
                cmdDelivery_Click(sender, e)
                WriteSum()
            Else
                lblDelivery.Visible = False
                cmdDelivery.Visible = False
                txtDelivery.Visible = False
            End If

            If e.Button = Windows.Forms.MouseButtons.Right Then
                ContxMnuStrPrihod.Show(MousePos)
            End If

        Catch Ex As Exception
            MsgBox(Ex.Message & " Ошибка в загрузке наполнения прихода")
            ConClose()
        End Try

    End Sub

	
    Public Sub cmdDelivery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelivery.Click
        Try
            Dim ID As String
            Dim CurSum As Double
            Dim cDelivery As Double
            Dim DeliveryStr As String = ""
            Dim NewPrice As Double
            Dim NewPriceStr As String
            Dim cPrice As Double
            Dim cNDS_inc As Boolean
            Dim cNDS As Double
            Dim CurTotal As Double

            DeliverySum = 0

            ID = GridViewPrihod.GetFocusedDataRow.Item("doc_code").ToString

            txtDelivery.Text = Replace$(txtDelivery.Text, ".", ",")
            If Not IsNumeric(txtDelivery.Text) Then txtDelivery.Text = "0"
            txtDelivery.Text = Trim(txtDelivery.Text)

            SQLEx("SELECT Sum(PriceMinusDelivery) FROM provodka WHERE doc_code = " & GridViewPrihod.GetFocusedDataRow.Item("doc_code").ToString)
            While reader.Read()
                If IsDBNull(reader.Item(0)) = True Then ConClose() : Exit Sub Else CurSum = CDbl(reader.Item(0))
            End While
            ConClose()

            cDelivery = CDbl(Trim(txtDelivery.Text))
            cNDS_inc = True
            cNDS = 0

            For i = 0 To GridViewPrihodClick.RowCount - 1

                cPrice = CDbl(GridViewPrihodClick.GetDataRow(i).Item("Стоимость"))

                If CurSum = 0 Then
                    NewPrice = 0
                    NewPriceStr = "0"
                Else
                    NewPrice = IIF_DBL(CurSum <> 0, cPrice + _
                                                    cPrice * Convert.ToDouble(IIF_DBL(cNDS_inc, cDelivery / (1 + cNDS / 100), cDelivery)) / CurSum, 0)
                    NewPriceStr = Replace(NewPrice.ToString, ",", ".")
                    DeliveryStr = Replace(CStr(Round(NewPrice - cPrice, 2)), ",", ".")
                End If

                CurTotal = CurTotal + CDbl(NewPrice)

                If DeliveryStr = "" Then DeliveryStr = "0"

                SQLExNQ(" UPDATE  provodka        " &
                        " SET    price         = " & NewPriceStr & ", " &
                        "       Delivery      = " & DeliveryStr & " " & _
                        " WHERE  provodka_code = " & GridViewPrihodClick.GetDataRow(i).Item("provodka_code").ToString)

            Next

            dt = New DataTable
            dt.Columns.Add("ware_code", GetType(String))
            dt.Columns.Add("ware_type", GetType(String))
            dt.Columns.Add("provodka_code", GetType(String))
            dt.Columns.Add("PriceMinusDelivery", GetType(String))
            dt.Columns.Add("Тип", GetType(String))
            dt.Columns.Add("Наименование", GetType(String))
            dt.Columns.Add("Короткое наим.", GetType(String))
            dt.Columns.Add("Кол-во", GetType(String))
            dt.Columns.Add("Ед", GetType(String))
            dt.Columns.Add("Стоимость", GetType(String))
            dt.Columns.Add("Цена", GetType(String))
            dt.Columns.Add("С доставкой", GetType(String))

            SQLEx("SELECT ware_code, ware_type, provodka_code, PriceMinusDelivery,  CASE WHEN Ware_type = 2 THEN 'Ст.' WHEN Ware_type = 1 THEN 'Мат.' END as Тип, Наименование, " &
                  " [Короткое наим.] as [Короткое наим.], Кол As [Кол-во], Ед, " &
                  " Round (PriceMinusDelivery, 5) as [Стоимость],  " &
                  " Round (PriceMinusDelivery / Кол, 5) As Цена, " &
                  " Стоимость as [С доставкой] FROM ft_Sklad_DocList_In_DocContent_Roman (" & GridViewPrihod.GetFocusedDataRow.Item("Doc_code").ToString & ") ORDER BY ware_type, Наименование")

            While reader.Read
                dt.Rows.Add()
                dt.Rows(dt.Rows.Count() - 1).Item(0) = reader.Item("ware_code")
                dt.Rows(dt.Rows.Count() - 1).Item(1) = reader.Item("ware_type")
                dt.Rows(dt.Rows.Count() - 1).Item(2) = reader.Item("provodka_code")
                dt.Rows(dt.Rows.Count() - 1).Item(3) = reader.Item("PriceMinusDelivery")
                dt.Rows(dt.Rows.Count() - 1).Item(4) = reader.Item("Тип")
                dt.Rows(dt.Rows.Count() - 1).Item(5) = reader.Item("Наименование")
                dt.Rows(dt.Rows.Count() - 1).Item(6) = reader.Item("Короткое наим.")
                dt.Rows(dt.Rows.Count() - 1).Item(7) = reader.Item("Кол-во")
                dt.Rows(dt.Rows.Count() - 1).Item(8) = reader.Item("Ед")
                dt.Rows(dt.Rows.Count() - 1).Item(9) = System.Math.Round(CDbl(reader.Item("Стоимость")), 2)
                dt.Rows(dt.Rows.Count() - 1).Item(10) = reader.Item("Цена")
                dt.Rows(dt.Rows.Count() - 1).Item(11) = System.Math.Round(CDbl(reader.Item("С доставкой")), 2)
            End While
            ConClose()

            Label14.Text = "Наполнение документа : " & GridViewPrihod.GetFocusedDataRow.Item("Название").ToString

            GridCtrlPrihodClick.DataSource = dt
            GridCtrlPrihodClick.Visible = True
            GridViewPrihodClick.Columns.Item("ware_code").Visible = False
            GridViewPrihodClick.Columns.Item("ware_type").Visible = False
            GridViewPrihodClick.Columns.Item("provodka_code").Visible = False
            GridViewPrihodClick.Columns.Item(3).Visible = False
            GridViewPrihodClick.BestFitColumns()
            lblSum.Visible = True
            WriteSum()

        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Public Sub Delivery(Optional ByVal id As String = Nothing)
        Try
            dt = New DataTable
            dt.Columns.Add("ware_code", GetType(String))
            dt.Columns.Add("ware_type", GetType(String))
            dt.Columns.Add("provodka_code", GetType(String))
            dt.Columns.Add("PriceMinusDelivery", GetType(String))
            dt.Columns.Add("Тип", GetType(String))
            dt.Columns.Add("Наименование", GetType(String))
            dt.Columns.Add("Короткое наим.", GetType(String))
            dt.Columns.Add("Кол-во", GetType(String))
            dt.Columns.Add("Ед", GetType(String))
            dt.Columns.Add("Стоимость", GetType(String))
            dt.Columns.Add("Цена", GetType(String))
            dt.Columns.Add("С доставкой", GetType(String))

            SQLEx("SELECT ware_code, ware_type, provodka_code, PriceMinusDelivery,  CASE WHEN Ware_type = 2 THEN 'Ст.' WHEN Ware_type = 1 THEN 'Мат.' END as Тип, Наименование, " &
                  " [Короткое наим.] as [Короткое наим.], Кол As [Кол-во], Ед, " &
                  " Round (PriceMinusDelivery, 5) as [Стоимость],  " &
                  " Round (PriceMinusDelivery / Кол, 5) As Цена, " &
                  " Стоимость as [С доставкой] FROM ft_Sklad_DocList_In_DocContent_Roman (" & GridViewPrihod.GetFocusedDataRow.Item("Doc_code").ToString & ") ORDER BY ware_type, Наименование")

            While reader.Read
                dt.Rows.Add()
                dt.Rows(dt.Rows.Count() - 1).Item(0) = reader.Item("ware_code")
                dt.Rows(dt.Rows.Count() - 1).Item(1) = reader.Item("ware_type")
                dt.Rows(dt.Rows.Count() - 1).Item(2) = reader.Item("provodka_code")
                dt.Rows(dt.Rows.Count() - 1).Item(3) = reader.Item("PriceMinusDelivery")
                dt.Rows(dt.Rows.Count() - 1).Item(4) = reader.Item("Тип")
                dt.Rows(dt.Rows.Count() - 1).Item(5) = reader.Item("Наименование")
                dt.Rows(dt.Rows.Count() - 1).Item(6) = reader.Item("Короткое наим.")
                dt.Rows(dt.Rows.Count() - 1).Item(7) = reader.Item("Кол-во")
                dt.Rows(dt.Rows.Count() - 1).Item(8) = reader.Item("Ед")
                dt.Rows(dt.Rows.Count() - 1).Item(9) = System.Math.Round(CDbl(reader.Item("Стоимость")), 2)
                dt.Rows(dt.Rows.Count() - 1).Item(10) = System.Math.Round(CDbl(reader.Item("Цена")), 2)
                dt.Rows(dt.Rows.Count() - 1).Item(11) = System.Math.Round(CDbl(reader.Item("С доставкой")), 2)
            End While
            ConClose()

            Label14.Text = "Наполнение документа : " & GridViewPrihod.GetFocusedDataRow.Item("Название").ToString

            GridCtrlPrihodClick.DataSource = dt
            GridCtrlPrihodClick.Visible = True
            GridViewPrihodClick.Columns.Item("ware_code").Visible = False
            GridViewPrihodClick.Columns.Item("ware_type").Visible = False
            GridViewPrihodClick.Columns.Item("provodka_code").Visible = False
            GridViewPrihodClick.Columns.Item(3).Visible = False
            GridViewPrihodClick.BestFitColumns()
            lblSum.Visible = True
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    'кнока Отчет за период
    Private Sub BtnReportPeriod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnReportPeriod.Click
        Try
            Panel3.Visible = False
            fraReport.Visible = False
            PanelOtchetPeriod.Visible = True
            PanelOtchetPeriod.Width = cScrWidth - 20
            PanelOtchetPeriod.Height = cScrHeight - 100
            GridCtrlPeriod.Height = cScrHeight - 268
            PanelOtchetPeriod.Top = 5
            PanelOtchetPeriod.Left = 5

            DateTimePicker1.Value = Convert.ToDateTime("01" & "." & Now.Month & "." & Now.Year)
            DateTimePicker2.Value = Convert.ToDateTime(Date.DaysInMonth(Now.Year, Now.Month) & "." & Now.Month & "." & Now.Year)
            CmbTypeRes.Text = "Материалы"

            отрисовкаОтчетаЗаПериод()
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub


    'отмена
    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Try
            Panel3.Visible = True
            fraReport.Visible = True
            PanelOtchetPeriod.Visible = False
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    ' выполнить
    Private Sub BtnGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnGo.Click, BtnReportPeriod.Click
        Try
            отрисовкаОтчетаЗаПериод()
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub отрисовкаОтчетаЗаПериод()
        Try
            Dim dtt As New DataTable
            Dim SQL As String

            If ChkBoxEmptyLine.Checked = False Then SQL = " WHERE [Остаток на начало] <> 0 OR [Приход] <> 0 OR [Расход] <> 0 OR [Остаток в конце] <> 0  " Else SQL = " "

            dtt.Clear()

            SQLEx(" SELECT warecode, Warename as [Наименование], wareshortname as [Короткое наим.], OKEI_Name as [Ед], [Остаток на начало] as [В начале], " &
                   " [Остаток на начало КГ] as [В начале КГ], Приход, [Приход КГ] as [Приход КГ], Расход, [Расход КГ] as [Расход КГ], [Остаток в конце] as [В конце], [Остаток в конце КГ] as [В конце КГ]   " &
                   " FROM dbo.ft_Sun_WareInOut ('" & DateTimePicker1.Value.ToString & "', '" & DateTimePicker2.Value.ToString & "', " & ListIndex & ", " & gWHS & " )  " &
                   SQL &
                   " ORDER BY WareName")

            dtt.Load(reader)
            ConClose()

            GridCtrlPeriod.DataSource = dtt
            GridViewPeriod.Columns.Item("warecode").Visible = False
            GridViewPeriod.BestFitColumns()
            GridViewPeriod.Columns("Наименование").AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            GridViewPeriod.Columns("Наименование").AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near

        Catch Ex As Exception
            ConClose()
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub ComboBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbTypeRes.TextChanged
        Try
            If CmbTypeRes.Text = "Материалы" Then ListIndex = 1 Else If CmbTypeRes.Text = "Стандартные изделия" Then ListIndex = 2
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub


    Sub GoToPrihod()
        Try
            Dim SQL As String
            SQL = ""
            SQL = SQL & "SELECT ISNULL(REPLACE ((SELECT CAST(ROW_NUMBER() OVER (ORDER BY W.WareType, W.WareName) AS VARCHAR) + ') ' + CASE "
            SQL = SQL & "                          WHEN W.WareType = 1 THEN 'Мат. - ' "
            SQL = SQL & "                          WHEN W.WareType = 2 THEN 'Ст. - ' "
            SQL = SQL & "                        END + W.WareName + '§' "
            SQL = SQL & "                 FROM   WareBalance WB "
            SQL = SQL & "                        LEFT JOIN CellWare CW ON CW.warecode = WB.warecode "
            SQL = SQL & "                                                 AND CW.smenaNo = WB.smena_no "
            SQL = SQL & "                        JOIN Wares W ON W.WareCode = WB.WareCode "
            SQL = SQL & "                 WHERE  WB.smena_no = " & gWHS & " "
            SQL = SQL & "                    AND WB.WareCode = " & GridViewOstatki.GetFocusedDataRow.Item("WareCode").ToString & " "
            SQL = SQL & "                 GROUP  BY WB.WareCode, "
            SQL = SQL & "                           W.WareType, "
            SQL = SQL & "                           W.WareName, "
            SQL = SQL & "                           WB.smena_no, "
            SQL = SQL & "                           WB.Num "
            SQL = SQL & "                 HAVING ROUND(WB.Num, 3) <> ROUND(SUM(CW.Num), 3) "
            SQL = SQL & "                         OR ROUND(SUM(CW.Num), 3) IS NULL ORDER BY W.WareType, W.WareName "
            SQL = SQL & "                 FOR XML PATH('')), '§', CHAR(13)), '') "

            SQLEx(SQL)
            While reader.Read
                If Len(reader.Item(0)) > 0 Then
                    MsgBox("Внимание - данным ресурсам не присвоена ячейка хранения на складе или распределено не все имеющееся количество: " & vbCr & vbCr & reader.Item(0).ToString, vbInformation, "Информация")
                    Exit Sub
                End If
            End While
            ConClose()

            Dim OKEI_Code As Integer

            SQLEx("SELECT TOP 1 OKEI_Code FROM Wares as W WHERE WareCode = '" & GridViewOstatki.GetFocusedDataRow.Item("WareCode").ToString)
            While reader.Read
                OKEI_Code = CInt(reader.Item(0))
            End While
            ConClose()
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    'клик по остаткам
    Private Sub GridViewOstatki_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridViewOstatki.Click
        rowOstatkiClick = 2 ' два если был произведен клик по названию столбцы
    End Sub

	
    'дабл клик по остаткам
    Private Sub GridViewOstatki_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridViewOstatki.DoubleClick
        Try
            If rowOstatkiClick = 1 Then 'если последний клик был на строке таблицы
                If GridViewOstatki.GetFocusedDataRow Is Nothing Then
                    Exit Sub
                Else
                    frmInOut.Text = "Накладные на передачу для ресурса: " & GridViewOstatki.GetFocusedDataRow.Item("Наименование").ToString
                    frmInOut.mMode = "Накладные на передачу для ресурса"
                    frmInOut.ShowDialog()
                End If
            End If
            rowOstatkiClick = 0 ' по умолчанию 0
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    'перетаскивание
    Private Sub GridViewOstatki_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles GridViewOstatki.MouseDown
        Try
            If opt = 1 Then
                Dim view As GridView = CType(sender, GridView)
                downHitInfo = Nothing
                Dim hitInfo As GridHitInfo = view.CalcHitInfo(New Point(e.X, e.Y))
                If Not Control.ModifierKeys = Keys.None Then Exit Sub
                If e.Button = MouseButtons.Right And hitInfo.RowHandle >= 0 Then downHitInfo = hitInfo
            End If
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub
	

    'перемещение остатков мышью
    Private Sub GridViewOstatki_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles GridViewOstatki.MouseMove
        Try
            Dim view As GridView = CType(sender, GridView)
            If e.Button = MouseButtons.Right And Not downHitInfo Is Nothing Then

                Dim dragSize As Size = SystemInformation.DragSize
                Dim DragRect As Rectangle = New Rectangle(New Point(downHitInfo.HitPoint.X - dragSize.Width / 2, downHitInfo.HitPoint.Y - dragSize.Height / 2), dragSize)

                If Not DragRect.Contains(New Point(e.X, e.Y)) Then
                    Dim row As DataRow = view.GetDataRow(downHitInfo.RowHandle)
                    view.GridControl.DoDragDrop(row, DragDropEffects.Move)
                    downHitInfo = Nothing
                    DevExpress.Utils.DXMouseEventArgs.GetMouseArgs(e).Handled = True
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, , "Ошибка при перемещении остатка.")
        End Try
    End Sub

	
    Private Sub BtnReportDay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnReportDay.Click
        Try
            Call showSumAndDelivery(False, 1)

            Label14.Visible = False
            lblDelivery.Visible = False
            cmdDelivery.Visible = False
            txtDelivery.Visible = False

            If BtnReportDay.Text = "Закрыть" Then

                XtraTabControl1.SelectedTabPage = XtraTabControl1.TabPages.Item(0)

                XtraTabControl1.TabPages.Item(1).PageVisible = True
                XtraTabControl1.TabPages.Item(3).PageVisible = True
                XtraTabControl1.TabPages.Item(5).PageVisible = True
                XtraTabControl1.TabPages.Item(6).PageVisible = False

                XtraTabControl1_Click(XtraTabControl1, e)

                GridCtrlPrihodClick.Visible = True
                GridCtrlRashodClick.Visible = True
                BtnReportDay.Text = "Отчет за день"

                GridViewPrihod.Columns.Clear()
                отрисовка_прихода()
                GridCtrlPrihod.DataSource = dt
                GridViewPrihod.Columns.Item("doc_code").Visible = False
                GridViewPrihod.Columns.Item("doc_type").Visible = False
                GridCtrlPrihod.Height = 190
                GridViewPrihod.Columns.Item("PartnerCode").Visible = False
                GridViewPrihod.BestFitColumns()

            Else

                XtraTabControl1.SelectedTabPage = XtraTabControl1.TabPages.Item(0)
                XtraTabControl1.TabPages.Item(1).PageVisible = False
                XtraTabControl1.TabPages.Item(3).PageVisible = False
                XtraTabControl1.TabPages.Item(5).PageVisible = False
                XtraTabControl1.TabPages.Item(6).PageVisible = True
                GridCtrlPrihodClick.Visible = False
                GridCtrlRashodClick.Visible = False
                GridViewPrihod.Columns.Clear()
                dt = New DataTable
                SQLEx(Sklad_Reports_Day_IN_GetSQL())
                dt.Load(reader)
                ConClose()
                GridCtrlPrihod.DataSource = dt
                BtnReportDay.Text = "Закрыть"
                GridCtrlPrihod.Height = 700

            End If
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub


    Private Sub ContxMnuStrPoZaiyvkeClick_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContxMnuStrPoZaiyvkeClick.ItemClicked
        Try
            If GridViewPoZaiyvkeClick.GetFocusedDataRow Is Nothing Then Exit Sub
            Select Case ContxMnuStrPoZaiyvkeClick.Items.IndexOf(e.ClickedItem)
                Case 0
                    If gIsTech = True Or gIsSklad = True Then
                        frmList.cMode = "Приход по заявке \ Список аналогов"
                        frmList.ShowDialog()
                    Else
                        MsgBox("ВНИМАНИЕ - Для выполнения замены ресурса на аналоги необходимо права " & gQ & " Технолога " & gQ, vbInformation, "Недостаточно прав")
                    End If
                Case 1
                    Dim xlApp As Microsoft.Office.Interop.Excel.Application
                    Dim PathString As String = ""
                    xlApp = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
                    PathString = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString & "Приход по заявке - наполнение " & Now.ToString("dd.MM.yyyy H.mm.ss") & ".xlsx"
                    GridViewPoZaiyvkeClick.ExportToXlsx(PathString)
                    xlApp.Workbooks.Open(PathString)
                    xlApp.Cells.Font.Name = "Times New Roman"
                    xlApp.Cells.Font.Size = 12
                    xlApp.Cells.Columns.AutoFit()
                    xlApp.Visible = True
            End Select
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Public Sub DTPicker1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTPicker1.ValueChanged
        Try
            отрисовка_накладных(True)
            GridCtrlNakl.DataSource = dt
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub GridControl1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles GridCtrlPrihod.MouseUp
        Try
            If e.Button = System.Windows.Forms.MouseButtons.Right Then
                If BtnReportDay.Text = "Закрыть" Then ContxMnuStrPrihod.Visible = False ' Else ContxMnuStrPrihod.Visible = True
            End If
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    'прибавление цены доставки
    Private Sub cmdDelivery2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDelivery2.Click
        Try
            Dim i As Integer

            'проверка значения поля
            txtDelivery2.Text = Replace(txtDelivery2.Text, ".", ",")
            If Not IsNumeric(txtDelivery2.Text) Then txtDelivery2.Text = "0"
            txtDelivery2.Text = Trim(Round(CDbl(txtDelivery2.Text), 2).ToString)

            Dim NewPrice As Double
            Dim CurSum As Double : CurSum = 0
            Dim CurTotal As Double

            Dim CurDelivery As Double
            Dim CurLineTotal As Double
            Dim CurCost As Double

            If GridViewPoZaiyvke.RowCount < 1 Then Exit Sub
            If GridViewPoZaiyvkeClick.RowCount < 1 Then Exit Sub

            'Вытаскиваем ИТОГО
            With GridViewPoZaiyvkeClick
                For i = 0 To .RowCount - 1
                    If CInt(.GetDataRow(i).Item("Warecode")) <> 0 Then
                        If Len(.GetDataRow(i).Item("Стоимость")) > 0 Then
                            CurSum = CurSum + CDbl(Ne(GP(.Tag.ToString, "[" & i & "_Cost]").ToString))
                        End If
                    End If
                Next i

                For i = 0 To .RowCount - 1
                    If Len(.GetDataRow(i).Item("Стоимость")) > 0 Then

                        If CurSum <> 0 Then
                            'функция Ne меняет NULL на 0
                            CurCost = CDbl(Ne(GP(.Tag.ToString, "[" & i & "_Cost]")))

                            CurDelivery = CDbl(IIf(GridViewPoZaiyvke.GetDataRow(i).Item("НДС?").ToString = "Да", CDbl(txtDelivery2.Text) / (1 + CDbl(NeS(GridViewPoZaiyvke.GetDataRow(i).Item("НДС(%)").ToString)) / 100), CDbl(txtDelivery2.Text)))
                            CurLineTotal = CurCost + CDbl(IIf(GridViewPoZaiyvke.GetDataRow(i).Item("НДС?").ToString = "Да", CDbl(txtDelivery2.Text) / (1 + CDbl(NeS(GridViewPoZaiyvke.GetDataRow(i).Item("НДС(%)").ToString)) / 100), CDbl(txtDelivery2.Text)))

                            NewPrice = CDbl(IIf(CurSum <> 0, CurCost + _
                                       CurCost * CDbl(IIf(GridViewPoZaiyvke.GetDataRow(i).Item("НДС?").ToString = "Да", CDbl(Trim(txtDelivery2.Text)) / (1 + CInt(GridViewPoZaiyvke.GetDataRow(i).Item("НДС(%)").ToString) / 100), CDbl(Trim(txtDelivery2.Text)))) / CurSum, 0))
                        Else
                            NewPrice = 0
                        End If

                        CurTotal = CurTotal + NewPrice

                        'Пишем новое значение "+Доставка"
                        .GetDataRow(i).Item("+Доставка") = Round(CDbl(NewPrice), 2)
                        .Tag = SETP(.Tag.ToString, "[" & i & "_Delivery]", NewPrice.ToString)

                        'Пишем новое значение "Цена"
                        If Not .GetDataRow(i).Item("Кол-во").ToString = "0" Then
                            .GetDataRow(i).Item("Цена") = Round(CDbl(.GetDataRow(i).Item("Стоимость")) / CDbl(.GetDataRow(i).Item("Кол-во")), 2)
                        Else
                            .GetDataRow(i).Item("Цена") = 0
                        End If
                    End If
                Next i
            End With

            GridViewPoZaiyvkeClick.RefreshData()
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try

        WriteSum2()
    End Sub

	
    'Замена Empty
    Function Ne(ByVal arg1 As Object, Optional ByVal arg2 As Object = 0) As Object
        Try
            If Trim(arg1.ToString) = "Empty" Then
                If TypeOf (arg1) Is DBNull Then
                    Ne = "0"
                Else
                    Ne = arg2
                End If
            Else
                Ne = Trim$(arg1.ToString)
            End If
        Catch Ex As Exception
            Ne = "0"
            MsgBox(Ex.Message)
        End Try
    End Function


    Private Sub ContxMnuStrNakl_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContxMnuStrNakl.ItemClicked
        Try
            If GridViewNakl.GetFocusedDataRow Is Nothing Then Exit Sub
            Select Case ContxMnuStrNakl.Items.IndexOf(e.ClickedItem)
                Case 0
                    ContxMnuStrNakl.Visible = False
                    Dim path As String = pathMyDocs & "\frM5.fr3"
                    IO.File.WriteAllBytes(path, My.Resources.frM5)
                    M5_ToFR(FRX, CInt(GridViewNakl.GetFocusedDataRow.Item("doc_code")), 1, , gWHS)
                Case 1
                    ContxMnuStrNakl.Visible = False
                    Dim path As String = pathMyDocs & "\frM112.fr3"
                    IO.File.WriteAllBytes(path, My.Resources.frM112)
                    M112_ToFR(FRX, CInt(GridViewNakl.GetFocusedDataRow.Item("doc_code")), 1)
                Case 2
                    ContxMnuStrNakl.Visible = False
                    Dim path As String = pathMyDocs & "\frM111.fr3"
                    IO.File.WriteAllBytes(path, My.Resources.frM111)
                    M111_ToFR(FRX, GridViewNakl.GetFocusedDataRow.Item("doc_code").ToString, 1)
                Case 3
                    frmAllClients.ShowDialog()
                    отрисовка_накладных(True)
                    GridCtrlNakl.DataSource = dt
                Case 5
                    отрисовка_накладных(True)
                    GridCtrlNakl.DataSource = dt
                Case 7
                    Dim xlApp As Microsoft.Office.Interop.Excel.Application
                    Dim PathString As String = ""
                    xlApp = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
                    PathString = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString & "Накладные " & Now.ToString("dd.MM.yyyy H.mm.ss") & ".xlsx"
                    GridViewNakl.ExportToXlsx(PathString)
                    xlApp.Workbooks.Open(PathString)
                    xlApp.Cells.Font.Name = "Times New Roman"
                    xlApp.Cells.Font.Size = 12
                    xlApp.Cells.Columns.AutoFit()
                    xlApp.Visible = True
            End Select
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub


    'контекстное меню на приходе
    Private Sub ContxMnuStrPrihod_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContxMnuStrPrihod.ItemClicked
        Try
            If GridViewPrihod.GetFocusedDataRow Is Nothing Then Exit Sub
            Dim i As Integer
            Select Case ContxMnuStrPrihod.Items.IndexOf(e.ClickedItem)
                Case 0   'изменить поставщика
                    i = GridViewPrihod.FocusedRowHandle
                    frmAllClients.ShowDialog()
                    отрисовка_прихода()
                    GridCtrlPrihod.DataSource = dt
                    GridViewPrihod.FocusedRowHandle = i
                Case 1   'обновление
                    отрисовка_прихода()
                    GridCtrlPrihod.DataSource = dt
                Case 2   'редактирование
                    i = GridViewPrihod.FocusedRowHandle
                    InEdit = True
                    frmNewDocument.ShowDialog()
                    GridViewPrihod.FocusedRowHandle = i
                Case 3   'удаление
                    ContxMnuStrPrihod.Close()
                    mnuDocDel(GridViewPrihod)
                Case 5  'экспорт в excel
                    Dim xlApp As Microsoft.Office.Interop.Excel.Application
                    Dim PathString As String = ""
                    xlApp = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
                    PathString = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString & "Приход " & Now.ToString("dd.MM.yyyy H.mm.ss") & ".xlsx"
                    GridViewPrihod.ExportToXlsx(PathString)
                    xlApp.Workbooks.Open(PathString)
                    xlApp.Cells.Font.Name = "Times New Roman"
                    xlApp.Cells.Font.Size = 12
                    xlApp.Cells.Columns.AutoFit()
                    xlApp.Visible = True
            End Select
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    'удаление документа из прихода
    Public Sub mnuDocDel(ByVal GV As DevExpress.XtraGrid.Views.Grid.GridView)

        Dim cmd As New SqlCommand
        Try
            If MsgBox("Действительно удалить документ: " & Chr(34) & GV.GetFocusedDataRow.Item("Название").ToString & Chr(34) & "?", vbQuestion + vbYesNo, "Удаление документа") = vbNo Then Exit Sub

            con.Open() 'открытие соединения
            cmd.Connection = con
            cmd.CommandText = "begin tran t1" 'начало транзакции
            cmd.ExecuteNonQuery()

            If XtraTabControl1.SelectedTabPage.Text = "Приход" Then
                SQLExNQ("DELETE FROM Documents WHERE Documents.doc_code=" & GV.GetFocusedDataRow.Item("doc_code").ToString, 1, cmd)
                SQLExNQ("DELETE FROM provodka WHERE provodka.doc_code=" & GV.GetFocusedDataRow.Item("doc_code").ToString, 1, cmd)
            Else
                SQLExNQ("DELETE FROM Documents WHERE Documents.doc_code=" & GV.GetFocusedDataRow.Item("doc_code").ToString, 1, cmd)
                SQLExNQ("DELETE FROM provodka WHERE provodka.doc_code=" & GV.GetFocusedDataRow.Item("doc_code").ToString, 1, cmd)
            End If

            cmd.CommandText = "commit tran t1"  ' закрепление транзакции
            cmd.ExecuteNonQuery()
            ConClose()
            showSumAndDelivery(False, 1)
            отрисовка_прихода()
            GridCtrlPrihod.DataSource = dt
            BigButton.Enabled = False
        Catch Ex As Exception
            cmd.CommandText = "rollback tran t1"  ' возврат транзакции
            cmd.ExecuteNonQuery()
            ConClose()
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub ContxMnuStrPoZaiyvke_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContxMnuStrPoZaiyvke.ItemClicked
        Try
            If GridViewPoZaiyvke.GetFocusedDataRow Is Nothing Then Exit Sub
            Select Case ContxMnuStrPoZaiyvke.Items.IndexOf(e.ClickedItem)
                Case 0
                    отрисовка_заказов()
                    GridCtrlPoZaiyvke.DataSource = dt
                Case 2
                    Dim xlApp As Microsoft.Office.Interop.Excel.Application
                    Dim PathString As String = ""
                    xlApp = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
                    PathString = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString & "Приход по заявке " & Now.ToString("dd.MM.yyyy H.mm.ss") & ".xlsx"
                    GridViewPoZaiyvke.ExportToXlsx(PathString)
                    xlApp.Workbooks.Open(PathString)
                    xlApp.Cells.Font.Name = "Times New Roman"
                    xlApp.Cells.Font.Size = 12
                    xlApp.Cells.Columns.AutoFit()
                    xlApp.Visible = True
            End Select
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    'контекстное меню на наполнение прихода
    Private Sub ContxMnuStrPrihodClick_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContxMnuStrPrihodClick.ItemClicked
        Try
            If GridViewPrihodClick.GetFocusedDataRow Is Nothing Then Exit Sub
            Select Case ContxMnuStrPrihodClick.Items.IndexOf(e.ClickedItem)
                Case 0
                    Dim iDocType As Integer
                    If XtraTabControl1.SelectedTabPage.Text = "Приход" Then
                        iDocType = CInt(GridViewPrihod.GetFocusedDataRow.Item("doc_type"))
                    Else
                        iDocType = 100
                    End If
                    InNumChange(GridViewPrihodClick, iDocType = 22)
                Case 1
                    mnuInDel()
                Case 3
                    Dim xlApp As Microsoft.Office.Interop.Excel.Application
                    Dim PathString As String = ""
                    xlApp = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
                    PathString = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString & "Приход наполнение " & Now.ToString("dd.MM.yyyy H.mm.ss") & ".xlsx"
                    GridViewPrihodClick.ExportToXlsx(PathString)
                    xlApp.Workbooks.Open(PathString)
                    xlApp.Cells.Font.Name = "Times New Roman"
                    xlApp.Cells.Font.Size = 12
                    xlApp.Cells.Columns.AutoFit()
                    xlApp.Visible = True
            End Select
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub GridViewVozvrat_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridViewVozvrat.Click
        Try
            If GridViewVozvrat.GetFocusedDataRow Is Nothing Then Exit Sub

            Dim SQL, sTMP As String

            SQL = ""
            SQL = SQL & "SELECT ISNULL(REPLACE ((SELECT CAST(ROW_NUMBER() OVER (ORDER BY W.WareType, W.WareName) AS VARCHAR) + ') ' + CASE "
            SQL = SQL & "                          WHEN W.WareType = 1 THEN 'Мат. - ' "
            SQL = SQL & "                          WHEN W.WareType = 2 THEN 'Ст. - ' "
            SQL = SQL & "                        END + W.WareName + '§' "
            SQL = SQL & "                 FROM   WareBalance WB "
            SQL = SQL & "                        LEFT JOIN CellWare CW ON CW.warecode = WB.warecode "
            SQL = SQL & "                                                 AND CW.smenaNo = WB.smena_no "
            SQL = SQL & "                        JOIN Provodka P ON P.ware_code = WB.WareCode "
            SQL = SQL & "                        JOIN Wares W ON W.WareCode = WB.WareCode "
            SQL = SQL & "                 WHERE  WB.smena_no = " & gWHS & " "
            SQL = SQL & "                    AND P.doc_code = " & GridViewVozvrat.GetFocusedDataRow.Item("doc_code").ToString & " "
            SQL = SQL & "                 GROUP  BY WB.WareCode, "
            SQL = SQL & "                           W.WareType, "
            SQL = SQL & "                           W.WareName, "
            SQL = SQL & "                           WB.smena_no, "
            SQL = SQL & "                           WB.Num "
            SQL = SQL & "                 HAVING ROUND(WB.Num, 3) <> ROUND(SUM(CW.Num), 3) "
            SQL = SQL & "                         OR ROUND(SUM(CW.Num), 3) IS NULL ORDER BY W.WareType, W.WareName "
            SQL = SQL & "                 FOR XML PATH('')), '§', CHAR(13)), '') "

            sTMP = CN(SQL).ToString
            If Len(sTMP) > 0 Then MsgBox("ВНИМАНИЕ - Данным ресурсам не присвоена ячейка хранения на складе или распределено не все имеющееся количество: " & vbCr & vbCr & sTMP, vbInformation, "Информация")

            'ПРОВЕРКА\УСТАНОВКА ФЛАГА ЮЗЕРА НА ДОКУМЕНТЕ
            If UserFlags_SET("doc_code", CInt(GridViewVozvrat.GetFocusedDataRow.Item("doc_code")), False) Then Exit Sub ' GoTo HELL

            FillSkladInLVW_Return(CBool(IIf(GridViewVozvrat.GetFocusedDataRow.Item("В том числе НДС?").ToString = "Да", True, False)))

            GridCtrlVozvratClick.DataSource = dt
            GridCtrlVozvratClick.Visible = True
            If GridViewVozvratClick.RowCount > 0 Then BigButton.Enabled = True Else BigButton.Enabled = False
            GridViewVozvratClick.Columns.Item("Warecode").Visible = False

        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub


    'РАСХОД СО СКЛАДА НА СМЕНУ ПО ТРЕБОВАНИЮ
    Public Sub close_treb()
        Try
            Dim V_code As String, SQL As String, ans As String, var_Ed As String, V_smena As String
            Dim lTMP As Integer
            Dim lDocType As String
            Dim fErr As Boolean
            Dim sTMP As String
            Dim idBlok As Integer
            Dim NewDocNote As String = ""

            ' если смена 486 - инструменты и ресурc, который передается измеряется в шт., значит нужно отдавать в заготовки
            If GridViewRashod.GetFocusedDataRow.Item("Код смены").ToString = "486" Then 'если смена 486 - инструменты

                Dim i As Integer

                'проверяем все ресурсы, которые передаем на смену
                For i = 0 To GridViewRashodClick.RowCount - 1

                    ' если этот материал измеряется в шт., то работаем с ним
                    If GridViewRashodClick.GetDataRow(i).Item("Ед").ToString = "шт." Then
                        If Int(CDbl(GridViewRashodClick.GetDataRow(i).Item("К выдаче"))) < CDbl(GridViewRashodClick.GetDataRow(i).Item("К выдаче")) Then
                            Dim style As String
                            Dim response As String
                            response = MsgBox("Нельзя выдавать не целое количество прутков ", vbCritical, "ВНИМАНИЕ").ToString
                            Exit Sub
                        End If
                    End If
                Next
            End If

            V_code = GridViewRashod.GetFocusedDataRow.Item("doc_code").ToString
            lDocType = GridViewRashod.GetFocusedDataRow.Item("doc_type").ToString

            'ПРОВЕРКА РЕГИСТРАЦИИ ДОКУМЕНТА
            If CInt(CN("SELECT COUNT(*) FROM Documents WHERE doc_code = " & V_code & " AND (CASE WHEN doc_type <> 22 AND data_reg IS NOT NULL THEN 1 WHEN doc_type = 22 AND data_make IS NOT NULL THEN 1 ELSE 0 END) = 1")) > 0 Then
                MsgBox("Внимание - Документ уже проведен\отправлен (возможно другим пользователем с другого компьютера)", vbExclamation, "Ошибка сохранения")
            End If

            'В СЛУЧАЕ ДОКУМЕНТА ПЕРЕДАЧИ СО СКЛАДА НА СКЛАД - УСТАНОВКА ДАТЫ СОЗДАНИЯ ДОКУМЕНТА
            If lDocType = "22" Then
                SQLExNQ("UPDATE Documents SET data_make = GETDATE() WHERE doc_code = " & V_code)

                'ПЕЧАТЬ НАКЛАДНОЙ
                If MsgBox("Печатать?", vbYesNo + vbQuestion, "Печать накладной на внутреннее перемещение ресурсов") = vbYes Then
                    Dim path As String = pathMyDocs & "\frNaklSend.fr3"
                    IO.File.WriteAllBytes(path, My.Resources.frNaklSend)
                    NaklSend_ToFR(FRX, 0, CInt(V_code), "", 1, , , "SELECT W.WareName AS EL, P.Num AS КОЛ, S_Client.StructureName AS Получатель, S_Master.StructureName AS Отправитель, '' doc_name, " & _
                                 "'Накладная на внутреннее перемещение №' + D.doc_text + ' от ' + CONVERT(VARCHAR, D.data_make, 104) NAKL, CASE W.WareType WHEN 1 THEN 'Мат.' ELSE 'Станд.' END Тип " & _
                                 " FROM Structure AS S_Master INNER JOIN Documents AS D INNER JOIN Provodka AS P ON D.doc_code = P.doc_code INNER JOIN Wares AS W ON P.ware_code = W.WareCode " & _
                                 " ON S_Master.StructureID = D.MasterCode INNER JOIN Structure AS S_Client ON D.cli_code = S_Client.StructureID WHERE D.doc_code = " & V_code & " AND D.data_make IS NOT NULL ORDER BY W.WareType, W.WareName")
                End If
                GoTo REFRESH
            End If

            SQL = ""
            SQL = SQL & "SELECT O.OKEI_Name                                            AS [Ед измерения], "
            SQL = SQL & "       P.NumFact                                              AS отпущено, "
            SQL = SQL & "       P.provodka_code, "
            SQL = SQL & "       W.WareType, "
            SQL = SQL & "       W.WareCode, "
            SQL = SQL & "       W.WareName                                             AS Название, "
            SQL = SQL & "       ROUND(ISNULL(WB.Num, 0) - ISNULL(qSend.NumSend, 0), 3) AS ОСТАТОК "
            SQL = SQL & "FROM   (SELECT SUM(P.Num) AS NumSend, "
            SQL = SQL & "               P.ware_code "
            SQL = SQL & "        FROM   Documents AS D "
            SQL = SQL & "               INNER JOIN Provodka AS P ON D.doc_code = P.doc_code "
            SQL = SQL & "        WHERE  ( D.doc_type = 22 ) "
            SQL = SQL & "           AND ( D.data_reg IS NULL ) "
            SQL = SQL & "        GROUP  BY P.ware_code) AS qSend "
            SQL = SQL & "       RIGHT OUTER JOIN OKEI AS O "
            SQL = SQL & "                        INNER JOIN Wares AS W ON O.OKEI_Code = W.OKEI_Code ON qSend.ware_code = W.WareCode "
            SQL = SQL & "       LEFT OUTER JOIN WareBalance AS WB ON W.WareCode = WB.WareCode "
            SQL = SQL & "       RIGHT OUTER JOIN Provodka AS P ON W.WareCode = P.ware_code "
            SQL = SQL & "WHERE  ( P.doc_code = " & V_code & " ) "
            SQL = SQL & "   AND ( ISNULL(WB.smena_no, " & gWHS & ") = " & gWHS & " ) "

            'проверка на пустые количества
            For i = 0 To GridViewRashodClick.RowCount - 1
                If Len(Trim(GridViewRashodClick.GetDataRow(i).Item("К выдаче").ToString)) = 0 Then
                    fErr = True
                End If
            Next

            If fErr Then
                If MsgBox("ВНИМАНИЕ - Не на всех позициях документа проставлено значение " & gQ & "К выдаче" & gQ & vbCr & vbCr & "Все равно провести документ, игнорируя эти позиции?", vbQuestion + vbYesNo, "Документ заполнен не полностью") = vbNo Then Exit Sub
                For i = 0 To GridViewRashodClick.RowCount - 1
                    If Len(Trim(GridViewRashodClick.GetDataRow(i).Item("К выдаче").ToString)) = 0 Then
                        GridViewRashodClick.GetDataRow(i).Item("К выдаче") = 0
                        ArrayTag(i) = "0"
                    End If
                Next
            End If

            SQLEx(SQL)

            Do While reader.Read
                For i = 0 To GridViewRashodClick.RowCount - 1
                    If GridViewRashodClick.GetDataRow(i).Item("Warecode").ToString = reader.Item("WareCode").ToString Then
                        '  проверяем на наличие остатков
                        If CDbl(NeS(ArrayTag(i))) > CDbl(reader.Item("Остаток")) Then
                            If (reader.Item("Ед измерения") Is DBNull.Value) Then var_Ed = "шт." Else var_Ed = reader.Item("Ед измерения").ToString
                            ans = MsgBox("ВНИМАНИЕ" & vbCr & vbCr & "Тип: " & IIF_S(reader.Item("WareType").ToString = "1", "Мат.", "Станд.") & vbCrLf & "Наименование: " & reader.Item("Название").ToString & vbCrLf & vbCrLf & "На отдающей смене нет " & reader.Item("отпущено").ToString & " " & var_Ed, vbInformation, reader.Item("Название").ToString).ToString
                            reader.Close()
                            ConClose()
                            Exit Sub
                        End If
                        Exit For
                    End If
                Next
            Loop
            reader.Close()
            ConClose()

            'ПРИСВАИВАНИЕ КОЛИЧЕСТВА МАТЕРИАЛА ВЗЯТОГО ИЗ ЯЧЕЙКИ/ПОЛОЖЕННОГО В ЯЧЕЙКУ
            sTMP = FrmCellWare.Show2(MsgBox("Ресурсы будут списаны с ячеек по умолчанию." & vbCr & "Хотите просмотреть\изменить список?", vbQuestion + vbYesNo, "Работа с ячейками склада") = vbYes, gWHS)
            If Len(sTMP) = 0 Then Exit Sub

            SQLExNQ("EXEC pr_Sklad_WareTransfer @DOCCODE = " & V_code & ", @XML = '" & Sklad_WareSend_MakeXML() & "', @UserID = " & gUserId & ", @WHS = " & gWHS & ", @CWXML = '" & sTMP & "', @NewDocNote='" & NewDocNote & "'")
            V_code = CN("select max(doc_code) from documents").ToString

            'ПЕЧАТЬ РАСХОДНОЙ НАКЛАДНОЙ (М-11)
            If MsgBox("Печатать?", vbYesNo + vbQuestion, "Печать расходной накладной (по форме М-11)") = vbYes Then
                Dim path As String = pathMyDocs & "\frM11.fr3"
                IO.File.WriteAllBytes(path, My.Resources.frM11)
                M11_ToFR(FRX, V_code, 1, , , , , gWHS) 'ПЕЧАТЬ ЧЕРЕЗ FastREport
            End If

            gWareClone = True
            'ПЕЧАТЬ ЛИСТА ЗАМЕНЫ НА АНАЛОГИ
            If gWareClone Then
                SQLExNQ("UPDATE  WCR SET WCR.doc_code = DD.ChildID FROM DocumentDocument DD JOIN WareCloneReplace WCR ON WCR.doc_code = DD.ParentID WHERE DD.ChildID = " & V_code)
            End If
            If gWareClone Then
                If CBool(CN("IF EXISTS (SELECT TOP 1 * FROM WareCloneReplace WHERE doc_code = " & V_code & " ) SELECT 1 ELSE SELECT 0 ")) Then
                    If MsgBox("В данном документе были произведены замены ресурсов на аналоги. Показать лист замены?", vbYesNo + vbQuestion, "Печать листа замены на аналоги") = vbYes Then
                        Dim path As String = pathMyDocs & "\frWareCloneReplace.fr3"
                        IO.File.WriteAllBytes(path, My.Resources.frWareCloneReplace)
                        WareCloneReplace_ToFR(FRX, CInt(V_code))
                    End If
                End If
            End If

            'если смена 486 - инструменты и ресурc, который передается измеряется в шт., значит нужно отдавать в заготовки
            If GridViewRashod.GetFocusedDataRow.Item("Код смены").ToString = "486" Then 'если смена 486 - инструменты

                'проверяем все ресурсы, которые передаем на смену
                For i = 0 To GridViewRashodClick.RowCount - 1

                    ' если этот материал измеряется в шт., то работаем с ним
                    If GridViewRashodClick.GetDataRow(i).Item("Ед").ToString = "шт." Then

                        'пускаем цикл по созданию записей (если количество прутков 3, то нужно записать 3 заготовки в 1 прутку)
                        For j = 0 To CInt(ArrayTag(i))

                            'смотрим есть ли в таблице прутки с остатком 0, то нужно остаток запиать на этот пруток, а не создавать новый
                            idBlok = CInt(CN("SELECT COUNT(idWaresInstrument) FROM WaresInstrument WHERE WareCode = " & GridViewRashodClick.GetDataRow(i).Item("Warecode").ToString & " and KolBalance = 0 "))    'то добавляем в таблицу WaresInctrument
                            If idBlok = 0 Then
                                'создаем новый пруток
                                SQLExNQ("INSERT INTO WaresInstrument([WareCode],[KolBalance]) Values (" & GridViewRashodClick.GetDataRow(i).Item("Warecode").ToString & ",1) ")
                            Else
                                'обновляем старую запись
                                SQLExNQ("UPDATE WaresInstrument SET [KolBalance] = 1 WHERE idWaresInstrument = (SELECT TOP 1 idWaresInstrument FROM WaresInstrument WHERE [WareCode] = " & GridViewRashodClick.GetDataRow(i).Item("Warecode").ToString & " and KolBalance = 0) ")
                            End If
                        Next
                    End If
                Next
            End If

            'ОБНОВЛЕНИЕ
REFRESH:
            отрисовка_расходов()
            GridCtrlRashod.DataSource = dt
            GridCtrlRashodClick.Visible = False

            opt = 1
            отрисовка_остатков(DirectCast(ComboBox1.SelectedItem, KeyValuePair(Of Integer, String)).Key)
            GridViewOstatki.Columns.Clear()
            GridCtrlOstatki.DataSource = dt
            GridViewOstatki.Columns.Item("Warecode").Visible = False
            GridViewOstatki.Columns.Item("WareType").Visible = False
            GridViewOstatki.BestFitColumns()
            Dim ij As Integer
            For ij = 0 To GridViewOstatki.Columns.Count - 4
                GridViewOstatki.Columns.Item(ij).BestFit()
            Next

            Exit Sub
        Catch
            MsgBox("ВНИМАНИЕ - Произошла ошибка. " & vbCr & "Пожалуйста, повторите попытку..." & vbCr & vbCr & "Описание ошибки: " & Err.Description, vbCritical, "Ошибка удаления   ")
        End Try
    End Sub

	
    Private Sub ContxMnuStrRashodClick_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContxMnuStrRashodClick.ItemClicked
        Try
            If GridViewRashodClick.GetFocusedDataRow Is Nothing Then Exit Sub
            Select Case ContxMnuStrRashodClick.Items.IndexOf(e.ClickedItem)
                Case 0
                    If gIsTech = True Or gIsSklad = True Then
                        frmList.cMode = "Расход \ Список аналогов"
                        frmList.ShowDialog()
                    Else
                        MsgBox("ВНИМАНИЕ - Для выполнения замены ресурса на аналоги необходимо права " & gQ & " Технолога " & gQ, vbInformation, "Недостаточно прав")
                    End If
                Case 1
                    IssueByDefault()
                Case 3
                    Dim xlApp As Microsoft.Office.Interop.Excel.Application
                    Dim PathString As String = ""
                    xlApp = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
                    PathString = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString & "Расход наполнение " & Now.ToString("dd.MM.yyyy H.mm.ss") & ".xlsx"
                    GridViewRashodClick.ExportToXlsx(PathString)
                    xlApp.Workbooks.Open(PathString)
                    xlApp.Cells.Font.Name = "Times New Roman"
                    xlApp.Cells.Font.Size = 12
                    xlApp.Cells.Columns.AutoFit()
                    xlApp.Visible = True
                Case 4
                    frGiveClick()
            End Select
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    'контекстное меню по остаткам
    Private Sub ContxMnuStrOstatki_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContxMnuStrOstatki.ItemClicked
        Try
            If GridViewOstatki.GetFocusedDataRow Is Nothing Then Exit Sub
            Select Case ContxMnuStrOstatki.Items.IndexOf(e.ClickedItem)
                Case 0
                    FrmCellWareSend.ShowDialog()
                Case 1
                    gWareCode = CInt(GridViewOstatki.GetFocusedDataRow("WareCode"))
                    XtraTabControl1.TabPages.Item(2).Show()
                    XtraTabControl1_Click(sender, e)
                    gWareCode = 0
                Case 2
                    gWareCode = CInt(GridViewOstatki.GetFocusedDataRow("WareCode"))
                    XtraTabControl1.TabPages.Item(1).Show()
                    XtraTabControl1_Click(sender, e)
                    gWareCode = 0
                Case 4
                    Dim xlApp As Microsoft.Office.Interop.Excel.Application
                    Dim PathString As String = ""
                    xlApp = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
                    PathString = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString & "Остатки " & Now.ToString("dd.MM.yyyy H.mm.ss") & ".xlsx"
                    GridViewOstatki.ExportToXlsx(PathString)
                    xlApp.Workbooks.Open(PathString)
                    xlApp.Cells.Font.Name = "Times New Roman"
                    xlApp.Cells.Font.Size = 12
                    xlApp.Cells.Columns.AutoFit()
                    xlApp.Visible = True
            End Select
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

    'окончание перемещения ресурса в приход
    Private Sub перетащить()
        Try
            Dim var_num As Double
            Dim ans2 As Object
            Dim cDocCode As Integer
            Dim ans As Object
            Dim curProvCode As Integer
            Dim ID As String

            Dim iWareCode As Integer

            Dim Sql As String
            Dim sTMP2 As String
            Dim sTMP As String : sTMP = "Тип: " & GridViewOstatki.GetFocusedDataRow.Item("Тип").ToString & vbCr & "Наименование: " & GridViewOstatki.GetFocusedDataRow.Item("Наименование").ToString & vbCr & "Остаток: " & GridViewOstatki.GetFocusedDataRow.Item("Ост.").ToString & " [" & GridViewOstatki.GetFocusedDataRow.Item("Ед").ToString & "]" & vbCr & "Цена: " & GridViewOstatki.GetFocusedDataRow.Item("Цена").ToString & "р." & vbCr & vbCr

            Dim i As Integer
            i = GridViewPrihod.FocusedRowHandle
            Sql = ""
            Sql = Sql & "SELECT ISNULL(REPLACE ((SELECT CAST(ROW_NUMBER() OVER (ORDER BY W.WareType, W.WareName) AS VARCHAR) + ') ' + CASE "
            Sql = Sql & "                          WHEN W.WareType = 1 THEN 'Мат. - ' "
            Sql = Sql & "                          WHEN W.WareType = 2 THEN 'Ст. - ' "
            Sql = Sql & "                        END + W.WareName + '§' "
            Sql = Sql & "                 FROM   WareBalance WB "
            Sql = Sql & "                        LEFT JOIN CellWare CW ON CW.warecode = WB.warecode "
            Sql = Sql & "                                                 AND CW.smenaNo = WB.smena_no "
            Sql = Sql & "                        JOIN Wares W ON W.WareCode = WB.WareCode "
            Sql = Sql & "                 WHERE  WB.smena_no = " & gWHS & " "
            Sql = Sql & "                    AND WB.WareCode = " & GridViewOstatki.GetFocusedDataRow.Item("WareCode").ToString
            Sql = Sql & "                 GROUP  BY WB.WareCode, "
            Sql = Sql & "                           W.WareType, "
            Sql = Sql & "                           W.WareName, "
            Sql = Sql & "                           WB.smena_no, "
            Sql = Sql & "                           WB.Num "
            Sql = Sql & "                 HAVING ROUND(WB.Num, 3) <> ROUND(SUM(CW.Num), 3) "
            Sql = Sql & "                         OR ROUND(SUM(CW.Num), 3) IS NULL ORDER BY W.WareType, W.WareName "
            Sql = Sql & "                 FOR XML PATH('')), '§', CHAR(13)), '') "

            sTMP2 = CN(Sql).ToString
            If Len(sTMP2) > 0 Then
                MsgBox("ВНИМАНИЕ - Данным ресурсам не присвоена ячейка хранения на складе или распределено не все имеющееся количество: " & vbCr & vbCr & sTMP, vbInformation, "Информация")
                Exit Sub
            End If

            If Not (XtraTabControl1.SelectedTabPage.Text = "Приход" Or XtraTabControl1.SelectedTabPage.Text = "Возврат поставщику") Then Exit Sub

            'Закладка на текущую запись
            If XtraTabControl1.SelectedTabPage.Text = "Приход" Then
                ID = GridViewPrihod.GetFocusedDataRow.Item("doc_code").ToString
            Else
                ID = GridViewVozvrat.GetFocusedDataRow.Item("doc_code").ToString
            End If
			
            'Спрашиваем Количество и Цену
            'изменено запрос цены, для возможности переводить в другие единицы (кг)
            Dim flagPrihod As Integer  '  1 - если пришло в др.ед и 0 - если пришло в тех ед, в которых заказывали
            Dim MyPrice As Double
            Dim MySum As Double
            Dim OKEI_Code As Integer

            OKEI_Code = CInt(CN("SELECT TOP 1 OKEI_Code FROM Wares as W WHERE WareCode = '" & GridViewOstatki.GetFocusedDataRow.Item("WareCode").ToString & " ' "))

            ans = frmKolWare.Show2(CInt(GridViewOstatki.GetFocusedDataRow.Item("WareType")), _
                                   CN("SELECT TOP 1 WareName FROM Wares WHERE WareCode = '" & GridViewOstatki.GetFocusedDataRow.Item("WareCode").ToString & " ' ").ToString, _
                                   "0", _
                                   CN("SELECT TOP 1 OKEI_Name FROM Wares as W, OKEI as O WHERE W.OKEI_Code = O.OKEI_Code AND WareCode = '" & GridViewOstatki.GetFocusedDataRow.Item("WareCode").ToString & " ' ").ToString, _
                                   CDbl(CN("SELECT TOP 1 ISNULL(Weight, 0) As Weight FROM Wares WHERE WareCode = '" & GridViewOstatki.GetFocusedDataRow.Item("WareCode").ToString & " ' ")), _
                                   0)

            flagPrihod = frmKolWare.MyShow(CInt(GridViewOstatki.GetFocusedDataRow.Item("WareType")), _
                                   CN("SELECT TOP 1 WareName FROM Wares WHERE WareCode = '" & GridViewOstatki.GetFocusedDataRow.Item("WareCode").ToString & " ' ").ToString, _
                                   "0", _
                                   CN("SELECT TOP 1 OKEI_Name FROM Wares as W, OKEI as O WHERE W.OKEI_Code = O.OKEI_Code AND WareCode = '" & GridViewOstatki.GetFocusedDataRow.Item("WareCode").ToString & " ' ").ToString, _
                                   CDbl(CN("SELECT TOP 1 ISNULL(Weight, 0) As Weight FROM Wares WHERE WareCode = '" & GridViewOstatki.GetFocusedDataRow.Item("WareCode").ToString & " ' ")))
            If ans Is Nothing Then Exit Sub
            If ans.ToString = "" Then Exit Sub
            If CDbl(ans) <= 0 Then Exit Sub

            'изменен алгоритм работы frmKolWare, поэтому он вернет в тех единицах, которые выбра пользователь. Чтоб не исправлять всюду ниже,
            '                                     изменяем значение поля ans
            var_num = CDbl(ans)
            If (flagPrihod = 1) Then
                ans = CDbl(ans) / (CDbl(CN("SELECT TOP 1 ISNULL(Weight, 0) As Weight FROM Wares WHERE WareCode = '" & GridViewOstatki.GetFocusedDataRow.Item("WareCode").ToString & " ' ")))
                OKEI_Code = 2
            End If

            'Спрашиваем "Цену"
            'MyPrice - берет цену из представления, если ее там нет, обнуляет
            'MySum - считает стоимость элементов: ans - умножает только чтовведеное количество на цену из бд
            MyPrice = CDbl(NeS(GridViewOstatki.GetFocusedDataRow.Item("Цена").ToString))
            MySum = Round(CDbl(ans) * MyPrice, 2) ' стоимость считается как старая единица на цену, потому что стоимость не меняется, меняется только цена

            'ans2 - выводит новую сумму
            'ans2 = CheckNum(InputBox(sTMP & "Стоимость: ", CN("SELECT TOP 1 WareName FROM Wares WHERE WareCode = '" & GridViewOstatki.GetFocusedDataRow.Item("WareCode").ToString & " ' ").ToString, MySum.ToString))
            Dim text As String = CN("SELECT TOP 1 WareName FROM Wares WHERE WareCode = '" & GridViewOstatki.GetFocusedDataRow.Item("WareCode").ToString & " ' ").ToString
            ans2 = CheckNum(frmInputBox.show2(sTMP, MySum.ToString, text))

            If ans2 Is Nothing Then Exit Sub
            If TypeOf (ans2) Is DBNull Then Exit Sub
            If CDbl(ans2) < 0 Then Exit Sub

            If CDbl(ans2) = 0 Then
                MsgBox("ВНИМАНИЕ - Приходование ресурсов по нулевым ценам запрещено.", vbInformation, "Неверный ввод")
                Exit Sub
            End If

            iWareCode = CInt(GridViewOstatki.GetFocusedDataRow.Item("WareCode"))

            If XtraTabControl1.SelectedTabPage.Text = "Приход" Then
                If Not Len(CN("select provodka_code from provodka where doc_code = " & GridViewPrihod.GetFocusedDataRow.Item("doc_code").ToString & " AND ware_type = " & GridViewOstatki.GetFocusedDataRow.Item("WareType").ToString & " AND ware_code = " & iWareCode).ToString) = 0 Then
                    MsgBox("Позиция уже присутствует", vbCritical, GridViewOstatki.GetFocusedDataRow.Item("Наименование").ToString)
                    Exit Sub
                End If

            Else
                If Not Len(CN("select provodka_code from provodka where doc_code = " & GridViewVozvrat.GetFocusedDataRow.Item("doc_code").ToString & " and ware_code = " & iWareCode).ToString) = 0 Then
                    MsgBox("Позиция уже присутствует", vbCritical, GridViewOstatki.GetFocusedDataRow.Item("Наименование").ToString)
                    Exit Sub
                End If
            End If

            'Создание проводки
            If var_num > 0 And CDbl(ans2) >= 0 Then

                If XtraTabControl1.SelectedTabPage.Text = "Приход" Then cDocCode = CInt(GridViewPrihod.GetFocusedDataRow.Item("doc_code")) _
                             Else cDocCode = CInt(GridViewVozvrat.GetFocusedDataRow.Item("doc_code"))

                Dim WareType As String = ""
                Select Case GridViewOstatki.GetFocusedDataRow.Item("Тип").ToString
                    Case "Мат."
                        WareType = "1"
                    Case "Ст."
                        WareType = "2"
                End Select

                'в INSERT запрос добавлено OKEI_Code
                SQLExNQ("INSERT INTO provodka (doc_code, ware_code, ware_type, Num, Price, PriceMinusDelivery, OKEI_Code) " & _
                           "VALUES (" & cDocCode & ", " & iWareCode & ", " & WareType & ", REPLACE('" & Round(var_num, 3) & "', ',', '.'), REPLACE('" & RCv(ans2) & "', ',', '.'), REPLACE('" & RCv(ans2) & "', ',', '.'), " & OKEI_Code & ") ", 1)

                curProvCode = CInt(CN("SELECT SCOPE_IDENTITY()", 1))
            End If

            If XtraTabControl1.SelectedTabPage.Text = "Приход" Then
                If GridViewPrihod.GetFocusedDataRow.Item("doc_type").ToString <> "22" Then
                    Delivery(ID)
                    WriteSum()
                Else
                    lblDelivery.Visible = False
                    cmdDelivery.Visible = False
                    txtDelivery.Visible = False
                End If
            ElseIf XtraTabControl1.SelectedTabPage.Text = "Возврат поставщику" Then
                Отрисовка_возврата()
                GridCtrlVozvrat.DataSource = dt
            End If

            GridViewPrihod.FocusedRowHandle = i

        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub IssueByDefault()
        Try
            Dim X As Double, V_provodka_code As Integer, i As Integer

            If GridViewRashodClick.RowCount = 0 Then Exit Sub
            If GridViewRashod.GetFocusedDataRow("doc_type").ToString = "22" Then Exit Sub

            For i = 0 To GridViewRashodClick.RowCount - 1
                If CDbl(GridViewRashodClick.GetDataRow(i).Item(6)) >= CDbl(GridViewRashodClick.GetDataRow(i).Item(3)) Then
                    X = CDbl(dec_sep(GridViewRashodClick.GetDataRow(i).Item(3).ToString))
                Else
                    X = CDbl(dec_sep(GridViewRashodClick.GetDataRow(i).Item(6).ToString))
                End If
                V_provodka_code = CInt(GridViewRashodClick.GetDataRow(i).Item("provodka_code")) 
                Out_fix(V_provodka_code, X, i)
            Next
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub frGiveClick()
        Try
            ContxMnuStrRashodClick.Close()
            Dim SQL As String
            Dim sParent As String = ""
            Dim sParent2 As String = ""
            Dim sParent3 As String = ""
            Dim sParent4 As String = ""

            sParent = GridViewRashod.GetFocusedDataRow.Item("№").ToString 
            sParent2 = " от " & GridViewRashod.GetFocusedDataRow.Item("Дата").ToString 
            sParent3 = vbCrLf & " заказ " & GridViewRashod.GetFocusedDataRow.Item("Заказ").ToString
            sParent4 = GridViewRashod.GetFocusedDataRow.Item("Получатель").ToString 

            SQL = "SELECT * FROM ft_Sklad_DocList_Out_DocContent ( " & gWHS & ", " & GridViewRashod.GetFocusedDataRow.Item("doc_code").ToString & " ) ORDER BY Вид, Наименование"

            Dim path As String = pathMyDocs & "\frGive.fr3"
            IO.File.WriteAllBytes(path, My.Resources.frGive)

            frGive_ToFR(FRX, 1, SQL, sParent2, sParent, sParent3, sParent4)   'ПЕЧАТЬ ЧЕРЕЗ FastREport
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub


    Public Sub Out_fix(ByVal V_code As Integer, ByVal v_count As Double, Optional ByVal i As Integer = -1)
        Try
            If i = -1 Then
                GridViewRashodClick.GetFocusedDataRow.Item(4) = v_count
                ArrayTag(GridViewRashodClick.GetFocusedDataSourceRowIndex) = v_count.ToString
            Else
                GridViewRashodClick.GetDataRow(i).Item(4) = v_count
                ArrayTag(i) = v_count.ToString
            End If
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub GridCtrlRashodClick_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles GridCtrlRashodClick.DragDrop
        Try
            Sklad_WareSend()
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub GridCtrlRashodClick_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles GridCtrlRashodClick.DragOver
        Try
            If e.Data.GetDataPresent(GetType(DataRow)) Then
                e.Effect = DragDropEffects.Move
            Else
                e.Effect = DragDropEffects.None
            End If
        Catch ex As Exception
            MsgBox(ex.Message, , "Ошибка при перемещении исполнителя над операциями.")
        End Try
    End Sub

    ' ДРОП НА СОДЕРЖАНИЕ РАСХОДНОГО ДОКУМЕНТА
    Sub Sklad_WareSend()  
        Dim ANSW As Object
        Dim i As Integer

        'ПРОВЕРКА НА ДУБЛЬ
        For i = 0 To GridViewRashodClick.RowCount - 1
            If GridViewRashodClick.GetDataRow(i).Item("Warecode").ToString = GridViewOstatki.GetFocusedDataRow.Item("WareCode").ToString Then
                MsgBox("ВНИМАНИЕ - Позиция уже присутствует в документе", vbInformation, "Неверный ввод")
            End If
        Next

        'ЗАПРОС КОЛИЧЕСТВА
        ANSW = CheckNum(InputBox("Укажите количество [" & GridViewOstatki.GetFocusedDataRow.Item("Наименование").ToString & "]: ", "Запрос количества ресурса"))
        If ANSW Is DBNull.Value Then Exit Sub
        If CDbl(ANSW) <= 0 Then Exit Sub

        If CDbl(ANSW) > CDbl(GridViewOstatki.GetFocusedDataRow.Item("Ост.")) Then
            MsgBox("ВНИМАНИЕ - В наличии нет " & ANSW & " [" & GridViewOstatki.GetFocusedDataRow.Item("Наименование").ToString & "], только " & GridViewOstatki.GetFocusedDataRow.Item("Ост.").ToString & " [" & GridViewOstatki.GetFocusedDataRow.Item("Наименование").ToString & "]", vbInformation, "Неверный ввод")
            Exit Sub
        End If

        '1) ЗАПИСЬ ПРОВОДКИ
        If ComboBox1.Text = "Материалы" Then
            SQLExNQ("INSERT INTO Provodka (doc_code, ware_code, Num, ware_type) " & _
                       "VALUES ( " & GridViewRashod.GetFocusedDataRow.Item("doc_code") & ", " & GridViewOstatki.GetFocusedDataRow.Item("WareCode") & ", " & Sng2DB(CDbl(ANSW), 3) & ", " & 1 & ") ")
        ElseIf ComboBox1.Text = "Стандартные изделия" Then
            SQLExNQ("INSERT INTO Provodka (doc_code, ware_code, Num, ware_type) " & _
                       "VALUES ( " & GridViewRashod.GetFocusedDataRow.Item("doc_code") & ", " & GridViewOstatki.GetFocusedDataRow.Item("WareCode") & ", " & Sng2DB(CDbl(ANSW), 3) & ", " & 2 & ") ")
        Else
            SQLExNQ("INSERT INTO Provodka (doc_code, ware_code, Num, ware_type) " & _
                       "VALUES ( " & GridViewRashod.GetFocusedDataRow.Item("doc_code") & ", " & GridViewOstatki.GetFocusedDataRow.Item("WareCode") & ", " & Sng2DB(CDbl(ANSW), 3) & ", " & 5 & ") ")
        End If

       '2) ОБНОВЛЕНИЕ КАРТИНКИ И ФОКУС НА НОВУЮ ПРОВОДКУ
        i = GridViewRashod.FocusedRowHandle
        GridViewRashod_RowClick(GridViewRashod, frmList.e2)
        GridViewRashodClick.FocusedRowHandle = i

    End Sub

    Private Sub GridCtrlRashodClick_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles GridCtrlRashodClick.MouseUp
        Try
            If e.Button = System.Windows.Forms.MouseButtons.Right And Len(GridViewRashodClick.GetFocusedDataRow.Item("Аналоги").ToString) > 0 Then
                ContxMnuStrRashodClick.Items.Item(0).Enabled = True
                ContxMnuStrRashodClick.Show()
            Else
                ContxMnuStrRashodClick.Items.Item(0).Enabled = False
            End If
        Catch Ex As Exception
            MsgBox(Ex.Message, vbCritical, "Ошибка")
        End Try
    End Sub

    Private Function lvwIN() As Object
        Throw New NotImplementedException
    End Function

    'Заполнение списка возврата
    Public Sub FillSkladInLVW_Return(Optional ByVal fNSD_inc As Boolean = False)
        Try
            Dim strSourceTypeName As String = ""
            Dim strSourceTypeICON As Integer
            Dim SumAndNDS As Double
            Dim SumMinusNDS As Double
            Dim SQL As String
            Dim cPrice As Double

            GridViewVozvratClick.Columns.Clear()

            dt = New DataTable
            dt.Columns.Add("Warecode", GetType(String))
            dt.Columns.Add("Тип", GetType(String))
            dt.Columns.Add("Наименование", GetType(String))
            dt.Columns.Add("Кол-во", GetType(String))
            dt.Columns.Add("Цена", GetType(String))
            If fNSD_inc Then
                dt.Columns.Add("Стоимость (без НДС)", GetType(String))
            Else
                dt.Columns.Add("Стоимость (без НДС)", GetType(String))
                dt.Columns.Add("Сумма налога", GetType(String))
                dt.Columns.Add("Стоимость (с НДС)", GetType(String))
            End If
            dt.Columns.Add("Ед", GetType(String))
            dt.Columns.Add("Короткое наим.", GetType(String))


            SQL = "SELECT DISTINCT Wares.WareName AS Наименование, ISNULL(Wares.WareShortName, '') AS [Короткое наим.], provodka.Num AS [Кол], OKEI.OKEI_Name AS [Ед], OK.OKEI_Name AS [Ед2], provodka.Price AS Стоимость, provodka.Delivery AS [Доставка], provodka.NumFact, provodka.provodka_code, provodka.ware_code, provodka.ware_type, provodka.PriceMinusDelivery " & _
          "FROM (provodka LEFT JOIN Wares ON provodka.ware_code = Wares.WareCode) LEFT JOIN Ware_Group ON Wares.WareCode = Ware_Group.WareCode " & _
          "LEFT JOIN [ОАОКЭТЗ].[dbo].[OKEI] ON OKEI.OKEI_Code =  CASE WHEN Provodka.OKEI_Code is null THEN Wares.OKEI_Code ELSE Provodka.OKEI_Code END " & _
          "LEFT JOIN [ОАОКЭТЗ].[dbo].[OKEI] As OK ON OK.OKEI_Code =  Wares.OKEI_Code  " & _
          "WHERE (((provodka.doc_code)=" & GridViewVozvrat.GetFocusedDataRow.Item("doc_code").ToString & " AND (provodka.ware_type=1 OR provodka.ware_type=2)))"

            SQLEx(SQL)

            While reader.Read
                Select Case CInt(reader.Item("ware_type")) : Case 1 : strSourceTypeName = "Мат." : strSourceTypeICON = 6
                    Case 2 : strSourceTypeName = "Станд." : strSourceTypeICON = 7 : End Select

                If CDbl(reader.Item("кол")) > 0 And SumAndNDS > 0 Then cPrice = Round(SumAndNDS / CDbl(reader.Item("кол")), 2) Else cPrice = 0
                If reader.Item("Ед").ToString <> reader.Item("Ед2").ToString Then
                    If reader.Item("ед").ToString = "кг." Or reader.Item("ед").ToString = "кг" Then 'значит принято в кг, надо записывать уч.ед
                        Dim kol As Double
                        kol = Round(CDbl(reader.Item("кол")) / CDbl(CN("SELECT Weight FROM  Wares WHERE (Wares.WareCode = '" & reader.Item("ware_code").ToString & "')")), 3)
                        Dim koll As String
                        kol = Round(kol, 3)
                        koll = Replace(CStr(kol), ",", ".") 'меняет все "," на "."
                    End If
                Else
                End If

                SumMinusNDS = CDbl(reader.Item("Стоимость")) / (1 + CDbl(GridViewVozvrat.GetFocusedDataRow.Item("НДС (%)")) / 100)

                dt.Rows.Add()
                dt.Rows(dt.Rows.Count() - 1).Item(0) = reader.Item("ware_code")
                dt.Rows(dt.Rows.Count() - 1).Item(1) = strSourceTypeName
                dt.Rows(dt.Rows.Count() - 1).Item(2) = reader.Item("Наименование")
                dt.Rows(dt.Rows.Count() - 1).Item(3) = Round(CDbl(reader.Item("кол")), 3)

                If CDbl(reader.Item("кол")) > 0 Then cPrice = Round(SumMinusNDS / CDbl(reader.Item("кол")), 2) Else cPrice = 0

                dt.Rows(dt.Rows.Count() - 1).Item(4) = cPrice

                If fNSD_inc Then
                    dt.Rows(dt.Rows.Count() - 1).Item(5) = Round(CDbl(reader.Item("Стоимость")), 2)
                    dt.Rows(dt.Rows.Count() - 1).Item(6) = reader.Item("ед")
                    dt.Rows(dt.Rows.Count() - 1).Item(7) = reader.Item("Короткое наим.")

                Else
                    dt.Rows(dt.Rows.Count() - 1).Item(5) = Round(SumMinusNDS, 2)
                    dt.Rows(dt.Rows.Count() - 1).Item(6) = Round(CDbl(reader.Item("Стоимость")) - SumMinusNDS, 2)
                    dt.Rows(dt.Rows.Count() - 1).Item(7) = Round(CDbl(reader.Item("Стоимость")), 2)
                    dt.Rows(dt.Rows.Count() - 1).Item(8) = reader.Item("ед")
                    dt.Rows(dt.Rows.Count() - 1).Item(9) = reader.Item("Короткое наим.")
                End If


            End While
            ConClose()

        Catch ex As Exception
            MsgBox(ex.Message, , "Ошибка")
            ConClose()
        End Try

    End Sub


    Private Sub ContxMnuStrPostavshiki_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContxMnuStrPostavshiki.ItemClicked
        Try
            If GridViewPostavshiki.GetFocusedDataRow Is Nothing Then Exit Sub
            Select Case ContxMnuStrPostavshiki.Items.IndexOf(e.ClickedItem)
                Case 0
                    ind = 0
                    frmNewDocument.LblDocument.Text = "Поставщик: " & GridViewPostavshiki.GetFocusedDataRow.Item("Наименование поставщика").ToString
                    frmNewDocument.ShowDialog()
                Case 1
                    'Запрашиваем НДС для возврата
                    ind = 3
                    frmNewDocument.ShowDialog()
                Case 2
                    отрисовка_поставщиков()
                    GridCtrlPostavshiki.DataSource = dt
                Case 4
                    Dim xlApp As Microsoft.Office.Interop.Excel.Application
                    Dim PathString As String = ""
                    xlApp = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
                    PathString = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString & "Поставщики " & Now.ToString("dd.MM.yyyy H.mm.ss") & ".xlsx"
                    GridViewPostavshiki.ExportToXlsx(PathString)
                    xlApp.Workbooks.Open(PathString)
                    xlApp.Cells.Font.Name = "Times New Roman"
                    xlApp.Cells.Font.Size = 12
                    xlApp.Cells.Columns.AutoFit()
                    xlApp.Visible = True
            End Select
        Catch ex As Exception
            MsgBox(ex.Message, , "Ошибка")
        End Try
    End Sub

    Private Sub ContxMnuStrRashod_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContxMnuStrRashod.ItemClicked
        Try
            If GridViewRashod.GetFocusedDataRow Is Nothing Then Exit Sub
            Select Case ContxMnuStrRashod.Items.IndexOf(e.ClickedItem)
                Case 1
                    If ContxMnuStrRashod.Items.Item(0).Enabled = False Then Exit Sub
                    Dim path As String = pathMyDocs & "\frM113.fr3"
                    IO.File.WriteAllBytes(path, My.Resources.frM113)
                    M113_ToFR(FRX, GridViewRashod.GetFocusedDataRow.Item("doc_code").ToString, 1, , gWHS)
                Case 3
                    ContxMnuStrRashod.Close()
                    SplashScreenManager.ShowForm(Me, GetType(SplashScreen), True, True, False)
                    отрисовка_расходов()
                    GridCtrlRashod.DataSource = dt
                    GridViewRashod.Columns.Item("Код смены").Visible = False
                    GridViewRashod.Columns.Item("doc_code").Visible = False
                    SplashScreenManager.CloseForm()
                Case 5
                    Dim xlApp As Microsoft.Office.Interop.Excel.Application
                    Dim PathString As String = ""
                    xlApp = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
                    PathString = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString & "Расход " & Now.ToString("dd.MM.yyyy H.mm.ss") & ".xlsx"
                    GridViewRashod.ExportToXlsx(PathString)
                    xlApp.Workbooks.Open(PathString)
                    xlApp.Cells.Font.Name = "Times New Roman"
                    xlApp.Cells.Font.Size = 12
                    xlApp.Cells.Columns.AutoFit()
                    xlApp.Visible = True
            End Select
        Catch ex As Exception
            MsgBox(ex.Message, , "Ошибка")
        End Try
    End Sub

    Private Sub ContxMnuStrRashod_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ContxMnuStrRashod.Opening
        'ПУНКТ МЕНЮ "ПЕРЕДАЧА РЕСУРСОВ НА ДРУГОЙ СКЛАД" - ПОДГРУЗКА СКЛАДОВ
        Try
            Res_Send()
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

    'ПЕРЕДАЧА РЕСУРСОВ НА ДРУГОЙ СКЛАД" - ПОДГРУЗКА СКЛАДОВ
    Sub Res_Send()
        Try
            Dim i As Integer
            i = 0
            'ОЧИСТКА
            ToolStripComboBox1.Items.Clear()
            ToolStripComboBox1.Tag = ""

            SQLEx("SELECT StructureID, StructureName FROM Structure WHERE StructureType=5 AND DTDel IS NULL AND StructureID <> " & gWHS)

            While reader.Read
                ToolStripComboBox1.Items.Add(reader.Item("StructureName"))
                i = i + 1
                ToolStripComboBox1.Tag = ToolStripComboBox1.Tag & i.ToString & "=" & reader.Item("structureID").ToString & vbCr
            End While

            ConClose()
        Catch ex As Exception
            ConClose()
            MsgBox(ex.Message, , "Ошибка")
        End Try
    End Sub

	
    Private Sub ToolStripComboBox1_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStripComboBox1.ItemClicked
        Try
            ToolStripComboBox1.Close()
            If MsgBox("ВНИМАНИЕ - Подтвердите создание новой накладной на внутреннее перемещение ресурсов между складами?" & vbCr & vbCr & "Отправитель:" & vbTab & gWHS_NAME & vbCr & "Получатель:" _
                      & vbTab & e.ClickedItem.ToString, vbQuestion + vbYesNo, "Запрос подтверждения") = vbNo Then Exit Sub

            Dim skladid As Integer = CInt(CN("SELECT StructureID FROM Structure WHERE StructureType=5 AND DTDel IS NULL AND StructureName = '" & e.ClickedItem.Text & "'"))
            Sklad_WareSend_NewDoc(skladid)
        Catch ex As Exception
            MsgBox(ex.Message, , "Ошибка")
        End Try
    End Sub

	
    Private Sub GridCtrlPrihodClick_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles GridCtrlPrihodClick.DragDrop
        Try
            перетащить()
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub GridCtrlPrihodClick_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles GridCtrlPrihodClick.DragOver
        Try
            If e.Data.GetDataPresent(GetType(DataRow)) Then
                e.Effect = DragDropEffects.Move
            Else
                e.Effect = DragDropEffects.None
            End If
        Catch ex As Exception
            MsgBox(ex.Message, , "Ошибка при перемещении исполнителя над операциями.")
        End Try
    End Sub

	
    Private Sub txtDelivery_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDelivery.KeyPress
        Try
            ' Вводить можно только цифры и символы управления
            If Char.IsNumber(e.KeyChar) Or Char.IsControl(e.KeyChar) Then
            Else
                e.KeyChar = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, , "Ошибка")
        End Try
    End Sub

	
    Private Sub txtDelivery_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDelivery.TextChanged
        Try
            Dim i As Double
            If Double.TryParse(txtDelivery.Text, i) = True And i >= 0 Then
                cmdDelivery.Enabled = True
            Else
                cmdDelivery.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, , "Ошибка")
        End Try
    End Sub

	
    'Удаление строки документа 
    Private Sub mnuInDel()
        Try
            SQLExNQ("DELETE FROM provodka WHERE provodka.provodka_code=" & GridViewPrihodClick.GetFocusedDataRow.Item("provodka_code").ToString)
            Delivery(GridViewPrihod.GetFocusedDataRow.Item("doc_code").ToString)
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub GridViewPrihodClick_RowClick(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.RowClickEventArgs) Handles GridViewPrihodClick.RowClick
        If e.Button = Windows.Forms.MouseButtons.Right Then
            ContxMnuStrPrihodClick.Show(MousePosition)
        End If
    End Sub

	
    Private Sub GridViewPoZaiyvkeClick_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridViewPoZaiyvkeClick.Click
        rowPoZaiyvkeClickClick = 2
    End Sub

	
    Private Sub GridViewPoZaiyvkeClick_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridViewPoZaiyvkeClick.DoubleClick
        Try
            If rowPoZaiyvkeClickClick = 1 Then
                If GridViewPoZaiyvkeClick.GetFocusedDataRow Is Nothing Or GridViewPoZaiyvkeClick.RowCount = 0 Then Exit Sub
                Dim ans As Object = ""
                Dim ans2 As Object = ""
                Dim var_num As Double
                Dim lTMP As Long : lTMP = GridViewPoZaiyvkeClick.RowCount
                Dim sTMP As String : sTMP = "Тип: " & IIF_S(GridViewPoZaiyvkeClick.GetFocusedDataRow.Item("Тип").ToString = "М", "Мат", "Станд") & vbCr & "Наименование: " & GridViewPoZaiyvkeClick.GetFocusedDataRow.Item("Название").ToString & vbCr & "Ед: " & GridViewPoZaiyvkeClick.GetFocusedDataRow.Item("Ед. изм.").ToString & vbCr & vbCr
                Dim MyPrice As Double
                Dim MySum As Double
                Dim TypeWare As Integer
                Dim cWeight As Double
                Dim flagOrder As Integer  ' 1 - если в заявке кг и 0 - если учетные единицы
                Dim flagPrihod As Integer ' 1 - если пришло в др. единицах и 0 если пришло в тех единицах в которых заказывали

                If (GridViewPoZaiyvkeClick.GetFocusedDataRow.Item("Тип").ToString = "Ст") Then
                    TypeWare = 2
                Else
                    TypeWare = 1 '"Мат"
                End If

                If Len(GridViewPoZaiyvke.GetFocusedDataRow.Item(5).ToString) = 0 Then 'Запрашиваем НДС
                    ind = 2
                    frmNewDocument.ShowDialog()
                    If ind = 0 Then Exit Sub
                End If

                'Спрашиваем (кол - во)
                If Len(mInDocOrder_SavedKey) > 0 Then
                    '  ans = GETP(mInDocOrder_SavedKey, "Num")
                    ' ans2 = GETP(mInDocOrder_SavedKey, "Price")
                Else
                    cWeight = CDbl(CN("SELECT TOP 1 ISNULL(Weight, 0) As Weight FROM Wares WHERE WareName = '" & GridViewPoZaiyvkeClick.GetFocusedDataRow.Item("Название").ToString & "' "))
                    '1 - если в заявке кг и 0 - если учетные единицы
                    If CDbl(CN("SELECT TOP 1 ISNULL(OKEI_Code, 0) As OKEI_Code FROM OKEI WHERE OKEI_Name = '" & GridViewPoZaiyvkeClick.GetFocusedDataRow.Item("Ед. Изм.").ToString & "' ")) = 2 Then
                        flagOrder = 1
                    Else
                        flagOrder = 0
                    End If
                    ans = frmKolWare.Show2(TypeWare, GridViewPoZaiyvkeClick.GetFocusedDataRow.Item("Название").ToString, GridViewPoZaiyvkeClick.GetFocusedDataRow.Item("Кол-во").ToString, GridViewPoZaiyvkeClick.GetFocusedDataRow.Item("Ед. изм.").ToString, cWeight, flagOrder)

                    '1 - если пришло в др.ед и 0 - если пришло в тех ед, в которых заказывали
                    flagPrihod = frmKolWare.MyShow(TypeWare, GridViewPoZaiyvkeClick.GetFocusedDataRow.Item("Название").ToString, GridViewPoZaiyvkeClick.GetFocusedDataRow.Item("Кол-во").ToString, GridViewPoZaiyvkeClick.GetFocusedDataRow.Item("Ед. изм.").ToString, cWeight)

                    If TypeOf (ans) Is DBNull Or CStr(ans) = "" Then Exit Sub
                    If CDbl(ans) < 0 Then Exit Sub

                    ' изменен алгоритм работы frmKolWare, поэтому он вернет в тех единицах, которые выбра пользователь. Чтоб не исправлять всюду ниже,
                    ' изменяем значение поля ans
                    If (flagOrder = 0) And (flagPrihod = 1) Then 'значит закупка была в уч.единицах  и приход в кг (поменялся флаг). ans - кг, надо вернуть к учетной
                        ans = CDbl(ans) / cWeight
                    Else
                        If (flagOrder = 1) And (flagPrihod = 1) Then
                            ans = CDbl(ans) * cWeight 'значит закупка была в кг и приход в учетных единицах. ans - уч, надо вернуть к кг
                        End If
                    End If

                    MyPrice = CDbl(NeS(GridViewPoZaiyvkeClick.GetFocusedDataRow.Item("Цена").ToString))
                    MySum = Round(CDbl(ans) * MyPrice, 2) 'стоимость считается как старая единица на цену, потому что стоимость не меняется, меняется только цена

                    'ans2 - выводит новую сумму
                    ans2 = CheckNum(InputBox(sTMP & "Стоимость: ", GridViewPoZaiyvkeClick.GetFocusedDataRow.Item("Название").ToString, MySum.ToString))
                    If TypeOf (ans2) Is DBNull Then Exit Sub
                    If CStr(ans2) = "" Then Exit Sub
                    If CDbl(ans2) < 0 Then Exit Sub
                End If

                If CDbl(ans2) = 0 Then
                    MsgBox("ВНИМАНИЕ - Приходование ресурсов по нулевым ценам запрещено.", vbInformation, "Неверный ввод")
                    Exit Sub
                End If

                var_num = CDbl(ans)
                If CDbl(var_num) >= 0 Then
                    If ((flagOrder = 0 And flagPrihod = 0) Or (flagOrder = 1 And flagPrihod = 0)) Then 'случаи 1) и 2)
                        GridViewPoZaiyvkeClick.SetFocusedRowCellValue("Кол-во", Round(CDbl(var_num), 3))
                    Else
                        If (flagOrder = 0 And flagPrihod = 1) Then 'случай 3)
                            GridViewPoZaiyvkeClick.SetFocusedRowCellValue("Кол-во", Round(CDbl(var_num * cWeight), 3))
                        Else 'flagOrder = 1 flagPrihod = 1  случай 4)
                            GridViewPoZaiyvkeClick.SetFocusedRowCellValue("Кол-во", Round(CDbl(var_num / cWeight), 3))
                        End If
                    End If

                End If

                'Пишем "Стоимость"
                GridViewPoZaiyvkeClick.SetFocusedRowCellValue("Стоимость", dec_sep(ans2.ToString))

                GridViewPoZaiyvkeClick.Tag = SETP(GridViewPoZaiyvkeClick.Tag.ToString, "[" & GridViewPoZaiyvkeClick.GetFocusedDataSourceRowIndex & "_Cost]", dec_sep(ans2.ToString))
                GridViewPoZaiyvkeClick.Tag = SETP(GridViewPoZaiyvkeClick.Tag.ToString, "[" & GridViewPoZaiyvkeClick.GetFocusedDataSourceRowIndex & "_InputCost]", dec_sep(ans2.ToString))

                'если требуется меняем расчетные единицы
                If (flagOrder = 0 And flagPrihod = 1) Then 'случай 3)
                    GridViewPoZaiyvkeClick.SetFocusedRowCellValue("Ед. изм.", "кг.")
                Else
                    If (flagOrder = 1 And flagPrihod = 1) Then 'случай 4)
                        GridViewPoZaiyvkeClick.SetFocusedRowCellValue("Ед. изм.", (CN("SELECT TOP 1 ISNULL(OKEI_Name, 0) As OKEI_Name FROM OKEI, Wares WHERE WareName = '" & GridViewPoZaiyvkeClick.GetFocusedDataRow.Item(1).ToString & " '  and (Wares.OKEI_Code = OKEI.OKEI_code)").ToString))
                    End If
                End If

                'Считаем цену с учетом доставки
                cmdDelivery2.PerformClick()
                WriteSum2()

            End If
            rowPoZaiyvkeClickClick = 0
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub


    Private Sub GridViewPoZaiyvkeClick_RowClick(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.RowClickEventArgs) Handles GridViewPoZaiyvkeClick.RowClick
        Try
            rowPoZaiyvkeClickClick = 1
            If e.Button = Windows.Forms.MouseButtons.Right Then
                If GridViewPoZaiyvkeClick.GetFocusedDataRow.Item("Аналоги").ToString = "+" Then
                    ContxMnuStrPoZaiyvkeClick.Items.Item(0).Enabled = True
                    ContxMnuStrPoZaiyvkeClick.Show(MousePosition)
                Else
                    ContxMnuStrPoZaiyvkeClick.Items.Item(0).Enabled = False
                    ContxMnuStrPoZaiyvkeClick.Show(MousePosition)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'клик по строке остатков
    Private Sub GridViewOstatki_RowClick(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.RowClickEventArgs) Handles GridViewOstatki.RowClick
        rowOstatkiClick = 1 'один если был клик по строке
        If GridViewOstatki.GetFocusedDataRow Is Nothing Then ContxMnuStrOstatki.Visible = False : Exit Sub
        If e.Button = System.Windows.Forms.MouseButtons.Right And opt = 1 Then
            ContxMnuStrOstatki.Items.Item(1).Visible = True
            ContxMnuStrOstatki.Items.Item(2).Visible = True
            ContxMnuStrOstatki.Items.Item(3).Visible = True
            If GridViewOstatki.GetFocusedDataRow.Item("Исх. треб-я").ToString = "x" And GridViewOstatki.GetFocusedDataRow.Item("Приход по заявке").ToString = "x" Then
                ContxMnuStrOstatki.Items.Item(1).Enabled = True
                ContxMnuStrOstatki.Items.Item(2).Enabled = True
            ElseIf GridViewOstatki.GetFocusedDataRow.Item("Исх. треб-я").ToString = "x" Then
                ContxMnuStrOstatki.Items.Item(1).Enabled = True
                ContxMnuStrOstatki.Items.Item(2).Enabled = False
            ElseIf GridViewOstatki.GetFocusedDataRow.Item("Приход по заявке").ToString = "x" Then
                ContxMnuStrOstatki.Items.Item(1).Enabled = False
                ContxMnuStrOstatki.Items.Item(2).Enabled = True
            Else
                ContxMnuStrOstatki.Items.Item(1).Enabled = False
                ContxMnuStrOstatki.Items.Item(2).Enabled = False
            End If
            ContxMnuStrOstatki.Show(MousePosition)
        ElseIf e.Button = System.Windows.Forms.MouseButtons.Right And (opt = 2 Or opt = 3) Then
            ContxMnuStrOstatki.Items.Item(1).Visible = False
            ContxMnuStrOstatki.Items.Item(2).Visible = False
            ContxMnuStrOstatki.Items.Item(3).Visible = False
            ContxMnuStrOstatki.Show(MousePosition)
        End If
    End Sub

    Public Sub GridViewRashod_RowClick(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.RowClickEventArgs) Handles GridViewRashod.RowClick
        Try
            Dim Str As String
            Dim Int As Integer
            Dim MousePos As Object = MousePosition
            Str = ""

            frmList.e2 = e
            frmList.sender2 = sender

            If GridViewRashod.GetFocusedDataRow Is Nothing Or BtnReportDay.Text = "Закрыть" Then Exit Sub
            Str = CN("SELECT ISNULL(REPLACE ((SELECT CAST(ROW_NUMBER() OVER (ORDER BY W.WareType, W.WareName) AS VARCHAR) + ') ' + CASE " &
             "                          WHEN W.WareType = 1 THEN 'Мат. - ' " &
             "                          WHEN W.WareType = 2 THEN 'Ст. - ' " &
             "                        END + W.WareName + '§' " &
             "                 FROM   WareBalance WB " &
             "                        LEFT JOIN CellWare CW ON CW.warecode = WB.warecode " &
             "                                                 AND CW.smenaNo = WB.smena_no " &
             "                        JOIN Provodka P ON P.ware_code = WB.WareCode " &
             "                        JOIN Wares W ON W.WareCode = WB.WareCode " &
             "                 WHERE  WB.smena_no = " & gWHS & " " &
             "                    AND P.doc_code = " & GridViewRashod.GetFocusedDataRow.Item("doc_code").ToString &
             "                 GROUP  BY WB.WareCode, " &
             "                           W.WareType, " &
             "                           W.WareName, " &
             "                           WB.smena_no, " &
             "                           WB.Num " &
             "                 HAVING ROUND(WB.Num, 3) <> ROUND(SUM(CW.Num), 3) " &
             "                         OR ROUND(SUM(CW.Num), 3) IS NULL ORDER BY W.WareType, W.WareName " &
             "                 FOR XML PATH('')), '§', CHAR(13)), '') ").ToString

            If Len(Str) > 0 Then MsgBox("ВНИМАНИЕ - Данным ресурсам не присвоена ячейка хранения на складе или распределено не все имеющееся количество: " & vbCr & vbCr & Str, vbInformation, "Информация")

            'ПРОВЕРКА\УСТАНОВКА ФЛАГА ЮЗЕРА НА ДОКУМЕНТЕ
            If UserFlags_SET("doc_code", CInt(GridViewRashod.GetFocusedDataRow.Item("doc_code")), False) Then GridCtrlRashodClick.Visible = False : Exit Sub

            Int = CInt(CN("SELECT COUNT(*) FROM Documents WHERE doc_code=" & GridViewRashod.GetFocusedDataRow.Item("doc_code").ToString & " AND (CASE WHEN doc_type<>22 AND data_reg IS NOT NULL THEN 1 WHEN doc_type=22 AND data_make IS NOT NULL THEN 1 ELSE 0 END) = 1"))

            If Int > 0 Then
                MsgBox("Внимание - Документ уже проведен\отправлен (возможно другим пользователем с другого компьютера)", vbExclamation, "Ошибка открытия документа")
            End If

            пересчет()
            GridCtrlRashodClick.DataSource = dt
            GridViewRashodClick.Columns.Item("Warecode").Visible = False
            GridViewRashodClick.Columns.Item("provodka_code").Visible = False
            GridCtrlRashodClick.Visible = True
            BigButton.Enabled = True
            If GridViewRashod.GetFocusedDataRow.Item("doc_type").ToString = "22" Then
                BigButton.Text = "Отправить"
            Else
                BigButton.Text = "Провести расходный документ"
            End If
            SmallButton.Enabled = True

            If e.Button = Windows.Forms.MouseButtons.Right Then
                ContxMnuStrRashod.Show(MousePos)
            End If

        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub GridViewPostavshiki_RowClick(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.RowClickEventArgs) Handles GridViewPostavshiki.RowClick
        Try
            If GridViewPostavshiki.GetFocusedDataRow Is Nothing Then
                ContxMnuStrPostavshiki.Visible = False
                Exit Sub
            ElseIf e.Button = System.Windows.Forms.MouseButtons.Right Then
                ContxMnuStrPostavshiki.Visible = True
                ContxMnuStrPostavshiki.Show(MousePosition)
            End If
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub GridViewNakl_RowClick(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.RowClickEventArgs) Handles GridViewNakl.RowClick
        Try
            Dim Sql As String
            Dim dtr As New DataTable

            If GridViewNakl.GetFocusedDataRow Is Nothing Then Exit Sub

            GridCtrlNaklClick.Visible = True
            GridViewNaklClick.Columns.Clear()
            dtr.Clear()
            Sql = ""
            Sql = Sql & " SELECT  "
            Sql = Sql & "      CASE P.ware_type "
            Sql = Sql & "               WHEN 5 THEN E.ElementName + ' (' + E.ElementCode + ')'"
            Sql = Sql & "               Else W.WareName"
            Sql = Sql & "             END                   AS Наименование,"
            Sql = Sql & "              CASE P.ware_type "
            Sql = Sql & "          WHEN 5 THEN 'шт'"
            Sql = Sql & "          Else O.Okei_name"
            Sql = Sql & "        END                        AS Ед, "
            Sql = Sql & "       P.NumFact as количество, "
            Sql = Sql & "       ISNULL(W.WareShortName, '') AS [Короткое наим.]"
            Sql = Sql & " FROM   Provodka P "
            Sql = Sql & "       LEFT JOIN Wares W ON P.ware_code = W.WareCode "
            Sql = Sql & "       LEFT JOIN OKEI O ON O.OKEI_Code = "
            Sql = Sql & "                                     Case "
            Sql = Sql & "                                        WHEN p.OKEI_Code Is Null"
            Sql = Sql & "                                        THEN W.OKEI_Code"
            Sql = Sql & "                                        Else p.OKEI_Code"
            Sql = Sql & "                                     END   "
            Sql = Sql & "        LEFT JOIN NaklContract NC ON NC.fact_no = P.ware_code AND P.ware_type = 5"
            Sql = Sql & "        LEFT JOIN ELEMENTS E ON E.ElementID = NC.ElementID"
            Sql = Sql & " WHERE  P.doc_code = " & GridViewNakl.GetFocusedDataRow.Item("doc_code").ToString
            Sql = Sql & " ORDER  BY W.WareType, "
            Sql = Sql & "        W.WareName "

            SQLEx(Sql)
            dtr.Load(reader)
            GridCtrlNaklClick.DataSource = dtr
            ConClose()
            BigButton.Enabled = True

            If GridViewNakl.GetFocusedDataRow Is Nothing Then
                ContxMnuStrNakl.Visible = False
                Exit Sub
            ElseIf e.Button = System.Windows.Forms.MouseButtons.Right Then
                ContxMnuStrNakl.Visible = True
                If GridViewNakl.GetFocusedDataRow.Item("Документ").ToString = "Расходная накл." Or GridViewNakl.GetFocusedDataRow.Item("Документ").ToString = "Возврат со смены" Or GridViewNakl.GetFocusedDataRow.Item("Документ").ToString = "Возврат поставщику" Then
                    ContxMnuStrNakl.Items.Item(0).Visible = False
                    ContxMnuStrNakl.Items.Item(1).Visible = False
                    ContxMnuStrNakl.Items.Item(3).Visible = False
                    ContxMnuStrNakl.Items.Item(4).Visible = False
                ElseIf GridViewNakl.GetFocusedDataRow.Item("Документ").ToString = "Приходная накл." Then
                    ContxMnuStrNakl.Items.Item(0).Visible = True
                    ContxMnuStrNakl.Items.Item(1).Visible = True
                    ContxMnuStrNakl.Items.Item(3).Visible = True
                    ContxMnuStrNakl.Items.Item(4).Visible = True
                End If
            End If

            If e.Button = Windows.Forms.MouseButtons.Right Then
                ContxMnuStrNakl.Show(MousePosition)
            End If

        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub GridViewNaklClick_RowClick(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.RowClickEventArgs) Handles GridViewNaklClick.RowClick
        Try
            If GridViewNaklClick.GetFocusedDataRow Is Nothing Then
                ContxMnuStrNaklClick.Visible = False
                Exit Sub
            ElseIf e.Button = System.Windows.Forms.MouseButtons.Right Then
                ContxMnuStrNaklClick.Visible = True
                ContxMnuStrNaklClick.Show(MousePosition)
            End If
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub ContxMnuStrNaklClick_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContxMnuStrNaklClick.ItemClicked
        Try
            If GridViewNaklClick.GetFocusedDataRow Is Nothing Then Exit Sub
            Select Case ContxMnuStrNaklClick.Items.IndexOf(e.ClickedItem)
                Case 0
                    Dim xlApp As Microsoft.Office.Interop.Excel.Application
                    Dim PathString As String = ""
                    xlApp = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
                    PathString = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString & "Накладная " & Now.ToString("dd.MM.yyyy H.mm.ss") & ".xlsx"
                    GridViewNaklClick.ExportToXlsx(PathString)
                    xlApp.Workbooks.Open(PathString)
                    xlApp.Cells.Font.Name = "Times New Roman"
                    xlApp.Cells.Font.Size = 12
                    xlApp.Cells.Columns.AutoFit()
                    xlApp.Visible = True
            End Select
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    'Проверка на потерянные копейки
    Public Sub CheckLostKopeck(ByVal ind As Integer)

        Dim i As Long
        Dim CountSum As Double
        Dim RealSum As Double
        Dim NewSum As Double

        Select Case ind

            Case 1 'Вкладка "Приход"

                    'Реальная цена

                If GridViewPrihod.RowCount = 0 Then
                    'Закрываемся
                    Call showSumAndDelivery(False, 1)
                    Exit Sub
                End If

                    Call showSumAndDelivery(True, 1)

                SQLEx("SELECT Sum(Price) " & _
                         "FROM provodka " & _
                         "WHERE doc_code = " & GridViewPrihod.GetFocusedDataRow.Item("doc_code").ToString)

                If reader.HasRows = True Then
                    While reader.Read
                        If reader.Item(0) Is DBNull.Value Then Exit Sub
                    End While

                    RealSum = Round(CDbl(reader.Item(0)), 2) + Round(CDbl(Trim$(txtDelivery.Text)), 2)
                End If

                'Расчетная цена
                SQLEx("SELECT Sum(provodka.Delivery) " & _
                         "FROM provodka " & _
                         "WHERE (((provodka.doc_code)= " & GridViewPrihod.GetFocusedDataRow.Item("doc_code").ToString)

                If reader.HasRows = False Then
                    While reader.Read
                        If reader.Item(0) Is DBNull.Value Then Exit Sub
                    End While

                    CountSum = Round(CDbl(reader.Item(0)), 2)
                End If

                'Проверка разницы
                If Abs(Round(RealSum - CountSum, 2)) > 0.5 Or Abs(RealSum - CountSum) = 0 Then Exit Sub

                'Прибавление потерянных копеек к последнему элементу таблицы
                For i = GridViewPrihodClick.RowCount - 1 To 0 Step -1
                    If GridViewPrihodClick.GetDataRow(i).Item("Стоимость") <> "" Then
                        NewSum = CDbl(CDbl(GridViewPrihodClick.GetDataRow(i).Item("Стоимость")) + CDbl(RealSum - CountSum))
                        GridViewPrihodClick.GetDataRow(i).Item("Стоимость") = Replace$(NewSum, ",", ".")
                        Exit Sub
                    End If
                Next i

            Case 2 'Вкладка "Приход по заявке"

                'Расчетная цена
                For i = 0 To GridViewPoZaiyvkeClick.RowCount - 1
                    If GridViewPoZaiyvkeClick.GetDataRow(i).Item("+Доставка").ToString <> "" Then
                        CountSum = CountSum + CDbl(GridViewPoZaiyvkeClick.GetDataRow(i).Item("+Доставка").ToString)
                    End If
                Next i

                'Реальная цена
                For i = 0 To GridViewPoZaiyvkeClick.RowCount - 1
                    If GridViewPoZaiyvkeClick.GetDataRow(i).Item("Стоимость").ToString <> "" Then
                        RealSum = RealSum + CDbl(GridViewPoZaiyvkeClick.GetDataRow(i).Item("Стоимость").ToString)
                    End If
                Next i
                RealSum = RealSum + CDbl(Replace$(Trim$(txtDelivery2.Text), ".", ","))
                CountSum = Round(CountSum, 2)


                'Проверка разницы
                If Abs(Round(RealSum - CountSum, 2)) > 0.5 Or Abs(RealSum - CountSum) = 0 Then Exit Sub

                'Прибавление потерянных копеек к последнему элементу таблицы
                For i = GridViewPoZaiyvkeClick.RowCount - 1 To 0 Step -1
                    If GridViewPoZaiyvkeClick.GetDataRow(i).Item("+Доставка").ToString <> "" Then
                        GridViewPrihodClick.GetDataRow(i).Item("Стоимость") = Round(CDbl(CDbl(GridViewPrihodClick.GetDataRow(i).Item("Стоимость")) + RealSum - CountSum), 2)
                        Exit Sub
                    End If
                Next i

        End Select
    End Sub

	
    Private Sub GridViewRashodClick_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridViewRashodClick.Click
        rowRashodClickClick = 2
    End Sub

	
    'РЕДАКТИРОВАНИЕ СОДЕРЖАНИЯ РАСХОДНОГО ДОКУМЕНТА
    Private Sub GridViewRashodClick_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridViewRashodClick.DoubleClick
        Try
            If rowRashodClickClick = 1 Then
                Dim yy As Double, V_provodka_code As Integer

                If GridViewRashodClick.RowCount = 0 Then Exit Sub
                If GridViewRashod.GetFocusedDataRow.Item("doc_type").ToString = "22" Then Exit Sub

                Dim kol As String
                If CDbl(GridViewRashodClick.GetFocusedDataRow.Item(3)) > CDbl(GridViewRashodClick.GetFocusedDataRow.Item(6)) Then
                    kol = GridViewRashodClick.GetFocusedDataRow.Item(6).ToString
                Else
                    kol = GridViewRashodClick.GetFocusedDataRow.Item(3).ToString
                End If

                Dim dd As Object
                Dim Wight As Double
                Dim sType As Integer

                If GridViewRashodClick.GetFocusedDataRow.Item(0).ToString = "Мат." Then sType = 1 Else sType = 2
                Wight = CDbl(CN("SELECT TOP 1 ISNULL(Weight, 0) As Weight FROM Wares WHERE WareName = '" & GridViewRashodClick.GetFocusedDataRow.Item(1).ToString & " ' "))
                dd = frmKolWare.Show2(sType, _
                                      GridViewRashodClick.GetFocusedDataRow.Item(1).ToString, _
                                      kol, _
                                      GridViewRashodClick.GetFocusedDataRow.Item(2).ToString, _
                                      Wight)

                If dd Is Nothing Then Exit Sub
                If dd.ToString = "" Then Exit Sub
                If CDbl(dd) < 0 Then Exit Sub

                If (1 = frmKolWare.MyShow(sType, GridViewRashodClick.GetFocusedDataRow.Item(1).ToString, kol, GridViewRashodClick.GetFocusedDataRow.Item(2).ToString, Wight)) Then
                    'если 1=1 значит менялась единица измерения. dd хранит кг!
                    GridViewRashodClick.GetFocusedDataRow.Item(5) = dd 'необходимо отобразить в колокне к кг, кг
                    dd = Round(CDbl(dd) / Wight, 3) 'необходимо перевести в уч = кг / коэф
                End If

                If dd.ToString <> "" Then
                    yy = CDbl(dec_sep(dd.ToString))
                    If yy < 0 Then Exit Sub

                    If CDbl(GridViewRashodClick.GetFocusedDataRow.Item(6)) < yy Then
                        MsgBox("ВНИМАНИЕ" & vbCr & "На складе нет - " & yy & " " & GridViewRashodClick.GetFocusedDataRow.Item(2).ToString, vbInformation, GridViewRashodClick.GetFocusedDataRow.Item(1).ToString)
                        Exit Sub
                    End If

                    V_provodka_code = CInt(GridViewRashodClick.GetFocusedDataRow.Item("provodka_code"))
                    Out_fix(V_provodka_code, yy)

                    'На след строку
                    If GridViewRashodClick.RowCount > GridViewRashodClick.GetFocusedDataSourceRowIndex Then GridViewRashodClick.SelectRow(GridViewRashodClick.GetFocusedDataSourceRowIndex + 1)
                End If
            End If
            rowRashodClickClick = 0
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub GridViewRashodClick_RowClick(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.RowClickEventArgs) Handles GridViewRashodClick.RowClick
        rowRashodClickClick = 1
    End Sub

	
    Private Sub ContxMnuStrPeriod_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContxMnuStrPeriod.ItemClicked
        Try
            If GridViewPeriod.GetFocusedDataRow Is Nothing Then Exit Sub
            Select Case ContxMnuStrPrihod.Items.IndexOf(e.ClickedItem)
                Case 0  'экспорт в excel
                    Dim xlApp As Microsoft.Office.Interop.Excel.Application
                    Dim PathString As String = ""
                    xlApp = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
                    PathString = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString & "Оборотная ведомость " & Now.ToString("dd.MM.yyyy H.mm.ss") & ".xlsx"
                    GridViewPeriod.ExportToXlsx(PathString)
                    xlApp.Workbooks.Open(PathString)
                    xlApp.Cells.Font.Name = "Times New Roman"
                    xlApp.Cells.Font.Size = 12
                    xlApp.Cells.Columns.AutoFit()
                    xlApp.Visible = True
            End Select
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub GridViewPeriod_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridViewPeriod.Click
        rowPeriodClick = 2
    End Sub

	
    Private Sub GridViewPeriod_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridViewPeriod.DoubleClick

        If rowPeriodClick = 1 Then
            If GridViewPeriod.GetFocusedDataRow Is Nothing Then Exit Sub
            ДвижениеЭлемента(GridViewPeriod, 2)
            frmInOut.ShowDialog()
        End If
        rowPeriodClick = 0

    End Sub

	
    Public Sub ДвижениеЭлемента(ByVal GV As DevExpress.XtraGrid.Views.Grid.GridView, Optional ByRef t As Integer = 0)
        Try

            Dim Sql As String
            Dim dt As New DataTable
            Dim dat As String = ""

            'если показать движение из клика по остаткам, то запрос без даты (t=1), если из отчетов то с датой (t=2)
            If t = 2 Then dat = " AND Documents.data_reg BETWEEN '" & DateTimePicker1.Value.ToString & "' AND '" & DateTimePicker2.Value.ToString & "' "

            dt.Columns.Add("Тип", GetType(String))
            dt.Columns.Add("Дата создания", GetType(Date))
            dt.Columns.Add("Пл. док/ ЛЗК", GetType(String))
            dt.Columns.Add("Документ", GetType(String))
            dt.Columns.Add("Приходный ордер", GetType(String))
            dt.Columns.Add("Кол. факт.", GetType(String))
            dt.Columns.Add("Кол. треб.", GetType(String))
            dt.Columns.Add("Контрагент", GetType(String))
            dt.Columns.Add("Заказы", GetType(String))
            dt.Columns.Add("doc_code", GetType(String))
            dt.Columns.Add("doc_type", GetType(String))

            Sql = ""
            Sql = Sql & " SELECT DISTINCT "
            Sql = Sql & "         provodka.ware_code                                                 ,"
            Sql = Sql & "         ISNULL(CONVERT(VARCHAR, Documents.data_make, 104), '') AS data_make,"
            Sql = Sql & "         documents.doc_code                                                 ,"
            Sql = Sql & "         CASE WHEN Documents.doc_type IN(1, 101 ) THEN Documents.DocPay ELSE A.ParentDocText  END as Cool,"
            Sql = Sql & "         Documents.doc_type                                                 ,"
            Sql = Sql & "         CAST(A.text AS VARCHAR)                        AS doc_text             ,"
            Sql = Sql & "         provodka.ware_type                                                 ,"
            Sql = Sql & "         clients.cli_name                                                   ,"
            Sql = Sql & "         provodka.NumFact                                                   ,"
            Sql = Sql & "         provodka.Num,"
            Sql = Sql & "         CASE WHEN Documents.doc_type IN (3, 101, 22) THEN  S.StructureName ELSE clients.cli_name END AS Partner,"
            Sql = Sql & "         ISNULL(( CASE Documents.doc_type WHEN 22 THEN ( CASE "
            Sql = Sql & "                                       WHEN Documents.MasterCode = " & gWHS & " THEN 'Расход'"
            Sql = Sql & "                                       WHEN Documents.cli_code = " & gWHS & " THEN 'Приход'"
            Sql = Sql & "                                     END )"
            Sql = Sql & "                                 END ), '')               AS Document_Type,"
            Sql = Sql & "         ISNULL(STUFF((SELECT DISTINCT ', ' + C.doc_name"
            Sql = Sql & "                     FROM   Documents D"
            Sql = Sql & "                            JOIN DocumentDocument DD ON DD.ChildID = D.doc_code"
            Sql = Sql & "                            JOIN Requirement R ON R.doc_code = DD.ParentID"
            Sql = Sql & "                            JOIN NaklContract NC ON NC.nakl_no = R.nakl_no"
            Sql = Sql & "                            JOIN COntracts C ON C.COntractID = NC.COntractID"
            Sql = Sql & "                     Where D.doc_code = Documents.doc_code"
            Sql = Sql & "                     FOR XML PATH('')), 1, 2, ''), '')        CList,"
            Sql = Sql & "         Documents.PayInSlip, "
            Sql = Sql & "         Documents.data_make "
            Sql = Sql & " FROM (provodka"
            Sql = Sql & "         LEFT JOIN Documents ON provodka.doc_code  = Documents.doc_code)"
            Sql = Sql & "         LEFT JOIN clients   ON Documents.cli_code = clients.cli_code"
            Sql = Sql & "         LEFT JOIN Structure S ON S.StructureID = (CASE WHEN Documents.doc_type IN (22) THEN "
            Sql = Sql & "                                                           (CASE WHEN Documents.MasterCode = " & gWHS & " THEN clients.cli_code "
            Sql = Sql & "                                                                                                 ELSE Documents.MasterCode END) "
            Sql = Sql & "                                                            ELSE Documents.MasterCode END ) "
            Sql = Sql & "         LEFT JOIN In_Out IO ON IO.ELID = (SELECT ElementID FROM Wares WHERE WareCode = " & GV.GetFocusedDataRow.Item("Warecode").ToString & ")" '[31.08.2016][31726][Мухаметов Р.И.] - Присоединил дополнительную таблицу.
            Sql = Sql & "         LEFT JOIN "
            Sql = Sql & "        (SELECT D.doc_code,"
            Sql = Sql & "                CAST(D.doc_text AS VARCHAR)          AS TEXT,"
            Sql = Sql & "                LZK.doc_text                         AS ParentDocText"
            Sql = Sql & "         FROM   Structure AS S"
            Sql = Sql & "                RIGHT OUTER JOIN Users AS U"
            Sql = Sql & "                                 RIGHT OUTER JOIN DocumentDocument"
            Sql = Sql & "                                                  LEFT OUTER JOIN Documents AS LZK ON DocumentDocument.ParentID = LZK.doc_code"
            Sql = Sql & "                                                  RIGHT OUTER JOIN Documents AS D ON DocumentDocument.ChildID = D.doc_code ON U.UserID = D.UserID ON S.StructureID = D.MasterCode)"
            Sql = Sql & "         As A On A.doc_code = documents.doc_code"
            Sql = Sql & "    WHERE    (provodka.ware_code            = " & GV.GetFocusedDataRow.Item("warecode").ToString & " OR provodka.ware_code = IO.fact_no)" '[31.08.2016][31726][Мухаметов Р.И.] - Добавил условие ОR для заполнения движения инструмента.
            Sql = Sql & dat
            Sql = Sql & "    AND Documents.doc_type IN (1  ,"
            Sql = Sql & "                               3  ,"
            Sql = Sql & "                               22 ,"
            Sql = Sql & "                               100,"
            Sql = Sql & "                               101) "
            Sql = Sql & "    AND CASE "
            Sql = Sql & "   WHEN Documents.doc_type = 1 "
            Sql = Sql & "        AND Documents.MasterCode = " & gWHS & " THEN 1 "
            Sql = Sql & "   WHEN Documents.doc_type = 100 "
            Sql = Sql & "        AND Documents.MasterCode = " & gWHS & " THEN 1 "
            Sql = Sql & "   WHEN Documents.doc_type = 3 "
            Sql = Sql & "        AND Documents.cli_code = " & gWHS & " THEN 1 "
            Sql = Sql & "   WHEN Documents.doc_type = 101 "
            Sql = Sql & "        AND Documents.cli_code =  " & gWHS & " THEN 1 "
            Sql = Sql & "   WHEN Documents.doc_type = 22"
            Sql = Sql & "        AND ( Documents.cli_code = " & gWHS & " "
            Sql = Sql & "              OR Documents.MasterCode = " & gWHS & ") THEN 1 "
            Sql = Sql & "   ELSE 0 "
            Sql = Sql & "   END = 1 "
            Sql = Sql & " ORDER BY Documents.data_make desc"

            SQLEx(Sql)

            While reader.Read
                dt.Rows.Add()
                If reader.Item("Document_Type") = "" Then
                    dt.Rows(dt.Rows.Count() - 1).Item(0) = IIF_S(reader.Item("doc_type") = 1 Or reader.Item("doc_type") = 101, "Приход", "Расход")
                Else
                    dt.Rows(dt.Rows.Count() - 1).Item(0) = reader.Item("Document_Type")
                End If
                dt.Rows(dt.Rows.Count() - 1).Item(1) = reader.Item("data_make")
                dt.Rows(dt.Rows.Count() - 1).Item(2) = reader.Item("Cool")
                dt.Rows(dt.Rows.Count() - 1).Item(3) = reader.Item("doc_text")
                dt.Rows(dt.Rows.Count() - 1).Item(4) = reader.Item("PayInSlip")
                dt.Rows(dt.Rows.Count() - 1).Item(5) = reader.Item("NumFact")
                dt.Rows(dt.Rows.Count() - 1).Item(6) = reader.Item("Num")
                dt.Rows(dt.Rows.Count() - 1).Item(7) = reader.Item("Partner")
                dt.Rows(dt.Rows.Count() - 1).Item(8) = reader.Item("cList")
                dt.Rows(dt.Rows.Count() - 1).Item(9) = reader.Item("doc_code")
                dt.Rows(dt.Rows.Count() - 1).Item(10) = reader.Item("doc_type")
            End While

            ConClose()

            frmInOut.GridView1.Columns.Clear()
            frmInOut.GridControl1.DataSource = dt
            frmInOut.GridView1.Columns.Item("doc_code").Visible = False
            frmInOut.GridView1.Columns.Item("doc_type").Visible = False
            frmInOut.GridView1.BestFitColumns()
            Dim i As Integer
            For i = 0 To 7
                frmInOut.GridView1.Columns.Item(i).BestFit()
            Next

            frmInOut.mMode = "Движение по элементу"
            frmInOut.Text = "Движение по элементу - " & GV.GetFocusedDataRow.Item("Наименование").ToString
            frmInOut.Label1.Text = "Движение по элементу - " & GV.GetFocusedDataRow.Item("Наименование").ToString
        Catch Ex As Exception
            ConClose()
            MsgBox(Ex.Message)
        End Try
    End Sub

	
    Private Sub GridViewPeriod_RowClick(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.RowClickEventArgs) Handles GridViewPeriod.RowClick
        rowPeriodClick = 1
        If e.Button = Windows.Forms.MouseButtons.Right Then
            ContxMnuStrPeriod.Show(MousePosition)
        End If
    End Sub

	
End Class

