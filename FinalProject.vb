Imports System.Data
Imports System.Data.SqlClient
Partial Class FinalProject
    Inherits System.Web.UI.Page

#Region "Data Connection"
    Public Shared strcon As String = System.Configuration.ConfigurationManager.ConnectionStrings("SJ13alex.huntConnectionString").ConnectionString
    Public Shared con As New SqlConnection(strcon)
#End Region

#Region "Global Variables"

    Public Shared User_Login As Boolean = False
    Public Shared LoginCount As Integer = 3
    'Declares the type of error that gets passed into the msgbox error function
    Public Shared CommonError = vbExclamation + vbMsgBoxSetForeground + vbApplicationModal
    Public Shared VerifyInput = vbInformation + vbYesNo + vbMsgBoxSetForeground + vbApplicationModal

    Public Shared DAContractsR As New SqlDataAdapter("SELECT TOP 1 * FROM [dbo].[Contracts] ORDER BY DeliveryID DESC", con)
    Public Shared DAExpensesR As New SqlDataAdapter("SELECT TOP 1 * FROM [dbo].[Expenses] ORDER BY ExpenseID DESC", con)
    Public Shared DAContracts As New SqlDataAdapter("SELECT * FROM [dbo].[Contracts] ORDER BY DeliveryID DESC", con)
    Public Shared DAExpenses As New SqlDataAdapter("SELECT * FROM [dbo].[Expenses] ORDER BY ExpenseID DESC", con)
    Public Shared DATruckers As New SqlDataAdapter("SELECT * FROM [dbo].[Truckers]", con)
    Public Shared DATrucks As New SqlDataAdapter("SELECT * FROM [dbo].[Trucks]", con)
    Public Shared DAClients As New SqlDataAdapter("SELECT * FROM [dbo].[Clients]", con)
    Public Shared DARoutes As New SqlDataAdapter("SELECT * FROM [dbo].[Routes]", con)

    Public Shared CBContracts As New SqlCommandBuilder(DAContracts)
    Public Shared CBExpenses As New SqlCommandBuilder(DAExpenses)
    Public Shared CBTruckers As New SqlCommandBuilder(DATruckers)
    Public Shared CBTrucks As New SqlCommandBuilder(DATrucks)
    Public Shared CBRoutes As New SqlCommandBuilder(DARoutes)
    Public Shared CBClients As New SqlCommandBuilder(DAClients)

    Public Shared TruckersDT As New DataTable
    Public Shared TrucksDT As New DataTable
    Public Shared ClientsDT As New DataTable
    Public Shared RoutesDT As New DataTable

    Public Shared ContractsDTR As New DataTable
    Public Shared ExpensesDTR As New DataTable
    Public Shared ContractsDT As New DataTable
    Public Shared ExpensesDT As New DataTable

#End Region

#Region "Multiview"
    Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
        MultiView1.ActiveViewIndex = 0
        Call SelectedView()
        LoginLink.ForeColor = Drawing.Color.DarkBlue
        LoginLink.BorderStyle = BorderStyle.Ridge
        Call Links()
        Call FillDataTables()
    End Sub

    Public Sub Links()
        Call CheckLogin()
        If User_Login = False Then
            MultiView1.ActiveViewIndex = 0
            DataLink.ForeColor = Drawing.Color.Gray
            ContractsLink.ForeColor = Drawing.Color.Gray
            ExpenseLink.ForeColor = Drawing.Color.Gray
            CurrentUser.Text = Nothing
        Else
            DataLink.ForeColor = Nothing
            ContractsLink.ForeColor = Nothing
            ExpenseLink.ForeColor = Nothing
        End If
    End Sub

    Public Sub CheckLogin()
        If User_Login <> True Then
            User_Login = False
        End If
    End Sub
    Protected Sub LoginLink_Click(sender As Object, e As EventArgs) Handles LoginLink.Click
        MultiView1.ActiveViewIndex = 0
        Call SelectedView()
        LoginLink.ForeColor = Drawing.Color.DarkBlue
        LoginLink.BorderStyle = BorderStyle.Ridge
        Call Links()
    End Sub

    Protected Sub ContractsLink_Click(sender As Object, e As EventArgs) Handles ContractsLink.Click
        If User_Login = True Then
            MultiView1.ActiveViewIndex = 1
            Call SelectedView()
            ContractsLink.ForeColor = Drawing.Color.DarkBlue
            ContractsLink.BorderStyle = BorderStyle.Ridge
        Else
            Response.Write("UserID and Password required")
        End If
    End Sub

    Protected Sub ExpenseLink_Click(sender As Object, e As EventArgs) Handles ExpenseLink.Click
        If User_Login = True Then
            MultiView1.ActiveViewIndex = 2
            Call SelectedView()
            ExpenseLink.ForeColor = Drawing.Color.DarkBlue
            ExpenseLink.BorderStyle = BorderStyle.Ridge
        Else
            Response.Write("UserID and Password required")
        End If
    End Sub

    'Only allows you to access the view when a user is logged in.
    Protected Sub DataLink_Click(sender As Object, e As EventArgs) Handles DataLink.Click
        If User_Login = True Then
            MultiView1.ActiveViewIndex = 3
            'Resets the forecolor and borderstyle for all link buttons
            Call SelectedView()
            DataLink.ForeColor = Drawing.Color.DarkBlue
            DataLink.BorderStyle = BorderStyle.Ridge
        Else
            Response.Write("UserID and Password required")
        End If
    End Sub

    Public Sub SelectedView()
        LoginLink.ForeColor = Nothing
        LoginLink.BorderStyle = Nothing
        ContractsLink.ForeColor = Nothing
        ContractsLink.BorderStyle = Nothing
        ExpenseLink.ForeColor = Nothing
        ExpenseLink.BorderStyle = Nothing
        DataLink.ForeColor = Nothing
        DataLink.BorderStyle = Nothing
    End Sub
    Protected Sub DataList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataList.SelectedIndexChanged
        If DataList.SelectedIndex = 1 Then
            MultiView2.ActiveViewIndex = 0
        ElseIf DataList.SelectedIndex = 2 Then
            MultiView2.ActiveViewIndex = 1
        ElseIf DataList.SelectedIndex = 3 Then
            MultiView2.ActiveViewIndex = 2
        ElseIf DataList.SelectedIndex = 4 Then
            MultiView2.ActiveViewIndex = 3
        End If
    End Sub

#End Region

#Region "Login"
    Protected Sub LoginButton_Click(sender As Object, e As EventArgs) Handles LoginButton.Click
        'Retrieves the user's name
        Dim LoginUserDA As New SqlDataAdapter("SELECT UserName FROM [dbo].[Users] WHERE UserID = @p1 AND Password = @p2", con)
        Dim Users As New DataTable
        Dim LoginCheck As New SqlCommand("SELECT COUNT(*) FROM [dbo].[Users] WHERE UserID = @p1 AND Password = @p2", con)
        Dim rows As Integer = 0
        With LoginCheck.Parameters
            .Clear()
            .AddWithValue("@p1", UserID.Text)
            .AddWithValue("@p2", Password.Text)
        End With
        With LoginUserDA.SelectCommand.Parameters
            .Clear()
            .AddWithValue("@p1", UserID.Text)
            .AddWithValue("@p2", Password.Text)
        End With
        Try
            If con.State = ConnectionState.Closed Then con.Open()
            rows = LoginCheck.ExecuteScalar
            If rows = 0 Then
                LoginCount -= 1
                Label1.Text = "Login failed, " & LoginCount & " attempts left"
                UserID.Text = Nothing
                Password.Text = Nothing
            End If
            If rows = 1 Then
                User_Login = True
                'Controls the color of the link buttons
                Call Links()
                'Takes the user to the contracts page upon login
                MultiView1.ActiveViewIndex = 1
                'Resets the color and borders of all the link buttons
                Call SelectedView()
                'Indicates that the button is selected
                ContractsLink.ForeColor = Drawing.Color.DarkBlue
                ContractsLink.BorderStyle = BorderStyle.Ridge
                UserID.Text = Nothing
                Password.Text = Nothing
                LoginCount = 3
                Label1.Text = "3 login attempts"
            End If
            LoginUserDA.Fill(Users)
            'Shows the current user at the top of the page
            CurrentUser.Text = Users.Rows(0).Item("UserName")
        Catch ex As Exception
            Response.Write(ex.Message)
        Finally
            con.Close()
        End Try
        If LoginCount = 0 Then
            Label1.Text = "You have been locked out, too many attempts"
            User_Login = False
            LoginButton.Enabled = False
            UserID.Enabled = False
            Password.Enabled = False
            Call Links()
        End If
    End Sub
#End Region

#Region "Load Data"
    Protected Sub FinalProject_Init(sender As Object, e As System.EventArgs) Handles Me.Init
        Call LoadLists()
        Call UpdateProfits()
    End Sub
    'Fills DropDownlists
    Public Sub LoadLists()

        Dim DATrucker As New SqlDataAdapter("SELECT * FROM [dbo].[Truckers]", con)
        Dim DATruckerContract As New SqlDataAdapter("SELECT * FROM [dbo].[Truckers] WHERE EmploymentStatus = 'Active'", con)
        Dim DATruck As New SqlDataAdapter("SELECT * FROM [dbo].[Trucks]", con)
        Dim DAClient As New SqlDataAdapter("SELECT * FROM [dbo].[Clients]", con)
        Dim DARoutes As New SqlDataAdapter("SELECT * FROM [dbo].[Routes]", con)

        Dim TruckerDT As New DataTable
        Dim TruckerDTA As New DataTable
        Dim TruckDT As New DataTable
        Dim ClientDT As New DataTable
        Dim RouteDT As New DataTable

        Try
            DATruckerContract.Fill(TruckerDTA)
            DATrucker.Fill(TruckerDT)
            With ContractTrucker
                .DataSource = TruckerDTA
                .DataTextField = "TruckerName"
                .DataValueField = "TruckerID"
                .DataBind()
                .Items.Insert(0, "Select a trucker")
            End With
            With ExpenseTrucker
                .DataSource = TruckerDTA
                .DataTextField = "TruckerName"
                .DataValueField = "TruckerID"
                .DataBind()
                .Items.Insert(0, "Select a trucker")
            End With
            With TruckerNameList
                .DataSource = TruckerDT
                .DataTextField = "TruckerName"
                .DataValueField = "TruckerID"
                .DataBind()
                .Items.Insert(0, "Select a trucker")
            End With
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try

        Try
            DATruck.Fill(TruckDT)
            With ContractTruck
                .DataSource = TruckDT
                .DataTextField = "Type"
                .DataValueField = "TruckID"
                .DataBind()
                .Items.Insert(0, "Select a truck")
            End With
            With ExpenseTruck
                .DataSource = TruckDT
                .DataTextField = "Type"
                .DataValueField = "TruckID"
                .DataBind()
                .Items.Insert(0, "Select a truck")
            End With
            With TruckTypeList
                .DataSource = TruckDT
                .DataTextField = "Type"
                .DataValueField = "TruckID"
                .DataBind()
                .Items.Insert(0, "Select a truck")
            End With
            With TruckSpeciality
                .DataSource = TruckDT
                .DataTextField = "Type"
                .DataValueField = "TruckID"
                .DataBind()
                .Items.Insert(0, "Select a truck")
            End With
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try

        Try
            DAClient.Fill(ClientDT)
            With ContractClient
                .DataSource = ClientDT
                .DataTextField = "ClientName"
                .DataValueField = "ClientID"
                .DataBind()
                .Items.Insert(0, "Select a client")
            End With
            With ClientNameList
                .DataSource = ClientDT
                .DataTextField = "ClientName"
                .DataValueField = "ClientID"
                .DataBind()
                .Items.Insert(0, "Select a client")
            End With
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try

        Try
            DARoutes.Fill(RouteDT)
            With ContractRoute
                .DataSource = RouteDT
                .DataTextField = "RouteName"
                .DataValueField = "RouteID"
                .DataBind()
                .Items.Insert(0, "Select a route")
            End With
            With ExpenseRoute
                .DataSource = RouteDT
                .DataTextField = "RouteName"
                .DataValueField = "RouteID"
                .DataBind()
                .Items.Insert(0, "Select a route")
            End With
            With RouteNameList
                .DataSource = RouteDT
                .DataTextField = "RouteName"
                .DataValueField = "RouteID"
                .DataBind()
                .Items.Insert(0, "Select a route")
            End With
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub
    'Contracts truck type data fill
    Protected Sub ContractTrucker_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ContractTrucker.SelectedIndexChanged
        If ContractTrucker.SelectedIndex <= 0 Then Exit Sub
        Dim TruckerDA As New SqlDataAdapter("SELECT * FROM [dbo].[Truckers] WHERE TruckerID = @p1", con)
        Dim TruckerDT As New DataTable
        With TruckerDA.SelectCommand.Parameters
            .Clear()
            .AddWithValue("@p1", ContractTrucker.SelectedValue)
        End With
        Try
            TruckerDA.Fill(TruckerDT)
            With TruckerDT.Rows(0)
                ContractTruck.SelectedValue = .Item("TruckID")
            End With
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub
    'Contracts payment type data fill
    Protected Sub ContractClient_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ContractClient.SelectedIndexChanged
        If ContractClient.SelectedIndex <= 0 Then Exit Sub
        Dim ClientDA As New SqlDataAdapter("SELECT * FROM [dbo].[Clients] WHERE ClientID = @p1", con)
        Dim ClientDT As New DataTable
        With ClientDA.SelectCommand.Parameters
            .Clear()
            .AddWithValue("@p1", ContractClient.SelectedValue)
        End With
        Try
            ClientDA.Fill(ClientDT)
            With ClientDT.Rows(0)
                PaymentTypeList.SelectedValue = .Item("PaymentType")
            End With
            Select Case PaymentTypeList.SelectedValue
                Case "Weight"
                    ShipmentLabel.Text = "Price per Ton"
                Case "Distance"
                    ShipmentLabel.Text = "Price per Mile"
                Case "Bid"
                    ShipmentLabel.Text = "Shipment Cost"
            End Select
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub
    'Payment type label control
    Protected Sub PaymentTypeList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PaymentTypeList.SelectedIndexChanged
        Select Case PaymentTypeList.SelectedValue
            Case "Weight"
                ShipmentLabel.Text = "Price per Ton"
            Case "Distance"
                ShipmentLabel.Text = "Price per Mile"
            Case "Bid"
                ShipmentLabel.Text = "Shipment Cost"
        End Select
    End Sub
    'Contracts mileage data fill
    Protected Sub ContractRoute_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ContractRoute.SelectedIndexChanged
        If ContractRoute.SelectedIndex <= 0 Then Exit Sub
        Dim RouteDA As New SqlDataAdapter("SELECT * FROM [dbo].[Routes] WHERE RouteID = @p1", con)
        Dim RouteDT As New DataTable
        With RouteDA.SelectCommand.Parameters
            .Clear()
            .AddWithValue("@p1", ContractRoute.SelectedValue)
        End With
        Try
            RouteDA.Fill(RouteDT)
            With RouteDT.Rows(0)
                ContractMileage.Text = .Item("Miles")
            End With
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub
    'Expense truck type data fill
    Protected Sub ExpenseTrucker_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ExpenseTrucker.SelectedIndexChanged
        If ExpenseTrucker.SelectedIndex <= 0 Then Exit Sub
        Dim TruckerDA As New SqlDataAdapter("SELECT * FROM [dbo].[Truckers] WHERE TruckerID = @p1", con)
        Dim TruckerDT As New DataTable
        With TruckerDA.SelectCommand.Parameters
            .Clear()
            .AddWithValue("@p1", ExpenseTrucker.SelectedValue)
        End With
        Try
            TruckerDA.Fill(TruckerDT)
            With TruckerDT.Rows(0)
                ExpenseTruck.SelectedValue = .Item("TruckID")
            End With
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub
    'Automatically fills the clients information. An autofill process is done for almost every selected index change
    Protected Sub ClientNameList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ClientNameList.SelectedIndexChanged
        If ClientNameList.SelectedIndex <= 0 Then Exit Sub
        Dim ClientDA As New SqlDataAdapter("SELECT * FROM [dbo].[Clients] WHERE ClientID = @p1", con)
        Dim ClientDT As New DataTable
        With ClientDA.SelectCommand.Parameters
            .Clear()
            .AddWithValue("@p1", ClientNameList.SelectedValue)
        End With
        Try
            ClientDA.Fill(ClientDT)
            With ClientDT.Rows(0)
                ClientName.Text = .Item("ClientName")
                ClientIndustry.Text = .Item("Industry")
                ClientRegion.SelectedValue = .Item("Region")
                ClientPaymentType.SelectedValue = .Item("PaymentType")
            End With
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub
    'Route data fill
    Protected Sub RouteNameList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RouteNameList.SelectedIndexChanged
        If RouteNameList.SelectedIndex <= 0 Then Exit Sub
        Dim RouteDA As New SqlDataAdapter("SELECT * FROM [dbo].[Routes] WHERE RouteID = @p1", con)
        Dim RouteDT As New DataTable
        With RouteDA.SelectCommand.Parameters
            .Clear()
            .AddWithValue("@p1", RouteNameList.SelectedValue)
        End With
        Try
            RouteDA.Fill(RouteDT)
            With RouteDT.Rows(0)
                RouteName.Text = .Item("RouteName")
                OriginState.Text = .Item("OriginState")
                OriginCity.Text = .Item("OriginCity")
                DestinationState.Text = .Item("DestinationState")
                DestinationCity.Text = .Item("DestinationCity")
                RouteDistance.Text = .Item("Miles")
                DistanceConversion.SelectedIndex = 0
            End With
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub
    'Trucker data fill
    Protected Sub TruckerNameList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TruckerNameList.SelectedIndexChanged
        If TruckerNameList.SelectedIndex <= 0 Then Exit Sub
        Dim TruckerDA As New SqlDataAdapter("SELECT * FROM [dbo].[Truckers] WHERE TruckerID = @p1", con)
        Dim TruckerDT As New DataTable
        With TruckerDA.SelectCommand.Parameters
            .Clear()
            .AddWithValue("@p1", TruckerNameList.SelectedValue)
        End With
        Try
            TruckerDA.Fill(TruckerDT)
            With TruckerDT.Rows(0)
                TruckSpeciality.SelectedValue = .Item("TruckID")
                TruckerName.Text = .Item("TruckerName")
                TruckerRegion.SelectedValue = .Item("Region")
                TruckerExperience.Text = .Item("YearsExperience")
                EmploymentType.SelectedValue = .Item("EmploymentType")
                EmploymentStatus.SelectedValue = .Item("EmploymentStatus")
            End With
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub
    'Truck type data fill
    Protected Sub TruckTypeList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TruckTypeList.SelectedIndexChanged
        If TruckTypeList.SelectedIndex <= 0 Then Exit Sub
        Dim TruckDA As New SqlDataAdapter("SELECT * FROM [dbo].[Trucks] WHERE TruckID = @p1", con)
        Dim TruckDT As New DataTable
        With TruckDA.SelectCommand.Parameters
            .Clear()
            .AddWithValue("@p1", TruckTypeList.SelectedValue)
        End With
        Try
            TruckDA.Fill(TruckDT)
            With TruckDT.Rows(0)
                TruckTypeTxt.Text = .Item("Type")
                TruckMileage.Text = .Item("Mileage")
                TruckLease.Text = .Item("LeaseAmount")
                TruckPayload.Text = .Item("PayloadRating")
            End With
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub
#End Region

#Region "Display Data"
    Public Sub ContractRecord()
        If ContractsDTR.Rows.Count > 0 Then ContractsDTR.Rows.Clear()
        Try
            DAContractsR.Fill(ContractsDTR)
            GridView1.DataSource = ContractsDTR
            GridView1.DataBind()
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub
    Public Sub ExpenseRecord()
        If ExpensesDTR.Rows.Count > 0 Then ExpensesDTR.Rows.Clear()
        Try
            DAExpensesR.Fill(ExpensesDTR)
            GridView3.DataSource = ExpensesDTR
            GridView3.DataBind()
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub
    Public Sub FillDataTables()
        Try
            If ContractsDT.Rows.Count > 0 Then ContractsDT.Rows.Clear()
            DAContracts.Fill(ContractsDT)
            GridView2.DataSource = ContractsDT
            GridView2.DataBind()
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try

        Try
            If ExpensesDT.Rows.Count > 0 Then ExpensesDT.Rows.Clear()
            DAExpenses.Fill(ExpensesDT)
            GridView4.DataSource = ExpensesDT
            GridView4.DataBind()
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try

        Try
            If ClientsDT.Rows.Count > 0 Then ClientsDT.Rows.Clear()
            DAClients.Fill(ClientsDT)
            GridView5.DataSource = ClientsDT
            GridView5.DataBind()
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try

        Try
            If RoutesDT.Rows.Count > 0 Then RoutesDT.Rows.Clear()
            DARoutes.Fill(RoutesDT)
            GridView6.DataSource = RoutesDT
            GridView6.DataBind()
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try

        Try
            If TruckersDT.Rows.Count > 0 Then TruckersDT.Rows.Clear()
            DATruckers.Fill(TruckersDT)
            GridView7.DataSource = TruckersDT
            GridView7.DataBind()
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try

        Try
            If TrucksDT.Rows.Count > 0 Then TrucksDT.Rows.Clear()
            DATrucks.Fill(TrucksDT)
            GridView8.DataSource = TrucksDT
            GridView8.DataBind()
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try

        Call UpdateProfits()
    End Sub
#End Region

#Region "Create Records"
    Protected Sub RecordInvoice_Click(sender As Object, e As EventArgs) Handles RecordInvoice.Click

        DAContracts.FillSchema(ContractsDT, SchemaType.Mapped)
        If ContractsDT.Rows.Count > 0 Then ContractsDT.Rows.Clear()

        Dim dr As DataRow = ContractsDT.NewRow
        Dim DeliveryDate As Date
        Dim Weight As Decimal
        Dim Cost As Decimal
        Dim Payout As Decimal
        Dim Experience As Integer
        Dim PayRate As Decimal

        If ContractTrucker.SelectedIndex <= 0 Then
            MsgBox("Cannot create invoice without a trucker", CommonError, "Missing Input")
            Exit Sub
        End If
        If ContractClient.SelectedIndex <= 0 Then
            MsgBox("Cannot create invoice without a client", CommonError, "Missing Input")
            Exit Sub
        End If
        If ContractRoute.SelectedIndex <= 0 Then
            MsgBox("Cannot create invoice without a delivery route", CommonError, "Missing Input")
            Exit Sub
        End If
        If ContractTruck.SelectedIndex <= 0 Then
            MsgBox("Cannot create invoice without a truck", CommonError, "Missing Input")
            Exit Sub
        End If
        If ContractProduct.Text = Nothing Then
            MsgBox("Cannot create invoice without a product", CommonError, "Missing Input")
            Exit Sub
        End If
        If PaymentTypeList.SelectedIndex < 0 Then
            MsgBox("Cannot create invoice without the client's payment type", CommonError, "Missing Input")
            Exit Sub
        End If
        If ContractDate.Text = Nothing Then
            MsgBox("Cannot create invoice without a date", CommonError, "Missing Input")
            Exit Sub
        End If
        'Error check for missing input
        If ContractMileage.Text = Nothing Then
            MsgBox("Cannot create invoice without the route's mileage", CommonError, "Missing Input")
            Exit Sub
        End If
        'Error check for faulty input
        If ContractMileage.Text < 0 Then
            MsgBox("Cannot create invoice with negative mileage", CommonError, "Faulty Input")
            Exit Sub
        End If
        If ContractWeight.Text = Nothing Then
            MsgBox("Cannot create invoice without the delivery weight", CommonError, "Missing Input")
        End If
        If ContractWeight.Text < 0 Then
            MsgBox("Cannot create invoice with negative weight", CommonError, "Faulty Input")
        End If
        If ContractCost.Text = Nothing Then
            MsgBox("Input delivery cost/price to create invoice", CommonError, "Missing Input")
        End If
        If ContractCost.Text < 0 Then
            MsgBox("Delivery cost cannot be negative", CommonError, "Faulty Input")
        End If
        If ContractCost.Text = 0 Then
            Dim msgBoxResult As MsgBoxResult = MsgBox("Are you sure you want to input a cost of 0", VerifyInput, "Input Verification")
            If msgBoxResult = MsgBoxResult.No Then
                Exit Sub
            End If
        End If

        If WeightConversion.SelectedIndex = 1 Then
            Weight = ContractWeight.Text / 2000
        Else
            Weight = ContractWeight.Text
        End If
        DeliveryDate = Date.Parse(ContractDate.Text)

        If PaymentTypeList.SelectedIndex = 0 Then
            Cost = Weight * ContractCost.Text
        ElseIf PaymentTypeList.SelectedIndex = 1 Then
            Cost = ContractMileage.Text * ContractCost.Text
        ElseIf PaymentTypeList.SelectedIndex = 2 Then
            Cost = ContractCost.Text
        End If

        Dim TruckerExperience As New SqlCommand("SELECT YearsExperience FROM [dbo].[Truckers] WHERE TruckerID = @p1", con)
        With TruckerExperience.Parameters
            .Clear()
            .AddWithValue("@p1", ContractTrucker.SelectedValue)
        End With
        Try
            If con.State = ConnectionState.Closed Then con.Open()
            Experience = TruckerExperience.ExecuteScalar
            If Experience <= 5 Then
                PayRate = 0.15
            ElseIf Experience > 5 And Experience <= 10 Then
                PayRate = 0.25
            Else
                PayRate = 0.35
            End If
        Catch ex As Exception
            Response.Write(ex.Message)
        Finally
            con.Close()
        End Try
        Payout = PayRate * Cost

        dr.Item("TruckerID") = ContractTrucker.SelectedValue
        dr.Item("ClientID") = ContractClient.SelectedValue
        dr.Item("RouteID") = ContractRoute.SelectedValue
        dr.Item("TruckID") = ContractTruck.SelectedValue
        dr.Item("ShipmentType") = PaymentTypeList.SelectedValue
        dr.Item("Commodity") = ContractProduct.Text
        dr.Item("Mileage") = ContractMileage.Text
        dr.Item("TonsDelivered") = Weight
        dr.Item("Payment") = Cost
        dr.Item("DeliveryDate") = DeliveryDate
        dr.Item("Month") = DeliveryDate.Month
        dr.Item("Year") = DeliveryDate.Year

        Dim TruckPayload As New SqlCommand("SELECT PayloadRating FROM [dbo].[Trucks] WHERE TruckID = @p1", con)
        Dim Payload As Decimal
        With TruckPayload.Parameters
            .Clear()
            .AddWithValue("@p1", ContractTruck.SelectedValue)
        End With
        'Checks if the truck can handle the shipment weight
        Try
            If con.State = ConnectionState.Closed Then con.Open()
            Payload = TruckPayload.ExecuteScalar
            If Weight > (Payload / 2000) Then
                Invoice.Text = "Cannot sign off on delivery, shipment weight exceeds the selected truck's payload rating"
                Exit Sub
            End If
        Catch ex As Exception
            Response.Write(ex.Message)
        Finally
            con.Close()
        End Try

        Try
            ContractsDT.Rows.Add(dr)
            DAContracts.Update(ContractsDT)

            Invoice.Text = "Delivery Invoice" _
                & "   " _
                & FormatDateTime(DeliveryDate, DateFormat.LongDate) _
                & vbNewLine _
                & vbNewLine _
                & "Shipment Cost: " _
                & FormatCurrency(Cost, 2) _
                & vbNewLine _
                & "Trucker's Rate: " _
                & FormatPercent(PayRate) _
                & vbNewLine _
                & "Trucker's Pay: " _
                & FormatCurrency(Payout, 2) _
                & vbNewLine _
                & vbNewLine _
                & "Delivery Information" _
                & vbNewLine _
                & vbNewLine _
                & "Trucker: " _
                & ContractTrucker.SelectedItem.ToString _
                & vbNewLine _
                & "Client: " _
                & ContractClient.SelectedItem.ToString _
                & vbNewLine _
                & "Product: " _
                & ContractProduct.Text.Trim _
                & vbNewLine _
                & "Weight (Tons): " _
                & FormatNumber(Weight) _
                & vbNewLine _
                & "Route: " _
                & ContractRoute.SelectedItem.ToString _
                & vbNewLine _
                & "Distance (Miles): " _
                & FormatNumber(ContractMileage.Text, 2) _
                & vbNewLine _
                & "Truck: " _
                & ContractTruck.SelectedItem.ToString
            '
            Call ContractRecord()
            Call UpdateContractData(DeliveryDate, Weight, Cost, Payout)
            Call CreateExpense(Payout, DeliveryDate)
            Call FillDataTables()
            Call ClearData()
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try

    End Sub

    Protected Sub TrackExpense_Click(sender As Object, e As EventArgs) Handles TrackExpense.Click
        DAExpenses.FillSchema(ExpensesDT, SchemaType.Mapped)
        If ExpensesDT.Rows.Count > 0 Then ExpensesDT.Rows.Clear()

        Dim dr As DataRow = ExpensesDT.NewRow
        Dim DeliveryDate As Date
        Dim DebtPayment As Decimal
        Dim Payout As Decimal

        If ExpenseTrucker.SelectedIndex <= 0 Then
            Response.Write("Missing Input")
            MsgBox("Cannot record an expense without a trucker", CommonError, "Missing Input")
            Exit Sub
        End If
        If ExpenseRoute.SelectedIndex <= 0 Then
            MsgBox("Cannot create an expense without a delivery route", CommonError, "Missing Input")
            Exit Sub
        End If
        If ExpenseTruck.SelectedIndex <= 0 Then
            MsgBox("Cannot create an expense without a truck", CommonError, "Missing Input")
            Exit Sub
        End If
        If ExpenseType.SelectedIndex = Nothing Then
            MsgBox("Cannot create an expense without the expense type", CommonError, "Missing Input")
            Exit Sub
        End If
        If ExpenseDate.Text = Nothing Then
            MsgBox("Cannot create an expense without a date", CommonError, "Missing Input")
            Exit Sub
        End If
        If ExpenseCost.Text = Nothing OrElse ExpenseCost.Text = 0 Then
            MsgBox("Cannot create an expense with no cost", CommonError, "Missing Input")
            Exit Sub
        End If
        'Adds a input verification box that asks if the user wants to add a comment when create an expense under the type 'Other' Comments are not required to make an expense but when selecting 'other', it is helpful to have more information for expense tracking
        If ExpenseComments.Text = Nothing And ExpenseType.SelectedIndex = 9 Then
            Dim result As MsgBoxResult = MsgBox("You have selected 'Other', do you want to write a comment describing the expense", VerifyInput, "Input Verification")
            If result = MsgBoxResult.Yes Then Exit Sub
        End If
        If ExpenseType.SelectedIndex = 3 Then
            DebtPayment = CDec(ExpenseCost.Text)
        Else
            DebtPayment = 0
        End If
        If ExpenseType.SelectedIndex = 2 Then
            Payout = CDec(ExpenseCost.Text)
        Else
            Payout = 0
        End If

        DeliveryDate = Date.Parse(ExpenseDate.Text)

        dr.Item("TruckerID") = ExpenseTrucker.SelectedValue
        dr.Item("RouteID") = ExpenseRoute.SelectedValue
        dr.Item("TruckID") = ExpenseTruck.SelectedValue
        dr.Item("Type") = ExpenseType.SelectedItem
        dr.Item("Comments") = ExpenseComments.Text
        dr.Item("Cost") = CDec(ExpenseCost.Text)
        dr.Item("Date") = DeliveryDate
        dr.Item("Month") = DeliveryDate.Month
        dr.Item("Year") = DeliveryDate.Year

        Dim Cost As Decimal = CDec(ExpenseCost.Text)
        Dim Trucker As Integer = ExpenseTrucker.SelectedValue
        Dim Route As Integer = ExpenseRoute.SelectedValue
        Dim Truck As Integer = ExpenseTruck.SelectedValue

        Try
            ExpensesDT.Rows.Add(dr)
            DAExpenses.Update(ExpensesDT)
            Call ExpenseRecord()
            Call UpdateExpenseData(Cost, Payout, DebtPayment, Trucker, Route, Truck)
            Call FillDataTables()
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
        Call ClearData()
    End Sub

    'Automatically creates a payroll expense. The invoice creation function call this with the parameters of trucker's payout and contract date
    Public Sub CreateExpense(Cost As Decimal, ExpenseDate As Date)
        DAExpenses.FillSchema(ExpensesDT, SchemaType.Mapped)
        If ExpensesDT.Rows.Count > 0 Then ExpensesDT.Rows.Clear()

        Dim dr As DataRow = ExpensesDT.NewRow
        dr.Item("TruckerID") = ContractTrucker.SelectedValue
        dr.Item("RouteID") = ContractRoute.SelectedValue
        dr.Item("TruckID") = ContractTruck.SelectedValue
        dr.Item("Type") = "Payroll"
        dr.Item("Comments") = "Auto payroll"
        dr.Item("Cost") = Cost
        dr.Item("Date") = ExpenseDate
        dr.Item("Month") = ExpenseDate.Month
        dr.Item("Year") = ExpenseDate.Year

        Dim Trucker As Integer = ContractTrucker.SelectedValue
        Dim Route As Integer = ContractRoute.SelectedValue
        Dim Truck As Integer = ContractTruck.SelectedValue

        Try
            ExpensesDT.Rows.Add(dr)
            DAExpenses.Update(ExpensesDT)
            'Fills in the single, most recent record data table for expenses
            Call ExpenseRecord()
            'Calls the update expense data where records associated with expenses get updated
            Call UpdateExpenseData(Cost, 0, 0, Trucker, Route, Truck)
            'Fills all data tables with the updated data
            Call FillDataTables()
            Call ClearData()
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub

#End Region

#Region "Update Data"
    Public Sub UpdateExpenseData(Payment As Decimal, Payout As Decimal, LeasePayment As Decimal, Trucker As Integer, Route As Integer, Truck As Integer)

        'Updates total expenses and total payments
        Dim UpdateTrucker As New SqlCommand("UPDATE [dbo].[Truckers] SET TotalExpenses += @p1, TotalPayments += @p2 WHERE TruckerID = @p0", con)
        'Updates total expenses
        Dim UpdateRoute As New SqlCommand("UPDATE [dbo].[Routes] SET TotalExpenses += @p1 WHERE RouteID = @p0", con)
        'Updates total expense, total paidoff
        Dim UpdateTruck As New SqlCommand("UPDATE [dbo].[Trucks] SET TotalExpenses += @p1, TotalPaidOff += @p4 WHERE TruckID = @p0", con)

        With UpdateTrucker.Parameters
            .Clear()
            .AddWithValue("@p0", Trucker)
            .AddWithValue("@p1", Payment)
            .AddWithValue("@p2", Payout)
        End With

        With UpdateRoute.Parameters
            .Clear()
            .AddWithValue("@p0", Route)
            .AddWithValue("@p1", Payment)
        End With

        With UpdateTruck.Parameters
            .Clear()
            .AddWithValue("@p0", Truck)
            .AddWithValue("@p1", Payment)
            .AddWithValue("@p4", LeasePayment)
        End With

        Try
            If con.State = ConnectionState.Closed Then con.Open()
            UpdateTrucker.ExecuteNonQuery()
            UpdateRoute.ExecuteNonQuery()
            UpdateTruck.ExecuteNonQuery()
            Call UpdateProfits()
        Catch ex As Exception
            Response.Write(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub
    'Only updates contract related expenses as well as payments. Parameters are passed in the UpdateExpenseData function so payments don't get duplicated
    Public Sub UpdateContractData(DeliveryDate As Date, Weight As Decimal, Cost As Decimal, Payout As Decimal)

        Dim UpdateTrucker As New SqlCommand("UPDATE [dbo].[Truckers] SET NumberDeliveries +=1, LastDelivery = @p1, Mileage += @p2, TotalTons += @p3, TotalIncome += @p4, TotalPayments += @p6 WHERE TruckerID = @p0", con)
        Dim UpdateClient As New SqlCommand("UPDATE [dbo].[Clients] SET NumberDeliveries += 1, LastDelivery = @p1, TotalTons += @p3, TotalPayments += @p4 WHERE ClientID = @p0", con)
        Dim UpdateRoute As New SqlCommand("UPDATE [dbo].[Routes] SET NumberDeliveries +=1, LastDelivery = @p1, TotalTons += @p3, TotalIncome += @p4 WHERE RouteID = @p0", con)
        Dim UpdateTruck As New SqlCommand("UPDATE [dbo].[Trucks] SET  NumberDeliveries +=1, LastDelivery = @p1, Mileage += @p2, TotalTons += @p3, TotalIncome += @p4 WHERE TruckID = @p0", con)

        With UpdateTrucker.Parameters
            .AddWithValue("@p0", ContractTrucker.SelectedValue)
            .AddWithValue("@p1", DeliveryDate)
            .AddWithValue("@p2", CDec(ContractMileage.Text))
            .AddWithValue("@p3", Weight)
            .AddWithValue("@p4", Cost)
            .AddWithValue("@p6", Payout)
        End With

        With UpdateClient.Parameters
            .AddWithValue("@p0", ContractClient.SelectedValue)
            .AddWithValue("@p1", DeliveryDate)
            .AddWithValue("@p4", Cost)
            .AddWithValue("@p3", Weight)
        End With

        With UpdateRoute.Parameters
            .AddWithValue("@p0", ContractRoute.SelectedValue)
            .AddWithValue("@p1", DeliveryDate)
            .AddWithValue("@p3", Weight)
            .AddWithValue("@p4", Cost)
        End With

        With UpdateTruck.Parameters
            .AddWithValue("@p0", ContractTruck.SelectedValue)
            .AddWithValue("@p1", DeliveryDate)
            .AddWithValue("@p2", CDec(ContractMileage.Text))
            .AddWithValue("@p3", Weight)
            .AddWithValue("@p4", Cost)
        End With

        Try
            If con.State = ConnectionState.Closed Then con.Open()
            UpdateTrucker.ExecuteNonQuery()
            UpdateClient.ExecuteNonQuery()
            UpdateRoute.ExecuteNonQuery()
            UpdateTruck.ExecuteNonQuery()
            Call UpdateProfits()
        Catch ex As Exception
            Response.Write(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub
    'Updates total profits for both trucks and routes and calculates the current debt for trucks. 
    Public Sub UpdateProfits()

        Dim UpdateProfitT As New SqlCommand("UPDATE [dbo].[Trucks] SET TotalProfit = (TotalIncome-TotalExpenses)", con)
        Dim UpdateProfitR As New SqlCommand("UPDATE [dbo].[Routes] SET TotalProfit = (TotalIncome-TotalExpenses)", con)
        Dim UpdateDebt As New SqlCommand("UPDATE [dbo].[Trucks] SET CurrentDebt = (LeaseAmount-TotalPaidoff)", con)

        Try
            If con.State = ConnectionState.Closed Then con.Open()
            UpdateProfitT.ExecuteNonQuery()
            UpdateProfitR.ExecuteNonQuery()
            UpdateDebt.ExecuteNonQuery()
        Catch ex As Exception
            Response.Write(ex.Message)
        Finally
            con.Close()
        End Try

    End Sub
#End Region

#Region "Manage Data"
    Protected Sub AddClient_Click(sender As Object, e As EventArgs) Handles AddClient.Click

        DAClients.FillSchema(ClientsDT, SchemaType.Mapped)

        Dim VerifyClient As New DataTable
        Dim DAClientName As New SqlDataAdapter("SELECT * FROM [dbo].[Clients] WHERE ClientName = @p1", con)
        With DAClientName.SelectCommand.Parameters
            .Clear()
            .AddWithValue("@p1", ClientName.Text.Trim)
        End With

        If ClientName.Text = Nothing Then
            MsgBox("Cannot add client without a name", CommonError, "Missing Input")
            Exit Sub
        End If
        If ClientIndustry.Text = Nothing Then
            MsgBox("Include the industry to add a client", CommonError, "Missing Input")
            Exit Sub
        End If
        If ClientRegion.SelectedIndex <= 0 Then
            MsgBox("Include the region to add a client", CommonError, "Missing Input")
            Exit Sub
        End If
        If ClientPaymentType.SelectedIndex < 0 Then
            MsgBox("Cannot add client without payment type", CommonError, "Missing Input")
            Exit Sub
        End If

        Try
            DAClientName.Fill(VerifyClient)
            If VerifyClient.Rows.Count > 0 Then
                MsgBox("Already a client with this name", CommonError, "Data Error")
                Exit Sub
            End If
            Dim dr As DataRow = ClientsDT.NewRow
            dr.Item("ClientName") = ClientName.Text.Trim
            dr.Item("Industry") = ClientIndustry.Text.Trim
            dr.Item("Region") = ClientRegion.SelectedValue
            dr.Item("PaymentType") = ClientPaymentType.SelectedValue
            dr.Item("LastDelivery") = DBNull.Value
            dr.Item("NumberDeliveries") = 0
            dr.Item("TotalTons") = 0
            dr.Item("TotalPayments") = 0
            ClientsDT.Rows.Add(dr)
            DAClients.Update(ClientsDT)
            Call LoadLists()
            Call FillDataTables()
            Call ClearData()
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub
    Protected Sub AddRoute_Click(sender As Object, e As EventArgs) Handles AddRoute.Click

        DARoutes.FillSchema(RoutesDT, SchemaType.Mapped)

        Dim Distance As Decimal
        Dim VerifyRoute As New DataTable
        Dim DARouteName As New SqlDataAdapter("SELECT * FROM [dbo].[Routes] WHERE RouteName = @p1", con)
        With DARouteName.SelectCommand.Parameters
            .Clear()
            .AddWithValue("@p1", RouteName.Text.Trim)
        End With

        If RouteName.Text = Nothing Then
            MsgBox("Cannot add new route without a name", CommonError, "Missing Input")
            Exit Sub
        End If
        If OriginState.Text = Nothing OrElse OriginCity.Text = Nothing Then
            MsgBox("Input the orgin location to add a route", CommonError, "Missing Input")
            Exit Sub
        End If
        If DestinationState.Text = Nothing OrElse DestinationCity.Text = Nothing Then
            MsgBox("Input the destination location to add a route", CommonError, "Missing Input")
            Exit Sub
        End If
        If RouteDistance.Text = Nothing Then
            MsgBox("Cannot add route without the distance", CommonError, "Missing Input")
            Exit Sub
        End If
        If RouteDistance.Text <= 0 Then
            MsgBox("Cannot have less than 1 mile", CommonError, "Faulty Input")
            Exit Sub
        End If

        'Manages the Miles to KM converter. This is the same process for the Tons and Lbs conversion
        If DistanceConversion.SelectedIndex = 1 Then
            Distance = RouteDistance.Text * 0.62
        Else
            Distance = RouteDistance.Text
        End If

        Try
            DARouteName.Fill(VerifyRoute)
            If VerifyRoute.Rows.Count > 0 Then
                MsgBox("Already a route with this name", CommonError, "Data Error")
                Exit Sub
            End If
            Dim dr As DataRow = RoutesDT.NewRow
            dr.Item("RouteName") = RouteName.Text.Trim
            dr.Item("OriginState") = OriginState.Text
            dr.Item("OriginCity") = OriginCity.Text
            dr.Item("DestinationState") = DestinationState.Text
            dr.Item("DestinationCity") = DestinationCity.Text
            dr.Item("Miles") = Distance
            'Sets the last delivery date of a new trucker/client/route/truck to a null value
            dr.Item("LastDelivery") = DBNull.Value
            dr.Item("NumberDeliveries") = 0
            dr.Item("TotalTons") = 0
            dr.Item("TotalIncome") = 0
            dr.Item("TotalExpenses") = 0
            dr.Item("TotalProfit") = 0
            RoutesDT.Rows.Add(dr)
            DARoutes.Update(RoutesDT)
            Call LoadLists()
            Call FillDataTables()
            Call ClearData()
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub
    Protected Sub AddTrucker_Click(sender As Object, e As EventArgs) Handles AddTrucker.Click

        DATruckers.FillSchema(TruckersDT, SchemaType.Mapped)

        Dim VerifyTrucker As New DataTable
        Dim DATruckerName As New SqlDataAdapter("SELECT * FROM [dbo].[Truckers] WHERE TruckerName = @p1", con)
        With DATruckerName.SelectCommand.Parameters
            .Clear()
            .AddWithValue("@p1", TruckerName.Text.Trim)
        End With

        If TruckerName.Text = Nothing Then
            MsgBox("Cannot add trucker without a name", CommonError, "Missing Input")
            Exit Sub
        End If
        If TruckerRegion.SelectedIndex <= 0 Then
            MsgBox("Include the region to add a trucker", CommonError, "Missing Input")
            Exit Sub
        End If
        If TruckSpeciality.SelectedIndex <= 0 Then
            MsgBox("Include a truck speciality to add trucker", CommonError, "Missing Input")
            Exit Sub
        End If
        If TruckerExperience.Text = Nothing Then
            MsgBox("Input experience to add trucker", CommonError, "Missing Input")
            Exit Sub
        End If
        If TruckerExperience.Text < 0 Then
            MsgBox("Cannot have negative experience", CommonError, "Faulty Input")
            Exit Sub
        End If
        If EmploymentType.SelectedIndex <= 0 Then
            MsgBox("Select employment type to add trucker", CommonError, "Missing Input")
            Exit Sub
        End If
        If EmploymentStatus.SelectedIndex <= 0 Then
            MsgBox("Select employment status to add trucker", CommonError, "Missing Input")
            Exit Sub
        End If

        Try
            DATruckerName.Fill(VerifyTrucker)
            If VerifyTrucker.Rows.Count > 0 Then
                MsgBox("Already a trucker with this name", CommonError, "Data Error")
                Exit Sub
            End If

            Dim dr As DataRow = TruckersDT.NewRow
            dr.Item("TruckerName") = TruckerName.Text.Trim
            dr.Item("Region") = TruckerRegion.SelectedValue
            dr.Item("TruckID") = TruckSpeciality.SelectedValue
            dr.Item("EmploymentType") = EmploymentType.SelectedValue
            dr.Item("EmploymentStatus") = EmploymentStatus.SelectedValue
            dr.Item("Mileage") = 0
            dr.Item("TotalIncome") = 0
            dr.Item("TotalExpenses") = 0
            dr.Item("LastDelivery") = DBNull.Value
            dr.Item("NumberDeliveries") = 0
            dr.Item("TotalTons") = 0
            dr.Item("TotalPayments") = 0
            dr.Item("YearsExperience") = TruckerExperience.Text
            TruckersDT.Rows.Add(dr)
            DATruckers.Update(TruckersDT)
            Call LoadLists()
            Call FillDataTables()
            Call ClearData()
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub
    Protected Sub AddTruck_Click(sender As Object, e As EventArgs) Handles AddTruck.Click

        DATrucks.FillSchema(TrucksDT, SchemaType.Mapped)

        Dim Payload As Decimal
        Dim VerifyTruck As New DataTable
        Dim DATruckName As New SqlDataAdapter("SELECT * FROM [dbo].[Trucks] WHERE Type = @p1", con)
        With DATruckName.SelectCommand.Parameters
            .Clear()
            .AddWithValue("@p1", TruckTypeTxt.Text.Trim)
        End With

        If TruckTypeTxt.Text = Nothing Then
            MsgBox("Cannot add truck without a type/name", CommonError, "Missing Input")
            Exit Sub
        End If
        If TruckMileage.Text = Nothing Then
            MsgBox("Include the starting mileage to add truck", CommonError, "Missing Input")
            Exit Sub
        End If
        If TruckMileage.Text < 0 Then
            MsgBox("Cannot have negative mileage", CommonError, "Faulty Data")
            Exit Sub
        End If
        If TruckLease.Text = Nothing Then
            MsgBox("Input lease amount or cost to add truck", CommonError, "Missing Input")
            Exit Sub
        End If
        If TruckLease.Text < 0 Then
            MsgBox("Cannot have negative input", CommonError, "Faulty Input")
            Exit Sub
        End If
        If TruckPayload.Text = Nothing Then
            MsgBox("For safety reasons, the payload rating is required", CommonError, "Missing Input")
            Exit Sub
        End If
        If TruckPayload.Text < 0 Then
            MsgBox("Cannot have negative input", CommonError, "Faulty Input")
            Exit Sub
        End If

        If (TruckPayload.Text < 1000 And PayloadConversion.SelectedIndex = 0) Or (TruckPayload.Text < 0.5 And PayloadConversion.SelectedIndex = 1) Then
            Dim result As MsgBoxResult = MsgBox("Is " & TruckPayload.Text & " the correct payload, this truck won't be able to carry many deliveries", VerifyInput, "Input Verification")
            If result = MsgBoxResult.No Then Exit Sub
        End If

        If PayloadConversion.SelectedIndex = 1 Then
            Payload = TruckPayload.Text * 2000
        Else
            Payload = TruckPayload.Text
        End If

        Try
            DATruckName.Fill(VerifyTruck)
            'Checks for same name when adding/updating a truck. Because there can be two of the same trucks, the program suggests adding a number
            If VerifyTruck.Rows.Count > 0 Then
                MsgBox("Already a truck with this name, use numbers to differentiate among trucks of the same type", CommonError, "Data Error")
                Exit Sub
            End If

            Dim dr As DataRow = TrucksDT.NewRow
            dr.Item("Type") = TruckTypeTxt.Text.Trim
            dr.Item("Mileage") = TruckMileage.Text
            dr.Item("LeaseAmount") = TruckLease.Text
            dr.Item("TotalIncome") = 0
            dr.Item("TotalExpenses") = 0
            dr.Item("LastDelivery") = DBNull.Value
            dr.Item("NumberDeliveries") = 0
            dr.Item("TotalTons") = 0
            dr.Item("TotalPaidOff") = 0
            dr.Item("CurrentDebt") = TruckLease.Text
            dr.Item("PayloadRating") = Payload
            dr.Item("TotalProfit") = 0
            TrucksDT.Rows.Add(dr)
            DATrucks.Update(TrucksDT)
            Call LoadLists()
            Call FillDataTables()
            Call ClearData()
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub

    Protected Sub UpdateClient_Click(sender As Object, e As EventArgs) Handles UpdateClient.Click
        If ClientNameList.SelectedIndex <= 0 Then
            MsgBox("Select a client to make updates", CommonError, "Missing Input")
            Exit Sub
        End If
        If ClientName.Text = Nothing Then
            MsgBox("Cannot update client with no name", CommonError, "Missing Input")
            Exit Sub
        End If
        If ClientIndustry.Text = Nothing Then
            MsgBox("Include the industry to update the client", CommonError, "Missing Input")
            Exit Sub
        End If
        If ClientRegion.SelectedIndex <= 0 Then
            MsgBox("Include the region to update client", CommonError, "Missing Input")
            Exit Sub
        End If
        If ClientPaymentType.SelectedIndex < 0 Then
            MsgBox("Cannot update client without payment type", CommonError, "Missing Input")
            Exit Sub
        End If

        Dim UpdateClient As New SqlCommand("UPDATE [dbo].[Clients] SET ClientName = @p1, Industry = @p2, Region = @p3, PaymentType = @p4 WHERE ClientID = @p0", con)
        With UpdateClient.Parameters
            .AddWithValue("@p0", ClientNameList.SelectedValue)
            .AddWithValue("@p1", ClientName.Text)
            .AddWithValue("@p2", ClientIndustry.Text)
            .AddWithValue("@p3", ClientRegion.SelectedValue)
            .AddWithValue("@p4", ClientPaymentType.SelectedValue)
        End With
        Try
            If con.State = ConnectionState.Closed Then con.Open()
            UpdateClient.ExecuteNonQuery()
            Call LoadLists()
            Call FillDataTables()
            Call ClearData()
        Catch ex As Exception
            Response.Write(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub
    Protected Sub UpdateTrucker_Click(sender As Object, e As EventArgs) Handles UpdateTrucker.Click
        If TruckerNameList.SelectedIndex <= 0 Then
            MsgBox("Select a trucker to make updates", CommonError, "Missing Input")
            Exit Sub
        End If
        If TruckerName.Text = Nothing Then
            MsgBox("Cannot update trucker without a name", CommonError, "Missing Input")
            Exit Sub
        End If
        If TruckerRegion.SelectedIndex <= 0 Then
            MsgBox("Include the region to update trucker", CommonError, "Missing Input")
            Exit Sub
        End If
        If TruckSpeciality.SelectedIndex <= 0 Then
            MsgBox("Include a truck speciality to update trucker", CommonError, "Missing Input")
            Exit Sub
        End If
        If TruckerExperience.Text = Nothing Then
            MsgBox("Input experience to update trucker", CommonError, "Missing Input")
            Exit Sub
        End If
        If TruckerExperience.Text < 0 Then
            MsgBox("Cannot have negative experience", CommonError, "Faulty Input")
            Exit Sub
        End If
        If EmploymentType.SelectedIndex <= 0 Then
            MsgBox("Select employment type to update trucker", CommonError, "Missing Input")
            Exit Sub
        End If
        If EmploymentStatus.SelectedIndex <= 0 Then
            MsgBox("Select employment status to update trucker", CommonError, "Missing Input")
            Exit Sub
        End If

        Dim UpdateTrucker As New SqlCommand("UPDATE [dbo].[Truckers] SET TruckerName = @p1, TruckID = @p2, Region = @p3, EmploymentType = @p4, EmploymentStatus = @p5 WHERE TruckerID = @p0", con)
        With UpdateTrucker.Parameters
            .AddWithValue("@p0", TruckerNameList.SelectedValue)
            .AddWithValue("@p1", TruckerName.Text)
            .AddWithValue("@p3", TruckerRegion.SelectedValue)
            .AddWithValue("@p2", TruckSpeciality.SelectedValue)
            .AddWithValue("@p4", EmploymentType.SelectedValue)
            .AddWithValue("@p5", EmploymentStatus.SelectedValue)
        End With
        Try
            If con.State = ConnectionState.Closed Then con.Open()
            UpdateTrucker.ExecuteNonQuery()
            Call LoadLists()
            Call FillDataTables()
            Call ClearData()
        Catch ex As Exception
            Response.Write(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub
    Protected Sub UpdateRoute_Click(sender As Object, e As EventArgs) Handles UpdateRoute.Click
        If RouteNameList.SelectedIndex <= 0 Then
            MsgBox("Select a route to update", CommonError, "Missing Input")
            Exit Sub
        End If
        If RouteName.Text = Nothing Then
            MsgBox("Cannot update route without a name", CommonError, "Missing Input")
            Exit Sub
        End If
        If OriginState.Text = Nothing OrElse OriginCity.Text = Nothing Then
            MsgBox("Input the orgin location to update route", CommonError, "Missing Input")
            Exit Sub
        End If
        If DestinationState.Text = Nothing OrElse DestinationCity.Text = Nothing Then
            MsgBox("Input the destination location to update route", CommonError, "Missing Input")
            Exit Sub
        End If
        If RouteDistance.Text = Nothing Then
            MsgBox("Cannot update route without the distance", CommonError, "Missing Input")
            Exit Sub
        End If
        If RouteDistance.Text <= 0 Then
            MsgBox("Cannot have less than 1 mile", CommonError, "Faulty Input")
            Exit Sub
        End If

        Dim Distance As Decimal
        If DistanceConversion.SelectedIndex = 1 Then
            Distance = RouteDistance.Text * 0.62
        Else
            Distance = RouteDistance.Text
        End If

        Dim UpdateRoute As New SqlCommand("UPDATE [dbo].[Routes] SET RouteName = @p1, OriginState = @p2, OriginCity = @p3, DestinationState = @p4, DestinationCity = @p5, Miles = @p6 WHERE RouteID = @p0", con)
        With UpdateRoute.Parameters
            .AddWithValue("@p0", RouteNameList.SelectedValue)
            .AddWithValue("@p1", RouteName.Text)
            .AddWithValue("@p2", OriginState.Text)
            .AddWithValue("@p4", DestinationState.Text)
            .AddWithValue("@p3", OriginCity.Text)
            .AddWithValue("@p5", DestinationCity.Text)
            .AddWithValue("@p6", Distance)
        End With
        Try
            If con.State = ConnectionState.Closed Then con.Open()
            UpdateRoute.ExecuteNonQuery()
            Call LoadLists()
            Call FillDataTables()
            Call ClearData()
        Catch ex As Exception
            Response.Write(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub
    Protected Sub UpdateTruck_Click(sender As Object, e As EventArgs) Handles UpdateTruck.Click
        If TruckTypeList.SelectedIndex <= 0 Then
            MsgBox("Select a truck to make updates", CommonError, "Missing Input")
            Exit Sub
        End If
        If TruckTypeTxt.Text = Nothing Then
            MsgBox("Cannot update truck without a type/name", CommonError, "Missing Input")
            Exit Sub
        End If
        If TruckMileage.Text = Nothing Then
            MsgBox("Include the mileage to update truck", CommonError, "Missing Input")
            Exit Sub
        End If
        If TruckMileage.Text < 0 Then
            MsgBox("Cannot have negative mileage", CommonError, "Faulty Data")
            Exit Sub
        End If
        If TruckLease.Text = Nothing Then
            MsgBox("Input lease amount or cost to update truck", CommonError, "Missing Input")
            Exit Sub
        End If
        If TruckLease.Text < 0 Then
            MsgBox("Cannot have negative input", CommonError, "Faulty Input")
            Exit Sub
        End If
        If TruckPayload.Text = Nothing Then
            MsgBox("For safety reasons, the payload rating is required", CommonError, "Missing Input")
            Exit Sub
        End If
        If TruckPayload.Text < 0 Then
            MsgBox("Cannot have negative input", CommonError, "Faulty Input")
            Exit Sub
        End If

        If (TruckPayload.Text < 1000 And PayloadConversion.SelectedIndex = 0) Or (TruckPayload.Text < 0.5 And PayloadConversion.SelectedIndex = 1) Then
            Dim result As MsgBoxResult = MsgBox("Is " & TruckPayload.Text & " the correct payload, this truck won't be able to carry many deliveries", VerifyInput, "Input Verification")
            If result = MsgBoxResult.No Then Exit Sub
        End If
        Dim Payload As Decimal
        If PayloadConversion.SelectedIndex = 1 Then
            Payload = 2000 * TruckPayload.Text
        Else
            Payload = TruckPayload.Text
        End If

        Dim UpdateTruck As New SqlCommand("UPDATE [dbo].[Trucks] SET Type = @p1, Mileage = @p2, LeaseAmount = @p3, PayloadRating = @p5 WHERE TruckID = @p0", con)
        With UpdateTruck.Parameters
            .AddWithValue("@p0", TruckTypeList.SelectedValue)
            .AddWithValue("@p1", TruckTypeTxt.Text)
            .AddWithValue("@p2", TruckMileage.Text)
            .AddWithValue("@p3", TruckLease.Text)
            .AddWithValue("@p5", Payload)
        End With
        Try
            If con.State = ConnectionState.Closed Then con.Open()
            UpdateTruck.ExecuteNonQuery()
            Call LoadLists()
            Call FillDataTables()
            Call ClearData()
        Catch ex As Exception
            Response.Write(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Protected Sub DeleteClient_Click(sender As Object, e As EventArgs) Handles DeleteClient.Click
        If ClientNameList.SelectedIndex <= 0 Then
            Response.Write("Select a client to delete")
            Exit Sub
        End If

        ClientDeletetxt.Visible = True
        ClientYes.Visible = True
        ClientNo.Visible = True
        ClientDeletetxt.Text = "Are you sure you want to delete this client, all data associated with this value will also be deleted, this process cannot be undone"

    End Sub
    Protected Sub DeleteRoute_Click(sender As Object, e As EventArgs) Handles DeleteRoute.Click
        If RouteNameList.SelectedIndex <= 0 Then
            Response.Write("Select a route to delete")
            Exit Sub
        End If

        RouteDeletetxt.Visible = True
        RouteYes.Visible = True
        RouteNo.Visible = True
        RouteDeletetxt.Text = "Are you sure you want to delete this route, all data associated with this value will also be deleted, this process cannot be undone"

    End Sub
    'Deletes all values associated with this trucker. This works by enabling cascade deletes in the properties of all primary to foreign key relations in SQL Server
    Protected Sub DeleteTrucker_Click(sender As Object, e As EventArgs) Handles DeleteTrucker.Click
        If TruckerNameList.SelectedIndex <= 0 Then
            Response.Write("Select a trucker to delete")
            Exit Sub
        End If

        truckerDeletetxt.Visible = True
        TruckerYes.Visible = True
        TruckerNo.Visible = True
        'Suggests changing trucker status instead. Only active truckers are filled into the contracts and expense drop down lists
        truckerDeletetxt.Text = "Are you sure you want to delete this trucker, all data associated with this value will also be deleted, this process cannot be undone. Consider changing employment status instead"

    End Sub
    Protected Sub DeleteTruck_Click(sender As Object, e As EventArgs) Handles DeleteTruck.Click
        If TruckTypeList.SelectedIndex <= 0 Then
            Response.Write("Select a truck to delete")
            Exit Sub
        End If

        TruckDeletetxt.Visible = True
        TruckYes.Visible = True
        TruckNo.Visible = True
        TruckDeletetxt.Text = "Are you sure you want to delete this truck, all data associated with this value will also be deleted, this process cannot be undone."
    End Sub
#End Region

#Region "Clear Forms"
    Public Sub ClearData()
        ContractTrucker.SelectedIndex = 0
        ContractClient.SelectedIndex = 0
        ContractRoute.SelectedIndex = 0
        ContractTruck.SelectedIndex = 0
        ContractProduct.Text = Nothing
        PaymentTypeList.SelectedIndex = -1
        ContractDate.Text = Nothing
        ContractMileage.Text = Nothing
        ContractWeight.Text = Nothing
        WeightConversion.SelectedIndex = 0
        ContractCost.Text = Nothing
        ExpenseComments.Text = Nothing
        ExpenseCost.Text = Nothing
        ExpenseTrucker.SelectedIndex = 0
        ExpenseTruck.SelectedIndex = 0
        ExpenseRoute.SelectedIndex = 0
        ExpenseType.SelectedIndex = 0
        ExpenseDate.Text = Nothing
        ClientNameList.SelectedIndex = 0
        ClientName.Text = Nothing
        ClientIndustry.Text = Nothing
        ClientRegion.SelectedIndex = 0
        ClientPaymentType.SelectedIndex = -1
        RouteNameList.SelectedIndex = 0
        RouteName.Text = Nothing
        OriginState.Text = Nothing
        OriginCity.Text = Nothing
        DestinationState.Text = Nothing
        DestinationCity.Text = Nothing
        RouteDistance.Text = Nothing
        DistanceConversion.SelectedIndex = 0
        TruckerNameList.SelectedIndex = 0
        TruckerName.Text = Nothing
        TruckerRegion.SelectedIndex = 0
        TruckSpeciality.SelectedIndex = 0
        TruckerExperience.Text = Nothing
        EmploymentType.SelectedIndex = 0
        EmploymentStatus.SelectedIndex = 0
        TruckTypeList.SelectedIndex = 0
        TruckTypeTxt.Text = Nothing
        TruckMileage.Text = Nothing
        TruckLease.Text = Nothing
        TruckPayload.Text = Nothing
        Call HideErrorSelection()
        Call FillDataTables()
    End Sub
    Protected Sub ClearClient_Click(sender As Object, e As EventArgs) Handles ClearClient.Click
        Call ClearData()
    End Sub
    Protected Sub ClearRoute_Click(sender As Object, e As EventArgs) Handles ClearRoute.Click
        Call ClearData()
    End Sub
    Protected Sub ClearTrucker_Click(sender As Object, e As EventArgs) Handles ClearTrucker.Click
        Call ClearData()
    End Sub
    Protected Sub ClearTruck_Click(sender As Object, e As EventArgs) Handles ClearTruck.Click
        Call ClearData()
    End Sub
    'ClearData is called after every action/calculation. Because we don't want the invoice information to be delete too, The clear button was added to clear the textbox when the user needs to
    Protected Sub ClearContracts_Click(sender As Object, e As EventArgs) Handles ClearContracts.Click
        Invoice.Text = Nothing
        Call ClearData()
    End Sub
#End Region

#Region "Delete Selection"
    Public Sub HideErrorSelection()
        ClientDeletetxt.Visible = False
        ClientYes.Visible = False
        ClientNo.Visible = False
        RouteDeletetxt.Visible = False
        RouteYes.Visible = False
        RouteNo.Visible = False
        truckerDeletetxt.Visible = False
        TruckerYes.Visible = False
        TruckerNo.Visible = False
        TruckDeletetxt.Visible = False
        TruckYes.Visible = False
        TruckNo.Visible = False
    End Sub
    Protected Sub ClientYes_Click(sender As Object, e As EventArgs) Handles ClientYes.Click
        If ClientNameList.SelectedIndex <= 0 Then
            Response.Write("Select a client to delete")
            Exit Sub
        End If

        Dim DeleteClient As New SqlCommand("DELETE FROM [dbo].[Clients] WHERE ClientID = @p1", con)
        With DeleteClient.Parameters
            .Clear()
            .AddWithValue("@p1", ClientNameList.SelectedValue)
        End With
        Try
            If con.State = ConnectionState.Closed Then con.Open()
            DeleteClient.ExecuteNonQuery()
            Call LoadLists()
            Call FillDataTables()
            Call ClearData()
        Catch ex As Exception
            Response.Write(ex.Message)
        Finally
            con.Close()
        End Try
        Call HideErrorSelection()
    End Sub
    Protected Sub RouteYes_Click(sender As Object, e As EventArgs) Handles RouteYes.Click
        If RouteNameList.SelectedIndex <= 0 Then
            Response.Write("Select a route to delete")
            Exit Sub
        End If

        Dim DeleteRoute As New SqlCommand("DELETE FROM [dbo].[Routes] WHERE RouteID = @p1", con)
        With DeleteRoute.Parameters
            .Clear()
            .AddWithValue("@p1", RouteNameList.SelectedValue)
        End With
        Try
            If con.State = ConnectionState.Closed Then con.Open()
            DeleteRoute.ExecuteNonQuery()
            Call LoadLists()
            Call FillDataTables()
            Call ClearData()
        Catch ex As Exception
            Response.Write(ex.Message)
        Finally
            con.Close()
        End Try
        Call HideErrorSelection()
    End Sub
    Protected Sub TruckerYes_Click(sender As Object, e As EventArgs) Handles TruckerYes.Click
        If TruckerNameList.SelectedIndex <= 0 Then
            Response.Write("Select a trucker to delete")
            Exit Sub
        End If

        Dim DeleteTrucker As New SqlCommand("DELETE FROM [dbo].[Truckers] WHERE TruckerID = @p1", con)
        With DeleteTrucker.Parameters
            .Clear()
            .AddWithValue("@p1", TruckerNameList.SelectedValue)
        End With
        Try
            If con.State = ConnectionState.Closed Then con.Open()
            DeleteTrucker.ExecuteNonQuery()
            Call LoadLists()
            Call FillDataTables()
            Call ClearData()
        Catch ex As Exception
            Response.Write(ex.Message)
        Finally
            con.Close()
        End Try
        Call HideErrorSelection()
        Call HideErrorSelection()
    End Sub
    Protected Sub TruckYes_Click(sender As Object, e As EventArgs) Handles TruckYes.Click
        If TruckTypeList.SelectedIndex <= 0 Then
            Response.Write("Select a truck to delete")
            Exit Sub
        End If

        Dim DeleteTruck As New SqlCommand("DELETE FROM [dbo].[Trucks] WHERE TruckID = @p1", con)
        With DeleteTruck.Parameters
            .Clear()
            .AddWithValue("@p1", TruckTypeList.SelectedValue)
        End With
        Try
            If con.State = ConnectionState.Closed Then con.Open()
            DeleteTruck.ExecuteNonQuery()
            Call LoadLists()
            Call FillDataTables()
            Call ClearData()
        Catch ex As Exception
            Response.Write(ex.Message)
        Finally
            con.Close()
        End Try
        Call HideErrorSelection()
    End Sub

    Protected Sub ClientNo_Click(sender As Object, e As EventArgs) Handles ClientNo.Click
        Call ClearData()
    End Sub
    Protected Sub RouteNo_Click(sender As Object, e As EventArgs) Handles RouteNo.Click
        Call ClearData()
    End Sub
    Protected Sub TruckerNo_Click(sender As Object, e As EventArgs) Handles TruckerNo.Click
        Call ClearData()
    End Sub
    Protected Sub TruckNo_Click(sender As Object, e As EventArgs) Handles TruckNo.Click
        Call ClearData()
    End Sub
#End Region

End Class
