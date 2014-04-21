Imports System.Data
Imports System.Configuration
Imports System.Data.SqlServerCe

Public Class Database
    Friend dbConnection As SqlCeConnection
    Friend dbCommand As SqlCeCommand
    Friend dbReader As SqlCeDataReader
    Friend dbResult As String
    Friend strConnectionString As String
    ' Access the connection string by referencing the app.config like this:
    ' ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString


    ' FUNCTION TO ADD NEW PROJECTS
    Friend Shared Function InsertNewProject(TeamId As Decimal, _
                                CustomerId As Decimal, _
                                ProjectDate As String, _
                                ProjectDesc As String) As Integer
        Dim rowsAffected As Integer
        Dim sql As String = "Insert into PRODUCTS " & _
            "VALUES (@teamid, @customerid, @projectDate, @ProjectDescription)"
        Try
            Using conn As New SqlCeConnection(ConfigurationManager.ConnectionStrings("SensationalDBConnectionString").ConnectionString)
                conn.Open()
                Using cmd As New SqlCeCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@ProductID", TeamId)
                    cmd.Parameters.AddWithValue("@ProductName", CustomerId)
                    cmd.Parameters.AddWithValue("@ProductDesc", ProjectDate)
                    cmd.Parameters.AddWithValue("@ProductPrice", ProjectDesc)
                    rowsAffected = cmd.ExecuteNonQuery()

                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error", ex.Message)
        End Try
        Return rowsAffected
    End Function
    ' This function returns the employee ID for a new employee. It increments by one, as that is what the DB does upon insert.
    Friend Shared Function GetNewEmployeeID() As Integer
        Dim employeeID As Integer
        ' Get the max number from the database
        Dim sql As String = "select max(empID) from EMPLOYEE"
        Try
            Using conn As New SqlCeConnection(ConfigurationManager.ConnectionStrings("SensationalDBConnectionString").ConnectionString)
                conn.Open()
                Using cmd As New SqlCeCommand(sql, conn)

                    employeeID = cmd.ExecuteScalar()

                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error", ex.Message)
        End Try
        ' Increment the max number by 1
        Return employeeID + 1
    End Function
    ' Finds a product ID
    Friend Shared Function GetProductID(ByVal productID As Double) As Double
        Dim result As Double
        ' Get the max number from the database
        Dim sql As String = "select productID from PRODUCTS where productID = " & productID
        Try
            Using conn As New SqlCeConnection(ConfigurationManager.ConnectionStrings("SensationalDBConnectionString").ConnectionString)
                conn.Open()
                Using cmd As New SqlCeCommand(sql, conn)

                    result = cmd.ExecuteScalar()

                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error", ex.Message)
        End Try
        Return productID
    End Function
    ' Sub for editing an inventory item record
    Friend Shared Sub EditItemRecord(productid As Int32, _
                         productName As String, _
                         productDesc As String, _
                         productPrice As Decimal, _
                         inventoryLevel As String, _
                         catagory As String, _
                         active As String)
        Dim rowsAffected As Integer
        'UPDATE <TABLENAME>
        'SET <COLUMNNAME> = V
        'WHERE <PRIMARYKEY> = X
        Dim sql As String = "UPDATE PRODUCTS " & _
            "SET productName = '" & productName & _
            "', productDesc = '" & productDesc & _
            "', productprice = '" & productPrice.ToString() & _
            "', inventorylevel = '" & inventoryLevel & _
            "', catagory = '" & catagory & _
            "', active = '" & active & _
            "' WHERE productId = '" & productid & "'"
        Try
            Using conn As New SqlCeConnection(ConfigurationManager.ConnectionStrings("SensationalDBConnectionString").ConnectionString)
                conn.Open()
                Using cmd As New SqlCeCommand(sql, conn)
                    rowsAffected = cmd.ExecuteNonQuery()
                End Using
            End Using
            MessageBox.Show("All changes made to this item have been saved in the Database" & " ROWS : " & rowsAffected)
        Catch ex As Exception
            MessageBox.Show("Error", ex.Message)
        End Try
    End Sub


    'For searching anything in the database and listing data in datagrid
    Friend Shared Sub SearchRecords(s As String, c As String, db As String, dg As DataGridView)

        Dim sql As String = "SELECT * FROM " & db & " WHERE " & c & " LIKE '%" & s & "%'"
        Dim sda As New SqlCeDataAdapter()
        Dim ds As New SensationalDBDataSet

        Try
            Using conn As New SqlCeConnection(ConfigurationManager.ConnectionStrings("SensationalDBConnectionString").ConnectionString)
                Using cmd As New SqlCeCommand(sql, conn)
                    sda.SelectCommand = New SqlCeCommand()
                    sda.SelectCommand.Connection = conn
                    sda.SelectCommand.CommandText = sql
                    sda.SelectCommand.CommandType = CommandType.Text
                    conn.Open()
                    sda.Fill(ds, db)
                    dg.DataSource = ds
                    dg.DataMember = db
                    conn.Close()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub
    ' lists all data in the table and fill data grid with data
    Friend Shared Sub SearchRecords(db, dg)
        Dim sql As String = "SELECT * FROM " + db
        Dim sda As New SqlCeDataAdapter()
        Dim ds As New SensationalDBDataSet
        Try
            Using conn As New SqlCeConnection(ConfigurationManager.ConnectionStrings("SensationalDBConnectionString").ConnectionString)
                Using cmd As New SqlCeCommand(sql, conn)
                    sda.SelectCommand = New SqlCeCommand()
                    sda.SelectCommand.Connection = conn
                    sda.SelectCommand.CommandText = sql
                    sda.SelectCommand.CommandType = CommandType.Text
                    conn.Open()
                    sda.Fill(ds, db)
                    dg.DataSource = ds
                    dg.DataMember = db
                    conn.Close()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub
    ' deletes records from table taken in argument.
    Friend Shared Sub DeleteRecords(s As String, db As String, dg As DataGridView, pk As String)

        Dim sql As String = "DELETE FROM " & db & " WHERE " & pk & " = '" & s & "'"
        Dim sda As New SqlCeDataAdapter()
        Dim ds As New SensationalDBDataSet
        Dim DataGridView As New DataGridView
        DataGridView = dg
        If MessageBox.Show("Do you really want to Delete record " + pk + " ?", "Delete", _
                           MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then
            MsgBox("Operation Cancelled")
            Exit Sub
        End If
        Try
            Using conn As New SqlCeConnection(ConfigurationManager.ConnectionStrings("SensationalDBConnectionString").ConnectionString)
                Using cmd As New SqlCeCommand(sql, conn)
                    sda.SelectCommand = New SqlCeCommand()
                    sda.SelectCommand.Connection = conn
                    sda.SelectCommand.CommandText = sql
                    sda.SelectCommand.CommandType = CommandType.Text
                    conn.Open()
                    sda.Fill(ds, db)
                    conn.Close()
                    DataGridView.DataSource = ds
                    DataGridView.DataMember = db
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try

    End Sub
    ' Returns SaleID
    Friend Shared Sub InsertSaleRecord(intSaleID As Integer, _
                                       grdCashier As DataGridView, _
                                       strTime As String, _
                                        intEmpID As Integer, _
                                        decSaleTax As Decimal, _
                                        decSaleTotal As Decimal)
        ' For SALES table
        'INSERT INTO [dbo].[SALES]
        '   ([timestamp]
        '   ,[empID]
        '   ,[sale_tax]
        '   ,[sale_total])
        'VALUES()
        '   (<timestamp, datetime2(7),>
        '   ,<empID, int,>
        '   ,<sale_tax, decimal(18,2),>
        '   ,<sale_total, decimal(18,2),>)
        ' For SALES_ITEMS table
        ' FOR SALES 
        Dim sql As String = "Insert into SALES (timestamp, empID, sale_tax, sale_total) " & _
            "VALUES (@TimePaid, @EmpID, @SaleTax, @SaleTotal);" '& "Select Scope_Identity()"
        Try
            Using conn As New SqlCeConnection(ConfigurationManager.ConnectionStrings("SensationalDBConnectionString").ConnectionString)
                conn.Open()
                Using cmd As New SqlCeCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@TimePaid", strTime)
                    cmd.Parameters.AddWithValue("@EmpID", intEmpID)
                    cmd.Parameters.AddWithValue("@SaleTax", decSaleTax)
                    cmd.Parameters.AddWithValue("@SaleTotal", decSaleTotal)
                    cmd.ExecuteScalar()
                    Debug.WriteLine("InsertSalesItems " & cmd.CommandText.ToString())
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Debug.WriteLine(ex.Message)
        End Try
        ' FOR SALES_ITEMS
        ' 1) Loop through all items and record their IDs, Quantity, and Price
        ' Price needs recorded as prices may change over time for specific items
        Try
            Dim decProductIDs(10000) As Decimal
            Dim decPriceEa(10000) As Decimal
            ' Dim intQuantity() As Integer
            For index As Integer = 0 To grdCashier.RowCount - 1
                decProductIDs(index) = Convert.ToDecimal(grdCashier.Rows(index).Cells(0).Value)
                decPriceEa(index) = Convert.ToDecimal(grdCashier.Rows(index).Cells(3).Value)
                ' TODO: Add a quantity field and record it
                ' intQuantity(index) = Convert.ToInt32(grdSaleDataGrid.Rows(index).Cells(0).Value)
                ' intQty += Convert.ToInteger(grdCashier.Rows(index).Cells(2).Value)
                ' MessageBox.Show(intProductIDs(index))
                InsertSalesItems(intSaleID, decProductIDs(index), 1, decPriceEa(index))
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    'INSERT INTO [dbo].[SALES_ITEMS]
    '           ([saleID]
    '           ,[productID]
    '           ,[quantity]
    '           ,[price_each])
    '        VALUES()
    '           (<saleID, int,>
    '           ,<productID, bigint,>
    '           ,<quantity, int,>
    '           ,<price_each, decimal(18,2),>)
    Friend Shared Sub InsertSalesItems(decSaleID As Decimal, _
                                            decProductID As Decimal, _
                                            intQuantity As Integer, _
                                            intPriceEa As Decimal)
        Dim sql As String = "Insert into SALES_ITEMS (saleID, productID, quantity, price_each) " & _
            "VALUES (@SaleID, @ProductID, @Quantity, @PriceEA);"
        Try
            Using conn As New SqlCeConnection(ConfigurationManager.ConnectionStrings("SensationalDBConnectionString").ConnectionString)
                conn.Open()
                Using cmd As New SqlCeCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@SaleID", decSaleID)
                    cmd.Parameters.AddWithValue("@ProductID", decProductID)
                    cmd.Parameters.AddWithValue("@Quantity", intQuantity)
                    cmd.Parameters.AddWithValue("@PriceEA", intPriceEa)
                    cmd.ExecuteNonQuery()
                    Debug.WriteLine("InsertSalesItems " & cmd.CommandText.ToString())
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
            Debug.WriteLine(ex.Message)
        End Try

    End Sub

End Class
