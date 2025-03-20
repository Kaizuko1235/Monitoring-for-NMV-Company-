Imports System.Data.SqlClient
Public Class Form1
    Dim connectionString As String = "Data Source=MICHAELPC\SQLEXPRESS;Initial Catalog=NMVDB;Integrated Security=True;Encrypt=False;"

    Private Sub LoadData()
        Try
            Using con As New SqlConnection(connectionString)
                Dim query As String = "SELECT * FROM [ Monitoringtb]"
                Dim adapter As New SqlDataAdapter(query, con)
                Dim dt As New DataTable()
                adapter.Fill(dt)
                DgvResults.DataSource = dt
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading data: " & ex.Message)
        End Try
    End Sub

    Private Sub Btninsert_Click(sender As Object, e As EventArgs) Handles Btninsert.Click
        Try
            Using con As New SqlConnection(connectionString)
                Using cmd As New SqlCommand("INSERT INTO [ Monitoringtb] ([DATE OF RECEIPT], [CLIENT'S NAME], [NATURE OF PROJECT], [TITLE No], [TITLE OWNER], [TITLE LOCATION], [LOT SIZE], [TAX DECLARATION No], [TAX DECLARATION OWNER], [TAX DECLOCATION LOCATION], [TECHNICAL DESCRIPTION], IMPROVEMENT, ENCUMB, STATUS, [LOCATION OF DOCUMENT], No) 
                                         VALUES (@DateOfReceipt, @ClientsName, @NatureOfProject, @TitleNo, @TitleOwner, @TitleLocation, @LotSize, @TaxDeclNo, @TaxDeclOwner, @TaxDeclLocation, @TechDesc, @Improvement, @Encumb, @Status, @LocationOfDoc, @No)", con)


                    cmd.Parameters.AddWithValue("@No", If(IsNumeric(TxtNo.Text) AndAlso TxtNo.Text <> "", CInt(TxtNo.Text), DBNull.Value))
                    cmd.Parameters.AddWithValue("@LotSize", If(IsNumeric(TxtLotsize.Text) AndAlso TxtLotsize.Text <> "", CDec(TxtLotsize.Text), DBNull.Value))

                    ' Add other parameters as strings
                    cmd.Parameters.AddWithValue("@DateOfReceipt", TxtDateofreceipt.Text)
                    cmd.Parameters.AddWithValue("@ClientsName", TxtClientsname.Text)
                    cmd.Parameters.AddWithValue("@NatureOfProject", TxtNatureofproject.Text)
                    cmd.Parameters.AddWithValue("@TitleNo", TxtTitleno.Text)
                    cmd.Parameters.AddWithValue("@TitleOwner", TxtTitleowner.Text)
                    cmd.Parameters.AddWithValue("@TitleLocation", TxtTitlelocation.Text)
                    cmd.Parameters.AddWithValue("@TaxDeclNo", Txttaxdeclarationlocation.Text)
                    cmd.Parameters.AddWithValue("@TaxDeclOwner", TxtTaxdeclarationowner.Text)
                    cmd.Parameters.AddWithValue("@TaxDeclLocation", Txttaxdeclarationlocation.Text)
                    cmd.Parameters.AddWithValue("@TechDesc", If(Cmbbtechnicaldescription.SelectedItem IsNot Nothing, Cmbbtechnicaldescription.SelectedItem.ToString(), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Improvement", Txtimprovement.Text)
                    cmd.Parameters.AddWithValue("@Encumb", If(Cmbbencumb.SelectedItem IsNot Nothing, Cmbbencumb.SelectedItem.ToString(), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Status", Txtstatus.Text)
                    cmd.Parameters.AddWithValue("@LocationOfDoc", TxtLocationOfDocument.Text)

                    con.Open()
                    cmd.ExecuteNonQuery()
                    MessageBox.Show("Record inserted successfully.")
                    LoadData()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error inserting record: " & ex.Message)
        End Try
    End Sub

    Private Sub Btndelete_Click(sender As Object, e As EventArgs) Handles Btndelete.Click
        If DgvResults.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a record to delete.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim selectedID As Integer
        Try
            selectedID = Convert.ToInt32(DgvResults.SelectedRows(0).Cells("ID").Value)
        Catch ex As Exception
            MessageBox.Show("Error retrieving the selected record ID: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

        Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete this record?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If result = DialogResult.Yes Then
            Try
                Using con As New SqlConnection(connectionString)
                    Using cmd As New SqlCommand("DELETE FROM [ Monitoringtb] WHERE ID = @ID", con)
                        cmd.Parameters.AddWithValue("@ID", selectedID)

                        con.Open()
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                        If rowsAffected > 0 Then
                            MessageBox.Show("Record deleted successfully.")
                            LoadData()
                        Else
                            MessageBox.Show("No matching record found to delete.")
                        End If
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("Error deleting record: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadData()
    End Sub

    Private Sub BtnLogout_Click(sender As Object, e As EventArgs) Handles BtnLogout.Click
        Dim confirmLogout As DialogResult = MessageBox.Show("Are you sure you want to logout?", "Logout", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If confirmLogout = DialogResult.Yes Then
            Application.Restart()
        End If
    End Sub

    Private Sub Btnsearch_Click(sender As Object, e As EventArgs) Handles Btnsearch.Click
        Dim searchTerm As String = Txtsearch.Text.Trim()
        If searchTerm = "" Then
            MessageBox.Show("Please enter a search term.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Try
            Using con As New SqlConnection(connectionString)
                Dim query As String = "SELECT * FROM [ Monitoringtb] WHERE 
                [CLIENT'S NAME] LIKE @Search OR 
                [TITLE No] LIKE @Search OR 
                [TITLE OWNER] LIKE @Search OR 
                [TITLE LOCATION] LIKE @Search OR 
                [TAX DECLARATION No] LIKE @Search OR 
                [TAX DECLARATION OWNER] LIKE @Search OR 
                [TAX DECLOCATION LOCATION] LIKE @Search OR 
                [TECHNICAL DESCRIPTION] LIKE @Search OR 
                IMPROVEMENT LIKE @Search OR 
                ENCUMB LIKE @Search OR 
                STATUS LIKE @Search OR 
                [LOCATION OF DOCUMENT] LIKE @Search OR 
                CAST(No AS VARCHAR) LIKE @Search"

                Using cmd As New SqlCommand(query, con)
                    cmd.Parameters.AddWithValue("@Search", "%" & searchTerm & "%")
                    Dim adapter As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    adapter.Fill(dt)
                    DgvResults.DataSource = dt

                    If dt.Rows.Count = 0 Then
                        MessageBox.Show("No records found matching your search.", "Search Results", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error searching records: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles BtnClear.Click
        LoadData()
        Txtsearch.Clear()
    End Sub

    Private Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnEdit.Click

        If DgvResults.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a record to edit.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim selectedID As Integer
        Try
            selectedID = Convert.ToInt32(DgvResults.SelectedRows(0).Cells("ID").Value)
        Catch ex As Exception
            MessageBox.Show("Error retrieving the selected record ID: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

        Try
            Using con As New SqlConnection(connectionString)
                Using cmd As New SqlCommand("UPDATE [ Monitoringtb] 
                SET [DATE OF RECEIPT] = @DateOfReceipt, [CLIENT'S NAME] = @ClientsName, 
                    [NATURE OF PROJECT] = @NatureOfProject, [TITLE No] = @TitleNo, 
                    [TITLE OWNER] = @TitleOwner, [TITLE LOCATION] = @TitleLocation, 
                    [LOT SIZE] = @LotSize, [TAX DECLARATION No] = @TaxDeclNo, 
                    [TAX DECLARATION OWNER] = @TaxDeclOwner, [TAX DECLOCATION LOCATION] = @TaxDeclLocation, 
                    [TECHNICAL DESCRIPTION] = @TechDesc, IMPROVEMENT = @Improvement, 
                    ENCUMB = @Encumb, STATUS = @Status, [LOCATION OF DOCUMENT] = @LocationOfDoc, No = @No 
                WHERE ID = @ID", con)

                    cmd.Parameters.AddWithValue("@DateOfReceipt", TxtDateofreceipt.Text)
                    cmd.Parameters.AddWithValue("@ClientsName", TxtClientsname.Text)
                    cmd.Parameters.AddWithValue("@NatureOfProject", TxtNatureofproject.Text)
                    cmd.Parameters.AddWithValue("@TitleNo", TxtTitleno.Text)
                    cmd.Parameters.AddWithValue("@TitleOwner", TxtTitleowner.Text)
                    cmd.Parameters.AddWithValue("@TitleLocation", TxtTitlelocation.Text)
                    cmd.Parameters.AddWithValue("@LotSize", If(IsNumeric(TxtLotsize.Text) AndAlso TxtLotsize.Text <> "", CDec(TxtLotsize.Text), DBNull.Value))
                    cmd.Parameters.AddWithValue("@TaxDeclNo", Txttaxdeclarationlocation.Text)
                    cmd.Parameters.AddWithValue("@TaxDeclOwner", TxtTaxdeclarationowner.Text)
                    cmd.Parameters.AddWithValue("@TaxDeclLocation", Txttaxdeclarationlocation.Text)
                    cmd.Parameters.AddWithValue("@TechDesc", If(Cmbbtechnicaldescription.SelectedItem IsNot Nothing, Cmbbtechnicaldescription.SelectedItem.ToString(), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Improvement", Txtimprovement.Text)
                    cmd.Parameters.AddWithValue("@Encumb", If(Cmbbencumb.SelectedItem IsNot Nothing, Cmbbencumb.SelectedItem.ToString(), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Status", Txtstatus.Text)
                    cmd.Parameters.AddWithValue("@LocationOfDoc", TxtLocationOfDocument.Text)
                    cmd.Parameters.AddWithValue("@No", If(IsNumeric(TxtNo.Text) AndAlso TxtNo.Text <> "", CInt(TxtNo.Text), DBNull.Value))
                    cmd.Parameters.AddWithValue("@ID", selectedID)

                    con.Open()
                    Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                    If rowsAffected > 0 Then
                        MessageBox.Show("Record updated successfully.")
                        LoadData()
                    Else
                        MessageBox.Show("No record was updated. Please check your input.")
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error updating record: " & ex.Message)
        End Try

    End Sub

End Class
