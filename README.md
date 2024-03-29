# CRUD_Purchase_System_Shp_Method
CRUD WEB Develop by .NET Framework DB Adventure ER Database table  Simple Purchasing.shipMethod
Imports System.Data.SqlClient

Public Class Ship_Method_CRUD
    Inherits System.Web.UI.Page

    Dim connectionString As String = "Data Source=DESKTOP-LJ58AB0\SQLEXPRESS;Initial Catalog=AdventureWorks2008R2;Integrated Security=True;"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            MultiView1.ActiveViewIndex = 0
            '---Hide Component by MultiView
            btnNewRow.Visible = True
            btnCancel.Visible = False
            btnInsert.Visible = False
            btnUpdate.Visible = False
            btnDelete.Visible = False
            DV1.DataSourceID = ""
            DV1.DataBind()
        End If
    End Sub

    Protected Sub btnNewRow_Click(sender As Object, e As EventArgs) Handles btnNewRow.Click
        DV1.ChangeMode(DetailsViewMode.Insert)
        DV1.DataBind()

        MultiView1.ActiveViewIndex = 1
        '---Hide Component by MultiView
        btnNewRow.Visible = False
        btnCancel.Visible = True
        btnInsert.Visible = True
        btnUpdate.Visible = False
        btnDelete.Visible = False
    End Sub

    Protected Sub btnInsert_Click(sender As Object, e As EventArgs) Handles btnInsert.Click
        If Page.IsValid Then
            ' Get values from DetailsView
            Dim txt_shipmethodid As TextBox = CType(DV1.FindControl("txt_shipmethodid"), TextBox)
            Dim txt_name As TextBox = CType(DV1.FindControl("txt_Name"), TextBox)
            Dim txt_shipbase As TextBox = CType(DV1.FindControl("txt_Shipbase"), TextBox)
            Dim txt_shiprate As TextBox = CType(DV1.FindControl("txt_Shiprate"), TextBox)
            Dim txt_rowguid As TextBox = CType(DV1.FindControl("txt_Rowguid"), TextBox)
            Dim txt_modifieddate As TextBox = CType(DV1.FindControl("txt_Modifieddate"), TextBox)

            ' Create SQL query
            Dim strSQL As String = "SET IDENTITY_INSERT Purchasing.ShipMethod ON;" &
                       "INSERT INTO Purchasing.ShipMethod (ShipMethodID, Name, ShipBase, ShipRate, rowguid, ModifiedDate) " &
                       "VALUES (@ShipMethodID, @Name, @ShipBase, @ShipRate, @Rowguid, @ModifiedDate);" &
                       "SET IDENTITY_INSERT Purchasing.ShipMethod OFF;"


            ' Create connection and command objects
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(strSQL, connection)
                    ' Add parameters
                    command.Parameters.AddWithValue("@ShipMethodID", txt_shipmethodid.Text)
                    command.Parameters.AddWithValue("@Name", txt_name.Text)
                    command.Parameters.AddWithValue("@ShipBase", (txt_shipbase.Text))
                    command.Parameters.AddWithValue("@ShipRate", (txt_shiprate.Text))
                    command.Parameters.AddWithValue("@Rowguid", (txt_rowguid.Text))
                    command.Parameters.AddWithValue("@ModifiedDate", (txt_modifieddate.Text))

                    ' Open connection and execute command
                    connection.Open()
                    command.ExecuteNonQuery()


                End Using
            End Using
            GV1.DataBind()

            ' Clear DetailsView and refresh GridView
            DV1.ChangeMode(DetailsViewMode.ReadOnly)
            DV1.DataBind()
            MultiView1.ActiveViewIndex = 0
            '---Hide Component by MultiView

            btnNewRow.Visible = True
            btnCancel.Visible = False
            btnInsert.Visible = False
            btnUpdate.Visible = False
            btnDelete.Visible = False
        End If

    End Sub
    Protected Function GetColIdxByHeader(wGv As GridView, findHeader As String, findField As String) As Integer
        Dim colIdx As Integer = -1
        For Each column As DataControlField In wGv.Columns
            If TypeOf column Is BoundField AndAlso DirectCast(column, BoundField).HeaderText = findHeader Then
                colIdx = wGv.Columns.IndexOf(column)
                Exit For
            End If
        Next
        Return colIdx
    End Function
    Function getData(ByVal connectionString As String, ByVal query As String) As DataTable
        Dim dataTable As New DataTable()
        Using connection As New SqlConnection(connectionString)
            Using adapter As New SqlDataAdapter(query, connection)
                connection.Open()
                adapter.Fill(dataTable)
            End Using
        End Using
        Return dataTable
    End Function

    Protected Function getSQL1Row(ByVal wshp_id As String) As String
        ' Construct the SQL query to retrieve the row based on the ShipMethodID
        Dim sqlQuery As String = String.Format("SELECT * FROM Purchasing.ShipMethod WHERE ShipMethodID = '{0}'", wshp_id)
        Return sqlQuery
    End Function

    Protected Sub Gv1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GV1.SelectedIndexChanged
        If GV1.SelectedIndex <> -1 Then ' Ensure a row is selected
            Dim selectedRow As GridViewRow = GV1.SelectedRow


            Dim shipMethodID As String = selectedRow.Cells(GetColIdxByHeader(GV1, "ShipMethodID", "")).Text
            Session("SelectedShipMethodID") = shipMethodID

            Dim query As String = getSQL1Row(shipMethodID)
            Dim dt As DataTable = getData(connectionString, query)

            If dt.Rows.Count > 0 Then
                DV1.ChangeMode(DetailsViewMode.Edit)
                DV1.DataSource = dt
                DV1.DataBind()

                ' Show relevant components and hide others
                btnNewRow.Visible = False
                btnCancel.Visible = True
                btnInsert.Visible = False
                btnUpdate.Visible = True
                btnDelete.Visible = True
                MultiView1.ActiveViewIndex = 1
            End If
        End If
    End Sub

    Protected Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        MultiView1.ActiveViewIndex = 0
        '---Hide Component by MultiView
        btnNewRow.Visible = True
        btnCancel.Visible = False
        btnInsert.Visible = False
        btnUpdate.Visible = False
        btnDelete.Visible = False
    End Sub

    Public Function ProcessData(ByVal connectionString As String, ByVal sqlQuery As String) As Boolean
        Dim _processData As Boolean = False
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Using command As New SqlCommand(sqlQuery, connection)
                    command.ExecuteNonQuery()
                End Using
                _processData = True
            End Using
        Catch ex As Exception
            _processData = False
            ' Handle exception if necessary
        End Try
        Return _processData
    End Function

    Protected Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click

        Dim selectedRow As GridViewRow = GV1.SelectedRow
        Dim shipMethodID As String = selectedRow.Cells(GetColIdxByHeader(GV1, "ShipMethodID", "")).Text
        Session("SelectedShipMethodID") = shipMethodID

        If Page.IsValid Then
            ProcessData(connectionString, getSQL_Update(shipMethodID))
            '---to Scrn-1
            btnNewRow.Visible = True
            btnCancel.Visible = False
            btnInsert.Visible = False
            btnUpdate.Visible = False
            btnDelete.Visible = False

            DV1.DataSourceID = ""
            DV1.DataBind()

            GV1.DataBind()
            MultiView1.ActiveViewIndex = 0
        End If

    End Sub
    Protected Function Dv_GetValue(wDv As DetailsView, wfindHeader As String, wfindField As String, wfindID As String) As String

        Return ""
    End Function
    Protected Function getSQL_Update(ByVal wPK_1 As String) As String
        Dim selectedRow As GridViewRow = GV1.SelectedRow
        Dim shipMethodID As String = selectedRow.Cells(GetColIdxByHeader(GV1, "ShipMethodID", "")).Text
        Session("SelectedShipMethodID") = shipMethodID

        Dim strName As String = CType(DV1.FindControl("txt_Name"), TextBox).Text
        Dim strShipbase As String = CType(DV1.FindControl("txt_Shipbase"), TextBox).Text
        Dim strShiprate As String = CType(DV1.FindControl("txt_Shiprate"), TextBox).Text
        Dim strRowguid As String = CType(DV1.FindControl("txt_Rowguid"), TextBox).Text
        Dim strModifieddate As String = CType(DV1.FindControl("txt_Modifieddate"), TextBox).Text

        Dim strSQL As New StringBuilder

        strSQL.AppendFormat("UPDATE Purchasing.ShipMethod SET Name='{0}', ShipBase={1}, Shiprate={2}, rowguid='{3}', ModifiedDate='{4}' ", strName, strShipbase, strShiprate, strRowguid, strModifieddate)
        strSQL.AppendFormat("WHERE ShipMethodID = '{0}'", shipMethodID)

        Return strSQL.ToString()
    End Function

    Protected Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        ' ตรวจสอบว่ามีแถวที่ถูกเลือกหรือไม่
        If GV1.SelectedIndex <> -1 Then
            ' ดึงค่า ShipMethodID จากแถวที่เลือก
            Dim shipMethodID As String = GV1.SelectedRow.Cells(GetColIdxByHeader(GV1, "ShipMethodID", "")).Text

            ' สร้างคำสั่ง SQL สำหรับลบข้อมูล
            Dim deleteQuery As String = "DELETE FROM Purchasing.ShipMethod WHERE ShipMethodID = @ShipMethodID"

            ' สร้าง connection และ command objects
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(deleteQuery, connection)
                    ' เพิ่ม parameter
                    command.Parameters.AddWithValue("@ShipMethodID", shipMethodID)

                    ' เปิดการเชื่อมต่อและ execute command
                    connection.Open()
                    command.ExecuteNonQuery()
                End Using
            End Using

            ' ล้างค่า Session
            Session.Remove("SelectedShipMethodID")

            ' ทำการอัปเดตหรือรีเฟรช GridView และ DetailsView
            GV1.DataBind()
            DV1.DataSourceID = ""
            DV1.DataBind()

            ' เปลี่ยนแสดง MultiView กลับไปที่หน้าที่ต้องการ
            MultiView1.ActiveViewIndex = 0

            ' ปิดปุ่มแก้ไข ลบ และยกเลิก และเปิดปุ่มเพิ่มข้อมูลใหม่
            btnNewRow.Visible = True
            btnCancel.Visible = False
            btnInsert.Visible = False
            btnUpdate.Visible = False
            btnDelete.Visible = False
        Else

        End If
    End Sub

End Class
