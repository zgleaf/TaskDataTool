Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class TaskDataTool

    Private Sub Btn_excel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_excel.Click

        OpenFileDialog1.Filter = "Excel File (*.xls)|*.xls"
        If Me.OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Me.TB_excel.Text = Me.OpenFileDialog1.FileName
        End If

    End Sub

    Private Sub Btn_db_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_db.Click

        OpenFileDialog1.Filter = "SQLServerExpress File (*.mdf)|*.mdf"
        If Me.OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Me.TB_db.Text = Me.OpenFileDialog1.FileName
        End If

    End Sub

    Private Sub Btn_import_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_import.Click
        If TB_excel.Text = "" Or TB_db.Text = "" Then
            myLog("please select excel file and db file.")
            Return
        End If

        Dim excel_file As String = Me.TB_excel.Text
        Dim db_file As String = Me.TB_db.Text
        Dim db_pre As String = Me.TB_dbpre.Text
        Dim db_post As String = Me.TB_dbpost.Text

        myLog("start import...")
        Dim dataset = getDataFromExcel(excel_file)
        setDataToDB(db_pre, db_file, db_post, dataset)

    End Sub

    Private Sub Btn_export_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_export.Click
        'myLog("not supported yet.")
        If TB_excel.Text = "" Or TB_db.Text = "" Then
            myLog("please select excel file and db file.")
            Return
        End If

        Dim excel_file As String = Me.TB_excel.Text
        Dim db_file As String = Me.TB_db.Text
        Dim db_pre As String = Me.TB_dbpre.Text
        Dim db_post As String = Me.TB_dbpost.Text

        Dim dataset = getDataFromDB(db_pre, db_file, db_post)
        setDataToExcel(excel_file, dataset)

    End Sub

    Public Function getDataFromExcel(ByVal filename As String) As DataSet

        Dim myDs As New DataSet
        Dim strConn As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'"
        Try
            Dim connection As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(strConn)
            Dim da As New OleDbDataAdapter("SELECT * FROM [Task$]", connection)
            myDs.Tables.Add("task")

            da.Fill(myDs, "task")

            myLog("read from excel succeed!!!")

            Dim myDsCol As DataColumn

            Dim out As String = ""
            For Each myDsCol In myDs.Tables("task").Columns
                out = out + myDsCol.ColumnName + " "
            Next
            Console.WriteLine(out)

            Dim myDsRow As DataRowCollection = myDs.Tables("task").Rows
            Dim items = myDsRow.Count
            myLog("item num: " + items.ToString())

        Catch ex As Exception
            myLog(ex.Message.ToString)
        End Try

        Return myDs

    End Function


    Public Sub setDataToExcel(ByVal filename As String, ByVal myDs As DataSet)

        Try
            Dim oExcel As New Excel.Application
            Dim oBook As Excel.Workbook
            Dim oSheet As Excel.Worksheet

            oBook = oExcel.Workbooks.Open(filename)
            oSheet = oBook.Worksheets().Add
            oSheet.Name = "Task_" + Now.Month.ToString() + "_" + Now.Day.ToString() + "_" + Now.Hour.ToString() + "_" + Now.Minute.ToString()

            myLog("Set sheet1 as " + oSheet.Name + " ok.")

            Dim myTable As Data.DataTable = myDs.Tables(0)
            For col As Integer = 0 To myTable.Columns.Count - 1
                oSheet.Cells(1, col + 1) = myTable.Columns(col).Caption
            Next
            myLog("Set caption ok.")

            For row As Integer = 0 To myTable.Rows.Count - 1
                For col As Integer = 0 To myTable.Columns.Count - 1
                    oSheet.Cells(row + 2, col + 1) = myTable.Rows(row)(col).ToString()
                Next
                myLog("export tem " + row.ToString + " ok.")
            Next
            oBook.Save()
            myLog("Save file ok.")
            oExcel.Quit()
            myLog("Excel quit ok.")

        Catch ex As Exception
            myLog(ex.Message.ToString)
        End Try

    End Sub

    Public Sub setDataToDB(ByVal dbpre As String, ByVal filename As String, ByVal dbpost As String, ByVal myDs As DataSet)

        Dim strConn As String = dbpre + filename + dbpost

        Dim task As String

        Dim append As Boolean = Me.CB_append.Checked

        Try
            Dim connect As SqlClient.SqlConnection = New SqlClient.SqlConnection(strConn)
            connect.Open()

            If Not append Then
                Dim cmd As New SqlClient.SqlCommand("SET IDENTITY_INSERT [Task] ON", connect)
                cmd.ExecuteNonQuery()
            End If

            If Not append Then
                Dim cmd As New SqlClient.SqlCommand("DELETE FROM [Task]", connect)
                cmd.ExecuteNonQuery()
                myLog("clear db.")
            End If

            Dim myDsRow As DataRow

            Dim info = "description, commiter, create_date, responsible, department, due_date, status, type, comment, update_date"
            If Not append Then
                info = "id, " + info
            End If
            Dim item_num As Integer
            For Each myDsRow In myDs.Tables("task").Rows
                Dim id As String = myDsRow("id").ToString()
                Dim descript As String = myDsRow("description").ToString()
                If Not append Then
                    task = id.ToString() + ","
                Else
                    task = ""
                End If
                task += "@description,"
                task += "'" + myDsRow("commiter").ToString() + "',"
                task += "'" + myDsRow("create_date").ToString() + "',"
                task += "'" + myDsRow("responsible").ToString() + "',"
                task += "'" + myDsRow("department").ToString() + "',"
                task += "'" + myDsRow("due_date").ToString() + "',"
                task += "'" + myDsRow("status").ToString() + "',"
                task += "'" + myDsRow("type").ToString() + "',"
                Dim comment As String = myDsRow("comment").ToString()
                task += "@comment,"
                task += "@update_date"

                Dim myinsert = "INSERT INTO [Task] (" + info + ") VALUES (" + task + ");"

                Dim cmd As New SqlClient.SqlCommand(myinsert, connect)
                cmd.Parameters.AddWithValue("@description", descript)
                cmd.Parameters.AddWithValue("@comment", comment)
                Dim update_date = Now.ToString("yyyy-MM-dd HH:mm:ss")
                cmd.Parameters.AddWithValue("@update_date", update_date)

                Try
                    cmd.ExecuteNonQuery()

                Catch ex As Exception
                    myLog("FAILED: " + myinsert)

                Finally
                    item_num += 1

                End Try
            Next

            If Not append Then
                myLog("import to db " + item_num.ToString + " items.")
            Else
                myLog("append to db " + item_num.ToString + " items.")
            End If

            If Not append Then
                Dim cmd As New SqlClient.SqlCommand("SET IDENTITY_INSERT [Task] OFF", connect)
                cmd.ExecuteNonQuery()
            End If
            connect.Close()

        Catch ex As Exception
            myLog(ex.Message.ToString)
        End Try

    End Sub

    Public Function getDataFromDB(ByVal dbpre As String, ByVal filename As String, ByVal dbpost As String) As DataSet

        Dim strConn As String = dbpre + filename + dbpost

        Dim myDs As New System.Data.DataSet

        Try
            Dim connect As SqlClient.SqlConnection = New SqlClient.SqlConnection(strConn)
            connect.Open()
            Dim cmd As New SqlClient.SqlCommand("SELECT * FROM [Task]", connect)
            Dim dataAdapter As New SqlClient.SqlDataAdapter(cmd)
            dataAdapter.Fill(myDs)

            connect.Close()

        Catch ex As Exception
            myLog(ex.Message.ToString)
        End Try

        Return myDs

    End Function

    Private Sub TaskDataTool_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TB_Log.Text = "please set sheet name as 'Task', "
        Me.TB_Log.Text += "and first row is " + vbCrLf
        Me.TB_Log.Text += "'id description commiter create_date responsible department due_date status type comment'" + vbCrLf

        Me.TB_excel.Text = "C:\Users\a3bh0zz\Documents\2.Project\Daily Task Management System in MRD\TaskDataTool\TEST.xls"
        Me.TB_db.Text = "D:\PROJ\TASKMANAGEMENT\TASKMANAGEMENT\APP_DATA\TASKMANAGEMENT.MDF"
        Me.TB_dbpre.Text = "Data Source=CNWSP3BH0ZZ;Initial Catalog="
        Me.TB_dbpost.Text = ";Persist Security Info=True;User ID=sa;Password=Jq609007;"

        'Me.TB_excel.Text = "D:\proj\TaskManagement\doc\TEST.xls"
        'Me.TB_db.Text = "D:\proj\TaskManagement\TaskManagement\App_Data\TaskManagement.mdf"

        'Me.TB_dbpre.Text = "Data Source=.\SQLEXPRESS;AttachDbFilename="
        'Me.TB_dbpost.Text = ";Integrated Security=True;User Instance=True"

    End Sub

    Private Sub myLog(ByVal log As String)
        Me.TB_Log.Text = log + vbCrLf + Me.TB_Log.Text
        Me.TB_Log.ScrollToCaret()
    End Sub
End Class
