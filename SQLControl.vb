Imports System.Data.SqlClient

Public Class SQLControl
    ' DEVELOPMENT SERVER:
    Private SQLCon As New SqlConnection("Server=LMR-R67;Database=All_Names;User=sa;Pwd=stud1o")

    ' DEMO SERVER:
    ' Private SQLCon As New SqlConnection("Server=mlmr-web;Database=All_Names;User=sa;Pwd=stud1o@2015")

    ' PRODUCTION SERVER:
    ' Private SQLCon As New SqlConnection("Server=Zonda;Database=All_Names;User=sa;Pwd=password")

    Private SQLCmd As New SqlCommand

    Public SQLDA As New SqlDataAdapter

    Public SQLDT As New DataTable

    Public Params As New List(Of SqlParameter)

    Public RecordCount As Integer
    Public Exception As String

    Public Sub New()

    End Sub

    Public Sub New(ConectioString As String)
        SQLCon = New SqlConnection(ConectioString)
    End Sub

    Public Sub RunQuery(Query As String, Optional ReturnIdentity As Boolean = False)
        RecordCount = 0
        Exception = ""

        Try
            SQLCon.Open()

            SQLCmd = New SqlCommand(Query, SQLCon)

            Params.ForEach(Sub(p) SQLCmd.Parameters.Add(p))
            Params.Clear()

            SQLDT = New DataTable
            SQLDA = New SqlDataAdapter(SQLCmd)
            RecordCount = SQLDA.Fill(SQLDT)

            If ReturnIdentity = True Then
                Dim ReturnQuery As String = ("SELECT @@IDENTITY As LastID;")
                SQLCmd = New SqlCommand(ReturnQuery, SQLCon)
                SQLDT = New DataTable
                SQLDA = New SqlDataAdapter(SQLCmd)
                RecordCount = SQLDA.Fill(SQLDT)
            End If
        Catch ex As Exception
            Exception = "RunQuery Error:" & vbNewLine & ex.Message
        Finally
            If SQLCon.State = ConnectionState.Open Then SQLCon.Close()
        End Try

    End Sub

    Public Sub AddParams(Name As String, Value As Object)
        Dim NewParams As New SqlParameter(Name, Value)
        Params.Add(NewParams)
    End Sub

    Public Function HasException(Optional Report As Boolean = False) As Boolean
        If String.IsNullOrEmpty(Exception) Then Return False
        If Report = True Then MsgBox(Exception, MsgBoxStyle.Critical, "Exception:")
        Return True
    End Function

End Class
