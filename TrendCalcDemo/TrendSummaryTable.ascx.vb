' <copyright file="TrendSummaryTable.ascx.vb" company="Colorado State University">
' Copyright (c) 2017 All Rights Reserved
' </copyright>
' <author>Shawn McClure</author>
' <date>03/23/2017 11:39:58 AM</date>

Public Class TrendSummaryTable : Inherits System.Web.UI.UserControl

    ' Trend-related properties
    Protected _ColumnsToExcludeFromOutput() As String = Nothing
    Protected _CommandFileID As String = "TrendsSqlCommandFile"
    Protected _ConnectionString As String = Nothing
    Protected _ConnectionStringID As String = "AIRDATA_CORE"
    Protected _CsvListOfColumnsToExcludeFromOutput As String = Nothing
    Protected _BaseCommandID As String = Nothing
    Protected _BaseCommandText As String = Nothing
    Protected _Markup As New StringBuilder()
    Protected _MaxNumberOfContiguousMissingValuesForValidTrend As Integer = 4
    Protected _Params As New Hashtable()
    Protected _RequiredPercentageDataCompletenessForTrending As Double = 60.0
    Protected _OutputFieldDelimiter As String = "|"
    Protected _OutputFormat As String = "table"
    Protected _RegressionVarXColumnName As String = "Year"
    Protected _RegressionVarYColumnName As String = "Value"
    Protected _RegressionVarZColumnName As String = "Parameter"
    Protected _TrendQueryCommandID As String = Nothing
    Protected _TrendQueryCommandText As String = Nothing

    Protected Overrides Sub OnLoad(e As EventArgs)

        MyBase.OnLoad(e)

        LoadParameters()

        ValidateDatabaseParams()

        Prepare()

        BuildTrendTableMarkup()

        Response.Write(_Markup.ToString())

    End Sub

    Public Sub BuildTrendTableMarkup(Optional ByVal aData As Object = Nothing)

        Dim cmdTextTrend As String
        Dim dc As DataColumn
        Dim drBase As DataRow
        Dim dtBase As DataTable = Nothing
        Dim dsTrendDataAll As DataSet = Nothing
        Dim dtTrendDataStream As DataTable = Nothing
        Dim dtTrendDataAll As DataTable = Nothing
        Dim dtTrendStats As DataTable = Nothing
        Dim drTrendSummary As DataRow = Nothing
        Dim dtTrendSummary As DataTable = Nothing
        Dim hasSplitField As Boolean = False
        Dim moduleName As String = Me.GetType().Name & "." & (New System.Diagnostics.StackFrame().GetMethod().Name)
        Dim msg As String = Nothing
        Dim splitField As String = Nothing
        Dim template As New XpFormatTemplate()
        Dim unresolvedTokens As StringCollection

        Try

            dtBase = DataAnalysis.GetDataTableFromQuery(_BaseCommandText, _ConnectionString, "BaseTable")

            If dtBase Is Nothing OrElse dtBase.Rows Is Nothing OrElse dtBase.Rows.Count = 0 Then
                Throw New System.Exception("There was a problem retrieving the base data from the database.")
            ElseIf dtBase.TableName = "Error" Then
                Throw New System.Exception("There was a problem retrieving the base data from the database.")
            End If

            splitField = Request.Params("splitfield")
            If Not String.IsNullOrEmpty(splitField) Then hasSplitField = True

            dtTrendSummary = dtBase.Clone()

            dtTrendSummary.Columns.Add(New DataColumn("StreamKey", GetType(String)))
            dtTrendSummary.Columns.Add(New DataColumn("Network", GetType(String)))
            dtTrendSummary.Columns.Add(New DataColumn("TableName", GetType(String)))
            dtTrendSummary.Columns.Add(New DataColumn("NumXValues", GetType(Integer)))
            dtTrendSummary.Columns.Add(New DataColumn("NumYValues", GetType(Integer)))
            dtTrendSummary.Columns.Add(New DataColumn("TrendSlope", GetType(Double)))
            dtTrendSummary.Columns.Add(New DataColumn("TrendPValue", GetType(Double)))
            dtTrendSummary.Columns.Add(New DataColumn("P", GetType(Integer)))
            dtTrendSummary.Columns.Add(New DataColumn("M", GetType(Integer)))
            dtTrendSummary.Columns.Add(New DataColumn("T", GetType(Integer)))
            dtTrendSummary.Columns.Add(New DataColumn("S", GetType(Integer)))
            dtTrendSummary.Columns.Add(New DataColumn("SigmaS", GetType(Double)))
            dtTrendSummary.Columns.Add(New DataColumn("ZetaS", GetType(Double)))
            dtTrendSummary.Columns.Add(New DataColumn("XValues", GetType(String)))
            dtTrendSummary.Columns.Add(New DataColumn("YValues", GetType(String)))
            dtTrendSummary.Columns.Add(New DataColumn("Status", GetType(String)))
            dtTrendSummary.Columns.Add(New DataColumn("Message", GetType(String)))

            For Each drBase In dtBase.Rows

                cmdTextTrend = _TrendQueryCommandText

                For Each dc In dtBase.Columns
                    cmdTextTrend = template.Resolve(cmdTextTrend, dc.ColumnName, drBase(dc.ColumnName).ToString())
                Next

                unresolvedTokens = template.GetUnresolvedTokens(cmdTextTrend)

                If unresolvedTokens.Count = 0 Then

                    dtTrendDataAll = DataAnalysis.GetDataTableFromQuery(cmdTextTrend, _ConnectionString, "TrendData")

                    If Not dtTrendDataAll Is Nothing AndAlso dtTrendDataAll.TableName <> "Error" AndAlso Not dtTrendDataAll.Rows Is Nothing AndAlso dtTrendDataAll.Rows.Count > 0 Then

                        If Not dsTrendDataAll Is Nothing Then dsTrendDataAll.Clear()
                        dsTrendDataAll = Nothing

                        If hasSplitField AndAlso dtTrendDataAll.Columns.Contains(splitField) Then
                            dsTrendDataAll = DataAnalysis.SplitDataTableByKey(dtTrendDataAll, splitField, splitField)
                        Else
                            dsTrendDataAll = New DataSet()
                            dsTrendDataAll.Tables.Add(dtTrendDataAll)
                        End If

                        For Each dtTrendDataStream In dsTrendDataAll.Tables

                            drTrendSummary = dtTrendSummary.NewRow()

                            For Each dc In dtBase.Columns
                                drTrendSummary(dc.ColumnName) = drBase(dc.ColumnName)
                            Next

                            If Not dtTrendDataStream Is Nothing AndAlso Not dtTrendDataStream.Rows Is Nothing AndAlso dtTrendDataStream.Rows.Count > 0 Then

                                dtTrendDataStream.TableName = dtTrendDataStream.Rows(0)("StreamKey")

                                drTrendSummary("StreamKey") = dtTrendDataStream.Rows(0)("StreamKey")
                                drTrendSummary("Network") = dtTrendDataStream.Rows(0)("Network")
                                drTrendSummary("TableName") = dtTrendDataStream.TableName

                                'If drTrendSummary("SiteID").ToString() = "25393" And drTrendSummary("ParamID").ToString() = "2547" Then
                                '    System.Diagnostics.Debug.Write("asdf")
                                'End If

                                dtTrendStats = DataAnalysis.GetTheilRegressionValues(dtTrendDataStream, _RegressionVarXColumnName, _RegressionVarYColumnName, _RegressionVarZColumnName, _RequiredPercentageDataCompletenessForTrending, _MaxNumberOfContiguousMissingValuesForValidTrend, Nothing, "", -1, msg, "TrendStats")

                                If Not dtTrendStats Is Nothing AndAlso Not dtTrendStats.Rows Is Nothing AndAlso dtTrendStats.Rows.Count > 0 Then
                                    drTrendSummary("NumXValues") = dtTrendStats.Rows(0)("NumXValues")
                                    drTrendSummary("NumYValues") = dtTrendStats.Rows(0)("NumYValues")
                                    drTrendSummary("TrendSlope") = dtTrendStats.Rows(0)("TrendSlope")
                                    drTrendSummary("TrendPValue") = dtTrendStats.Rows(0)("TrendPValue")
                                    drTrendSummary("P") = dtTrendStats.Rows(0)("P")
                                    drTrendSummary("M") = dtTrendStats.Rows(0)("M")
                                    drTrendSummary("T") = dtTrendStats.Rows(0)("T")
                                    drTrendSummary("S") = dtTrendStats.Rows(0)("S")
                                    drTrendSummary("SigmaS") = dtTrendStats.Rows(0)("SigmaS")
                                    drTrendSummary("ZetaS") = dtTrendStats.Rows(0)("ZetaS")
                                    drTrendSummary("XValues") = dtTrendStats.Rows(0)("XValues")
                                    drTrendSummary("YValues") = dtTrendStats.Rows(0)("YValues")
                                    drTrendSummary("Status") = dtTrendStats.Rows(0)("Status")
                                    drTrendSummary("Message") = msg
                                Else
                                    drTrendSummary("TableName") = dtTrendDataStream.TableName
                                    drTrendSummary("Message") = "The trend statistics table was empty."
                                End If

                            Else
                                drTrendSummary("Message") = "The data stream table was empty."
                            End If

                            dtTrendSummary.Rows.Add(drTrendSummary)

                        Next

                    Else
                        drTrendSummary = dtTrendSummary.NewRow()
                        For Each dc In dtBase.Columns
                            drTrendSummary(dc.ColumnName) = drBase(dc.ColumnName)
                        Next
                        drTrendSummary("TableName") = "n/a"
                        If dtTrendDataAll.TableName <> "Error" Then
                            drTrendSummary("Message") = "FAILURE: No trend data was found"
                        ElseIf Not dtTrendDataAll.Rows Is Nothing AndAlso dtTrendDataAll.Rows.Count > 0 AndAlso dtTrendDataAll.Columns.Contains("Error") Then
                            drTrendSummary("Message") = "FAILURE: " & dtTrendDataAll.Rows(0)("Error").ToString().Replace(_OutputFieldDelimiter, " ")
                        Else
                            drTrendSummary("Message") = "FAILURE: There was a problem retrieving the base trend data."
                        End If
                        dtTrendSummary.Rows.Add(drTrendSummary)
                    End If

                Else
                    drTrendSummary = dtTrendSummary.NewRow()
                    For Each dc In dtBase.Columns
                        drTrendSummary(dc.ColumnName) = drBase(dc.ColumnName)
                    Next
                    drTrendSummary("TableName") = "n/a"
                    drTrendSummary("Message") = "Could not resolve all tokens in _TrendQueryCommandText. Unresolved tokens: " & DataAnalysis.JoinStringCollection(unresolvedTokens)
                    dtTrendSummary.Rows.Add(drTrendSummary)
                End If

            Next

            If Not _ColumnsToExcludeFromOutput Is Nothing AndAlso _ColumnsToExcludeFromOutput.Length > 0 Then
                For Each colname As String In _ColumnsToExcludeFromOutput
                    dtTrendSummary.Columns.Remove(colname)
                Next
            End If

            If _OutputFormat = "table" Then
                _Markup.Append(DataAnalysis.GetHtmlTableFromDataTable(dtTrendSummary))
            ElseIf _OutputFormat = "delimited" Then
                _Markup.Append(DataAnalysis.GetDelimitedStringFromDataTable(dtTrendSummary, -1, True, "<br />", _OutputFieldDelimiter))
            End If

        Catch ex As Exception
            _Markup.Append(ex)
        End Try

    End Sub

    Public Sub LoadParameters(Optional ByVal aParameters As Hashtable = Nothing, Optional aReset As Boolean = False)

        Dim val As String

        Try

            val = Request.Params("cmdfileid") : If Not String.IsNullOrEmpty(val) Then _CommandFileID = val
            val = Request.Params("cmdidbase") : If Not String.IsNullOrEmpty(val) Then _BaseCommandID = val
            val = Request.Params("cmdidtrend") : If Not String.IsNullOrEmpty(val) Then _TrendQueryCommandID = val
            val = Request.Params("connstrid") : If Not String.IsNullOrEmpty(val) Then _ConnectionStringID = val

            val = Request.Params("excludecols") : If Not String.IsNullOrEmpty(val) Then _ColumnsToExcludeFromOutput = val.Split(",")

            val = Request.Params("format")
            If Not String.IsNullOrEmpty(val) Then
                Select Case val.Trim().ToLower()
                    Case "delimited" : _OutputFormat = val.Trim().ToLower()
                    Case "table" : _OutputFormat = val.Trim().ToLower()
                End Select
            End If

            val = Request.Params("fdelimiter") : If Not String.IsNullOrEmpty(val) Then _OutputFieldDelimiter = val

        Catch ex As Exception
            Response.Write(ex.Message)
        End Try

    End Sub

    Public Sub Prepare()

        Dim unresolvedTokens As New Specialized.StringCollection()

        Try

            _BaseCommandText = DataAnalysis.ResolveCommandTextParameters(_BaseCommandText, DataAnalysis.Convert(Request.Params), unresolvedTokens)

            unresolvedTokens.Clear()

            _TrendQueryCommandText = DataAnalysis.ResolveCommandTextParameters(_TrendQueryCommandText, DataAnalysis.Convert(Request.Params), unresolvedTokens)

        Catch ex As Exception
            Response.Write(ex.Message)
        End Try

    End Sub

    Public Sub ValidateDatabaseParams(Optional ByVal ResolveCommandTextParameters As Boolean = True)

        Try

            If String.IsNullOrEmpty(_ConnectionString) Then ' An explicitly-set _ConnectionString has priority, so no need to go deeper if it's available
                If Not String.IsNullOrEmpty(_ConnectionStringID) Then ' Try the _ConnectionStringID property next
                    _ConnectionString = DataAnalysis.ResolveConnectionString(_ConnectionStringID) ' Attempt to resolve the connection string from Web.config using _ConnectionStringID
                    If String.IsNullOrEmpty(_ConnectionString) Then
                        Throw New System.Exception("ConnectionStringID """ & _ConnectionStringID & """ was not found in Web.config.")
                    Else
                        _Params.Add("connstr", _ConnectionString)
                    End If
                Else
                    Throw New System.Exception("Either the ""ConnectionString"" or the ""ConnectionStringID"" property must be provided, but both are empty.")
                End If
            End If

            If String.IsNullOrEmpty(_CommandFileID) Then
                Throw New System.Exception("Required property ""_CommandFileID"" was empty.")
            End If

            If String.IsNullOrEmpty(_BaseCommandID) Then
                Throw New System.Exception("Required property ""_BaseCommandID"" was empty.")
            End If

            If String.IsNullOrEmpty(_TrendQueryCommandID) Then
                Throw New System.Exception("Required property ""_TrendQueryCommandID"" was empty.")
            End If

            _BaseCommandText = DataAnalysis.ResolveCommandText(_BaseCommandID, _CommandFileID)

            If String.IsNullOrEmpty(_BaseCommandText) Then
                Throw New System.Exception("CommandID """ & _BaseCommandID & """ was not found in Web.config.")
            End If

            _TrendQueryCommandText = DataAnalysis.ResolveCommandText(_TrendQueryCommandID, _CommandFileID)

            If String.IsNullOrEmpty(_TrendQueryCommandText) Then
                Throw New System.Exception("CommandID """ & _TrendQueryCommandID & """ was not found in Web.config.")
            End If

        Catch ex As SystemException
            Response.Write(ex.Message)
        End Try

    End Sub

End Class
