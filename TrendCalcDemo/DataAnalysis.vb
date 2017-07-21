' <copyright file="DataAnalysis.vb" company="Colorado State University">
' Copyright (c) 2017 All Rights Reserved
' </copyright>
' <author>Shawn McCLure</author>
' <date>03/23/2017 11:39:58 AM</date>

Imports System.Data.SqlClient
Imports System.IO
Imports System.Xml

Public Class DataAnalysis

    Public Shared KendallPValueLookup As Dictionary(Of System.Tuple(Of Integer, Integer), Double) = New Dictionary(Of Tuple(Of Integer, Integer), Double) From {
        {New Tuple(Of Integer, Integer)(3, 1), 0.5},
        {New Tuple(Of Integer, Integer)(3, 3), 0.167},
        {New Tuple(Of Integer, Integer)(4, 0), 0.625},
        {New Tuple(Of Integer, Integer)(4, 2), 0.375},
        {New Tuple(Of Integer, Integer)(4, 4), 0.167},
        {New Tuple(Of Integer, Integer)(4, 6), 0.042},
        {New Tuple(Of Integer, Integer)(5, 0), 0.592},
        {New Tuple(Of Integer, Integer)(5, 2), 0.408},
        {New Tuple(Of Integer, Integer)(5, 4), 0.242},
        {New Tuple(Of Integer, Integer)(5, 6), 0.117},
        {New Tuple(Of Integer, Integer)(5, 8), 0.042},
        {New Tuple(Of Integer, Integer)(5, 10), 0.0083},
        {New Tuple(Of Integer, Integer)(6, 1), 0.5},
        {New Tuple(Of Integer, Integer)(6, 3), 0.36},
        {New Tuple(Of Integer, Integer)(6, 5), 0.235},
        {New Tuple(Of Integer, Integer)(6, 7), 0.136},
        {New Tuple(Of Integer, Integer)(6, 9), 0.068},
        {New Tuple(Of Integer, Integer)(6, 11), 0.028},
        {New Tuple(Of Integer, Integer)(6, 13), 0.0083},
        {New Tuple(Of Integer, Integer)(6, 15), 0.0014},
        {New Tuple(Of Integer, Integer)(7, 1), 0.5},
        {New Tuple(Of Integer, Integer)(7, 3), 0.386},
        {New Tuple(Of Integer, Integer)(7, 5), 0.281},
        {New Tuple(Of Integer, Integer)(7, 7), 0.191},
        {New Tuple(Of Integer, Integer)(7, 9), 0.119},
        {New Tuple(Of Integer, Integer)(7, 11), 0.068},
        {New Tuple(Of Integer, Integer)(7, 13), 0.035},
        {New Tuple(Of Integer, Integer)(7, 15), 0.015},
        {New Tuple(Of Integer, Integer)(7, 17), 0.0054},
        {New Tuple(Of Integer, Integer)(7, 19), 0.0014},
        {New Tuple(Of Integer, Integer)(7, 21), 0.0002},
        {New Tuple(Of Integer, Integer)(8, 0), 0.548},
        {New Tuple(Of Integer, Integer)(8, 2), 0.452},
        {New Tuple(Of Integer, Integer)(8, 4), 0.36},
        {New Tuple(Of Integer, Integer)(8, 6), 0.274},
        {New Tuple(Of Integer, Integer)(8, 8), 0.199},
        {New Tuple(Of Integer, Integer)(8, 10), 0.138},
        {New Tuple(Of Integer, Integer)(8, 12), 0.089},
        {New Tuple(Of Integer, Integer)(8, 14), 0.054},
        {New Tuple(Of Integer, Integer)(8, 16), 0.031},
        {New Tuple(Of Integer, Integer)(8, 18), 0.0156},
        {New Tuple(Of Integer, Integer)(8, 20), 0.0071},
        {New Tuple(Of Integer, Integer)(8, 22), 0.0028},
        {New Tuple(Of Integer, Integer)(8, 24), 0.0009},
        {New Tuple(Of Integer, Integer)(8, 26), 0.0002},
        {New Tuple(Of Integer, Integer)(8, 28), 0.00009},
        {New Tuple(Of Integer, Integer)(9, 0), 0.54},
        {New Tuple(Of Integer, Integer)(9, 2), 0.46},
        {New Tuple(Of Integer, Integer)(9, 4), 0.381},
        {New Tuple(Of Integer, Integer)(9, 6), 0.306},
        {New Tuple(Of Integer, Integer)(9, 8), 0.238},
        {New Tuple(Of Integer, Integer)(9, 10), 0.179},
        {New Tuple(Of Integer, Integer)(9, 12), 0.13},
        {New Tuple(Of Integer, Integer)(9, 14), 0.09},
        {New Tuple(Of Integer, Integer)(9, 16), 0.06},
        {New Tuple(Of Integer, Integer)(9, 18), 0.038},
        {New Tuple(Of Integer, Integer)(9, 20), 0.022},
        {New Tuple(Of Integer, Integer)(9, 22), 0.0124},
        {New Tuple(Of Integer, Integer)(9, 24), 0.0063},
        {New Tuple(Of Integer, Integer)(9, 26), 0.0029},
        {New Tuple(Of Integer, Integer)(9, 28), 0.0012},
        {New Tuple(Of Integer, Integer)(9, 30), 0.0004},
        {New Tuple(Of Integer, Integer)(9, 32), 0.0001},
        {New Tuple(Of Integer, Integer)(9, 34), 0.00009},
        {New Tuple(Of Integer, Integer)(9, 36), 0.00009},
        {New Tuple(Of Integer, Integer)(10, 1), 0.5},
        {New Tuple(Of Integer, Integer)(10, 3), 0.431},
        {New Tuple(Of Integer, Integer)(10, 5), 0.364},
        {New Tuple(Of Integer, Integer)(10, 7), 0.3},
        {New Tuple(Of Integer, Integer)(10, 9), 0.242},
        {New Tuple(Of Integer, Integer)(10, 11), 0.19},
        {New Tuple(Of Integer, Integer)(10, 13), 0.146},
        {New Tuple(Of Integer, Integer)(10, 15), 0.108},
        {New Tuple(Of Integer, Integer)(10, 17), 0.078},
        {New Tuple(Of Integer, Integer)(10, 19), 0.054},
        {New Tuple(Of Integer, Integer)(10, 21), 0.036},
        {New Tuple(Of Integer, Integer)(10, 23), 0.023},
        {New Tuple(Of Integer, Integer)(10, 25), 0.0143},
        {New Tuple(Of Integer, Integer)(10, 27), 0.0083},
        {New Tuple(Of Integer, Integer)(10, 29), 0.0046},
        {New Tuple(Of Integer, Integer)(10, 31), 0.0023},
        {New Tuple(Of Integer, Integer)(10, 33), 0.0011},
        {New Tuple(Of Integer, Integer)(10, 35), 0.0005},
        {New Tuple(Of Integer, Integer)(10, 37), 0.0002},
        {New Tuple(Of Integer, Integer)(10, 39), 0.00009},
        {New Tuple(Of Integer, Integer)(10, 41), 0.00009},
        {New Tuple(Of Integer, Integer)(10, 43), 0.00009},
        {New Tuple(Of Integer, Integer)(10, 45), 0.00009}
    }

    Public Shared StandardNormalDistribution As Dictionary(Of Double, Double) = New Dictionary(Of Double, Double) From {
        {0, 0.5},
        {0.01, 0.504},
        {0.02, 0.508},
        {0.03, 0.512},
        {0.04, 0.516},
        {0.05, 0.5199},
        {0.06, 0.5239},
        {0.07, 0.5279},
        {0.08, 0.5319},
        {0.09, 0.5359},
        {0.1, 0.5398},
        {0.11, 0.5438},
        {0.12, 0.5478},
        {0.13, 0.5517},
        {0.14, 0.5557},
        {0.15, 0.5596},
        {0.16, 0.5636},
        {0.17, 0.5675},
        {0.18, 0.5714},
        {0.19, 0.5753},
        {0.2, 0.5793},
        {0.21, 0.5832},
        {0.22, 0.5871},
        {0.23, 0.591},
        {0.24, 0.5948},
        {0.25, 0.5987},
        {0.26, 0.6026},
        {0.27, 0.6064},
        {0.28, 0.6103},
        {0.29, 0.6141},
        {0.3, 0.6179},
        {0.31, 0.6217},
        {0.32, 0.6255},
        {0.33, 0.6293},
        {0.34, 0.6331},
        {0.35, 0.6368},
        {0.36, 0.6406},
        {0.37, 0.6443},
        {0.38, 0.648},
        {0.39, 0.6517},
        {0.4, 0.6554},
        {0.41, 0.6591},
        {0.42, 0.6628},
        {0.43, 0.6664},
        {0.44, 0.67},
        {0.45, 0.6736},
        {0.46, 0.6772},
        {0.47, 0.6808},
        {0.48, 0.6844},
        {0.49, 0.6879},
        {0.5, 0.6915},
        {0.51, 0.695},
        {0.52, 0.6985},
        {0.53, 0.7019},
        {0.54, 0.7054},
        {0.55, 0.7088},
        {0.56, 0.7123},
        {0.57, 0.7157},
        {0.58, 0.719},
        {0.59, 0.7224},
        {0.6, 0.7257},
        {0.61, 0.7291},
        {0.62, 0.7324},
        {0.63, 0.7357},
        {0.64, 0.7389},
        {0.65, 0.7422},
        {0.66, 0.7454},
        {0.67, 0.7486},
        {0.68, 0.7517},
        {0.69, 0.7549},
        {0.7, 0.758},
        {0.71, 0.7611},
        {0.72, 0.7642},
        {0.73, 0.7673},
        {0.74, 0.7704},
        {0.75, 0.7734},
        {0.76, 0.7764},
        {0.77, 0.7794},
        {0.78, 0.7823},
        {0.79, 0.7852},
        {0.8, 0.7881},
        {0.81, 0.791},
        {0.82, 0.7939},
        {0.83, 0.7967},
        {0.84, 0.7995},
        {0.85, 0.8023},
        {0.86, 0.8051},
        {0.87, 0.8078},
        {0.88, 0.8106},
        {0.89, 0.8133},
        {0.9, 0.8159},
        {0.91, 0.8186},
        {0.92, 0.8212},
        {0.93, 0.8238},
        {0.94, 0.8264},
        {0.95, 0.8289},
        {0.96, 0.8315},
        {0.97, 0.834},
        {0.98, 0.8365},
        {0.99, 0.8389},
        {1, 0.8413},
        {1.01, 0.8438},
        {1.02, 0.8461},
        {1.03, 0.8485},
        {1.04, 0.8508},
        {1.05, 0.8531},
        {1.06, 0.8554},
        {1.07, 0.8577},
        {1.08, 0.8599},
        {1.09, 0.8621},
        {1.1, 0.8643},
        {1.11, 0.8665},
        {1.12, 0.8686},
        {1.13, 0.8708},
        {1.14, 0.8729},
        {1.15, 0.8749},
        {1.16, 0.877},
        {1.17, 0.879},
        {1.18, 0.881},
        {1.19, 0.883},
        {1.2, 0.8849},
        {1.21, 0.8869},
        {1.22, 0.8888},
        {1.23, 0.8907},
        {1.24, 0.8925},
        {1.25, 0.8944},
        {1.26, 0.8962},
        {1.27, 0.898},
        {1.28, 0.8997},
        {1.29, 0.9015},
        {1.3, 0.9032},
        {1.31, 0.9049},
        {1.32, 0.9066},
        {1.33, 0.9082},
        {1.34, 0.9099},
        {1.35, 0.9115},
        {1.36, 0.9131},
        {1.37, 0.9147},
        {1.38, 0.9162},
        {1.39, 0.9177},
        {1.4, 0.9192},
        {1.41, 0.9207},
        {1.42, 0.9222},
        {1.43, 0.9236},
        {1.44, 0.9251},
        {1.45, 0.9265},
        {1.46, 0.9279},
        {1.47, 0.9292},
        {1.48, 0.9306},
        {1.49, 0.9319},
        {1.5, 0.9332},
        {1.51, 0.9345},
        {1.52, 0.9357},
        {1.53, 0.937},
        {1.54, 0.9382},
        {1.55, 0.9394},
        {1.56, 0.9406},
        {1.57, 0.9418},
        {1.58, 0.9429},
        {1.59, 0.9441},
        {1.6, 0.9452},
        {1.61, 0.9463},
        {1.62, 0.9474},
        {1.63, 0.9484},
        {1.64, 0.9495},
        {1.65, 0.9505},
        {1.66, 0.9515},
        {1.67, 0.9525},
        {1.68, 0.9535},
        {1.69, 0.9545},
        {1.7, 0.9554},
        {1.71, 0.9564},
        {1.72, 0.9573},
        {1.73, 0.9582},
        {1.74, 0.9591},
        {1.75, 0.9599},
        {1.76, 0.9608},
        {1.77, 0.9616},
        {1.78, 0.9625},
        {1.79, 0.9633},
        {1.8, 0.9641},
        {1.81, 0.9649},
        {1.82, 0.9656},
        {1.83, 0.9664},
        {1.84, 0.9671},
        {1.85, 0.9678},
        {1.86, 0.9686},
        {1.87, 0.9693},
        {1.88, 0.9699},
        {1.89, 0.9706},
        {1.9, 0.9713},
        {1.91, 0.9719},
        {1.92, 0.9726},
        {1.93, 0.9732},
        {1.94, 0.9738},
        {1.95, 0.9744},
        {1.96, 0.975},
        {1.97, 0.9756},
        {1.98, 0.9761},
        {1.99, 0.9767},
        {2, 0.9772},
        {2.01, 0.9778},
        {2.02, 0.9783},
        {2.03, 0.9788},
        {2.04, 0.9793},
        {2.05, 0.9798},
        {2.06, 0.9803},
        {2.07, 0.9808},
        {2.08, 0.9812},
        {2.09, 0.9817},
        {2.1, 0.9821},
        {2.11, 0.9826},
        {2.12, 0.983},
        {2.13, 0.9834},
        {2.14, 0.9838},
        {2.15, 0.9842},
        {2.16, 0.9846},
        {2.17, 0.985},
        {2.18, 0.9854},
        {2.19, 0.9857},
        {2.2, 0.9861},
        {2.21, 0.9864},
        {2.22, 0.9868},
        {2.23, 0.9871},
        {2.24, 0.9875},
        {2.25, 0.9878},
        {2.26, 0.9881},
        {2.27, 0.9884},
        {2.28, 0.9887},
        {2.29, 0.989},
        {2.3, 0.9893},
        {2.31, 0.9896},
        {2.32, 0.9898},
        {2.33, 0.9901},
        {2.34, 0.9904},
        {2.35, 0.9906},
        {2.36, 0.9909},
        {2.37, 0.9911},
        {2.38, 0.9913},
        {2.39, 0.9916},
        {2.4, 0.9918},
        {2.41, 0.992},
        {2.42, 0.9922},
        {2.43, 0.9925},
        {2.44, 0.9927},
        {2.45, 0.9929},
        {2.46, 0.9931},
        {2.47, 0.9932},
        {2.48, 0.9934},
        {2.49, 0.9936},
        {2.5, 0.9938},
        {2.51, 0.994},
        {2.52, 0.9941},
        {2.53, 0.9943},
        {2.54, 0.9945},
        {2.55, 0.9946},
        {2.56, 0.9948},
        {2.57, 0.9949},
        {2.58, 0.9951},
        {2.59, 0.9952},
        {2.6, 0.9953},
        {2.61, 0.9955},
        {2.62, 0.9956},
        {2.63, 0.9957},
        {2.64, 0.9959},
        {2.65, 0.996},
        {2.66, 0.9961},
        {2.67, 0.9962},
        {2.68, 0.9963},
        {2.69, 0.9964},
        {2.7, 0.9965},
        {2.71, 0.9966},
        {2.72, 0.9967},
        {2.73, 0.9968},
        {2.74, 0.9969},
        {2.75, 0.997},
        {2.76, 0.9971},
        {2.77, 0.9972},
        {2.78, 0.9973},
        {2.79, 0.9974},
        {2.8, 0.9974},
        {2.81, 0.9975},
        {2.82, 0.9976},
        {2.83, 0.9977},
        {2.84, 0.9977},
        {2.85, 0.9978},
        {2.86, 0.9979},
        {2.87, 0.9979},
        {2.88, 0.998},
        {2.89, 0.9981},
        {2.9, 0.9981},
        {2.91, 0.9982},
        {2.92, 0.9982},
        {2.93, 0.9983},
        {2.94, 0.9984},
        {2.95, 0.9984},
        {2.96, 0.9985},
        {2.97, 0.9985},
        {2.98, 0.9986},
        {2.99, 0.9986},
        {3, 0.9987},
        {3.01, 0.9987},
        {3.02, 0.9987},
        {3.03, 0.9988},
        {3.04, 0.9988},
        {3.05, 0.9989},
        {3.06, 0.9989},
        {3.07, 0.9989},
        {3.08, 0.999},
        {3.09, 0.999},
        {3.1, 0.999},
        {3.11, 0.9991},
        {3.12, 0.9991},
        {3.13, 0.9991},
        {3.14, 0.9992},
        {3.15, 0.9992},
        {3.16, 0.9992},
        {3.17, 0.9992},
        {3.18, 0.9993},
        {3.19, 0.9993},
        {3.2, 0.9993},
        {3.21, 0.9993},
        {3.22, 0.9994},
        {3.23, 0.9994},
        {3.24, 0.9994},
        {3.25, 0.9994},
        {3.26, 0.9994},
        {3.27, 0.9995},
        {3.28, 0.9995},
        {3.29, 0.9995},
        {3.3, 0.9995},
        {3.31, 0.9995},
        {3.32, 0.9995},
        {3.33, 0.9996},
        {3.34, 0.9996},
        {3.35, 0.9996},
        {3.36, 0.9996},
        {3.37, 0.9996},
        {3.38, 0.9996},
        {3.39, 0.9997},
        {3.4, 0.9997},
        {3.41, 0.9997},
        {3.42, 0.9997},
        {3.43, 0.9997},
        {3.44, 0.9997},
        {3.45, 0.9997},
        {3.46, 0.9997},
        {3.47, 0.9997},
        {3.48, 0.9997},
        {3.49, 0.9998}
    }

    Public Shared Function ArrayFromColumnValues_Double(ByVal aDT As DataTable, ByVal aColumnName As String, Optional ByVal aUnique As Boolean = True, Optional ByVal aNullValueSubstitute As Nullable(Of Double) = Nothing) As Double()

        If aDT Is Nothing OrElse aDT.Rows Is Nothing OrElse aDT.Rows.Count = 0 OrElse String.IsNullOrEmpty(aColumnName) OrElse Not aDT.Columns.Contains(aColumnName) Then Return Nothing

        Dim isNotEmpty As Boolean = True
        Dim isNumber As Boolean = False
        Dim obj As Object
        Dim val As Double
        Dim vals() As Double = Nothing

        Try

            Select Case aDT.Columns(aColumnName).DataType.Name
                Case "Int16", "Int32", "Int64", "Single", "Decimal", "Double", "Float", "Real" : isNumber = True
            End Select

            If Not isNumber Then Throw New System.Exception("The data type of column """ & aColumnName & """ must be numeric. Instead, it is of type """ & aDT.Columns(aColumnName).DataType.Name & """.")

            For Each DR As DataRow In aDT.Rows

                obj = DR(aColumnName)

                isNotEmpty = True

                If Not IsDBNull(obj) AndAlso CStr(obj).Trim().Length > 0 Then
                    val = CDbl(obj)
                ElseIf aNullValueSubstitute.HasValue Then
                    val = CDbl(aNullValueSubstitute)
                Else
                    isNotEmpty = False
                End If

                If isNotEmpty Then

                    If vals Is Nothing Then
                        ReDim vals(0)
                        vals(0) = val
                    ElseIf aUnique Then
                        If Array.IndexOf(vals, val) = -1 Then
                            ReDim Preserve vals(UBound(vals) + 1)
                            vals(UBound(vals)) = obj
                        End If
                    Else
                        ReDim Preserve vals(UBound(vals) + 1)
                        vals(UBound(vals)) = val
                    End If

                End If

            Next

            Return vals

        Catch ex As Exception
            Throw New System.Exception(ex.Message & " (Source: " & New System.Diagnostics.StackFrame().GetMethod().Name & ")")
        End Try

    End Function

    Public Shared Function ArrayMaxDouble(ByVal aArray() As Double, Optional ByVal aErrorValue As Double = -999.0) As Double

        If aArray Is Nothing OrElse aArray.Length = 0 Then Return aErrorValue

        Dim max As Double

        Try

            For i As Integer = 0 To UBound(aArray)
                max = IIf(i > 0, IIf(aArray(i) > max, aArray(i), max), aArray(i))
            Next

            Return max

        Catch ex As Exception
            Return aErrorValue
        End Try

    End Function

    Public Shared Function ArrayMedianDouble(ByVal aArray() As Double, Optional ByVal aErrorValue As Double = -999.0) As Double

        'Definition: "Middle value" of a list. The smallest number such that at least half the numbers in the list are no greater than it. If the
        'list has an odd number of entries, the median is the middle entry in the list after sorting the list into increasing order. If the list
        'has an even number of entries, the median is equal to the sum of the two middle (after sorting) numbers divided by two.

        If aArray Is Nothing OrElse aArray.GetLength(0) = 0 Then Return aErrorValue
        If aArray.GetLength(0) = 1 Then Return aArray(0) ' A one-element array (array with length = 1)

        Dim median As Double = aErrorValue
        Dim n As Integer = aArray.GetLength(0)

        'We need to sort the numbers to get the median.
        Array.Sort(aArray)

        'If the array length is divisible by two, add the two middle numbers together and return the average (mean!) of those.
        If n Mod 2 = 0 Then
            Dim middleElement1 As Double = aArray((n / 2) - 1)
            Dim middleElement2 As Double = aArray((n / 2))
            median = (middleElement1 + middleElement2) / 2
        Else
            median = aArray(n \ 2) ' The backslash operator ("\") performs "integer division", while the forward slash operator ("/") performs floating point division. See: http://msdn.microsoft.com/en-us/library/b6ex274z.aspx
        End If

        Return median

    End Function

    Public Shared Function ArrayMinDouble(ByVal aArray() As Double, Optional ByVal aErrorValue As Double = -999.0) As Double

        If aArray Is Nothing OrElse aArray.Length = 0 Then Return aErrorValue

        Dim min As Double

        Try

            For i As Integer = 0 To UBound(aArray)
                min = IIf(i > 0, IIf(aArray(i) < min, aArray(i), min), aArray(i))
            Next

            Return min

        Catch ex As Exception
            Return aErrorValue
        End Try

    End Function

    Public Shared Function Convert(ByVal aNVC As Specialized.NameValueCollection) As Hashtable
        Dim ht As New Hashtable()
        If Not aNVC Is Nothing Then
            For Each key As String In aNVC.AllKeys  'Use the "AllKeys" property (collection) when iterating over a NameValueCollection
                ht.Add(key, aNVC.Get(key))
            Next
        End If
        Return ht
    End Function

    Public Shared Function GetDataTableFromQuery(
            ByVal CommandText As String,
            ByVal ConnectionString As String,
            Optional ByVal DataTableName As String = "DataTable",
            Optional ByVal CommandTimeout As Integer = -1) As DataTable

        Dim DA As SqlDataAdapter
        Dim DR As DataRow
        Dim DT As DataTable = Nothing

        Try

            If String.IsNullOrEmpty(CommandText) Then Throw New System.Exception("Required argument ""CommandText"" is empty.")
            If String.IsNullOrEmpty(ConnectionString) Then Throw New System.Exception("Required argument ""ConnectionString"" is empty.")

            DA = New SqlDataAdapter(CommandText, ConnectionString)

            If CommandTimeout <> -1 Then DA.SelectCommand.CommandTimeout = CommandTimeout

            DT = New DataTable(IIf(Not String.IsNullOrEmpty(DataTableName), DataTableName, "DataTable"))

            Try
                DA.Fill(DT)
            Catch ex As Exception
                DT = New DataTable("Error")
                DT.Columns.Add(New DataColumn("Error", GetType(String)))
                DR = DT.NewRow()
                DR("Error") = ex.Message
                DT.Rows.Add(DR)
            End Try

            Return DT

        Catch ex As System.Exception
            Throw New System.Exception(ex.Message & " (Source: " & New System.Diagnostics.StackFrame().GetMethod().Name & ")")
        End Try

    End Function

    Public Shared Function GetDelimitedStringFromDataTable(
            ByVal aDT As DataTable,
            Optional ByVal aNumRowsToShow As Integer = -1,
            Optional ByVal aIncludeColumnHeader As Boolean = False,
            Optional ByVal aRecordDelimiter As String = "!",
            Optional ByVal aFieldDelimiter As String = "|") As String

        Dim DC As DataColumn
        Dim DR As DataRow
        Dim colCount As Integer
        Dim rowCount As Integer = 0
        Dim sb As New StringBuilder()

        Try

            If aDT Is Nothing Then
                sb.Append("Source data table was empty in XpData.GetHtmlTableFromDataTable()")
                Return sb.ToString()
            End If

            If aIncludeColumnHeader Then
                colCount = 0
                For Each DC In aDT.Columns
                    If colCount > 0 Then sb.Append(aFieldDelimiter)
                    sb.Append(DC.ColumnName)
                    colCount += 1
                Next
            End If

            If Not aDT.Rows Is Nothing AndAlso aDT.Rows.Count > 0 Then

                If aIncludeColumnHeader Then sb.Append(aRecordDelimiter)

                For Each DR In aDT.Rows

                    If rowCount > 0 Then sb.Append(aRecordDelimiter)

                    colCount = 0
                    For Each DC In aDT.Columns
                        If colCount > 0 Then sb.Append(aFieldDelimiter)
                        sb.Append(DR(DC.ColumnName).ToString())
                        colCount += 1
                    Next

                    rowCount += 1

                    If aNumRowsToShow > 0 AndAlso rowCount >= aNumRowsToShow Then Exit For

                Next

            End If

            Return sb.ToString()

        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try

    End Function

    Public Shared Function GetHtmlTableFromDataTable(
            ByVal aDT As DataTable,
            Optional ByVal aNumRowsToShow As Integer = -1,
            Optional ByVal aCssClass As String = "tblData",
            Optional ByVal aUseFieldNumberAsCssClass As Boolean = True,
            Optional ByVal aEmptyTableMessage As String = "No records returned") As String

        Dim DC As DataColumn
        Dim DR As DataRow
        Dim obj As Object
        Dim colCount As Integer
        Dim rowCount As Integer = 0
        Dim sb As New StringBuilder()
        Dim tableID As String = "tbl" & GetRandomKey2(4)
        Dim val As String

        Try

            If aDT Is Nothing Then
                sb.Append("<table id=""") : sb.Append(tableID) : sb.Append(""" class=""") : sb.Append(aCssClass) : sb.Append("""><thead><tr><th>Empty Source Table</th></tr></thead><tbody><tr><td>Source data table was empty in XpData.GetHtmlTableFromDataTable()</td></tr></tbody></table>")
                Return sb.ToString()
            ElseIf aDT.TableName.Trim().Length > 0 Then
                tableID = aDT.TableName
            End If

            If String.IsNullOrEmpty(aCssClass) Then aCssClass = "tblData"

            sb.Append("<table id=""") : sb.Append(tableID) : sb.Append(""" class=""") : sb.Append(aCssClass) : sb.Append(""">")
            sb.Append("<thead>")
            sb.Append("<tr>")

            colCount = 0
            For Each DC In aDT.Columns
                sb.Append("<th")
                If aUseFieldNumberAsCssClass Then
                    sb.Append(" class=""c") : sb.Append(colCount.ToString) : sb.Append("""")
                End If
                sb.Append(">")
                sb.Append(DC.ColumnName)
                sb.Append("</th>")
                colCount += 1
            Next

            sb.Append("</tr>")
            sb.Append("</thead>")
            sb.Append("<tbody>")

            If Not aDT.Rows Is Nothing AndAlso aDT.Rows.Count > 0 Then

                For Each DR In aDT.Rows

                    sb.Append("<tr>")

                    colCount = 0
                    For Each DC In aDT.Columns
                        obj = DR(DC.ColumnName)
                        val = "" : If Not IsDBNull(obj) Then val = obj.ToString()
                        sb.Append("<td")
                        If aUseFieldNumberAsCssClass Then
                            sb.Append(" class=""c") : sb.Append(colCount.ToString) : sb.Append("""")
                        End If
                        sb.Append(">")
                        sb.Append(val)
                        sb.Append("</td>")
                        colCount += 1
                    Next

                    sb.Append("</tr>")

                    rowCount += 1

                    If aNumRowsToShow > 0 AndAlso rowCount >= aNumRowsToShow Then Exit For

                Next

            Else
                sb.Append("<tr><td colspan=""") : sb.Append(aDT.Columns.Count.ToString()) : sb.Append(""">") : sb.Append(aEmptyTableMessage) : sb.Append("</td></tr>")
            End If

            sb.Append("</tbody>")
            sb.Append("</table>")

            Return sb.ToString()

        Catch ex As Exception
            Return "<table><tr><td>Error: " & ex.Message & "</td></tr></table>"
        End Try

    End Function

    Public Shared Function GetKendallTheilTrendStatistics(ByVal aXValues() As Double, ByVal aYValues() As Double, Optional ByRef aMsg As String = "") As Double()

        ' Example X Values: 2003,2004,2005,2006,2007,2008,2009,2010,2011,2012
        ' Example Y Values: 17.1979,12.2746,12.899,17.1795,14.8332,19.8577,11.3793,9.9303,8.2203,11.4125

        ' This method is based upon the approach described in this resource:
        ' Helsel, D.R. and R. M. Hirsch, 2002. Statistical Methods in Water Resources Techniques of Water Resources Investigations, Book 4, chapter A3. U.S. Geological Survey. 522 pages.
        ' Specifically, section 8.2: pages 212-216, section 10.1: various excerpts, and Answers to questions from chapter 10, pages 484-485, and Table B8 in the appendix.
        ' The URL for the above resource is: http://pubs.usgs.gov/twri/twri4a3/pdf/twri4a3-new.pdf

        If aXValues Is Nothing OrElse aXValues.Length = 0 Then
            aMsg = "Error: The required aXValues argument was empty."
            Return Nothing
        ElseIf aYValues Is Nothing OrElse aYValues.Length = 0 Then
            aMsg = "Error: The required aYValues argument was empty."
            Return Nothing
        ElseIf aXValues.Length <> aYValues.Length Then
            aMsg = "Error: The X-values array has a different length (" & aXValues.Length.ToString() & ") than the Y-values array (" & aYValues.Length.ToString() & ") - they must be equal in length."
            Return Nothing
        End If

        Dim aryStats(7) As Double ' Slope,P-value,P,M,T,KendallS
        Dim diffX As Double
        Dim diffY As Double
        Dim i As Integer, j As Integer
        Dim idx As Integer = 0
        Dim KendallS As Integer = -999
        Dim lookupTableIndex As System.Tuple(Of Integer, Integer)
        Dim numTies As Integer = 0
        Dim numXValues As Integer = aXValues.Length
        Dim numYValues As Integer = aYValues.Length
        Dim P As Integer = 0, M As Integer = 0, T As Integer = 0
        Dim sigmaS As Double = -999.0
        Dim slope As Double
        Dim slopes() As Double = Nothing
        Dim sb As New System.Text.StringBuilder()
        Dim sumInner As Integer = 0
        Dim ties As New Hashtable()
        Dim tieExtents As New Hashtable()
        Dim zetaS As Double = -999.0

        Try

            For i = 0 To numXValues - 2 ' Loop over the first through the second-to-last value
                For j = i + 1 To numXValues - 1 ' Loop over the i+1th through the last value
                    diffX = aXValues(j) - aXValues(i)
                    diffY = aYValues(j) - aYValues(i)
                    If diffY > 0 Then
                        P += 1 ' This is a "concordant pair"
                    ElseIf diffY < 0 Then
                        M += 1 ' This is a "discordant pair"
                    Else ' diffY = 0
                        T += 1 ' This is a tie
                        If aYValues.Length > 10 Then ' Keep track of ties and their extent for the large-sample approach, if needed
                            If ties.ContainsKey(aYValues(i)) Then
                                ties(aYValues(i)) = CInt(ties(aYValues(i)) + 1)
                            Else
                                ties(aYValues(i)) = CInt(1)
                            End If
                        End If
                    End If
                    slope = Double.NaN ' Assign NaN by default, i.e. undefined slopes will have this value
                    If diffX <> 0 Then slope = diffY / diffX ' Assign the slope, if the denominator is non-zero. Otherwise, keep NaN
                    If slopes Is Nothing Then ReDim slopes(0) Else ReDim Preserve slopes(UBound(slopes) + 1) ' Increase the length of the slope() array
                    slopes(UBound(slopes)) = slope ' Add this slope to the slope() array
                    idx = idx + 1
                Next
            Next

            KendallS = System.Math.Abs(P - M) ' Kendall's S statistic is defined as the absolute value of the difference between the concordant and discordant value pairs

            Array.Sort(slopes) ' Sort the slopes in preparation for finding the median slope

            aryStats(0) = ArrayMedianDouble(slopes) ' The median slope
            aryStats(2) = P ' The number of concordant pairs
            aryStats(3) = M ' The number of discordant pairs
            aryStats(4) = T ' The number of ties
            aryStats(5) = KendallS ' Kendall's 'S' statistic
            aryStats(6) = sigmaS ' SigmaS for the large sample approach
            aryStats(7) = zetaS ' ZetaS for the large sample approach

            If aYValues.Length <= 10 Then ' The number of data points is 10 or less

                lookupTableIndex = New System.Tuple(Of Integer, Integer)(numXValues, KendallS)

                If KendallPValueLookup.ContainsKey(lookupTableIndex) Then

                    aryStats(1) = KendallPValueLookup(lookupTableIndex)

                Else ' If the index combination of numXValues and KendallS is not found in the lookup table, then use this alternative approach

                    'From TRC Environmental Corporation | Tecumseh Products Company, 2012. Statistical Evaluation of Groundwater Stability, RCRA-05-2010-0012.
                    'URL: http://www.epa.gov/region5/cleanup/rcra/tecumseh/pdfs/20120621_groundwater_stability.pdf
                    '"In some cases where the data set included 10 or fewer values, tied values within the data set resulted in a
                    'Mann-Kendall statistic (S) that is not included in the tabulated values (Appendix A). In these
                    'cases the results of the Mann-Kendall trend tests were manually approximated by
                    'conservatively comparing the tabulated significance level of the next lower Mann-Kendall
                    'statistic (|S|-1) to the 95-percent (one-tailed) confidence level (i.e. an α of 0.05)."

                    KendallS = KendallS - 1 ' Use a value for KendallS that is one less than the calculated KendallS

                    'If subtracting one from KendallS causes the index to fall outside the lookup table, then use the first entry in the lookup table for N
                    If Array.IndexOf(New Integer() {4, 5, 8, 9}, numXValues) >= 0 And KendallS < 0 Then
                        KendallS = 0
                    ElseIf Array.IndexOf(New Integer() {3, 6, 7, 10}, numXValues) >= 0 And KendallS < 1 Then
                        KendallS = 1
                    End If

                    lookupTableIndex = New System.Tuple(Of Integer, Integer)(numXValues, KendallS)

                    If KendallPValueLookup.ContainsKey(lookupTableIndex) Then
                        aryStats(1) = KendallPValueLookup(lookupTableIndex)
                    Else
                        Throw New System.Exception("An element corresponding to numDataPoints = " & numXValues.ToString() & ", P = " & P.ToString() & ", M = " & M.ToString() & ", T = " & T.ToString() & ", and KendallS = " & System.Math.Abs(KendallS) & " does not exist in the standard KendallPValueLookup table.")
                    End If

                End If

            Else ' The number of data points is greater than 10

                For i = 1 To numXValues
                    tieExtents.Add(i, 0) ' Initialize the lookup table of tie "extents"
                Next

                ' Populate the lookup table of tie "extents"
                For Each key As Object In ties.Keys
                    numTies = CInt(ties.Item(key))
                    tieExtents(numTies) = CInt(tieExtents(numTies)) + 1
                Next

                ' Perform the "inner" sum of the equation (the 2nd half of the equation beneath the square root symbol
                For i = 1 To numXValues
                    sumInner = tieExtents(i) * i * (i - 1) * ((2 * i) + 5)
                Next

                ' Calculate sigmaS
                sigmaS = System.Math.Sqrt(((numXValues * (numXValues - 1) * ((2 * numXValues) + 5)) - sumInner) / 18)

                ' Calculate zetaS
                If KendallS > 0 Then
                    zetaS = (KendallS - 1) / sigmaS
                ElseIf KendallS < 0 Then
                    zetaS = (KendallS + 1) / sigmaS
                Else ' KendallS = 0
                    zetaS = 0
                End If

                zetaS = System.Math.Round(zetaS, 2)

                aryStats(6) = sigmaS
                aryStats(7) = zetaS

                ' Lookup the proper value in the "standard normal distribution" table using zetaS
                If StandardNormalDistribution.ContainsKey(zetaS) Then
                    aryStats(1) = 2 * (1 - StandardNormalDistribution(zetaS))
                Else
                    If zetaS > 3.49 Then zetaS = 3.49
                    aryStats(1) = 2 * (1 - StandardNormalDistribution(zetaS))
                End If

            End If

            Return aryStats

        Catch ex As Exception
            aMsg = ex.Message
            Return Nothing
        End Try

    End Function

    Public Shared Function GetRandomKey(ByVal aSupplementalKey As String, ByVal aNumAuxDigits As Integer) As String

        Dim dateNow As Date = Date.Now()

        Dim uniqueKey As String =
            Date.Now().Year().ToString() &
            Date.Now().Month().ToString() &
            Date.Now().Day().ToString() &
            Date.Now().Hour().ToString() &
            Date.Now().Minute().ToString() &
            Date.Now().Second().ToString() &
            Date.Now().Millisecond.ToString() &
            GetRandomString(6) &
            aSupplementalKey

        Return uniqueKey

    End Function

    Public Shared Function GetRandomKey2(Optional ByVal aNumAuxDigits As Integer = 0, Optional ByVal aSupplementalKey As String = Nothing) As String

        Dim dateNow As Date = Date.Now()
        Dim ran1 As New System.Random(dateNow.Millisecond)
        Dim sb As New StringBuilder()

        sb.Append(dateNow.Second().ToString())
        sb.Append(dateNow.Millisecond.ToString())

        If aNumAuxDigits <> 0 Then sb.Append(GetRandomString(aNumAuxDigits))
        If Not String.IsNullOrEmpty(aSupplementalKey) Then sb.Append(aSupplementalKey)

        Return sb.ToString()

    End Function

    Public Shared Function GetRandomString(ByVal aLength As Integer) As String

        Dim ran1 As New System.Random(Date.Now().Millisecond), ran2 As New System.Random(Date.Now().Millisecond)
        Dim sb As New StringBuilder()
        Dim bin As Integer, asciicode As Integer, i As Integer

        For i = 0 To aLength - 1

            bin = ran1.Next(1, 4) ' Generates a number between 1 and 3, including 1 and 3 (upper bound - 4 - is *exclusive*)

            Select Case bin
                Case 1 : asciicode = ran2.Next(48, 57) ' ASCII range for digits 0-9
                Case 2 : asciicode = ran2.Next(65, 90) ' ASCII range for uppercase letters A-Z
                Case 3 : asciicode = ran2.Next(97, 122) ' ASCII range for lowercase letters a-z
                Case Else : asciicode = ran2.Next(65, 90) ' Should never be executed
            End Select

            sb.Append(Chr(asciicode))

        Next

        Return sb.ToString()

    End Function

    Public Shared Function GetTheilRegressionValues(
                ByRef aSourceTable As DataTable,
                ByVal aVarXColumnName As String,
                ByVal aVarYColumnName As String,
                ByVal aVarZColumnName As String,
                Optional ByVal aRequiredPercentageDataCompleteness As Double = 60.0,
                Optional ByVal aMaxNumberOfContiguousMissingValues As Nullable(Of Integer) = Nothing,
                Optional ByVal aVarZColumnValueFilter As Specialized.StringCollection = Nothing,
                Optional ByVal aVarZColumnToken As String = "#TLS",
                Optional ByVal aTrendValueDecimalPlaces As Integer = -1,
                Optional ByRef aMessage As String = Nothing,
                Optional ByVal aTableName As String = Nothing) As DataTable

        ' Since this method returns a specifically-formatted table, we construct the return table before doing anything else
        ' so that the expected structure can be returned (even if empty) in the case of an early error.

        If String.IsNullOrEmpty(aTableName) Then aTableName = "TheilStats_" & GetRandomKey(Date.Now().Millisecond.ToString(), 6)

        Dim dtStats As New DataTable(aTableName)

        dtStats.Columns.Add(New DataColumn("Parameter", GetType(String)))
        dtStats.Columns.Add(New DataColumn("NumXValues", GetType(Integer)))
        dtStats.Columns.Add(New DataColumn("NumYValues", GetType(Integer)))
        dtStats.Columns.Add(New DataColumn("NumValid", GetType(Integer)))
        dtStats.Columns.Add(New DataColumn("TrendSlope", GetType(Double)))
        dtStats.Columns.Add(New DataColumn("TrendPValue", GetType(Double)))
        dtStats.Columns.Add(New DataColumn("P", GetType(Integer)))
        dtStats.Columns.Add(New DataColumn("M", GetType(Integer)))
        dtStats.Columns.Add(New DataColumn("T", GetType(Integer)))
        dtStats.Columns.Add(New DataColumn("S", GetType(Integer)))
        dtStats.Columns.Add(New DataColumn("SigmaS", GetType(Double)))
        dtStats.Columns.Add(New DataColumn("ZetaS", GetType(Double)))
        dtStats.Columns.Add(New DataColumn("XValues", GetType(String)))
        dtStats.Columns.Add(New DataColumn("YValues", GetType(String)))
        dtStats.Columns.Add(New DataColumn("Status", GetType(String)))
        dtStats.Columns.Add(New DataColumn("Message", GetType(String)))

        Dim B0 As Double
        Dim count As Integer = 0
        Dim DTVarZ As DataTable
        Dim DS As DataSet
        Dim dr As DataRow, drNew As DataRow
        Dim dc As DataColumn
        Dim drAry() As DataRow
        Dim drStats As DataRow = Nothing
        Dim firstNonEmptyDataValueIndex As Integer = -1
        Dim i As Integer
        Dim idx As Integer
        Dim lastNonEmptyDataValueIndex As Integer = -1
        Dim nullValueSubstitute As Double = -999.0
        Dim msg As String = ""
        Dim numMissingValues As Integer = 0
        Dim numYValues As Integer = 0
        Dim numYValuesClean As Integer = 0
        Dim param As String = ""
        Dim processTable As Boolean = True
        Dim sb As New StringBuilder()
        Dim template1 As New XpFormatTemplate()
        Dim TrendStats() As Double
        Dim val As String = Nothing
        Dim valDbl As Double
        Dim xvalMin As Double
        Dim xvalMax As Double
        Dim XValues() As Double, XValues_Cleaned() As Double = Nothing
        Dim YValues() As Double, YValues_Cleaned() As Double = Nothing
        Dim yfit As Double
        Dim yr As Integer

        Try

            If aSourceTable Is Nothing OrElse aSourceTable.Rows Is Nothing OrElse aSourceTable.Rows.Count = 0 Then
                Throw New System.Exception("Required argument SourceTable was empty. Cannot calculate Theil regression.")
            End If

            If String.IsNullOrEmpty(aVarXColumnName) Then
                Throw New System.Exception("Required argument VarXColumnName was empty. Cannot calculate Theil regression.")
            End If

            If String.IsNullOrEmpty(aVarYColumnName) Then
                Throw New System.Exception("Required argument VarYColumnName was empty. Cannot calculate Theil regression.")
            End If

            If String.IsNullOrEmpty(aVarZColumnName) Then
                Throw New System.Exception("Required argument VarZColumnName was empty. Cannot calculate Theil regression.")
            End If

            If aRequiredPercentageDataCompleteness <= 0.0 Then
                Throw New System.Exception("Optional argument aRequiredPercentageDataCompleteness must be greater than zero. Cannot calculate Theil regression.")
            End If

            If aMaxNumberOfContiguousMissingValues.HasValue AndAlso aMaxNumberOfContiguousMissingValues < 0 Then
                Throw New System.Exception("Optional argument aMaxNumberOfContiguousMissingValues must be greater than zero. Cannot calculate Theil regression.")
            End If

            DS = SplitDataTableByKey(aSourceTable, aVarZColumnName, aVarZColumnName)

            For Each DTVarZ In DS.Tables

                processTable = True : If Not aVarZColumnValueFilter Is Nothing AndAlso aVarZColumnValueFilter.Count > 0 AndAlso Not aVarZColumnValueFilter.Contains(DTVarZ.TableName) Then processTable = False

                If Not processTable Then Continue For

                drStats = dtStats.NewRow()

                dtStats.Rows.InsertAt(drStats, dtStats.Rows.Count)

                If DTVarZ.Rows.Count > 0 Then
                    param = DTVarZ.Rows(0)(aVarZColumnName)
                Else
                    param = DTVarZ.TableName
                End If

                drStats("Parameter") = param
                drStats("NumXValues") = 0
                drStats("NumYValues") = 0
                drStats("NumValid") = 0
                drStats("Status") = "Pending"

                ' The following check is based upon the comment of "A sample size of five XY points is, algebraically, the minimum sample size that will produce meaningful ranks."
                ' found in the following paper: http://pubs.usgs.gov/tm/2006/tm4a7/pdf/TM4-A7_KTRLine_report.pdf
                If DTVarZ.Rows.Count < 5 Then
                    aMessage = "A Theil regression cannot be calculated because the minimum required sample size is 5 data records and only " & DTVarZ.Rows.Count.ToString() & " records are available."
                    drStats("Status") = "Failed"
                    drStats("Message") = aMessage
                    Continue For
                End If

                ' Get arrays of numeric (double) values from the "X" and "Y" columns
                XValues = DataAnalysis.ArrayFromColumnValues_Double(DTVarZ, aVarXColumnName, False, nullValueSubstitute)
                YValues = DataAnalysis.ArrayFromColumnValues_Double(DTVarZ, aVarYColumnName, False, nullValueSubstitute)

                numYValues = YValues.Length

                '-------------------------------------------------->>>
                ' Find and remove any empty data values from the X and Y arrays in preparation for performing the trend calculations
                '-------------------------------------------------->>>

                idx = 0
                For Each valDbl In YValues
                    If valDbl <> nullValueSubstitute Then
                        If XValues_Cleaned Is Nothing Then ReDim XValues_Cleaned(0) Else ReDim Preserve XValues_Cleaned(UBound(XValues_Cleaned) + 1)
                        If YValues_Cleaned Is Nothing Then ReDim YValues_Cleaned(0) Else ReDim Preserve YValues_Cleaned(UBound(YValues_Cleaned) + 1)
                        XValues_Cleaned(UBound(XValues_Cleaned)) = XValues(idx)
                        YValues_Cleaned(UBound(YValues_Cleaned)) = valDbl
                        If firstNonEmptyDataValueIndex = -1 Then firstNonEmptyDataValueIndex = idx
                        lastNonEmptyDataValueIndex = idx
                    End If
                    idx += 1
                Next

                numYValuesClean = 0
                If Not YValues_Cleaned Is Nothing AndAlso YValues_Cleaned.GetLength(0) > 0 Then numYValuesClean = YValues_Cleaned.GetLength(0)

                drStats("NumXValues") = XValues.GetLength(0)
                drStats("NumYValues") = YValues.GetLength(0)
                drStats("NumValid") = numYValuesClean
                drStats("XValues") = String.Join(",", XValues)
                drStats("YValues") = String.Join(",", YValues)

                If numYValuesClean = 0 Then
                    aMessage = "A Theil regression cannot be calculated because there are no valid data values."
                    drStats("Status") = "Failed"
                    drStats("Message") = aMessage
                    Continue For
                End If

                If YValues(UBound(YValues)) = nullValueSubstitute Then
                    aMessage = "A Theil regression cannot be calculated because the final year of data is missing."
                    drStats("Status") = "Failed"
                    drStats("Message") = aMessage
                    Continue For
                End If

                '-------------------------------------------------->>>
                ' Determine whether the data exceeds the "maximum number of contiguous missing values" threshold
                '-------------------------------------------------->>>

                If aMaxNumberOfContiguousMissingValues.HasValue Then ' A check needs to be done only if the caller has provided a value

                    i = 0
                    While i < YValues.Length And numMissingValues <= aMaxNumberOfContiguousMissingValues ' Traverse the array until the end is reached or the threshold is exceeded
                        If YValues(i) = nullValueSubstitute Then
                            numMissingValues += 1 ' Keep a count of contiguous empty values
                        Else
                            numMissingValues = 0 ' Set the count back to zero as soon as a non-empty value is encountered
                        End If
                        i += 1
                    End While

                    If numMissingValues > aMaxNumberOfContiguousMissingValues Then
                        aMessage = "A Theil regression cannot be calculated because there are more contiguous missing values than the maximum allowed (" & aMaxNumberOfContiguousMissingValues.ToString() & ")."
                        drStats("Status") = "Failed"
                        drStats("Message") = aMessage
                        Continue For
                    End If

                End If

                '-------------------------------------------------->>>
                ' Determine whether the data passes or fails the "required percentage data completeness" criteria
                '-------------------------------------------------->>>

                valDbl = (numYValuesClean / numYValues) * 100.0

                If valDbl < aRequiredPercentageDataCompleteness Then
                    aMessage = valDbl & "% of the data is non-empty, but " & aRequiredPercentageDataCompleteness.ToString() & "% completeness is required to calculate a Theil regression."
                    drStats("Status") = "Failed"
                    drStats("Message") = aMessage
                    Continue For
                End If

                TrendStats = GetKendallTheilTrendStatistics(XValues_Cleaned, YValues_Cleaned, msg)

                If TrendStats Is Nothing OrElse InStr(msg, "Error") > 0 Then
                    aMessage = msg
                    drStats("Status") = "Error"
                    drStats("Message") = aMessage
                    Continue For
                End If

                'get the yintercept by forcing the yfit median to the observed median
                B0 = DataAnalysis.ArrayMedianDouble(YValues_Cleaned) - TrendStats(0) * DataAnalysis.ArrayMedianDouble(XValues_Cleaned)

                drStats("TrendSlope") = System.Math.Round(TrendStats(0), 3)
                drStats("TrendPValue") = System.Math.Round(TrendStats(1), 3)
                drStats("P") = CInt(TrendStats(2))
                drStats("M") = CInt(TrendStats(3))
                drStats("T") = CInt(TrendStats(4))
                drStats("S") = CInt(TrendStats(5))
                drStats("SigmaS") = CDbl(TrendStats(6))
                drStats("ZetaS") = CDbl(TrendStats(7))
                drStats("Status") = "Valid"

                xvalMin = DataAnalysis.ArrayMinDouble(XValues_Cleaned)
                xvalMax = DataAnalysis.ArrayMaxDouble(XValues_Cleaned)

                'drAry = DTMain.Select("Year = " & CInt(xvalMax).ToString())
                drAry = DTVarZ.Select(aVarXColumnName & " = " & CInt(xvalMax).ToString())

                If drAry.Length <> 1 Then
                    msg = IIf(drAry.Length = 0, "Could not locate the last data row by the maximum X axis value. The regression cannot be calculated.", "There is more than one row of data per X axis value in the base dataset. The regression cannot be calculated.")
                    aMessage = msg
                    drStats("Status") = "Error"
                    drStats("Message") = aMessage
                    Continue For
                End If

                idx = 0
                For Each dr In DTVarZ.Rows

                    yr = dr("Year")

                    drNew = aSourceTable.NewRow()

                    yfit = TrendStats(0) * dr("Year") + B0

                    For Each dc In aSourceTable.Columns

                        drNew(dc.ColumnName) = DTVarZ.Rows(0)(dc.ColumnName)
                        drNew(aVarZColumnName) = DTVarZ.Rows(0)(aVarZColumnName) & aVarZColumnToken
                        drNew("Date") = "7/2/" & yr.ToString() & " 12:00:00 AM"
                        drNew("Year") = yr

                        If idx >= firstNonEmptyDataValueIndex And idx <= lastNonEmptyDataValueIndex Then
                            If aTrendValueDecimalPlaces >= 0 Then
                                drNew("Value") = System.Math.Round(yfit, aTrendValueDecimalPlaces)
                            Else
                                drNew("Value") = yfit
                            End If
                        Else
                            drNew("Value") = DBNull.Value
                        End If

                    Next

                    aSourceTable.Rows.InsertAt(drNew, aSourceTable.Rows.Count)

                    idx += 1

                Next

            Next

            Return dtStats

        Catch ex As Exception
            aMessage = ex.Message
            If drStats Is Nothing Then
                drStats = dtStats.NewRow()
                drStats("Parameter") = param
            End If
            drStats("Status") = "Error"
            drStats("Message") = aMessage
            Return dtStats
        End Try

    End Function

    Public Shared Function GetXmlNode(ByVal XmlDocPath As String, ByVal XPathExpression As String, Optional ByRef Message As String = "AOK") As XmlNode

        Dim xnodes As XmlNodeList = DataAnalysis.GetXmlNodes(XmlDocPath, XPathExpression, Message)

        If xnodes Is Nothing OrElse xnodes.Count = 0 Then
            Message = "No XmlNodes were found in XmlDocument """ & XmlDocPath & """ for XPath expression """ & XPathExpression & """." & IIf(Not String.IsNullOrEmpty(Message), " (" & Message & ")", "")
            Return Nothing
        End If

        Return xnodes.Item(0)

    End Function

    Public Shared Function GetXmlNodes(ByVal XmlDocPath As String, ByVal XPathExpression As String, Optional ByRef Message As String = "AOK") As XmlNodeList

        Dim ThisModuleKey As String = "XpXml.GetXmlNodes"

        If String.IsNullOrEmpty(XmlDocPath) Then
            Message = "Expected argument 'XmlDocPath' was empty. (Source: " & ThisModuleKey & "())"
            Return Nothing
        End If

        If String.IsNullOrEmpty(XPathExpression) Then
            Message = "Expected argument 'XPathExpression' was empty. (Source: " & ThisModuleKey & "())"
            Return Nothing
        End If

        Dim xmldoc As New XmlDocument
        Dim xnodes As XmlNodeList
        Dim fi As FileInfo

        Try

            fi = New FileInfo(XmlDocPath)

            If Not fi.Exists Then
                Message = "XML file """ & XmlDocPath & """ does not exist. (Source: " & ThisModuleKey & "())"
                Return Nothing
            End If

            xmldoc.Load(XmlDocPath)

            xnodes = xmldoc.SelectNodes(XPathExpression)

            If Not xnodes Is Nothing Then
                Return xnodes
            Else
                Message = "No XmlNode found for XPath expression """ & XPathExpression & """. (Source: " & ThisModuleKey & "())"
                Return Nothing
            End If

            Return Nothing

        Catch ex As Exception
            Message = ex.Message & " (Source: " & ThisModuleKey & "())"
            Return Nothing
        End Try

    End Function

    Public Shared Function GetXmlNodeInnerText(ByVal XmlDocPath As String, ByVal XPathExpression As String, Optional ByRef Message As String = "AOK") As String
        Dim xnode As XmlNode = DataAnalysis.GetXmlNode(XmlDocPath, XPathExpression, Message)
        If Not xnode Is Nothing Then Return xnode.InnerText
        Return Nothing
    End Function

    Public Shared Function JoinStringCollection(ByVal aSC As Specialized.StringCollection, Optional ByVal aDelimiter As String = ",") As String

        Dim count As Integer = 0
        Dim sb As New StringBuilder()
        Try

            For Each str As String In aSC
                If count > 0 Then sb.Append(aDelimiter)
                sb.Append(str)
                count += 1
            Next

            Return sb.ToString()

        Catch ex As Exception
            Return ex.Message
        End Try

    End Function

    Public Shared Function ResolveCommandText(ByVal aCommandTextID As String, ByVal aCommandFileID As String, Optional ByVal aCommandTextXPathExpression As String = "//CommandText[@ID=""{CommandTextID}""]") As String

        'Return literal connection strings as-is
        If String.IsNullOrEmpty(aCommandFileID) Then Return aCommandTextID

        Dim cmdText As String
        Dim fi As FileInfo
        Dim physicalPath As String
        Dim virtualPath As String = ConfigurationManager.AppSettings(aCommandFileID)

        If String.IsNullOrEmpty(aCommandTextID) Then Throw New System.Exception("Required argument ""aCommandTextID"" was missing or empty. Cannot resolve command text.")

        If String.IsNullOrEmpty(virtualPath) Then Throw New System.Exception("Application setting """ & aCommandFileID & """ was not found in Web.config. Cannot resolve command text.")

        physicalPath = Hosting.HostingEnvironment.MapPath(virtualPath)
        fi = New FileInfo(physicalPath)
        If Not fi.Exists Then Throw New System.Exception("The file path specified by application setting """ & aCommandFileID & """ does not exist.")

        cmdText = GetXmlNodeInnerText(physicalPath, aCommandTextXPathExpression.Replace("{CommandTextID}", aCommandTextID))
        If String.IsNullOrEmpty(cmdText) Then Throw New System.Exception("The SQL command indicated by command ID """ & aCommandTextID & """ was empty or could not be resolved in SQL command file """ & physicalPath & """.")

        Return cmdText.Trim()

    End Function

    Public Shared Function ResolveCommandTextParameters(ByVal aCommandText As String, ByVal aParams As Hashtable, Optional ByRef aUnresolvedParameters As Specialized.StringCollection = Nothing) As String
        Dim template1 As New XpFormatTemplate()
        Dim cmdText As String = template1.Resolve(aCommandText, aParams, aUnresolvedParameters, False)
        Return cmdText
    End Function

    Public Shared Function ResolveConnectionString(ByVal aConnectionString As String) As String

        'Remove any HTML tags from the connection string ID in order to guard against XSS (cross-site scripting) attempts
        Dim connStringCleaned As String = aConnectionString

        'If aConnectionString has a semicolon in it, then we assume that this is a literal connection string and simply return it as-is.
        'If there's no semicolon we assume that it's a "connection string ID" instead, and attempt to retrieve the literal connection
        'string from Web.config using the aConnectionString text as an ID.

        If InStr(connStringCleaned, ";") > 0 Then Return connStringCleaned

        Dim connStringSettings As ConnectionStringSettings
        Dim connStringID As String = connStringCleaned
        Dim connString As String

        '-------------------->>>
        ' Determine if the argument "ConnectionString" is being used an identifier (key) for a connection
        ' string defined in a Web.config file.
        '-------------------->>>
        connStringSettings = ConfigurationManager.ConnectionStrings(connStringID)
        If connStringSettings Is Nothing Then Throw New System.Exception("Connection string """ & connStringID & """ was not found in Web.config.")
        connString = connStringSettings.ToString()
        If String.IsNullOrEmpty(connString) Then Throw New System.Exception("Connection string """ & connStringID & """ was empty.")

        Return connString

    End Function

    Public Shared Function SplitDataTableByKey(
              ByRef aSourceDT As DataTable,
              ByVal aKeyColumn As String,
              Optional ByVal aTableNameColumn As String = Nothing) As DataSet

        If aSourceDT Is Nothing OrElse aSourceDT.Rows Is Nothing OrElse aSourceDT.Rows.Count = 0 Then
            Return New DataSet("Empty")
        End If

        Dim count As Integer = 0
        Dim dsNew As New DataSet()
        Dim dtNew As New DataTable()
        Dim drSource As DataRow
        Dim keyColumnExists As Boolean = False
        Dim keyValueCurrent As String = ""
        Dim tableName As String

        If Not String.IsNullOrEmpty(aTableNameColumn) Then keyColumnExists = aSourceDT.Columns.Contains(aTableNameColumn)

        For Each drSource In aSourceDT.Rows

            If drSource(aKeyColumn).ToString() <> keyValueCurrent Or count = 0 Then

                If keyColumnExists Then
                    tableName = drSource(aTableNameColumn).ToString().Trim()
                Else
                    tableName = count.ToString()
                End If

                dtNew = aSourceDT.Clone()
                dtNew.TableName = tableName
                dsNew.Tables.Add(dtNew)

                keyValueCurrent = drSource(aKeyColumn).ToString()

                count += 1

            End If

            dtNew.ImportRow(drSource)

        Next

        Return dsNew

    End Function

End Class
