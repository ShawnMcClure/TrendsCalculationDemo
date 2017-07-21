Imports System.Collections.Specialized
Imports System.Text.RegularExpressions

Public Class XpFormatTemplate

    Protected _BeginTag As String = "{"
    Protected _EndTag As String = "}"
    Protected _Pattern As String = "\XP_BEGIN_TAG(?<key>\w+(\.\w+)*)(?<key>:[^\XP_END_TAG]*)?\XP_END_TAG"

#Region "Properties"

    Public ReadOnly Property BeginTag() As String
        Get
            Return _BeginTag
        End Get
    End Property

    Public ReadOnly Property EndTag() As String
        Get
            Return _EndTag
        End Get
    End Property

#End Region

#Region "Constructors and Initializers"

    Public Sub New()
        _Pattern = _Pattern.Replace("XP_BEGIN_TAG", _BeginTag).Replace("XP_END_TAG", _EndTag)
    End Sub

    Public Sub New(ByVal aBeginTag As String, ByVal aEndTag As String)
        If Not String.IsNullOrEmpty(aBeginTag) Then _BeginTag = aBeginTag
        If Not String.IsNullOrEmpty(aEndTag) Then _EndTag = aEndTag
        _Pattern = _Pattern.Replace("XP_BEGIN_TAG", _BeginTag).Replace("XP_END_TAG", _EndTag)
    End Sub

#End Region

#Region "Methods"

    Public Function GetKeys(ByVal aText As String) As StringCollection

        Dim sc As New StringCollection()
        Dim matches As MatchCollection

        Try

            matches = (New Regex(_Pattern)).Matches(aText)

            For Each m As Match In matches
                sc.Add(m.Value.Replace(Me.BeginTag, "").Replace(Me.EndTag, ""))
            Next

            Return sc

        Catch ex As Exception
            Return sc
        End Try

    End Function

    Public Function GetTokens(ByVal aText As String) As Hashtable

        Dim ary() As String
        Dim ht As New Hashtable()
        Dim key As String
        Dim matches As MatchCollection
        Dim val As String

        Try

            If String.IsNullOrEmpty(aText) Then Return ht

            matches = (New Regex(_Pattern)).Matches(aText)

            For Each m As Match In matches

                key = m.Value.Replace(Me.BeginTag, "").Replace(Me.EndTag, "")
                val = ""

                If InStr(key, ":") > 0 Then
                    ary = key.Split(":")
                    If ary.Length >= 2 Then
                        key = ary(0)
                        val = ary(1)
                    End If
                End If

                If ht.ContainsKey(key) Then
                    ht(key) = val
                Else
                    ht.Add(key, val)
                End If

            Next

            Return ht

        Catch ex As Exception
            Return ht
        End Try

    End Function

    Public Function GetMatches(ByVal aText As String) As MatchCollection
        If String.IsNullOrEmpty(aText) Then Return Nothing
        Return (New Regex(_Pattern)).Matches(aText)
    End Function

    Public Function GetUnresolvedTokens(ByVal aText As String) As StringCollection

        Dim sc As New StringCollection()
        Dim matches As MatchCollection

        Try

            matches = (New Regex(_Pattern)).Matches(aText)

            For Each m As Match In matches
                sc.Add(m.Value.Replace(_BeginTag, "").Replace(_EndTag, ""))
            Next

            Return sc

        Catch ex As Exception
            Return sc
        End Try

    End Function

    Public Function IsResolved(ByVal aText As String) As Boolean

        Dim exp As New Regex(_Pattern)
        Dim matches As MatchCollection

        Try

            matches = exp.Matches(aText)

            Return (matches.Count = 0)

        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function Resolve( _
            ByVal aText As String, _
            ByVal aSubstitutions As Hashtable, _
            Optional ByRef aUnresolvedTokens As StringCollection = Nothing, _
            Optional ByRef aTreatEmptySubstitutionsAsValid As Boolean = True, _
            Optional ByRef aResolutionSteps As ArrayList = Nothing) As String

        Dim count As Integer = 0
        Dim result As String = aText
        Dim resultPrev As String = aText
        Dim sc As New StringCollection()

        Try

            If String.IsNullOrEmpty(aText) Then Return aText

            If Not aResolutionSteps Is Nothing Then aResolutionSteps.Add(result)

            If Not aSubstitutions Is Nothing Then
                For Each key As Object In aSubstitutions.Keys
                    result = Me.Resolve(result, key, aSubstitutions.Item(key), aTreatEmptySubstitutionsAsValid)
                    If Not aResolutionSteps Is Nothing AndAlso (count = 0 Or result <> resultPrev) AndAlso Not aResolutionSteps.Contains(result) Then aResolutionSteps.Add(result)
                    resultPrev = result
                    count += 1
                Next
            End If

            If Not Me.IsResolved(result) Then result = Me.ResolveDefaults(result)

            sc = GetUnresolvedTokens(result)

            If aUnresolvedTokens Is Nothing Then aUnresolvedTokens = New StringCollection()

            For Each str As String In sc
                If Not aUnresolvedTokens.Contains(str) Then aUnresolvedTokens.Add(str)
            Next

            Return result

        Catch ex As Exception
            Return "ERROR: " & ex.Message
        End Try

    End Function

    Public Overridable Function Resolve( _
            ByVal aText As String, _
            ByVal aKey As String, _
            ByVal aValue As String, _
            Optional ByRef aTreatEmptySubstitutionsAsValid As Boolean = True) As String

        Dim result As String
        Dim matches As MatchCollection
        Dim key As String
        Dim ary() As String

        Try

            If String.IsNullOrEmpty(aText) Or String.IsNullOrEmpty(aKey) Then Return aText

            result = aText

            matches = (New Regex(_Pattern)).Matches(result)

            For Each m As Match In matches

                key = m.Value.Replace(Me.BeginTag, "").Replace(Me.EndTag, "")

                If InStr(key, ":") > 0 Then
                    ary = key.Split(":")
                    If ary.Length >= 2 Then key = ary(0)
                End If

                If key.ToLower() = aKey.ToLower() Then

                    If Not String.IsNullOrEmpty(aValue) OrElse aTreatEmptySubstitutionsAsValid Then
                        result = result.Replace(m.Value, aValue)
                    End If

                End If

            Next

            Return result

        Catch ex As Exception
            Return "ERROR: " & ex.Message
        End Try

    End Function

    Public Function ResolveDefaults(ByVal aText As String) As String

        Dim result As String = aText
        Dim matches As MatchCollection
        Dim key As String
        Dim ary() As String

        Try

            If String.IsNullOrEmpty(aText) Then Return aText

            matches = (New Regex(_Pattern)).Matches(result)

            For Each m As Match In matches

                key = m.Value.Replace(_BeginTag, "").Replace(_EndTag, "")

                If InStr(key, ":") > 0 Then
                    ary = key.Split(":")
                    If ary.Length >= 2 Then result = result.Replace(m.Value, ary(1))
                End If

            Next

            Return result

        Catch ex As Exception
            Return "ERROR: " & ex.Message
        End Try

    End Function

    Public Function ResolveUntilDone( _
            ByVal aText As String, _
            ByVal aSubstitutions As Hashtable, _
            Optional ByRef aUnresolvedTokens As StringCollection = Nothing, _
            Optional ByRef aTreatEmptySubstitutionsAsValid As Boolean = True, _
            Optional ByRef aResolutionSteps As ArrayList = Nothing, _
            Optional ByRef aNumIterations As Integer = 0) As String

        If String.IsNullOrEmpty(aText) Then Return Nothing

        Dim doneResolvingText As Boolean = False
        Dim template As New XpFormatTemplate()
        Dim textToResolve As String = aText
        'Dim unresolvedKeys As Specialized.StringCollection = Nothing
        Dim prevTextToResolve As String = aText
        Dim prevUnresolvedKeysCount As Integer = 0
        Dim showSteps As Boolean = IIf(aResolutionSteps Is Nothing, False, True)

        Try

            If showSteps Then aResolutionSteps.Add(textToResolve)

            aNumIterations = 0

            While Not doneResolvingText

                If Not aUnresolvedTokens Is Nothing Then aUnresolvedTokens.Clear()

                textToResolve = template.Resolve(textToResolve, aSubstitutions, aUnresolvedTokens, True)

                If showSteps AndAlso textToResolve <> prevTextToResolve Then aResolutionSteps.Add(textToResolve)

                ' If (there are still tokens to be resolved) *OR* (1. the number of unresolved tokens AND 2. the base text are the same as last time through the loop),
                ' then we can consider the text to be as fully resolved as it's going to get with the current collection of parameters.
                If aUnresolvedTokens.Count = 0 Or ((aUnresolvedTokens.Count = prevUnresolvedKeysCount) And (textToResolve = prevTextToResolve)) Then
                    doneResolvingText = True
                End If

                prevUnresolvedKeysCount = aUnresolvedTokens.Count
                prevTextToResolve = textToResolve

                aNumIterations += 1

            End While

            Return textToResolve

        Catch ex As Exception
            Return "ERROR: " & ex.Message
        End Try

    End Function

#End Region

End Class
