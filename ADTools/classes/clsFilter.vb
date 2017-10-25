Imports System.Collections.ObjectModel

Public Class clsFilter

    Private _name As String
    Private _filter As String
    Private _pattern As String

    Sub New()

    End Sub

    Sub New(name As String, filter As String)
        _name = name
        _filter = filter
    End Sub

    Sub New(filter As String)
        _filter = filter
    End Sub

    Sub New(pattern As String,
            attributes As ObservableCollection(Of clsAttribute),
            searchobjectclasses As clsSearchObjectClasses)

        Me.Pattern = pattern

        Dim patterns() As String = If(Not String.IsNullOrEmpty(pattern), pattern.Split({"|", "/", vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries), {""})

        attributes = If(attributes, attributesForSearchDefault)
        searchobjectclasses = If(searchobjectclasses, New clsSearchObjectClasses)

        Dim attrfilter = ""
        attributes.ToList.ForEach(
            Sub(a As clsAttribute)
                patterns.ToList.ForEach(
                    Sub(p)
                        Dim freeprefix As Boolean = False
                        Dim freesuffix As Boolean = True

                        If p.Contains("*") Then
                            freeprefix = False
                            freesuffix = False
                        End If
                        If p.Contains("""") Or p.Contains("'") Then
                            p = p.Replace("""", "").Replace("'", "")
                            freeprefix = False
                            freesuffix = False
                        End If

                        attrfilter &= String.Format("({0}={1}{2}{3})", a.Name, If(freeprefix, "*", ""), Trim(p), If(freesuffix, "*", ""))
                    End Sub)
            End Sub)

        Me.Filter = "(&" +
                     "(|" +
                         If(searchobjectclasses.User, "(&(objectCategory=person)(objectClass=user)(!(objectClass=inetOrgPerson)))", "") +
                         If(searchobjectclasses.Contact, "(&(objectCategory=person)(objectClass=contact))", "") +
                         If(searchobjectclasses.Computer, "(objectClass=computer)", "") +
                         If(searchobjectclasses.Group, "(objectClass=group)", "") +
                         If(searchobjectclasses.OrganizationalUnit, "(objectClass=organizationalunit)", "") +
                     ")" +
                     If(Not String.IsNullOrEmpty(attrfilter), "(|" & attrfilter & ")", "") +
                  ")"
    End Sub


    Public Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Public Property Filter As String
        Get
            Return _filter
        End Get
        Set(value As String)
            _filter = value
        End Set
    End Property

    Public Property Pattern As String
        Get
            Return _pattern
        End Get
        Set(value As String)
            _pattern = value
        End Set
    End Property

End Class
