Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports System.DirectoryServices
Imports System.Text.RegularExpressions
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Data
Imports SSP.ActiveDirectoryX.Data.Repositories
Imports SSP.Data.StatementBuildersAD.Core
#End Region

Namespace Core

	Partial Public Class DistinguishedName

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private _baseProperties As DistinguishedNameBaseProperties
#End Region  '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Public Sub New(ByVal baseProperties As DistinguishedNameBaseProperties)
			_baseProperties = baseProperties
		End Sub

		'''' <summary>
		'''' Erstellt ein neues DistinguishedName-Objekt 
		'''' und setzt den DistinguishedName auf den Wert von value.
		'''' </summary>
		'Public Sub New(ByVal distinguishedName As String)
		'  'Me.Value = distinguishedName
		'End Sub
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public ReadOnly Property BaseProperties As DistinguishedNameBaseProperties
			Get
				Return _baseProperties
			End Get
		End Property

    ''' <summary>
    ''' Liefert den DistinguishedName.
    ''' Bsp: ou=test, dc=domain, dc=net
    ''' </summary>
    Public ReadOnly Property Value As String
			Get
				Return Me.BaseProperties.DistinguishedName
			End Get
		End Property

    ''' <summary>
    ''' Liefert den Namen des DistinguishedNames.
    ''' Bsp: ou=test, dc=domain, dc=net -> test
    ''' </summary>
    Public ReadOnly Property Name As String
			Get
				Try
					Return Me.SplitDistinguishedName.First.Split("="c).Last.Replace("\", "")
				Catch ex As Exception
					Return ""
				End Try
			End Get
		End Property

    ''' <summary>
    ''' Liefert den DistinguishedName des hierachisch höher angeordneten Objektes.
    ''' Bsp: ou=test, dc=domain, dc=net => dc=domain, dc=net
    ''' </summary>
    Public ReadOnly Property Parent As DistinguishedName
			Get
				Dim parts = Me.SplitDistinguishedName
				If parts.Count <= 1 Then
					Return Nothing
				Else
          '.RemoveAt(0)
          Dim distinguishedName = String.Join(",", parts.Skip(1))
					Dim bp = New DistinguishedNameBaseProperties(distinguishedName, "", "", New Guid, "")
					Return New DistinguishedName(bp)
				End If
			End Get
		End Property

    ''' <summary>
    ''' Liefert alle Kind-Objekte des DistinguishedName-Objektes.
    ''' </summary>
    Public ReadOnly Property Children As IEnumerable(Of DistinguishedName)
			Get
				Return GetChildren()
			End Get
		End Property

    ''' <summary>
    ''' Liefert sämtliche Werte der Eigenschaft ObjectClass.
    ''' </summary>
    Public ReadOnly Property ObjectClassValues As ObjectClasses()
			Get
				Return AdInformation.GetObjectClassValues(Me)
			End Get
		End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
		Private Function GetChildren() As List(Of DistinguishedName)

			Dim temp = New List(Of String)

			If AdRepositoryHelper.Instance.ExistPath(Me) Then
				For Each child As DirectoryEntry In Me.ToDirectoryEntry(False).Children
					temp.Add(child.Path.Replace("LDAP://", ""))
				Next child
			End If

			Return DistinguishedNameRepository.Instance.GetByDistinguishedNames(temp).ToList
		End Function
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		''' <summary>
		''' Liefert den DistinguishedName.
		''' Bsp: ou=test, dc=domain, dc=net
		''' </summary>
		Public Overrides Function ToString() As String

			Return Me.Name
		End Function

    ''' <summary>
    ''' Liefert den DistinguishedName als URL.
    ''' Bsp: LDAP://ou=test, dc=domain, dc=net
    ''' </summary>
    Public Function ToUrl() As String

			Return String.Concat(My.Settings.LdapProtocolName, Me.Value)
		End Function

    ''' <summary>
    ''' Liefert den relativen DistinguishedName.
    ''' Bsp: cn=test
    ''' </summary>
    Public Function ToRelativeDistinguishedName() As String

			Return Me.SplitDistinguishedName.FirstOrDefault
		End Function

    ''' <summary>
    ''' Liefert den relativen DistinguishedName ohne Maskierung.
    ''' Bsp: ou=test\, test, dc=domain, dc=net => ou=test, test, dc=domain, dc=net
    ''' </summary>
    Public Function ToDeEscapedRelativeDistinguishedName() As String

			Return DbResultAD.DeEscapeValue(Me.ToRelativeDistinguishedName)
		End Function

    ''' <summary>
    ''' Liefert den relativen Namen.
    ''' Bsp: ou=test, dc=domain, dc=net => test
    ''' </summary>
    Public Function ToRelativeName() As String

			Return Me.Name
		End Function

    ''' <summary>
    ''' Liefert den relativen Namen ohne Maskierung.
    ''' Bsp: ou=test\, test, dc=domain, dc=net => test, test
    ''' </summary>
    Public Function ToDeEscapedRelativeName() As String

			Return DbResultAD.DeEscapeValue(Me.ToRelativeName)
		End Function

    ''' <summary>
    ''' Zerlegt den DistinguishedName in seine Einzelteile und gibt diese zurück.
    ''' Bsp: ou=test, dc=domain, dc=net => Liste mit den 3 Items ou=test - dc=domain - dc=net
    ''' </summary>
    Public Shared Function SplitDistinguishedName _
		(ByVal distinguishedName As DistinguishedName) As List(Of String)

			Const pattern = "[^\\],"

			Dim parts = New List(Of String)
			Dim start = 0

			Try
				With distinguishedName
					For Each match As Match In Regex.Matches(.Value, pattern)
						parts.Add(.Value.Substring(start, (match.Index + 1) - start).Trim)
						start = match.Index + 2
					Next match
					parts.Add(.Value.Substring(start).Trim)
				End With
			Catch ex As Exception
				parts = New List(Of String)
			End Try

			Return parts
		End Function

    ''' <summary>
    ''' Zerlegt den DistinguishedName in seine Einzelteile und gibt diese zurück.
    ''' Bsp: ou=test, dc=domain, dc=net => Liste mit den 3 Items ou=test - dc=domain - dc=net
    ''' </summary>
    Public Function SplitDistinguishedName() As List(Of String)

			Return SplitDistinguishedName(Me)
		End Function

    ''' <summary>
    ''' Maskiert Sonderzeichen mit einem voranstehendem "\"
    ''' Bsp: Dampf, Hans => Dampf\, Hans
    ''' </summary>
    Public Function EscapeValue() As String

			Return DbResultAD.EscapeValue(Me.Value)
		End Function

    ''' <summary>
    ''' Demaskiert Sonderzeichen im einem voranstehendem "\"
    ''' Bsp: Dampf\, Hans => Dampf, Hans
    ''' </summary>
    Public Function DeEscapeValue() As String

			Return DbResultAD.DeEscapeValue(Me.Value)
		End Function

    ''' <summary>
    ''' Liefert ein DirectoryEntry-Objekt auf Grundlage des DistinguishedNames.
    ''' Das Objekt wird im Berechtigungskontext des ManipulationUsers erzeugt,
    ''' wenn useManipulationUser den Wert true besitzt.
    ''' </summary>
    Public Function ToDirectoryEntry(ByVal useManipulationUser As Boolean) As DirectoryEntry

			If useManipulationUser Then
				Return New DirectoryEntry(Me.ToUrl, Settings.Instance.ManipulationUserName, Settings.Instance.ManipulationUserPassword)
			Else
				Return New DirectoryEntry(Me.ToUrl)
			End If
		End Function

    ''' <summary>
    ''' Liefert den DistinguishedName eines Unterelementes.
    ''' Bsp: GetDistinguishedNameFromChild("ou=test2, ou=test") -> ou=test2, ou=test, dc=domain, dc=net
    ''' </summary>
    Public Function GetDistinguishedNameFromChild(ByVal relativeDistingusehdName As String) As DistinguishedName

			Dim baseProperties = New DistinguishedNameBaseProperties(relativeDistingusehdName, "", "", New Guid, "")
			Dim list = SplitDistinguishedName(New DistinguishedName(baseProperties))

			Dim parts = list.Select(Function(s) If((s.Trim.ToLower.StartsWith("cn=")) _
			OrElse (s.Trim.ToLower.StartsWith("ou=")), s, "ou=" & s).ToString).ToArray

			Dim dn = String.Join(",", parts) & "," & Me.Value

			Return DistinguishedName.GetByDistinguishedName(dn)
		End Function

    ''' <summary>
    ''' Prüft rekursiv, ob das Objekt des DistinguishedNames Mitglied der angegebenen Gruppe ist.
    ''' Dauert unter Umständen relativ lange.
    ''' Es sollte IsMemberOf der GrantTree-Klasse verwendet werden.
    ''' </summary>
    Public Function IsMemberOf(ByVal groupDistinguishedName As DistinguishedName) As Boolean

			Dim distinguishedNames = Me.ToDirectoryEntry(False).Properties.Item _
			(AdProperties.memberOf.ToString).Cast(Of String)

			If Not distinguishedNames.Any Then Return False

			Dim memberOf = DistinguishedNameRepository.Instance.GetByDistinguishedNames(distinguishedNames)
			If memberOf.Where(Function(dn) dn.IsEqualTo(groupDistinguishedName)).Any Then
				Return True
			Else
				Return memberOf.Any(Function(dn) dn.IsMemberOf(groupDistinguishedName))
			End If
		End Function

    ''' <summary>
    ''' Ruft den Wert der Eigenschaft propertyName aus dem referenzierten DirectoryEntry ab.
    ''' </summary>
    Public Function GetProperty(ByVal propertyName As String) As Object
			Return Me.ToDirectoryEntry(False).Properties.Item(propertyName).Value
		End Function

    ''' <summary>
    ''' Ruft den Wert der Eigenschaft propertyName aus dem referenzierten DirectoryEntry ab.
    ''' </summary>
    Public Function GetProperty(ByVal propertyName As AdProperties) As Object

			Dim sb = AdRepositoryHelper.Instance.CreateDefaultSelectBuilder
			sb.Select.Add(propertyName.ToString)
			sb.Where.Add(String.Format("{0}='{1}'", AdProperties.distinguishedName.ToString, Me.Value))
			Dim value As Object = sb.ExecuteScalar

			Return sb.ExecuteScalar
		End Function

    ''' <summary>
    ''' Ruft den Wert der Eigenschaft propertyName aus dem referenzierten DirectoryEntry ab.
    ''' Bei DBNull-Werten wird defaultValueForDBNull als Rückgabe geliefert.
    ''' </summary>
    Public Function GetProperty(ByVal propertyName As AdProperties, ByVal defaultValueForDBNull As Object) As Object

			Dim value As Object = GetProperty(propertyName)
			Return If(Convert.IsDBNull(value), defaultValueForDBNull, value)
		End Function

    ''' <summary>
    ''' Prüft, ob das übergebene dn-Objekt, die angegebene objectClass beinhaltet.
    ''' </summary>
    Public Function IsObjectClass(ByVal objectClass As ObjectClasses) As Boolean
			Return AdInformation.IsObjectClass(Me, objectClass)
		End Function

    ''' <summary>
    ''' Prüft, ob das übergebene dn-Objekt, die objectClass Group beinhaltet.
    ''' Jetzt schneller Manko, alles was kein User ist liefert true ... naja!
    ''' </summary>
    Public Function IsGroup() As Boolean
			Return _baseProperties.PersonId = 0
		End Function

    ''' <summary>
    ''' Prüft, ob das übergebene dn-Objekt, die objectClass User beinhaltet.
    ''' </summary>
    Public Function IsUser() As Boolean
			Return _baseProperties.PersonId > 0
		End Function

    ''' <summary>
    ''' Prüft, ob das übergebene dn-Objekt, die objectClass OrganizationalUnit beinhaltet.
    ''' </summary>
    Public Function IsOrganizationalUnit() As Boolean
			Return AdInformation.IsOrganizationalUnit(Me)
		End Function

    ''' <summary>
    ''' Prüft, ob die aktuelle Instanz auf das Ad-Objekt von dn verweist.
    ''' </summary>
    Public Function IsEqualTo(ByVal dn As DistinguishedName) As Boolean

			Return IsEqualTo(Me, dn)
		End Function

    ''' <summary>
    ''' Prüft, ob zwei DistinguishedNames auf ein Ad-Objekt verweisen.
    ''' </summary>
    Public Shared Function IsEqualTo _
		(ByVal dn As DistinguishedName, ByVal dn2 As DistinguishedName) As Boolean

			Return dn.ToDeEscapedRelativeDistinguishedName.ToLower.Replace(" ", "") _
			= dn2.ToDeEscapedRelativeDistinguishedName.ToLower.Replace(" ", "")
		End Function

    ''' <summary>
    ''' Prüft, ob die aktuelle Instanz auf das selbe Objekt im AD verweist, wie das DirectoryEntry de.
    ''' </summary>
    Public Function IsDirectoryEntryEqualTo(ByVal de As DirectoryEntry) As Boolean

			Return IsDirectoryEntryEqualTo(Me.ToDirectoryEntry(False), de)
		End Function

    ''' <summary>
    ''' Prüft, ob zwei DestinguishedName-Objekte auf das selbe Objekt im AD verweisen.
    ''' </summary>
    Public Shared Function IsDirectoryEntryEqualTo _
		(ByVal de As DirectoryEntry, ByVal de2 As DirectoryEntry) As Boolean

			Return de.Guid.ToString = de2.Guid.ToString
		End Function

    ''' <summary>
    ''' Liefert alle Mitglieder des DinstinguishedName-Objektes.
    ''' </summary>
    Public Function GetMembers() As DistinguishedName()

			Dim sb = AdRepositoryHelper.Instance.CreateDefaultSelectBuilder
			sb.Select.Add(AdProperties.member.ToString)
			sb.Where.Add(String.Format("{0}='{1}'", AdProperties.distinguishedName.ToString, Me.Value))

			Dim temp = New List(Of String)

			Using dr = sb.ExecuteReader
				While dr.Read
					If Not IsDBNull(dr.Item(AdProperties.member.ToString)) Then
						temp.AddRange(CType(dr.Item(AdProperties.member.ToString), Object()).Cast(Of String))
					End If
				End While
			End Using

			Return DistinguishedNameRepository.Instance.GetByDistinguishedNames(temp)
		End Function

    ''' <summary>
    ''' Liefert alle Mitglieder des DinstinguishedName-Objektes rekursiv aus allen Unterobjekten.
    ''' </summary>
    Public Function GetMembersRecursive() As List(Of DistinguishedName)
			Return GetMembersRecursive(GetMembersRecursiveTypes.AllMembers)
		End Function

    ''' <summary>
    ''' Liefert alle Mitglieder des DinstinguishedName-Objektes rekursiv aus allen Unterobjekten.
    ''' </summary>
    Public Function GetMembersRecursive(ByVal memberType As GetMembersRecursiveTypes) As List(Of DistinguishedName)

			Dim results = New List(Of DistinguishedName)
			Dim members = Me.GetMembers.ToList
			Dim groups = members.Where(Function(dn) dn.IsGroup).ToList

			If (memberType = GetMembersRecursiveTypes.AllMembers) _
			OrElse (memberType = GetMembersRecursiveTypes.GroupsOnly) Then
				results.AddRange(groups)
			End If

			groups.ForEach(Sub(dn) results.AddRange(dn.GetMembersRecursive(memberType)))

			If (memberType = GetMembersRecursiveTypes.AllMembers) _
			OrElse (memberType = GetMembersRecursiveTypes.UsersOnly) Then
				results.AddRange(members.Where(Function(dn) dn.IsUser).ToList())
			End If

			Return results.GroupBy(Function(dn) dn.Value).Select(Function(dnGroup) dnGroup.First).ToList
		End Function

    ''' <summary>
    ''' Liefert alle Mitgliedschaften des DinstinguishedName-Objektes.
    ''' </summary>
    Public Function GetMemberOf() As DistinguishedName()

			Dim sb = AdRepositoryHelper.Instance.CreateDefaultSelectBuilder
			sb.Select.Add(AdProperties.memberOf.ToString)
			sb.Where.Add(String.Format("{0}='{1}'", AdProperties.distinguishedName.ToString, Me.Value))

			Dim temp = New List(Of String)

			Using dr = sb.ExecuteReader
				While dr.Read
					If Not IsDBNull(dr.Item(AdProperties.memberOf.ToString)) Then
						temp.AddRange(CType(dr.Item(AdProperties.memberOf.ToString), Object()).Cast(Of String))
					End If
				End While
			End Using

			Return DistinguishedNameRepository.Instance.GetByDistinguishedNames(temp)
		End Function

		''' <summary>
		''' Prüft, ob der aktuelle DistinguishedName den übergebenen DistinguishedName
		''' als Parent-Path beinhaltet.
		''' </summary>
		Public Function ContainsDn(ByVal dn As DistinguishedName) As Boolean
			Try
				Return ContainsDn(dn.Value)
			Catch ex As Exception
				Return False
			End Try
		End Function

		''' <summary>
		''' Prüft, ob der aktuelle DistinguishedName den übergebenen DistinguishedName
		''' als Parent-Path beinhaltet.
		''' </summary>
		Public Function ContainsDn(ByVal dn As String) As Boolean
			Try
				Return Me.Value.ToLower.Contains(dn.ToLower)
			Catch ex As Exception
				Return False
			End Try
		End Function

#End Region '{Öffentliche Methoden der Klasse}

	End Class

End Namespace
