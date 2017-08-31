Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Data
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core
Imports SSP.Data.StatementBuildersAD.Core
#End Region

Namespace Grants

	Public Class GroupTree

		Inherits AdTree

#Region " --------------->> Enumerationen der Klasse "
#End Region  '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private Const MaxQueryResults As Int32 = 40
#End Region  '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Public Sub New(ByVal groupName As String)
			Me.New(DistinguishedName.GetByGroupName(groupName))
		End Sub

		Public Sub New(ByVal distinguishedName As DistinguishedName)
			MyBase.New(distinguishedName)
		End Sub
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		''' <summary>
		''' Liefert rekursiv alle Personen-Ids aller User-Mitglieder einer Gruppe.
		''' </summary>
		Public ReadOnly Property AllUserMembersPersonIds As Int64()
			Get
				Return Me.GetMembersPersonIds(AdProperties.user)
			End Get
		End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
		'Private Function GetMembersStatement _
		'(ByVal distinguishedNames As IEnumerable(Of String), ByVal objectClass As AdProperties) As String

		'	If distinguishedNames.Count = 0 Then Return ""

		'	Dim whereParts = distinguishedNames.Select(Function(s) _
		'	String.Format("{0}='{1}'", AdProperties.distinguishedName.ToString, s)).ToArray
		'    Dim temp = "AND (" & String.Join(String.Format("{0}OR ", vbCrLf), whereParts) & ")"

		'    Dim sb = AdSelectBuilder.CreateDefaultSelectBuilder
		'    sb.Select.Add(AdProperties.distinguishedName.ToString)

		'    If objectClass = AdProperties.user Then
		'      sb.Select.Add(AdProperties.employeeID.ToString)
		'    End If

		'    sb.Where.Add("objectClass='{0}'", objectClass.ToString)
		'    sb.Where.Add(temp)
		'    Return sb.ToString
		'End Function

		Private Function GetMemberDistinguishedNames(ByVal statement As String) As List(Of DistinguishedName)

			Dim list = New List(Of String)

			Using dr = DbResultAD.Instance.ExecuteReader(statement)
				While dr.Read
					list.Add(dr.Item(AdProperties.distinguishedName.ToString).ToString)
				End While
			End Using

			Return Repositories.DistinguishedNameRepository.Instance.GetByDistinguishedNames(list).ToList
		End Function

		Private Function GetMemberPersonIds(ByVal statement As String) As List(Of Int64)

			Dim list = New List(Of Int64)

			Using dr = DbResultAD.Instance.ExecuteReader(statement)
				While dr.Read
					If Not Convert.IsDBNull(dr.Item(AdProperties.employeeID.ToString)) Then
						list.Add(Convert.ToInt64(dr.Item(AdProperties.employeeID.ToString)))
					End If
				End While
			End Using

			Return list
		End Function

		Private Function GetMembersPersonIds(ByVal objectClass As AdProperties) As Int64()

			Return Me.Items.Select(Function(x) x.BaseProperties.PersonId).ToArray
		End Function
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		''' <summary>
		''' Wandelt den GroupTree in ein AdministrationGroup-Objekt um und gibt es zurück.
		''' </summary>
		Public Function ToAdministrationGroup() As AdministrationGroup
			Return New AdministrationGroup(Me.DistinguishedNameContext)
		End Function

		''' <summary>
		''' Erzeugt den GrantTree und lädt die rekursiven Members
		''' </summary>
		Public Overloads Sub Generate()

			MyBase.GenerateTree(AdProperties.member)
		End Sub
#End Region  '{Öffentliche Methoden der Klasse}

	End Class

End Namespace
