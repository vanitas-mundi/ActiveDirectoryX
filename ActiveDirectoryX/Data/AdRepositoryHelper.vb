Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports/ usings "
Imports SSP.ActiveDirectoryX.Core
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.Data.StatementBuildersAD.Core
#End Region

Namespace Data

	Public Class AdRepositoryHelper

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private Shared _instance As AdRepositoryHelper
#End Region '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Shared Sub New()
			_instance = New AdRepositoryHelper
		End Sub

		Private Sub New()
		End Sub
#End Region '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public Shared ReadOnly Property Instance As AdRepositoryHelper
			Get
				Return _instance
			End Get
		End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		Public Function CreateDefaultSelectBuilder(ByVal fromDn As SpecialDistinguishedNameKeys) As SelectBuilderAD
			Dim sb = New SelectBuilderAD
			sb.From.Add("'" & SpecialDistinguishedNames.Item(fromDn).ToUrl & "'")
			Return sb
		End Function

		Public Function CreateDefaultSelectBuilder(ByVal fromDn As DistinguishedName) As SelectBuilderAD
			Dim sb = New SelectBuilderAD
			sb.From.Add("'" & fromDn.ToUrl & "'")
			Return sb
		End Function

		Public Function CreateDefaultSelectBuilder(ByVal fromDn As String) As SelectBuilderAD
			Dim sb = New SelectBuilderAD
			sb.From.Add("'" & String.Concat(My.Settings.LdapProtocolName, fromDn) & "'")
			Return sb
		End Function

		Public Function CreateDefaultSelectBuilder() As SelectBuilderAD
			Return CreateDefaultSelectBuilder(SpecialDistinguishedNameKeys.Domain)
		End Function

		'''<summary>Prüft ob der übergebene distinguishedName existiert.</summary>
		Public Function ExistPath _
		(ByVal connectionString As String, ByVal distinguishedName As String) As Boolean

			Dim sb = AdRepositoryHelper.Instance.CreateDefaultSelectBuilder
			sb.Select.Add(AdProperties.distinguishedName.ToString)
			sb.Select.Add(AdProperties.memberOf.ToString)
			sb.Where.Add(String.Format("{0}='{1}'", AdProperties.distinguishedName.ToString, distinguishedName))

			Return DbResultAD.Instance.ExecuteReader(connectionString, sb.ToString).HasRows
		End Function

		'''<summary>Prüft ob der übergebene distinguishedName existiert.</summary>
		Public Function ExistPath(ByVal distinguishedName As String) As Boolean
			Return ExistPath(DbResultAD.DefaultConnectionString, distinguishedName)
		End Function

		'''<summary>Prüft ob der übergebene distinguishedName existiert.</summary>
		Public Function ExistPath _
		(ByVal connectionString As String, ByVal distinguishedName As DistinguishedName) As Boolean
			Return ExistPath(connectionString, distinguishedName.Value)
		End Function

		'''<summary>Prüft ob der übergebene distinguishedName existiert.</summary>
		Public Function ExistPath(ByVal distinguishedName As DistinguishedName) As Boolean
			Return ExistPath(DbResultAD.DefaultConnectionString, distinguishedName.Value)
		End Function

		Public Sub SelectAddAdProperty(ByVal sb As SelectBuilderAD, ByVal adProperty As AdProperties)

			sb.Select.Add(adProperty.ToString)
		End Sub

		Public Sub SetSingleWhereAdPropertyCondition _
		(ByVal sb As SelectBuilderAD, ByVal adProperty As AdProperties, ByVal value As String)

			sb.Where.Clear()
			sb.Where.Add("{0}='{1}'", adProperty.ToString, value)
		End Sub

		Public Sub SetFromDn(ByVal sb As SelectBuilderAD, ByVal fromDn As SpecialDistinguishedNameKeys)
			sb.From.Clear()
			sb.From.Add("'" & SpecialDistinguishedNames.Item(fromDn).ToUrl & "'")
		End Sub
#End Region '{Öffentliche Methoden der Klasse}

	End Class

End Namespace

