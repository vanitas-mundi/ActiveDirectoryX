Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports/ usings "
Imports SSP.ActiveDirectoryX.Core
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.Data.StatementBuildersAD.Core
#End Region

Namespace Data.Repositories

	Public Class DistinguishedNameRepository

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private Shared _instance As DistinguishedNameRepository
#End Region '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Private Sub New()
		End Sub

		Shared Sub New()
			_instance = New DistinguishedNameRepository
		End Sub
#End Region '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public Shared ReadOnly Property Instance As DistinguishedNameRepository
			Get
				Return _instance
			End Get
		End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
		Private Function GetDistinguishedNameBase(ByVal whereCondition As String) As DistinguishedName()
			Return GetDistinguishedNameBase(New String() {whereCondition})
		End Function

		Private Function GetDistinguishedNameBase(ByVal whereConditions As String()) As DistinguishedName()
			Return GetDistinguishedNameBase(whereConditions, AdProperties.sAMAccountName)
		End Function

		Private Function GetDistinguishedNameBase _
		(ByVal whereConditions As String() _
		, ByVal nameProperty As AdProperties) As DistinguishedName()

			Dim sb = AdRepositoryHelper.Instance.CreateDefaultSelectBuilder(SpecialDistinguishedNames.DomainDn)
			AdRepositoryHelper.Instance.SelectAddAdProperty(sb, AdProperties.distinguishedName)
			AdRepositoryHelper.Instance.SelectAddAdProperty(sb, nameProperty)
			AdRepositoryHelper.Instance.SelectAddAdProperty(sb, AdProperties.description)
			AdRepositoryHelper.Instance.SelectAddAdProperty(sb, AdProperties.objectGUID)
			AdRepositoryHelper.Instance.SelectAddAdProperty(sb, AdProperties.employeeID)

			sb.Where.AddRange(whereConditions)

			Dim distinguishedNames = New List(Of DistinguishedName)

			Using dr = sb.ExecuteReader
				While dr.Read
					Dim values(dr.FieldCount - 1) As Object
					dr.GetValues(values)
					Dim properties = New DistinguishedNameBaseProperties(values.ToArray.Reverse.ToArray)
					distinguishedNames.Add(New DistinguishedName(properties))
				End While
			End Using

			Return distinguishedNames.ToArray
		End Function

		Private Function BuildSingleWhereCondition(ByVal adProperty As AdProperties, ByVal value As String) As String()
			Return New String() {String.Format("{0}='{1}'", adProperty.ToString, value)}
		End Function

    '''<summary>Liefert ein DistinguishedName-Objekte anhan von DistinguishedNames.</summary>
    Private Function GetByStatementConditionsMaximum(ByVal values As IEnumerable(Of String) _
		, ByVal propertyName As AdProperties) As DistinguishedName()

			Dim results = New List(Of DistinguishedName)
			Dim steps = values.Count \ My.Settings.StatementConditionsMaximum

			For stepCounter = 0 To steps

				Dim ret = values.Skip _
				(stepCounter * My.Settings.StatementConditionsMaximum).Take _
				(My.Settings.StatementConditionsMaximum).ToList

				If ret.Any Then
					Dim delimiter = String.Format("' OR {0} = '", propertyName.ToString)
					Dim where = New String() {propertyName.ToString & " = '" & String.Join(delimiter, ret) & "'"}
					results.AddRange(GetDistinguishedNameBase(where))
				End If
			Next stepCounter

			Return results.ToArray
		End Function
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		''' <summary>
		''' Liefert ein DistinguishedName-Objekt anhand einer PersonenID.
		''' </summary>
		Public Function GetByPersonId(ByVal personId As Int64) As DistinguishedName
			Return GetDistinguishedNameBase(BuildSingleWhereCondition _
			(AdProperties.employeeID, personId.ToString)).FirstOrDefault
		End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand einer Guid.
    ''' </summary>
    Public Function GetByGuid(ByVal guid As Guid) As DistinguishedName
			Return GetByGuid(guid.ToString)
		End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand eines Guid-Strings.
    ''' </summary>
    Public Function GetByGuid(ByVal guid As String) As DistinguishedName
			Return GetDistinguishedNameBase(BuildSingleWhereCondition _
			(AdProperties.objectGUID, AdTypeConverter.GuidToEscapedNativeGuid(guid))).FirstOrDefault
		End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand eines samAccountNames.
    ''' </summary>
    Public Function GetByUserName(ByVal userName As String) As DistinguishedName
			Return GetDistinguishedNameBase(BuildSingleWhereCondition(AdProperties.sAMAccountName, userName)).FirstOrDefault
		End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand eines common names (cn).
    ''' </summary>
    Public Function GetByCn(ByVal cn As String) As DistinguishedName
			Return GetDistinguishedNameBase(BuildSingleWhereCondition(AdProperties.cn, cn)).FirstOrDefault
		End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand eines common names (cn).
    ''' </summary>
    Public Function GetByOu(ByVal ou As String) As DistinguishedName
			Return GetDistinguishedNameBase(BuildSingleWhereCondition(AdProperties.ou, ou), AdProperties.name).FirstOrDefault
		End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand eines DisplayNames.
    ''' </summary>
    Public Function GetByDisplayName(ByVal displayName As String) As DistinguishedName
			Return GetDistinguishedNameBase(BuildSingleWhereCondition(AdProperties.displayName, displayName)).FirstOrDefault
		End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand eines samAccountNames.
    ''' </summary>
    Public Function GetByGroupName(ByVal groupName As String) As DistinguishedName

			Dim temp = New String() _
			{
				BuildSingleWhereCondition(AdProperties.sAMAccountName, groupName).First,
				"AND " & BuildSingleWhereCondition(AdProperties.objectClass, ObjectClasses.group.ToString).First
			}

			Return GetDistinguishedNameBase(temp).FirstOrDefault
      'oder lieber doch nach 'Name' suchen, ich weiß es nicht ...
      'Return GetDistinguishedNameBase(String.Format("{0}='{1}' AND {2}='{3}'" _
      ', AdProperties.Name.ToString, groupName, AdProperties.objectClass.ToString, ObjectClasses.group.ToString))
    End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand einer Berechtigung.
    ''' </summary>
    Public Function GetByGrant(ByVal appName As String, ByVal grantName As String) As DistinguishedName
			Dim groupName = String.Concat(appName, ".", grantName)
			Return GetByGroupName(groupName)
		End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand eines DistinguishedNames.
    ''' </summary>
    Public Function GetByDistinguishedName(ByVal distinguishedName As String) As DistinguishedName
			Return GetDistinguishedNameBase(BuildSingleWhereCondition(AdProperties.distinguishedName, distinguishedName)).FirstOrDefault
		End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand eines AdSelectBuilder, welcher als ersten Wert einen DistinguishedName liefert.
    ''' </summary>
    Public Function GetByDistinguishedName(ByVal sb As SelectBuilderAD) As DistinguishedName

			Return GetByDistinguishedName(sb.ExecuteStringScalar)
		End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekte anhand von DistinguishedNames.
    ''' </summary>
    Public Function GetByDistinguishedNames(ByVal distinguishedNames As IEnumerable(Of String)) As DistinguishedName()

			Return GetByStatementConditionsMaximum(distinguishedNames, AdProperties.distinguishedName)
		End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekte anhand eines AdSelectBuilder, welcher als erstes Feld einen DistinguishedName beinhaltet.
    ''' </summary>
    Public Function GetByDistinguishedNames(ByVal sb As SelectBuilderAD) As DistinguishedName()
			Dim temp = New List(Of String)

			Using dr = sb.ExecuteReader
				While dr.Read
					temp.Add(dr.GetString(0))
				End While
			End Using

			Return GetByDistinguishedNames(temp)
		End Function
#End Region '{Öffentliche Methoden der Klasse}

	End Class

End Namespace
