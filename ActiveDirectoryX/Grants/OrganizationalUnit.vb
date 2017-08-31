Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "

Imports SSP.ActiveDirectoryX.Grants.Administration
Imports SSP.ActiveDirectoryX.Core
Imports System.DirectoryServices
Imports SSP.ActiveDirectoryX.Grants.Exceptions
Imports System.Text
Imports SSP.ActiveDirectoryX.Grants.Enums
Imports SSP.ActiveDirectoryX.Data.Repositories
#End Region

Namespace Grants

	Public Class OrganizationalUnit

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private _organizationalUnitDn As DistinguishedName
		Private _administration As OrganizationalUnitAdministration
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Public Sub New(ByVal organizationalUnitDn As DistinguishedName)
			If Not AdministrationTypeResolver.Instance.IsOrganizationalUnit(organizationalUnitDn) Then
				Throw New WrongAdministrationTypeException
			End If

			_administration = New OrganizationalUnitAdministration(Me)
			_organizationalUnitDn = organizationalUnitDn
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		''' <summary>
		''' Stellt Routinen zur Administration von Administrations-Organisationseinheiten zur verfügung.
		''' </summary>
		Public ReadOnly Property Administration As OrganizationalUnitAdministration
		Get
			Return _administration
		End Get
		End Property

    ''' <summary>
    ''' Liefert ein Array mit den untergeordneten Organisationseinheiten der Organisationseinheit.
    ''' </summary>
    Public ReadOnly Property ChildrenOrganizationalUnits As OrganizationalUnit()
      Get
        Dim distinguishedNames = Me.OrganizationalUnitDn.ToDirectoryEntry(False).Children.OfType _
        (Of DirectoryEntry).Where(Function(de) de.Path.StartsWith _
        ("ldap://ou=", StringComparison.CurrentCultureIgnoreCase)).Select(Function(de) de.Path)

        Dim ous = DistinguishedNameRepository.Instance.GetByDistinguishedNames _
        (distinguishedNames).Select(Function(dn) New OrganizationalUnit(dn))

        Return ous.ToArray
      End Get
    End Property

    ''' <summary>
    ''' Liefert die übergeordnete Organisationseinheit der Organisationseinheit.
    ''' Liefert NULL für Elemente oberhalb der Administrationsstruktur.
    ''' </summary>
    Public ReadOnly Property Parent As OrganizationalUnit
		Get
			Dim administrationDn = SpecialDistinguishedNames.Item(Core.Enums.SpecialDistinguishedNameKeys.Administration)
			Dim parentDn = Me.OrganizationalUnitDn.Parent
			If parentDn.ContainsDn(administrationDn) Then
				Return New OrganizationalUnit(parentDn)
			Else
				Return Nothing
			End If
		End Get
		End Property

		''' <summary>
		''' Liefert den Namen der Organisationseinheit.
		''' </summary>
		Public ReadOnly Property Name As String
		Get
			Return _organizationalUnitDn.ToRelativeName
		End Get
		End Property

		''' <summary>
		''' Liefert den Namen der Organisationseinheit ohne führendem Organisationseinheittypnamen.
		''' </summary>
		Public ReadOnly Property NameSimple As String
		Get
			Return Me.Name.Split("."c).LastOrDefault
		End Get
		End Property

		''' <summary>
		''' Liefert den DistinguishedName der Organisationseinheit.
		''' </summary>
		Public ReadOnly Property OrganizationalUnitDn As DistinguishedName
		Get
			Return _organizationalUnitDn
		End Get
		End Property
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region	'{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		''' <summary>
		''' Liefert einen StringBuilder mit allen untergeordneten OU-Namen, rekursiv bis in die letzte Hierachie.
		''' </summary>
		Public Function GetTreeStringBuilder _
		(ByVal parentLine As String, ByVal regularElement As String _
		, ByVal endElement As String, ByVal whiteSpace As String) As StringBuilder

			Dim builder = New TreeBuilder(TreeBuilderRootLineTypes.Custom)
			builder.Rootlines.EndElement = endElement
			builder.Rootlines.RegularElement = regularElement
			builder.Rootlines.ParentLine = parentLine
			builder.Rootlines.WhiteSpace = whiteSpace
			Return builder.GetTreeStringBuilder(Me)
		End Function

		''' <summary>
		''' Liefert einen StringBuilder mit allen untergeordneten OU-Namen, rekursiv bis in die letzte Hierachie.
		''' </summary>
		Public Function GetTreeStringBuilder() As StringBuilder
			Return GetTreeStringBuilder(TreeBuilderRootLineTypes.Regular)
		End Function

		''' <summary>
		''' Liefert einen StringBuilder mit allen untergeordneten OU-Namen, rekursiv bis in die letzte Hierachie.
		''' </summary>
		Public Function GetTreeStringBuilder(ByVal rootLines As TreeBuilderRootLineTypes) As StringBuilder
			Dim builder = New TreeBuilder(rootLines)

			Return builder.GetTreeStringBuilder(Me)
		End Function

		''' <summary>
		''' Liefert einen String mit allen untergeordneten OU-Namen, rekursiv bis in die letzte Hierachie.
		''' </summary>
		Public Function GetTreeString _
		(ByVal parentLine As String, ByVal regularElement As String _
		, ByVal endElement As String, ByVal whiteSpace As String) As String
			Return GetTreeStringBuilder(parentLine, regularElement, endElement, whiteSpace).ToString
		End Function

		''' <summary>
		''' Liefert einen String mit allen untergeordneten OU-Namen, rekursiv bis in die letzte Hierachie.
		''' </summary>
		Public Function GetTreeString() As String
			Return GetTreeStringBuilder(TreeBuilderRootLineTypes.Regular).ToString
		End Function

		''' <summary>
		''' Liefert einen String mit allen untergeordneten OU-Namen, rekursiv bis in die letzte Hierachie.
		''' </summary>
		Public Function GetTreeString(ByVal rootLines As TreeBuilderRootLineTypes) As String
			Return GetTreeStringBuilder(rootLines).ToString
		End Function

#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace

