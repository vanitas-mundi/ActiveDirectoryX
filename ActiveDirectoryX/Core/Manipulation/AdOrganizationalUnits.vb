Option Explicit On
Option Strict On
Option Infer On

#Region " --------------->> Imports/ usings "

Imports SSP.ActiveDirectoryX.Core.Enums

#End Region

Namespace Core.Manipulation

	Public Class AdOrganizationalUnits

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region	'{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		''' <summary>
		''' Legt eine neue Organisationseinheit im Active-Directory an.
		''' </summary>
		Public Shared Sub CreateOrganizationalUnit(ByVal parentOuName As String _
		, ByVal newOuName As String, ByVal useManipulationUser As Boolean)

			CreateOrganizationalUnit(DistinguishedName.GetByOu _
			(parentOuName), newOuName, "", useManipulationUser)
		End Sub

		''' <summary>
		''' Legt eine neue Organisationseinheit im Active-Directory an.
		''' </summary>
		Public Shared Sub CreateOrganizationalUnit(ByVal parentOuName As String, ByVal newOuName As String _
		, ByVal newOuDescription As String, ByVal useManipulationUser As Boolean)

			CreateOrganizationalUnit(DistinguishedName.GetByOu(parentOuName) _
			, newOuName, newOuDescription, useManipulationUser)
		End Sub

		''' <summary>
		''' Legt eine neue Organisationseinheit im Active-Directory an.
		''' </summary>
		Public Shared Sub CreateOrganizationalUnit _
		(ByVal parentOuDn As DistinguishedName, ByVal newOuName As String, ByVal useManipulationUser As Boolean)

			CreateOrganizationalUnit(parentOuDn, newOuName, "", useManipulationUser)
		End Sub

		''' <summary>
		''' Legt eine neue Organisationseinheit im Active-Directory an.
		''' </summary>
		Public Shared Sub CreateOrganizationalUnit _
		(ByVal parentOuDn As DistinguishedName, ByVal newOuName As String _
		, ByVal newOuDescription As String, ByVal useManipulationUser As Boolean)

			Dim newOuEntry = parentOuDn.ToDirectoryEntry(useManipulationUser).Children.Add _
			(String.Concat("OU=", newOuName), ObjectClasses.organizationalUnit.ToString)

			If newOuDescription.Trim.Length > 0 Then
				newOuEntry.Properties.Item(AdProperties.description.ToString).Value = newOuDescription.Trim
			End If

			newOuEntry.CommitChanges()
		End Sub

		''' <summary>
		''' Löscht die angegebene Organisationseinheit im Active-Directory.
		''' </summary>
		Public Shared Sub DeleteOrganizationalUnit(ByVal ouName As String, ByVal useManipulationUser As Boolean)

			DeleteOrganizationalUnit(DistinguishedName.GetByOu(ouName), useManipulationUser)
		End Sub

		''' <summary>
		''' Löscht die angegebene Organisationseinheit im Active-Directory.
		''' </summary>
		Public Shared Sub DeleteOrganizationalUnit _
		(ByVal ouDn As DistinguishedName, ByVal useManipulationUser As Boolean)

			Dim ouEntry = ouDn.ToDirectoryEntry(useManipulationUser)
			Dim parent = ouDn.Parent.ToDirectoryEntry(useManipulationUser)
			parent.Children.Remove(ouEntry)
			ouEntry.CommitChanges()
			parent.CommitChanges()
		End Sub
#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace
