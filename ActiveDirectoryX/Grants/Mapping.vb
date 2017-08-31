Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports/ usings "
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core
Imports SSP.ActiveDirectoryX.Grants.Exceptions
#End Region

Namespace Grants

	Public Class Mapping

		Inherits AdministrationGroup

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private Const Delimiter As Char = "|"c
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Public Sub New(ByVal distinguishedName As DistinguishedName)
			MyBase.New(distinguishedName)

			If Not AdministrationTypeResolver.Instance.IsMapping(distinguishedName) Then
				Throw New WrongAdministrationTypeException
			End If
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public Property Unc As String
		Get
			Return NullToEmptyString(Me.Description.Split(Delimiter).FirstOrDefault)
		End Get
		Set(value As String)
			SetMappingInfo(value, Me.DriveLetter)
		End Set
		End Property

		Public Property DriveLetter As String
		Get
			Return NullToEmptyString(Me.Description.Split(Delimiter).LastOrDefault)
		End Get
		Set(value As String)
			SetMappingInfo(Me.Unc, value)
		End Set
		End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
		Private Sub MapNetworkDriveBase(ByVal mapping As Mapping)

			Dim info = New ProcessStartInfo("net", "use " & mapping.DriveLetter.Replace("\", "") & " " & mapping.Unc)
			info.WindowStyle = ProcessWindowStyle.Hidden
			Process.Start(info)
		End Sub

		Private Function NullToEmptyString(ByVal value As String) As String
			Return If(value Is Nothing, "", value.Trim)
		End Function

		Private Sub SetMappingInfo(ByVal unc As String, ByVal driveLetter As String)

			Dim value = String.Concat(NullToEmptyString(unc), " ", Delimiter, " ", NullToEmptyString(driveLetter))
			Dim entry = Me.GroupDistinguishedName.ToDirectoryEntry(True)
			entry.Properties.Item(AdProperties.description.ToString).Value = value
			entry.CommitChanges()
		End Sub
#End Region  '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		Public Overrides Function ToString() As String
			Return Me.NameSimple
		End Function

		''' <summary>Bindet das Mapping ins System des aufrufenden Benutzers ein.</summary>
		Public Function MapNetworkDrive() As AdManipulationResults

			Select Case True
				Case Not Me.GroupDistinguishedName.IsGroup
					Return AdManipulationResults.IsNotGroup
				Case Not Me.GroupDistinguishedName.ContainsDn _
				(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Mappings))
					Return AdManipulationResults.GroupIsNotMapping
				Case Else
					Try
						MapNetworkDriveBase(New Mapping(Me.GroupDistinguishedName))
					Catch ex As System.Exception
						Return AdManipulationResults.MappingError
				End Try
				Return AdManipulationResults.Successful
			End Select
		End Function
#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace

