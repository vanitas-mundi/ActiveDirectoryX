Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
#End Region

Namespace Core

	<DataContract()>
	Public Class GrantInfo

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private _appName As String
		Private _personId As Int32
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		<DataMember()>
		Public Property AppName() As String
		Get
			Return _appName
		End Get
		Set(value As String)
			_appName = value
		End Set
		End Property

		<DataMember()>
		Public Property PersonId() As Int32
		Get
			Return _personId
		End Get
		Set(value As Int32)
			_personId = value
		End Set
		End Property
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region	'{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace
