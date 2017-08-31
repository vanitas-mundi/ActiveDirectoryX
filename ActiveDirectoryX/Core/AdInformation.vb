Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "

Imports System.DirectoryServices
Imports System.Security
Imports SSP.ActiveDirectoryX.Core.Enums
Imports System.Runtime.InteropServices
Imports SSP.Data.StatementBuildersAD.Core
Imports SSP.ActiveDirectoryX.Data

#End Region

Namespace Core

	Public Class AdInformation

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
#End Region '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
#End Region '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "

		'''<summary>Prüft, ob ein User-Objekt mit der angegebenen EmployeeId in AD existiert.</summary>
		Public Shared Function ExistsEmployeeId(ByVal id As String) As Boolean
			Dim sb = AdRepositoryHelper.Instance.CreateDefaultSelectBuilder
			sb.Select.Add(AdProperties.distinguishedName.ToString)
			sb.Where.Add(String.Format("{0}='{1}'", AdProperties.employeeID, id))
			Return sb.ExecuteScalar IsNot Nothing
		End Function

		'''<summary>Prüft, ob ein User-Objekt mit der angegebenen PersonenId in AD existiert.</summary>
		Public Shared Function ExistsPersonId(ByVal id As String) As Boolean
			Return ExistsEmployeeId(id)
		End Function

		'''<summary>Prüft, ob ein User-Objekt mit der angegebenen PersonenId in AD existiert.</summary>
		Public Shared Function ExistsPersonId(ByVal id As Int64) As Boolean
			Return ExistsEmployeeId(id.ToString)
		End Function

		''' <summary>
		''' Prüft, ob ein Account userName im AD existiert.
		''' </summary>
		Public Shared Function ExistsUserName(ByVal userName As String) As Boolean
			Dim sb = AdRepositoryHelper.Instance.CreateDefaultSelectBuilder
			sb.Select.Add(AdProperties.distinguishedName.ToString)
			sb.Where.Add(String.Format("{0}='{1}'", AdProperties.sAMAccountName.ToString, userName))
			Return sb.ExecuteScalar IsNot Nothing
		End Function

		''' <summary>
		''' Prüft das Kennwort von userName
		''' </summary>
		Public Shared Function IsPasswordCorrect(ByVal userName As String, ByVal password As String) As Boolean
			Dim pwd = New SecureString()
			password.ToCharArray.ToList.ForEach(Sub(c) pwd.AppendChar(c))
			pwd.MakeReadOnly()

			'Passwort wird einem Pointer übergeben, damit dieser später "entschlüsselt" werden kann 
			Dim pPwd = Marshal.SecureStringToBSTR(pwd)

			Try
				Dim domain = SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Domain)
				Using entry = New DirectoryEntry(domain.ToUrl, userName, Marshal.PtrToStringBSTR(pPwd))
					Dim nativeObject = entry.NativeObject
					Return True
				End Using
			Catch ex As Exception
				Return False
			End Try
		End Function

		''' <summary>
		''' Überprüft, ob eine Anmeldung mit angegebenem userName und password möglich ist.
		''' </summary>
		Public Shared Function IsLoginSuccessful _
		(ByVal userName As String, password As String) As LoginResults

			Select Case True
				Case Not AdInformation.ExistsUserName(userName)
					Return LoginResults.InvalidUserName
				Case AdInformation.IsAccountDeactive(userName)
					Return LoginResults.AccountDeactivated
				Case AdInformation.IsAccountExpired(userName)
					Return LoginResults.AccountExpired
				Case Not IsPasswordCorrect(userName, password)
					Return LoginResults.InvalidPwd
				Case Else
					Return LoginResults.Successful
			End Select
		End Function

		''' <summary>
		''' Überprüft, ob eine SingleSignOn Anmeldung mit angegebenem userName möglich ist.
		''' </summary>
		Public Shared Function IsLoginSingleSignOnSuccessful _
		(ByVal userName As String) As LoginResults

			Select Case True
				Case Not ExistsUserName(userName)
					Return LoginResults.InvalidUserName
				Case IsAccountDeactive(userName)
					Return LoginResults.AccountDeactivated
				Case IsAccountExpired(userName)
					Return LoginResults.AccountExpired
				Case IsAccountExpired(userName)
					Return LoginResults.AccountExpired
				Case Else
					Return LoginResults.Successful
			End Select
		End Function

		''' <summary>
		''' Prüft, ob der Account von userName abgelaufen ist.
		''' </summary>
		Public Shared Function IsAccountExpired(ByVal userName As String) As Boolean
			Dim sb = AdRepositoryHelper.Instance.CreateDefaultSelectBuilder
			sb.Select.Add(AdProperties.accountExpires.ToString)
			sb.Where.Add(String.Format("{0}='{1}'", AdProperties.sAMAccountName.ToString, userName))
			Dim expireDate = AdTypeConverter.LargeIntegerToDate(sb.ExecuteScalar)

			Select Case True
				Case Not expireDate.HasValue
					Return False
				Case Else
					Return expireDate.Value < DateTime.Today
			End Select
		End Function

		''' <summary>
		''' Prüft, ob der Account des angegebenen Usernames deaktiv ist.
		''' </summary>
		Public Shared Function IsAccountDeactive(ByVal userName As String) As Boolean

			Dim sb = AdRepositoryHelper.Instance.CreateDefaultSelectBuilder
			sb.Select.Add(AdProperties.userAccountControl.ToString)
			sb.Where.Add(String.Format("{0}='{1}'", AdProperties.sAMAccountName.ToString, userName))

			Select Case CType(sb.ExecuteScalar, UserAccountControlTypes)
				Case UserAccountControlTypes.AccountEnabled _
				, UserAccountControlTypes.AccountEnabledPasswordNotRequired _
				, UserAccountControlTypes.AccountEnabledPasswordNotExpire _
				, UserAccountControlTypes.AccountEnabledPasswordNotExpireAndNotRequired _
				, UserAccountControlTypes.AccountEnabledSmartcardRequired _
				, UserAccountControlTypes.AccountEnabledSmartcardRequiredPasswordNotRequired _
				, UserAccountControlTypes.AccountEnabledSmartcardRequiredPasswordNotExpire _
				, UserAccountControlTypes.AccountEnabledSmartcardRequiredPasswordNotExpireAndNotRequired
					Return False
				Case Else
					Return True
			End Select
		End Function

		''' <summary>
		''' Prüft, ob der Account des angegebenen Usernames aktiv ist.
		''' </summary>
		Public Shared Function IsAccountActive(ByVal userName As String) As Boolean

			Return Not IsAccountDeactive(userName)
		End Function

		''' <summary>
		''' Liefert alle Werte der objectClass-Eigenschaft für das übergebene dn-Objekt.
		''' </summary>
		Public Shared Function GetObjectClassValues(ByVal dn As DistinguishedName) As ObjectClasses()

			Dim sb = AdRepositoryHelper.Instance.CreateDefaultSelectBuilder
			sb.Select.Add(AdProperties.objectClass.ToString)
			sb.Where.Add(String.Format("{0}='{1}'", AdProperties.distinguishedName.ToString, dn.Value))

			Try
				Return DirectCast(sb.ExecuteScalar, Object()).Cast(Of String).Select _
				(Function(s) CType([Enum].Parse(GetType(ObjectClasses), s), ObjectClasses)).ToArray
			Catch ex As Exception
				Return New ObjectClasses() {}
			End Try
		End Function

		''' <summary>
		''' Prüft, ob das übergebene dn-Objekt, die angegebene objectClass beinhaltet.
		''' </summary>
		Public Shared Function IsObjectClass(ByVal dn As DistinguishedName, ByVal objectClass As ObjectClasses) As Boolean
			Return GetObjectClassValues(dn).Contains(objectClass)
		End Function

		''' <summary>
		''' Prüft, ob das übergebene dn-Objekt, die objectClass Group beinhaltet.
		''' </summary>
		Public Shared Function IsGroup(ByVal dn As DistinguishedName) As Boolean
			Return IsObjectClass(dn, ObjectClasses.group)
		End Function

		''' <summary>
		''' Prüft, ob das übergebene dn-Objekt, die objectClass User beinhaltet.
		''' </summary>
		Public Shared Function IsUser(ByVal dn As DistinguishedName) As Boolean
			Return IsObjectClass(dn, ObjectClasses.user)
		End Function

		''' <summary>
		''' Prüft, ob das übergebene dn-Objekt, die objectClass OrganizationalUnit beinhaltet.
		''' </summary>
		Public Shared Function IsOrganizationalUnit(ByVal dn As DistinguishedName) As Boolean
			Return IsObjectClass(dn, ObjectClasses.organizationalUnit)
		End Function

#End Region '{Öffentliche Methoden der Klasse}

	End Class

End Namespace




