Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Data.Repositories
#End Region

Namespace Core

  Partial Public Class DistinguishedName

#Region " --------------->> Private Methoden der Klasse "
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand einer PersonenID.
    ''' </summary>
    Public Shared Function GetByPersonId(ByVal personId As Int64) As DistinguishedName
      Return DistinguishedNameRepository.Instance.GetByPersonId(personId)
    End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand einer Guid.
    ''' </summary>
    Public Shared Function GetByGuid(ByVal guid As Guid) As DistinguishedName
      Return GetByGuid(guid.ToString)
    End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand eines Guid-Strings.
    ''' </summary>
    Public Shared Function GetByGuid(ByVal guid As String) As DistinguishedName
      Return DistinguishedNameRepository.Instance.GetByGuid(guid)
    End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand eines samAccountNames.
    ''' </summary>
    Public Shared Function GetByUserName(ByVal userName As String) As DistinguishedName
      Return DistinguishedNameRepository.Instance.GetByUserName(userName)
    End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand eines common names (cn).
    ''' </summary>
    Public Shared Function GetByCn(ByVal cn As String) As DistinguishedName
      Return DistinguishedNameRepository.Instance.GetByCn(cn)
    End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand eines common names (cn).
    ''' </summary>
    Public Shared Function GetByOu(ByVal ou As String) As DistinguishedName
      Return DistinguishedNameRepository.Instance.GetByOu(ou)
    End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand eines DisplayNames.
    ''' </summary>
    Public Shared Function GetByDisplayName(ByVal displayName As String) As DistinguishedName
      Return DistinguishedNameRepository.Instance.GetByDisplayName(displayName)
    End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand eines samAccountNames.
    ''' </summary>
    Public Shared Function GetByGroupName(ByVal groupName As String) As DistinguishedName
      Return DistinguishedNameRepository.Instance.GetByGroupName(groupName)
    End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand einer Berechtigung.
    ''' </summary>
    Public Shared Function GetByGrant(ByVal appName As String, ByVal grantName As String) As DistinguishedName
      Return DistinguishedNameRepository.Instance.GetByGrant(appName, grantName)
    End Function

    ''' <summary>
    ''' Liefert ein DistinguishedName-Objekt anhand eines DistinguishedName.
    ''' </summary>
    Public Shared Function GetByDistinguishedName(ByVal distinguishedName As String) As DistinguishedName
      Return DistinguishedNameRepository.Instance.GetByDistinguishedName(distinguishedName)
    End Function
#End Region '{Öffentliche Methoden der Klasse}

  End Class

End Namespace