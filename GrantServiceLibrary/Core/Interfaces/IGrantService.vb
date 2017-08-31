Option Explicit On
Option Infer On
Option Strict On

Namespace Core.Interfaces

  <ServiceContract()>
  Public Interface IGrantService

    <OperationContract()>
    Function GetGrantString(ByVal appName As String, ByVal userName As String, ByVal pwd As String) As String

    <OperationContract()>
    Function GetGrantStringByPersonId(ByVal appName As String, ByVal personId As Int32, ByVal pwd As String) As String

    <OperationContract()>
    Function GetGrantStringSingleSignOn(ByVal appName As String, ByVal userName As String) As String

    <OperationContract()>
    Function GetGrantStringSingleSignOnByPersonId(ByVal appName As String, ByVal personId As Int32) As String

    <OperationContract()>
    Function GetAppNames() As String()

    <OperationContract()>
    Function GetUserProperty(ByVal userName As String, ByVal propertyName As String) As String

    <OperationContract()>
    Function GetUserPropertyByPersonId(ByVal personId As Int32, ByVal propertyName As String) As String

    <OperationContract()>
    Function GetHolidayGroupNamesFromUser(ByVal personId As Int32) As String

    <OperationContract()>
    Function IsUserInHolidayGroup(ByVal personId As Int32, ByVal holidayGroupName As String) As Boolean

    <OperationContract()>
    Function GetUsersOfHolidayGroup(ByVal holidayGroupName As String) As String

    <OperationContract()>
    Function GetHolidayGroupManagerPersonIds(ByVal holidayGroupName As String) As String

    <OperationContract()>
    Function GetGroupNamesByManagerPersonId(ByVal personId As Int32) As String

    <OperationContract()>
    Function GetAllHolidayGroups() As String()
  End Interface

End Namespace
