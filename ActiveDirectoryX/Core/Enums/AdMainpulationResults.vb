Option Explicit On
Option Infer On
Option Strict On

Namespace Core.Enums

	Public Enum AdManipulationResults As Int32
		Successful
		SuccessfulWithSomeErrors
		UnknownError
		UnknownUserPrincipal
		UnknownGroupPrincipal
		MemberAlreadyExist
		MemberNotExist
		MemberOfAlreadyExist
		MemberOfNotExist
		GroupIsNotRole
		GroupIsNotOrganizationGroup
		GroupIsNotAdminstrationGroup
		GroupIsNotDomainGroup
		GroupIsNotGrant
		IsNotGroup
		IsNotUser
		GroupIsNotOrganizationUnit
		SetDefaultAttributesError
		SetUnixAttributesError
		GenerateMailboxError
		SetPasswordError
		SetExpirePasswordNowError
		SetUserEnabledError
		IsNotApplication
		GroupIsNotMapping
		MappingError
		AccesDenied
		InvalidMappingName
		InvalidOrganizationGroupType
		InvalidGroupManagerName
		UserNotExists
	End Enum

End Namespace


