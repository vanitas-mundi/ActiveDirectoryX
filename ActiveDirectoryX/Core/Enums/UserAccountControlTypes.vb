Option Explicit On
Option Infer On
Option Strict On

Namespace Core.Enums

	Public Enum UserAccountControlTypes
		AccountEnabled = 512
		AccountEnabledPasswordNotRequired = 544
		AccountEnabledPasswordNotExpire = 66048
		AccountEnabledPasswordNotExpireAndNotRequired = 66080
		AccountEnabledSmartcardRequired = 262656
		AccountEnabledSmartcardRequiredPasswordNotRequired = 262688
		AccountEnabledSmartcardRequiredPasswordNotExpire = 328192
		AccountEnabledSmartcardRequiredPasswordNotExpireAndNotRequired = 328224

		AccountDisabled = 514
		AccountDisabledPasswordNotRequired = 546
		AccountDisabledPasswordNotExpire = 66050
		AccountDisabledPasswordNotExpireAndNotRequired = 66082
		AccountDisabledSmartcardRequired = 262658
		AccountDisabledSmartcardRequiredPasswordNotRequired = 262690
		AccountDisabledSmartcardRequiredPasswordNotExpire = 328194
		AccountDisabledSmartcardRequiredPasswordNotExpireAndNotRequired = 328226
	End Enum

End Namespace
