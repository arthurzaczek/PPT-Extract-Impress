Extracts power point slides to impress js using powershell

**ITS UNDER DEVELOPMENT!**

	PS> Get-Help .\Extract-Powerpoint.ps1 -detailed

	NAME
		Extract-Powerpoint.ps1

	SYNOPSIS
		Extracts power point slides to impress js


	SYNTAX
		P:\Impress-Extract-Powerpoint\Extract-Powerpoint.ps1 [[-file] <String>] [-Simple] [-Position] [-FontStyle] [-Source
		Code] [-Open] [<CommonParameters>]


	DESCRIPTION
		Supports font-size, syntax highlighing, positioning and css classes for theming


	PARAMETERS
		-file <String>

		-Simple [<SwitchParameter>]

		-Position [<SwitchParameter>]

		-FontStyle [<SwitchParameter>]

		-SourceCode [<SwitchParameter>]

		-Open [<SwitchParameter>]

		-------------------------- EXAMPLE 1 --------------------------

		C:\PS>.\Extract-Powerpoint.ps1 'MyPresentation.pptx'


		-------------------------- EXAMPLE 2 --------------------------

		C:\PS>ls *.pptx | % { .\Extract-Powerpoint.ps1 $_ }