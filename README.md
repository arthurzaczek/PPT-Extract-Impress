Extracts power point slides to impress js using powershell

**ITS UNDER DEVELOPMENT!**

	C:\PS> Get-Help .\PPT-Extract-Impress.ps1 -detailed

	NAME
		PPT-Extract-Impress.ps1

	SYNOPSIS
		Extracts power point slides to impress js


	SYNTAX
		PPT-Extract-Impress.ps1 [[-file] <String>] [-Simple] [-Position] [-FontStyle] [-SourceCode] [-Open] [<CommonParameters>]


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

		C:\PS> .\PPT-Extract-Impress.ps1 'MyPresentation.pptx'


		-------------------------- EXAMPLE 2 --------------------------

		C:\PS> ls *.pptx | % { .\PPT-Extract-Impress.ps1 $_ }