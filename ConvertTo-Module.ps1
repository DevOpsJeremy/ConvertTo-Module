function ConvertTo-Module
{
    <# 
    .SYNOPSIS
        Converts a PowerShell script of functions into a module.
    .DESCRIPTION
        This function takes a simple PowerShell script (.ps1) of functions, creates the directory structure for a PowerShell module, generates the module manifest (`.psd1) and script module (`.psm1), and populates the module directories with the functions.

        The directory structure is as follows:
            <Module Name>\
                private\
                    functions\
                        <Private Function Files (`.ps1`)>
                    Types.ps1
                public\
                    functions\
                        <Public Function Files (`.ps1`)>
                <Module Name>.psd1
                <Module Name>.psm1
    .PARAMETER Name
        The name of the module. This will be used as the directory, mainifest, and module script names.
    .PARAMETER Source
        Source script from which to create the module.
    .PARAMETER Destination
        Destination directory. A new sub-directory will be created using the name provided to the `-Name` parameter. Default is the current directory.
    .PARAMETER PrivateFunctions
        Any functions from the `-Source` script which do not need to be exported for use. This is typically for any functions which are only used by other functions in the module and do not need to be available to users at the console, etc.
    .PARAMETER CompatiblePSEditions
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-compatiblepseditions
    .PARAMETER NestedModules
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-nestedmodules
    .PARAMETER Guid
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-guid
    .PARAMETER Author
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-author
    .PARAMETER CompanyName
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-companyname
    .PARAMETER Copyright
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-copyright
    .PARAMETER ModuleVersion
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-moduleversion
    .PARAMETER Description
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-description
    .PARAMETER ProcessorArchitecture
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-processorarchitecture
    .PARAMETER PowerShellVersion
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-powershellversion
    .PARAMETER CLRVersion
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-clrversion
    .PARAMETER DotNetFrameworkVersion
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-dotnetframeworkversion
    .PARAMETER PowerShellHostName
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-powershellhostname
    .PARAMETER PowerShellHostVersion
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-powershellhostversion
    .PARAMETER RequiredModules
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-requiredmodules
    .PARAMETER TypesToProcess
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-typestoprocess
    .PARAMETER FormatsToProcess
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-formatstoprocess
    .PARAMETER ScriptsToProcess
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-scriptstoprocess
    .PARAMETER RequiredAssemblies
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-requiredassemblies
    .PARAMETER FileList
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-filelist
    .PARAMETER ModuleList
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-modulelist
    .PARAMETER FunctionsToExport
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-functionstoexport
    .PARAMETER AliasesToExport
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-aliasestoexport
    .PARAMETER VariablesToExport
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-variablestoexport
    .PARAMETER CmdletsToExport
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-cmdletstoexport
    .PARAMETER DscResourcesToExport
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-dscresourcestoexport
    .PARAMETER PrivateData
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-privatedata
    .PARAMETER Tags
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-tags
    .PARAMETER ProjectUri
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-projecturi
    .PARAMETER LicenseUri
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-licenseuri
    .PARAMETER IconUri
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-iconuri
    .PARAMETER ReleaseNotes
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-releasenotes
    .PARAMETER Prerelease
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-prerelease
    .PARAMETER ExternalModuleDependencies
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-externalmoduledependencies
    .PARAMETER HelpInfoUri
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-helpinfouri
    .PARAMETER DefaultCommandPrefix
        Pass-through parameter for `New-ModuleManifest`. See Microsoft documentation: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest#-defaultcommandprefix
    .OUTPUTS
        System.IO.DirectoryInfo
    .LINK
        https://devopsjeremy.github.io/documentation/powershell/ConvertTo-Module.html
    .LINK
        About PowerShell Modules: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_modules
    .LINK
        New-ModuleManifest: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/new-modulemanifest
    .EXAMPLE
        ConvertTo-Module -Name Confluence -Source Confluence.ps1 -PrivateFunctions ConfluenceExpandPropertyArgumentCompleter -Author 'Jeremy Watkins'
        
        This command creates the Confluence PowerShell module and directory structure. It exports the private function(s) (ConfluenceExpandPropertyArgumentCompleter) into the Confluence\private\functions directory and sets the author as 'Jeremy Watkins'.
    .EXAMPLE
        ConvertTo-Module -Name ConfiForms -Source ConfiForms.ps1 -Author 'Jeremy Watkins' -Description 'ConfiForms functions' `
        Import-Module ConfiForms\ConfiForms.psd1
        
        These commands create the ConfiForms PowerShell module and directory structure, sets the author and description, then imports the module.
    .NOTES
            Ver
    #>
    param (
        [Parameter(
            Mandatory,
            Position = 0
        )]
        [string] $Source,
        [String] $Name,
        [string] $Destination = $PWD.Path,
        [String[]] $PrivateFunctions,
        [ValidateSet(
            'Desktop',
            'Core'
        )]
        [String[]] $CompatiblePSEditions = @(
            'Destkop',
            'Core'
        ),
        [Object[]] $NestedModules,
        [System.Guid] $Guid,
        [String] $Author,
        [String] $CompanyName,
        [String] $Copyright,
        [System.Version] $ModuleVersion,
        [String] $Description,
        [System.Reflection.ProcessorArchitecture] $ProcessorArchitecture,
        [System.Version] $PowerShellVersion,
        [System.Version] $CLRVersion,
        [System.Version] $DotNetFrameworkVersion,
        [String] $PowerShellHostName,
        [System.Version] $PowerShellHostVersion,
        [Object[]] $RequiredModules,
        [String[]] $TypesToProcess,
        [String[]] $FormatsToProcess,
        [String[]] $ScriptsToProcess,
        [String[]] $RequiredAssemblies,
        [String[]] $FileList,
        [Object[]] $ModuleList,
        [String[]] $FunctionsToExport,
        [String[]] $AliasesToExport,
        [String[]] $VariablesToExport,
        [String[]] $CmdletsToExport,
        [String[]] $DscResourcesToExport,
        [Object] $PrivateData,
        [String[]] $Tags,
        [System.Uri] $ProjectUri,
        [System.Uri] $LicenseUri,
        [System.Uri] $IconUri,
        [String] $ReleaseNotes,
        [String] $Prerelease,
        [String[]] $ExternalModuleDependencies,
        [String] $HelpInfoUri,
        [String] $DefaultCommandPrefix
    )
    #region Functions
    function New-TypeDefinitionFile {
        param (
            [String] $Path,
            [System.Management.Automation.Language.ScriptBlockAst] $Parser
        )
        $TypeDefAstList = $Parser.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.TypeDefinitionAst] } | Sort-Object -Property IsEnum -Descending
        Set-Content `
            -Path $Path `
            -Value $TypeDefAstList.Extent.Text
    }
    function Get-Assemblies {
        <#
            .DESCRIPTION
            Searches through the script (ScriptBlockAst) for instances of the 'Add-Type' command, where the type being added is an AssemblyName, Path, or LiteralPath. Returns a string array of assembly types.
        #>
        param (
            [System.Management.Automation.Language.ScriptBlockAst] $ScriptBlockAst
        )
        # Getting all Parameter value objects from script where the parameter starts with 'a' ('AssemblyName'), 'pat' ('Path'), or 'li' ('LiteralPath')
        $Assemblies = $ScriptBlockAst.FindAll(
            {
                param ($node)
                $node.Parent -is [System.Management.Automation.Language.CommandAst] -and
                $node.Parent.GetCommandName() -eq 'Add-Type' -and
                (
                    $index = $node.Parent.CommandElements.IndexOf(
                        $node
                    )
                ) -and
                $node.Parent.CommandElements[$index - 1] -is [System.Management.Automation.Language.CommandParameterAst] -and
                $node.Parent.CommandElements[$index - 1].ParameterName -match '^(a|pat|li)'
            },
            $true
        )
        # Return only the value
        if ($Assemblies){
            $Assemblies.SafeGetValue()
        }
    }
    function New-ModuleFile {
        param (
            [String] $Path,
            [System.IO.FileInfo] $Script
        )
        function Get-ScriptHelpBlock {
            param (
                [System.IO.FileInfo] $Script
            )
            $Content = Get-Content $Script
            $line = 0
            while ($Content[$line] -notmatch '<#' -and $line -le $Content.Count){
                $line++
            }
            $commentBlockArray = @()
            do {
                $commentBlockArray += $Content[$line]
                $line++
            } until ($Content[$line] -match '#>' -or $line -gt $Content.Count)
            $commentBlockArray += $Content[$line++]
            return $commentBlockArray
        }
        Set-Content -Path $Path -Value @(
            (Get-ScriptHelpBlock -Script $Script),
            "Get-ChildItem (Split-Path `$script:MyInvocation.MyCommand.Path) -Filter '*.ps1' -Recurse | ForEach-Object { ",
            "    . `$_.FullName ",
            "}"
        )
    }
    #endregion Functions

    # Resolve paths
    $SourceObj = Get-Item $(Resolve-Path -Path $Source)
    $DestinationObj = Get-Item $(Resolve-Path -Path $Destination)

    # Parse functions from script
    if (!$Name){
        $Name = $SourceObj.BaseName
    }

    $scriptParser = [System.Management.Automation.Language.Parser]::ParseFile(
        $SourceObj.FullName,
        [ref] $null,
        [ref] $null
    )
    $functionParser = $scriptParser.EndBlock.Statements.Where(
        { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] }
    )

    # Create directory structure
    $rootPath = "$($DestinationObj.FullName)\$Name"
    @{
        public = @(
            'functions'
        )
        private = @(
            'functions'
        )
    }.GetEnumerator() | ForEach-Object {
        $varName = "{0}Path" -f $_.Key
        $varValue = "{0}\{1}" -f $rootPath, $_.Key
        New-Variable -Name $varName -Value $varValue -Force
        foreach ($subDir in $_.Value){
            $varName = "{0}{1}{2}" -f $_.Key, $subDir, 'Path'
            $varValue = "{0}\{1}\{2}" -f $rootPath, $_.Key, $subDir
            New-Variable -Name $varName -Value $varValue -Force
        }
    }
    @{
        manifest = 'psd1'
        module = 'psm1'
    }.GetEnumerator() | ForEach-Object {
        New-Variable -Name "$($_.Key)Path" -Value "$rootPath\$Name.$($_.Value)" -Force
    }
    $folders = @(
        $rootPath,
        $publicFunctionsPath,
        $privateFunctionsPath
    )
    $files = @(
        $manifestPath,
        $modulePath
    )
    New-Item -Path $folders -ItemType Directory | Out-Null
    New-Item -Path $files -ItemType File | Out-Null

    # Export functions to individual script files
    foreach ($function in $functionParser){
        $functionDir = if ($function.Name -in $PrivateFunctions){
            $privateFunctionsPath
        } else {
            $publicFunctionsPath
        }
        Set-Content -Path "$functionDir\$($function.Name).ps1" -Value $function.Extent.Text
    }

    # Create types file 
    New-TypeDefinitionFile -Path "$privatePath\Types.ps1" -Parser $scriptParser

    # Create module file 
    New-ModuleFile -Path $modulePath -Script $SourceObj # -ScriptBlockAst $scriptParser

    # Create module manifest
    $manifestParams = @{
        Path = $manifestPath
        RootModule = Split-Path $modulePath -Leaf
        FunctionsToExport = $functionParser.Name | Where-Object { $_ -notin $PrivateFunctions }
        RequiredAssemblies = Get-Assemblies -ScriptBlockAst $scriptParser
    }
    foreach (
        $parameter in $PSBoundParameters.GetEnumerator() | 
            Where-Object {
                $_.Key -notin @(
                    'Name',
                    'Source',
                    'Destination',
                    'PrivateFunctions'
                )
            }
    ){
        $manifestParams[$parameter.Key] = $parameter.Value
    }
    New-ModuleManifest @manifestParams | Out-Null
    return (Get-Item $rootPath)
}
