
function Get-SolutionDir {
    if($dte.Solution -and $dte.Solution.IsOpen) {
        return Split-Path $dte.Solution.Properties.Item("Path").Value
    }
    else {
        throw "Solution not avaliable"
    }
}

function Get-SolutionDir {
    if($dte.Solution -and $dte.Solution.IsOpen) {
        return Split-Path $dte.Solution.Properties.Item("Path").Value
    }
    else {
        throw "Solution not avaliable"
    }
}

function Resolve-ProjectName {
    param(
        [parameter(ValueFromPipelineByPropertyName = $true)]
        [string[]]$ProjectName
    )
    
    if($ProjectName) {
        $projects = Get-Project $ProjectName
    }
    else {
        # All projects by default
        $projects = Get-Project -All
    }
    
    $projects
}

function Get-MSBuildProject {
    param(
        [parameter(ValueFromPipelineByPropertyName = $true)]
        [string[]]$ProjectName
    )
    Process {
        (Resolve-ProjectName $ProjectName) | % {
            $path = $_.FullName
            @([Microsoft.Build.Evaluation.ProjectCollection]::GlobalProjectCollection.GetLoadedProjects($path))[0]
        }
    }
}

function Set-MSBuildProperty {
    param(
        [parameter(Position = 0, Mandatory = $true)]
        $PropertyName,
        [parameter(Position = 1, Mandatory = $true)]
        $PropertyValue,
        [parameter(Position = 2, ValueFromPipelineByPropertyName = $true)]
        [string[]]$ProjectName
    )
    Process {
        (Resolve-ProjectName $ProjectName) | %{
            $buildProject = $_ | Get-MSBuildProject
            $buildProject.SetProperty($PropertyName, $PropertyValue) | Out-Null
            $_.Save()
        }
    }
}

function Get-MSBuildProperty {
    param(
        [parameter(Position = 0, Mandatory = $true)]
        $PropertyName,
        [parameter(Position = 2, ValueFromPipelineByPropertyName = $true)]
        [string]$ProjectName
    )
    
    $buildProject = Get-MSBuildProject $ProjectName
    $buildProject.GetProperty($PropertyName)
}

function Install-NuSpec {
    param(
        [parameter(ValueFromPipelineByPropertyName = $true)]
        [string[]]$ProjectName,
    	[switch]$EnableIntelliSense,
        [string]$TemplatePath
    )
    
    Process {
    
        $projects = (Resolve-ProjectName $ProjectName)
        
        if(!$projects) {
            Write-Error "Unable to locate project. Make sure it isn't unloaded."
            return
        }
		
		$profileDirectory = Split-Path $profile -parent
		$profileModulesDirectory = (Join-Path $profileDirectory "Modules")
		$moduleDir = (Join-Path $profileModulesDirectory "NuSpec")
		
        if($EnableIntelliSense){
            Enable-NuSpecIntelliSense            
        }
        
        # Add NuSpec file for project(s)
        $projects | %{ 
            $project = $_
            
            # Set the nuspec target path
            $projectFile = Get-Item $project.FullName
            $projectDir = [System.IO.Path]::GetDirectoryName($projectFile)
            $projectNuspec = "$($project.Name).tmpl.nuspec"
            $projectNuspecPath = Join-Path $projectDir $projectNuspec
            
            # Get the nuspec template source path
            if($TemplatePath) {
                $nuspecTemplatePath = $TemplatePath
            }
            else {
                $nuspecTemplatePath = Join-Path $moduleDir NuSpecTemplate.xml
            }
            
            # Copy the templated nuspec to the project nuspec if it doesn't exist
            if(!(Test-Path $projectNuspecPath)) {
                Copy-Item $nuspecTemplatePath $projectNuspecPath
            }
            else {
                Write-Warning "Failed to install nuspec '$projectNuspec' into '$($project.Name)' because the file already exists."
            }
            
            try {
                # Add nuspec file to the project
                $project.ProjectItems.AddFromFile($projectNuspecPath) | Out-Null
                $project.Save()
				
				Set-MSBuildProperty NuSpecFile $projectNuspec $project.Name
                
                Write-Verbose "Updated '$($project.Name)' to use nuspec '$projectNuspec'"
            }
            catch {
                Write-Warning "Failed to install nuspec '$projectNuspec' into '$($project.Name)'"
            }
        }
    }
}

function Enable-NuSpecIntelliSense {
    Process {		
		$profileDirectory = Split-Path $profile -parent
		$profileModulesDirectory = (Join-Path $profileDirectory "Modules")
		$moduleDir = (Join-Path $profileModulesDirectory "NuSpec")

        $solutionDir = Get-SolutionDir
        $solution = Get-Interface $dte.Solution ([EnvDTE80.Solution2])
        
        # Set up solution folder "Solution Items"
        $solutionItemsProject = $dte.Solution.Projects | Where-Object { $_.ProjectName -eq "Solution Items" }
        if(!($solutionItemsProject)) {
            $solutionItemsProject = $solution.AddSolutionFolder("Solution Items")
        }        
        
        # Copy the XSD in the solution directory
        try {
            $xsdInstallPath = Join-Path $solutionDir 'nuspec.xsd'
            $xsdToolsPath = Join-Path $moduleDir 'nuspec.xsd'
                
            if(!(Test-Path $xsdInstallPath)) {
                Copy-Item $xsdToolsPath $xsdInstallPath
            }
                
            $alreadyAdded = $solutionItemsProject.ProjectItems | Where-Object { $_.Name -eq 'nuspec.xsd' }
            if(!($alreadyAdded)) {
                $solutionItemsProject.ProjectItems.AddFromFile($xsdInstallPath) | Out-Null
            }
        }
        catch {
            Write-Warning "Failed to install nuspec.xsd into 'Solution Items'"
        }
        $solution.SaveAs($solution.FullName)
    }
}


function Resolve-ProjectName2 {
    param(
        [parameter(ValueFromPipelineByPropertyName = $true)]
        [string[]]$ProjectName
    )
    
    if($ProjectName) {
        $projects = Get-Project $ProjectName
    }
	
	$projects | %{ 
		$path = $_.FullName 	
		$content = (Get-Content $path)
		$content
		$xcontent =[xml] $content
		Write-Host "-------------xx-------------"
		$xcontent.InnerText
		$xcontent.LastChild.ChildNodes | %{ Write-Host "  " $_; $_.ChildNodes | %{ Write-Host "    " $_}; Write-Host "  /" $_ }
		
		Write-Host "-------------yy-------------"
		
		$xcontent.LastChild.ItemGroup | %{ Write-Host "Boo"; $_.Reference | %{ Write-Host $_ } ; }
				
		Write-Host "-------------yy-------------"
		$xcontent.Project.ItemGroup | %{ $_.Reference |% { Write-Host "Reference ", $_.Include } }
		
		Write-Host "-------------zz-------------"
		$xcontent.Project.ItemGroup.Reference |% { Write-Host "Reference ", $_.Include } 
				
		Write-Host "-------------aa-------------"
		$xcontent.Project.ItemGroup.Reference.Include | %{Write-Host "Reference ", $_ }

		Write-Host "-------------bb-------------"
		$xcontent.Project.ItemGroup.ProjectReference.Include | %{Write-Host "ProjectReference ", $_ }		
	}
	
    $projects
}

$global:processedProjects = @{}
Function Write-Nuspec {
	param(
        [parameter(Position=0, 
			ValueFromPipelineByPropertyName = $true, 
			ParameterSetName="Default")]
        [string[]]$ProjectName,
		
        [parameter( 
			ValueFromPipelineByPropertyName = $true, 
			ParameterSetName="Path")]
        [string]$ProjectPath
    )
    
	$originalPath = pwd
	$directory = $originalPath
	
    if(![string]::IsNullOrWhiteSpace($ProjectName)) {
        Write-Debug "ProjectName: $ProjectName"
		$projects = Get-Project $ProjectName
		$ProjectPath = $projects.FullName
    }
	elseif (![string]::IsNullOrWhiteSpace($ProjectPath)) {
		Write-Debug "ProjectPath: $ProjectPath"
	}
	else {
		Write-Verbose "No path to work with"
		return;
	}

	$ProjectPath | %{ 

		# Get the directory
		$projectFolder = (Split-Path $_ -Parent)
		if ([System.IO.Path]::IsPathRooted($projectFolder)) {
			$directory = $projectFolder
		}
		elseif (![string]::IsNullOrWhiteSpace($projectFolder)) {		
			$directory = Join-Path $directory $projectFolder
		}
	
		# 
		$projectXml =[xml] (Get-Content $_)
		$projectReferencePaths = $projectXml.Project.ItemGroup.ProjectReference.Include 
		
		$targetProjectName = [System.IO.Path]::GetFileNameWithoutExtension($_)
		$templateName = "{0}.tmpl.nuspec" -f $targetProjectName 
		$templatePath = Join-Path $directory $templateName 
		if (!(Test-Path $templatePath)) { 
			Install-Nuspec $targetProjectName 
		}
		$outputNuspecName = "{0}.nuspec" -f $targetProjectName 
		$nuspecPath = Join-Path $directory $outputNuspecName 

		$nuspecXml = [xml] (Get-Content $templatePath)
	
		$dependencies = $nuspecXml.Package.Metadata.Dependencies
		if ($dependencies -eq $null)
		{
			$dependencies = $nuspecXml.CreateElement("dependencies")
			$nuspecXml.Package.Metadata.AppendChild($dependencies) | Out-Null
		}
	
		if ($projectReferencePaths.Length -ne 0) {
			$projectReferencePaths | %{
				$projectReferencePath = $_
				$projectName = $_ | Split-Path -Leaf | %{ [System.IO.Path]::GetFileNameWithoutExtension($_) }
					
				if (!$global:processedProjects.ContainsKey($projectName))
				{					
					$version = $nuspecXml.Package.Metadata.Version

					Write-Debug "Adding project $projectName version ----------------> $version"
					Write-Debug "This may have been added multiple times. I do not know why"

					$global:processedProjects.Add($projectName, $version)
					Write-Nuspec -ProjectPath $projectReferencePath
				}					
				else {
					$version = $global:processedProjects[$projectName]
					Write-Debug "$projectName version ----------------> $version"
				}

				$dependency = $nuspecXml.CreateElement("dependency")
				$dependency.SetAttribute("id", $projectName)
				$dependency.SetAttribute("version", $version)
				$dependencies.AppendChild($dependency) | Out-Null
				
			}
		}
				
		$nuspecXml.Save($nuspecPath )	
		Write-Verbose ("Saved changes to $nuspecPath" )
	}
}


Function Publish-SolutionProject {
	param(
        [parameter(Position=0, 
			ValueFromPipelineByPropertyName = $true, 
			ParameterSetName="Default")]
        [string[]]$ProjectName,
		
        [parameter( 
			ValueFromPipelineByPropertyName = $true, 
			ParameterSetName="Path")]
        [string]$ProjectPath,

		[parameter(Mandatory=$true,	ParameterSetName="Default")]
		[parameter(Mandatory=$true,	ParameterSetName="Path")]
        [string]$NugetServer,

		[parameter(Mandatory=$true,	ParameterSetName="Default")]
		[parameter(Mandatory=$true,	ParameterSetName="Path")]
        [string]$ApiKey
    )
    
	Write-Verbose "Beginning to publish project to nuget server"

	$originalPath = pwd
	$directory = $originalPath
	
    if(![string]::IsNullOrWhiteSpace($ProjectName)) {
        Write-Debug "ProjectName: $ProjectName"
		$projects = Get-Project $ProjectName
		$ProjectPath = $projects.FullName
    }
	elseif (![string]::IsNullOrWhiteSpace($ProjectPath)) {
		Write-Debug "ProjectPath: $ProjectPath"
		$pushed = $true
	}
	else {
		Write-Verbose "No path to work with"
		return;
	}

	$ProjectPath | %{ 

		# Get the directory
		$projectFolder = (Split-Path $_ -Parent)
		if ([System.IO.Path]::IsPathRooted($projectFolder)) {
			$directory = $projectFolder
		}
		elseif (![string]::IsNullOrWhiteSpace($projectFolder)) {		
			$directory = Join-Path $directory $projectFolder
		}

		pushd $directory
	
		# 
		$projectXml =[xml] (Get-Content $_)
		$projectReferencePaths = $projectXml.Project.ItemGroup.ProjectReference.Include 
		
		$targetProjectName = [System.IO.Path]::GetFileNameWithoutExtension($_)

		# Differences below		
		$outputNuspecName = "{0}.nuspec" -f $targetProjectName 

		$nuspecPath = Join-Path $directory $outputNuspecName 
		if (!(Test-Path $nuspecPath)) { 
			Write-Nuspec $targetProjectName 
		}

		Write-Debug "nuspecPath $nuspecPath"
		$nuspecXml = [xml](Get-Content $nuspecPath)
		$nuspecVersion = $nuspecXml.Package.Metadata.Version
		$outputNupkgName = "{0}.{1}.nupkg" -f $targetProjectName, $nuspecVersion
					
		if ($projectReferencePaths.Length -ne 0) {
			$projectReferencePaths | %{
				$projectReferencePath = $_
				$projectName = $_ | Split-Path -Leaf | %{ [System.IO.Path]::GetFileNameWithoutExtension($_) }
				
				# Insert only once 
				if(!$Global:processedProjects.ContainsKey($projectName))
				#($Global:processedProjects | Where-Object { $_.Name -match $projectName }).Length -eq 0)
				{				
					$Global:processedProjects.Set_Item("$projectName", "$_")
					Publish-SolutionProject -ProjectPath $projectReferencePath `
						-NugetServer $NugetServer -ApiKey $ApiKey
				}					
			}
		}

		Write-Verbose ("nuget Pack $outputNuspecName")
		Write-Verbose (nuget Pack $outputNuspecName | Out-String)
		
		Write-Verbose ('#nuget Push "{0}" -Source "{1}" -ApiKey {2}' -f $outputNupkgName, $NugetServer, $ApiKey )
		Write-Verbose (nuget Push $outputNupkgName -Source $NugetServer -ApiKey $ApiKey | Out-String)
		popd
	}
}


Function Update-ProjectVersion
{
	param(	[parameter(ParameterSetName="Increment")]
			[string]$Increment,

			[parameter(ParameterSetName="Absolute")]
			[string]$Absolute
			)
	Write-Verbose "Beginning to update project versions"

	$projects = Get-Project -All

	$projects | %{
		$projectFile = $_.FullName
	
		$targetProjectName	 = [System.IO.Path]::GetFileNameWithoutExtension($projectFile)
		$targetProjectFolder = [System.IO.Path]::GetDirectoryName($projectFile)
		$templateFile = "{0}.tmpl.nuspec" -f $targetProjectName
		$templatePath = Join-Path $targetProjectFolder $templateFile

		$nuspecXml = [xml](Get-Content $templatePath)
		$nuspecVersion = $nuspecXml.Package.Metadata.Version

		if (![string]::IsNullOrWhitespace($Increment))  {
			$versionPartInts = $nuspecVersion.Split(".") | %{[int]$_}
			$incrementPartInts = $Increment.Split(".") | %{[int]$_}
			$version = ""; 0..3 | %{ 
				$versionPartInts[$_] = $versionPartInts[$_] + $incrementPartInts[$_]
			}
			$version = [string]::Join(".", $versionPartInts)
			$nuspecXml.Package.Metadata.Version = "$version"
		}
		elseif (![string]::IsNullOrWhitespace($Absolute)) {
			if ($Absolute -Match "\d+\.\d+\.\d+\.\d+") {
				$nuspecXml.Package.Metadata.Version = "$Absolute"
			}
		}
		else{
			do
			{
				$version = Read-Host ("Please enter a version for {0} [{1}]" -f $_.Name, $nuspecVersion)
			} while ($version -NotMatch "\d+\.\d+\.\d+\.\d+" -and ![string]::IsNullOrWhitespace($version))
			if (![string]::IsNullOrWhitespace($version))
			{
				$nuspecXml.Package.Metadata.Version = "$version"
			}
		}
				
		$nuspecXml.Save($templatePath)
		Write-Verbose "$templateFile now has version $version"
	}
}


#Install-Nuspec #"SampleProject1"
#Update-ProjectVersion -Absolute 0.0.0.1
#Write-Nuspec "SampleProject1" 
#Publish-SolutionProject "SampleProject1" -NugetServer "http://nuget.phillipgivens.com/" -ApiKey "keyofphillip"






# Statement completion for project names
'Install-Nuspec', 'Write-Nuspec', 'Publish-SolutionProject' | %{ 
    Register-TabExpansion $_ @{
        ProjectName = { Get-Project -All | Select -ExpandProperty Name }
    }
}

Export-ModuleMember Install-NuSpec, Enable-NuSpecIntelliSense, Get-SolutionDir, Install-Nuspec, `
	Write-Nuspec, Publish-SolutionProject, Update-ProjectVersion

















