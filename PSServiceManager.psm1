# We need to call Get-Service here to add ServiceController assemblies to PS v3
# sessions to make the [ServiceStartMode] type available for cmdlet parameters.
Get-Service | Out-Null

. $PSScriptRoot\PSServiceManager.ServiceDefinition.ConfigurationCommand.ps1

<#
.SYNOPSIS
   Retrieve to object a PSServiceDefinition as emitted by PowerShell code saved
   to file. The "New-PSServiceDefinition" cmdlet is made available to the code
   to facilitate this.
.EXAMPLE
   Get-PSServiceDefinition C:\Path\To\DefinitionFile.ps1
.EXAMPLE
   "\\SERVER\Share\Path\To\DefinitionFile.ps1" | Get-PSServiceDefinition
.INPUTS
   System.String
.OUTPUTS
   System.Management.Automation.PSCustomObject w/ PSTypeName "PSServiceDefinition"
#>
function Get-PSServiceDefinition {
  [CmdletBinding(
    PositionalBinding = $false
  )]
  param(
    # Path to a file that emits a PSServiceDefinition. A PSServiceDefinition
    # retrieved by this cmdlet may be piped to this cmdlet as well.
    [Parameter(
      Mandatory = $true,
      Position = 0,
      ValueFromPipeline = $true,
      ValueFromPipelineByPropertyName = $true
    )]
    [Alias("DefinitionSourcePath")]
    [string]
    $LiteralPath
  )

  try {
    . $PSScriptRoot\PSServiceManager.ServiceDefinition.ConfigurationCommand.ps1

    & $LiteralPath |
      Add-Member -MemberType NoteProperty -Name DefinitionSourcePath -Value $LiteralPath -PassThru
  } catch {
    $PSCmdlet.ThrowTerminatingError($_)
  }
}

<#
.SYNOPSIS
   Install the Windows Service described by the PSServiceDefinition retrieved
   from the provided file path.
.EXAMPLE
   Get-PSServiceDefinition C:\Path\To\DefinitionFile.ps1 | Install-PSService
.INPUTS
   System.String
.OUTPUTS
   System.ServiceProcess.ServiceController
#>
function Install-PSService {
  [CmdletBinding(
    PositionalBinding = $false
  )]
  param(
    # Path to a file that emits a PSServiceDefinition. A PSServiceDefinition
    # retrieved by Get-ServiceDefinition may be piped to this cmdlet as well.
    [Parameter(
      Mandatory = $true,
      Position = 0,
      ValueFromPipeline = $true,
      ValueFromPipelineByPropertyName = $true
    )]
    [Alias("DefinitionSourcePath")]
    [string]
    $LiteralPath
  )
  try {
    $svcDef = Get-PSServiceDefinition -LiteralPath $LiteralPath
    
    $exe = @{}

    # As is semantically appropriate, all values replaced in this manner are
    # readonly properties in the c# source.
    $exe.SourceCode = (
      Get-Content -LiteralPath $PSScriptRoot\PSServiceManager.ExecutableSource.cs -Raw
    ).Replace(
      "%ServiceName%",
      $svcDef.ServiceName
    ).Replace(
      "%InstallPath%",
      $svcDef.InstallPath
    ).Replace(
      "%DataPath%",
      $svcDef.DataPath
    ).Replace(
      "%LogRoot%",
      $svcDef.LogPath
    ).Replace(
      "%SourceIsAvailable%",
      $svcDef.SourceIsAvailable.ToString().ToLower()
    )

    $exe.ResourcePaths = @()

    # The service definition will be added to the executable as a resource, and
    # must have a "known" name so we can find it from the compiled code.
    $exe.ResourcePaths += $sdPath = Join-Path -Path $env:TMP -ChildPath PSServiceManager.ServiceDefinition.ps1
    $exe.ResourcePaths += "$PSScriptRoot\PSServiceManager.ServiceDefinition.ConfigurationCommand.ps1"

    # Copy the service definition to the known location.
    Copy-Item -LiteralPath $LiteralPath -Destination $sdPath

    if (-not (Test-Path -LiteralPath $svcDef.InstallPath)) {
      New-Item -Path $svcDef.InstallPath -ItemType Directory |
        Out-Null
    }

    $exe.OutputPath = Join-Path -Path $svcDef.InstallPath -ChildPath "$($svcDef.ServiceName).exe"

    # Unless compilation occurs in a separate process, an open handle to the
    # resulting executable will preclude (e.g.) uninstall. CompilerParameters
    # will not survive a serialization round-trip; thus, they must be built
    # from component primitives passed as parameters.
    Start-Job -ScriptBlock {
      param($exe)

      $params = New-Object -TypeName System.CodeDom.Compiler.CompilerParameters

      $exe.ResourcePaths |
        ForEach-Object {
          $params.EmbeddedResources.Add($_) | Out-Null
        }

      $params.GenerateExecutable = $true
      $params.OutputAssembly = $exe.OutputPath

      # Due to some quirk of context, this is needed to make ServiceController
      # assemblies available within the scriptblock even as of PS v5.1
      Get-Service | Out-Null

      $params.ReferencedAssemblies.Add([System.ComponentModel.Component].Assembly.Location) | Out-Null
      $params.ReferencedAssemblies.Add([System.Dynamic.IDynamicMetaObjectProvider].Assembly.Location) | Out-Null
      $params.ReferencedAssemblies.Add([System.Management.Automation.PowerShell].Assembly.Location) | Out-Null
      $params.ReferencedAssemblies.Add([System.ServiceProcess.ServiceBase].Assembly.Location) | Out-Null

      Add-Type -TypeDefinition $exe.SourceCode -CompilerParameters $params
    } -ArgumentList $exe | Wait-Job | Remove-Job

    Remove-Item -LiteralPath $sdPath

    if ($svcDef.SourceIsAvailable) {
      $srcPath = Join-Path -Path $svcDef.InstallPath -ChildPath "$($svcDef.ServiceName).ServiceDefinition.ps1"

      Copy-Item -LiteralPath $LiteralPath -Destination $srcPath

      $readMePath = Join-Path -Path $svcDef.InstallPath -ChildPath "$($svcDef.ServiceName).!README.txt"

      $readMeText = @"
Since this service definition was compiled with the SourceIsAvailable setting
enabled, the definition as embedded in the service executable has been copied
to the service install path. If the executable detects a mismatch between the
embedded resource and exposed definition it will terminate on start.
"@

      New-Item -Path $readMePath -ItemType File -Value $readMeText |
        Out-Null

    }

    $service = New-Service `
    -BinaryPathName $exe.OutputPath `
    -Name $svcDef.ServiceName `
    -DisplayName $svcDef.ServiceDisplayName `
    -Description $svcDef.ServiceDescription `
    -StartupType $svcDef.StartupType `
    -Credential  $svcDef.Credential |
      Get-Service

    if ($svcDef.AutoStart) {
      $service |
        Start-Service
    }

    $service
  } catch {
    $PSCmdlet.ThrowTerminatingError($_)
  }
}

<#
.SYNOPSIS
   Uninstall the Windows Service described by the PSServiceDefinition retrieved
   from the provided file path, and remove associated program and data paths.
.EXAMPLE
   $svcDef = Get-PSServiceDefinition C:\Path\To\DefinitionFile.ps1

   $svcDef | Install-PSService | Out-Null

   $svcDef | Uninstall-PSService
.INPUTS
   System.String
.OUTPUTS
   None
#>
function Uninstall-PSService {
  [CmdletBinding(
    PositionalBinding = $false
  )]
  param(
    # Path to a file that emits a PSServiceDefinition. A PSServiceDefinition
    # retrieved by Get-ServiceDefinition may be piped to this cmdlet as well.
    [Parameter(
      Mandatory = $true,
      Position = 0,
      ValueFromPipeline = $true,
      ValueFromPipelineByPropertyName = $true
    )]
    [Alias("DefinitionSourcePath")]
    [string]
    $LiteralPath
  )
  try {
    $svcDef = Get-PSServiceDefinition -LiteralPath $LiteralPath
    
    $svc = Get-Service -Name $svcDef.ServiceName

    if ($svc.Status -ne "Stopped") {
      $svc |
        Stop-Service
    }

    $delResult = & sc.exe delete $svcDef.ServiceName

    Write-Verbose $delResult

    if ($LASTEXITCODE -ne 0) {
      throw "Service removal via sc.exe command-line failed with error $LastExitCode."
    }

    # On S2016, a 30-second wait is needed after service deletion before the
    # executable may be removed. Cause is unclear, but the specificity of a
    # 30-second wait (29 seconds is not enough) is interesting.
    Start-Sleep -Seconds 30

    $svcDef.LogPath,
    $svcDef.DataPath,
    $svcDef.InstallPath |
      ForEach-Object {
        if (Test-Path -LiteralPath $_) {
          Remove-Item -LiteralPath $_ -Recurse
        }
      }
  } catch {
    $PSCmdlet.ThrowTerminatingError($_)
  }
}

<#
.SYNOPSIS
   Send a control message to a service using a named pipe.
.EXAMPLE
   $svcDef = Get-PSServiceDefinition C:\Path\To\DefinitionFile.ps1

   $svcDef | Install-PSService | Out-Null

   $svcDef | Send-PSServiceMessage MyControlMessage
.EXAMPLE
   $service = Get-PSServiceDefinition C:\Path\To\DefinitionFile.ps1 | Install-PSService

   $service | Send-PSServiceMessage MyControlMessage
.EXAMPLE
   $service = Get-PSServiceDefinition C:\Path\To\DefinitionFile.ps1 | Install-PSService

   Send-PSServiceMessage -PipeName svcPipe_($service.Name) -Message MyControlMessage
.INPUTS
   System.String or System.ServiceProcess.ServiceController
.OUTPUTS
   None
#>
function Send-PSServiceMessage {
  [CmdletBinding(
    DefaultParameterSetName = "UsingSvcDef",
    PositionalBinding = $false
  )]
  param(
    # Path to a file that emits a PSServiceDefinition. A PSServiceDefinition
    # retrieved by Get-ServiceDefinition may be piped to this cmdlet as well.
    [Parameter(
      ParameterSetName = "UsingServiceDefinition",
      Mandatory = $true,
      ValueFromPipeline = $true,
      ValueFromPipelineByPropertyName = $true
    )]
    [Alias("DefinitionSourcePath")]
    [string]
    $LiteralPath,

    # A ServiceController object, as retrieved by Get-Service or emitted by
    # Install-PSService on service install.
    [Parameter(
      ParameterSetName = "UsingServiceController",
      Mandatory = $true,
      ValueFromPipeline = $true
    )]
    [System.ServiceProcess.ServiceController]
    $Service,

    # A named pipe to which to send the message.
    [Parameter(
      ParameterSetName = "UsingPipeName",
      Mandatory = $true
    )]
    [string]
    $PipeName,

    # The message content to send.
    [Parameter(
      Mandatory = $true,
      Position  = 0
    )]
    [string]
    $Message
  )
  try {
    if ($PSCmdlet.ParameterSetName -eq "UsingServiceDefinition") {
      $svcDef = Get-PSServiceDefinition -LiteralPath $LiteralPath

      $PipeName = "svcPipe_$($svcDef.ServiceName)"
    }
    elseif ($PSCmdlet.ParameterSetName -eq "UsingServiceController") {
      $PipeName = "svcPipe_$($service.Name)"
    }

    $pipe = New-Object -TypeName System.IO.Pipes.NamedPipeClientStream -ArgumentList @(
      ".",
      $PipeName,
      [System.IO.Pipes.PipeDirection]::Out
    )
    $pipe.Connect(1000)

    $sw = New-Object -TypeName System.IO.StreamWriter -ArgumentList $pipe
    $sw.AutoFlush = $true

    $sw.WriteLine($Message)

    $sw.Dispose()
    $pipe.Dispose()
  } catch {
    $PSCmdlet.ThrowTerminatingError($_)
  }
}

Export-ModuleMember New-PSServiceDefinition,
                    Get-PSServiceDefinition,
                    Install-PSService,
                    Uninstall-PSService,
                    Send-PSServiceMessage