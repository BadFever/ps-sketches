param(
    [Parameter(Mandatory, Position = 1)]
    [String]$VIServer,

    [Parameter(Mandatory, Position = 2)]
    [String]$User,

    [Parameter(Mandatory, Position = 3)]
    [String]$PW,

    [Parameter(Position = 4)]
    [String[]]$ExcludedList = @(),

    [Parameter(Position = 5)]
    [int]$CriticalThreshold = 10,

    [Parameter(Position = 6)]
    [int]$WarningThreshold = 20,

    [Parameter(Position = 5)]
    [String[]]$Modules = @("VMware.VimAutomation.Core")
)

. ./New-HtmlReport

#region import modules
try {
    foreach($Module in $Modules) {
        Import-Module -Name $Module | Out-Null
    }
} catch {
    Write-Error "[$(Get-Date)][$($MyInvocation.MyCommand)] Failed to import modules: '$Modules'"
    exit 1
}
#endregion import modules

#region connect to viserver
try {
    Connect-VIServer -Server $VIServer -User $User -Password $PW | Out-Null
}catch {
    Write-Error "[$(Get-Date)][$($MyInvocation.MyCommand)] Could not connect to vi server '$VIServer': $_"
    exit 1
}
#endregion connect to viserver

#region execute task
try {

    $Fragments = @()
    $Datacenters = @(Get-View -ViewType Datacenter -Verbose:$false)

    foreach($Datacenter in $Datacenters) {

        $StartTime = Get-Date
        $Details = @()
        $Datastores = Get-View -ViewType Datastore -SearchRoot $Datacenter.MoRef -Verbose:$false
        $Datastores = $Datastores|Where-Object{$_.Name -notin $ExcludedList}

        foreach($Datastore in $Datastores) {

            $VMCount = @(Get-View -ViewType VirtualMachine -SearchRoot $Datastore.MoRef -Property Name -Verbose:$false).length
            $CapacityGB = ([Math]::round($Datastore.Summary.Capacity / 1GB,2))
            $FreeSpaceGB = ([Math]::round($Datastore.Summary.FreeSpace / 1GB,2))
            $UsageGB = ([Math]::round(($Datastore.Summary.Capacity - $Datastore.Summary.FreeSpace) / 1GB, 2))
            $FreeSpacePercent = ([Math]::round($Datastore.Summary.FreeSpace / $Datastore.Summary.Capacity * 100, 0))

            $Detail = [PSCustomObject]@{
                Name = $Datastore.Name
                CapacityGb = $CapacityGB
                UsageGB = $UsageGB
                FreeSpaceGB = $FreeSpaceGB
                FreeSpacePercent = $FreeSpacePercent
                VmCount = $VMCount
            }
            $Details += $Detail

        }

        $EndTime = Get-Date

        $RuntimeInfo = [PSCustomObject]@{
            StartTime = $StartTime
            EndTime = $EndTime
            Duration = $EndTime - $StartTime
        }

        $StatusInfo = [PSCustomObject]@{
            Critical = $Details|Where-Object{$_.FreeSpacePercent -lt $CriticalThreshold}
            Warning = $Details|Where-Object{$_.FreeSpacePercent -lt $WarningThreshold -and $_.FreeSpacePercent -gt $CriticalThreshold}
            Normal = $Details|Where-Object{$_.FreeSpacePercent -gt $WarningThreshold}
        }

        if($Details.length -gt 0) {
            $Report = [PSCustomObject]@{
                Title = "Datastore Usage Report"
                Description = "$($Datacenter.Name)"
                Status = "Success"
                StatusEx = "$($Datastores.length) of $($Datastores.length) datastores processed"
                Date = (Get-Date).toString()
                Overview = @($StatusInfo, $RuntimeInfo)
                Details = $Details|Sort-Object -Property FreeSpaceGB
            }
    
            $Fragments += New-HtmlReport -Json ($Report|ConvertTo-Json) -Fragment
        }
    }

    $Report = "<html><body>" + $Fragments + "</body></html>"
    $Report|Out-File -FilePath "./report.html"

    if($false) {
        # send mail
    }


}catch {
    Write-Warning "error: $_"
    exit 1
}finally {
    if($global:DefaultVIServer) {
        Disconnect-VIServer -Server $VIServer -Force -Confirm:$false -Verbose:$false
    }
}
#endregion execute task