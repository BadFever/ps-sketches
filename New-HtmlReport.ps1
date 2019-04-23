function New-HtmlReport {

    [CmdletBinding()]
    
    param (
        [Parameter(Mandatory, Position = 1)]
        [String]$Json,

        [Parameter(Position = 2)]
        [Switch]$Fragment
    )
    
    Write-Verbose "[$(Get-Date)][$($MyInvocation.MyCommand)] Started execution"

    #region validate json
    $JsonObject = $Json | ConvertFrom-Json -Verbose:$false
    $Properties = @("title", "description", "status", "statusex", "date", "overview", "details")
    foreach ($Property in $Properties) {
        if (@($JsonObject.psobject.Properties.Name) -notcontains $Property) {
            throw "Invalid json format. Property '$Property' missing."
        }
    }
    #endregion validate json

    $Title = $JsonObject.title
    $Description = $JsonObject.description
    $Status = $JsonObject.status
    $StatusEx = $JsonObject.StatusEx
    $Date = Get-date $JsonObject.date
    $Overview = $JsonObject.overview
    $Details = $JsonObject.details

        
    #region header color
    switch ($Status) {
        "Success" {
            $Color = "#00B050"
        }

        "Warning" {
            $Color = "yellow"
        }
        
        "Error" {
            $Color = "red"
        }

        default {
            $Color = "#00B050"
        }
    }
    #endregion header color
    
    # outer table
    $Report = '<div><table cellspacing="0" cellpadding="0" width="100%" style="border-collapse:collapse;border:0"><tbody>' + "`n"
    

    # inner table
    $Report += '<tr><td style="border:none;padding:0px;font-family:Tahoma;font-size:12px"><table cellspacing="0" cellpadding="0" width="100%" style="border-collapse:collapse;border:0"><tbody>' + "`n"
    


    # head
    $Report += '<tr style="height:70px">' + "`n"
    $Report += '<td style="width:80%;border:none;background-color:' + $Color + ';color:White;font-weight:bold;font-size:16px;height:70px;vertical-align:bottom;padding:0 0 17px 15px;font-family:Tahoma">'
    $Report += $Title + '<div style="margin-top:5px; font-size:12px">' + $Description + '</div></td>'
    $Report += '<td style="border:none; padding:0px; font-family:Tahoma; font-size:12px;background-color:' + $Color + ';color:White;font-weight:bold;font-size:16px;height:70px;vertical-align:bottom;padding:0 0 17px 15px;font-family:Tahoma">'
    $Report += $Status + '<div style="margin-top:5px; font-size:12px">' + $StatusEx + '</div></td></tr>'

    # date
    $Report += '<tr><td colspan="2" style="border:none;padding:0px;font-family:Tahoma;font-size:12px"><table width="100%" cellspacing="0" cellpadding="0" style="margin:0px; border-collapse:collapse;border:0"><tbody>' + "`n"
    $Report += '<tr style="height:17px"><td colspan="' + @($Details[0].psobject.properties.name).length + '" style="border-style:solid; border-color:#a7a9ac; border-width:1px 1px 0 1px; height:35px; background-color:#f3f4f4; font-size:16px; vertical-align:middle; padding:5px 0 0 15px; color:#626365; font-family:Tahoma">'
    $Report += '<span>' + $Date + '</span></td></tr>' + "`n"

    # overview
    if ($Overview.length -gt 0) {
        $IsFirstRow = $true
        $ColSpan = @($Details[0].psobject.properties.name).length - (@($Overview[0].psobject.properties.name).length * 2)
        Write-Warning "Colspan: $ColSpan"
        foreach ($Object in $Overview) {
            $IsFirstCol = $true
            $Report += '<tr style="height:17px">'

            foreach ($PropertyName in $Object.psobject.properties.Name) {
                if ($IsFirstCol) {
                    $Report += '<td nowrap="" style="width:1%;padding:2px 3px 2px 3px;vertical-align:top;border:1px solid #a7a9ac;font-family:Tahoma;font-size:12px"><b>' + $PropertyName + '</b></td>'
                    $IsFirstCol = $false
                }
                else {
                    $Report += '<td nowrap="" style="width:85px;padding:2px 3px 2px 3px;vertical-align:top;border:1px solid #a7a9ac;font-family:Tahoma;font-size:12px"><b>' + $PropertyName + '</b></td>'
                }

                $Report += '<td nowrap="" style="width:85px;padding:2px 3px 2px 3px;vertical-align:top;border:1px solid #a7a9ac;font-family:Tahoma;font-size:12px">' + $Object.$PropertyName + '</td>'

            }

            # only first row
            if ($IsFirstRow -and $ColSpan -gt 0) {
                $Report += '<td rowspan="' + $Overview.length + '" colspan="' + $ColSpan + '"'
                $Report += 'style="border:1px solid #a7a9ac;font-family:Tahoma;font-size:12px;vertical-align:top"><span style="font-size:10px">&nbsp;</span></td>'
                $IsFirstRow = $false
            }
            $Report += "</tr>`n"       
        }
    }

    # details
    if ($Details.length -gt 0) {
        $IsFirst = $true
        foreach ($Object in $Details) {
            if ($IsFirst) {
                $Report += '<tr style="height:17px"><td colspan="' + $(@($Object.PSObject.Properties).length) + '" nowrap="" style="height:35px;background-color:#f3f4f4;font-size:16px;vertical-align:middle;padding:5px 0 0 15px;color:#626365;font-family:Tahoma;border:1px solid #a7a9ac">Details</td></tr>'

                $Report += '<tr style="height:23px">'
                $Counter = 1
                foreach ($PropertyName in $Object.psobject.properties.Name) {
    
                    $Report += '<td '
                    if ($Counter -lt @($Object.psobject.properties).length) {
                        $Report += 'nowrap="" '
                        
                    }
                    $Report += 'style="background-color:#e3e3e3; padding:2px 3px 2px 3px;vertical-align:top; border:1px solid #a7a9ac; border-top:none; font-family:Tahoma; font-size:12px">'
                    $Report += '<b>' + $PropertyName + '</b></td>'
    
                    $Counter++
                }

                $Report += '</tr>'
                $IsFirst = $false
            }
    
            $Report += '<tr style="height:17px">'
            $Counter = 1
            foreach ($PropertyName in $Object.psobject.properties.Name) {
                $Report += "<td "
                if ($Counter -lt @($Object.psobject.properties).length) {
                    $Report += 'nowrap="" '
                }
                $Report += 'style="padding:2px 3px 2px 3px;vertical-align:top;border:1px solid #a7a9ac;font-family:Tahoma;font-size:12px">'
                $Report += "$($Object.$PropertyName)" + '</td>'
            }
            $Report += "</tr>"
        }
    }

        $Report += '</tbody></table></td></tr></tbody></table></td></tr></tbody></table></div>'

        if ($Fragment) {
            return $Report
        }
        return '<html><body>' + $Report + '</body></html>'
  
    }