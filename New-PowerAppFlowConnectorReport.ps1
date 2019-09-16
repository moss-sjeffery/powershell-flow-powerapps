param(
    # Note: Please ensure that you provide the file extension .csv, example: c:\reports\powerappsreport.csv
    [string]$CsvReport
)

#region Begin
    Install-Module -Name Microsoft.PowerApps.Administration.PowerShell
    Install-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber
    If (Test-Path $CsvReport)
    {
        Remove-Item $CsvReport
    }
    "Type" + "," + "DisplayName" + "," + "OwnerEmail" | Out-File $CsvReport -Encoding ascii
    Add-PowerAppsAccount
    $flows = Get-AdminFlow
    $powerApps = Get-AdminPowerApp
#endregion

#region Process Flows
    foreach ($flow in $flows)
    {
        $flowDetails = $flow | Get-AdminFlow
        $flowName = $flow.DisplayName -replace '[,]'
        $flowOwnerObj = Get-UsersOrGroupsFromGraph -ObjectId $flow.createdBy.objectId
        $ownerUPN = $flowOwnerObj.UserPrincipalName
        foreach ($connector in $flowDetails.Internal.properties.connectionReferences)
        {
            foreach ($connection in ($connector | gm -MemberType NoteProperty).Name)
            {
                if ($connection -eq "shared_Office365")
                {
                    "Flow" + "," + $flowName + "," + $ownerUPN | Out-File $CsvReport -Encoding ascii -Append
                }
            }
        }
    }
#endregion

#region Process PowerApps
    foreach ($app in $powerApps)
    {
        $appName = $app.DisplayName
        $ownerObject = $app.owner.id
        $appOwnerObj = Get-UsersOrGroupsFromGraph -ObjectId $ownerObject
        $ownerUPN = $appOwnerObj.UserPrincipalName
        foreach ($connector in $app.Internal.properties.connectionReferences)
        {
            foreach ($connection in ($connector | gm -MemberType NoteProperty).Name)
            {
                $connectionProperties = $($connector.$connection)
                write-host $connectionProperties.DisplayName
                if ($connectionProperties.DisplayName -eq "Office 365 Outlook")
                {
                    "PowerApp" + "," + $appName + "," + $ownerUPN | Out-File $CsvReport -Encoding ascii -Append
                }
            }
        }
    }
#endregion
