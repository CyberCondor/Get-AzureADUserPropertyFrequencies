<#
.SYNOPSIS
Analyze Azure Active Directory users by frequency of assigned properties. 
Find out how many users are assigned a specific property like "LastDirSyncTime" or "Department".
Find out how many unique property assignments exist for a specific property.
.DESCRIPTION
Analyze Azure Active Directory users by frequency of assigned properties.
.EXAMPLE
PS C:\> Get-AzureADUserPropertyFrequencies.ps1
#>
Write-Host "`n`t`tAttempting to query Azure Active Directory." -BackgroundColor Black -ForegroundColor Yellow
try{Get-AzureADUser -All $true > $null -ErrorAction stop
}
catch{$errMsg = $_.Exception.message
    if($errMsg.Contains("is not recognized as the name of a cmdlet")){
        Write-Warning "`t $_.Exception"
        Write-Output "Ensure 'AzureAD PS Module is installed. 'Install-Module AzureAD'"
        break
    }
    elseif($_.Exception -like "*Connect-AzureAD*"){
        Write-Warning "`t $_.Exception"
        Write-Output "Calling Connect-AzureAD"
        try{Connect-AzureAD -ErrorAction stop
        }
        catch{$errMsg = $_.Exception.message
            Write-Warning "`t $_.Exception"
            break
        }
    }
    else{Write-Warning "`t $_.Exception" ; break}
}

function Get-ExistingUsers_AzureAD{
    try{$ExistingUsers = Get-AzureADUser -All $true -ErrorAction Stop
        return $ExistingUsers
    }
    catch{$errMsg = $_.Exception.message
        Write-Warning "`t $_.Exception"
        return $null
    }
}

function Get-PropertyFrequencies($Property, $Object){
    $Total = ($Object).count
    $ProgressCount = 0
    $AllUniquePropertyValues = $Object | select $Property | sort $Property | unique -AsString # Get All Uniques
    $PropertyFrequencies = @()                                                                # Init empty Object
    $isDate = $false                                                                                                                                                          
    foreach($UniqueValue in $AllUniquePropertyValues){
        if(!($isDate -eq $true)){
            if([string]$UniqueValue.$Property -as [DateTime]){$isDate = $true}
        }
        $PropertyFrequencies += New-Object -TypeName PSobject -Property @{$Property=$($UniqueValue.$Property);Count=0;Frequency="100%"} # Copy Uniques to Object Array and Init Count as 0
    }
    if(($isDate -eq $true) -and (($Object | Select $Property | Get-Member).Definition -like "*datetime*")){
        foreach($PropertyFrequency in $PropertyFrequencies){
            if(($PropertyFrequency.$Property) -and ([string]$PropertyFrequency.$Property -as [DateTime])){
                try{$PropertyFrequency.$Property = $PropertyFrequency.$Property.ToString("yyyy-MM")}
                catch{# Nothing
                }
            }
        }
        foreach($PropertyName in $Object.$Property){                                                            # For each value in Object
            if($Total -gt 0){Write-Progress -id 1 -Activity "Finding $Property Frequencies -> ( $([int]$ProgressCount) / $Total )" -Status "$(($ProgressCount++/$Total).ToString("P")) Complete"}
            foreach($PropertyFrequency in $PropertyFrequencies){                                                # Search through all existing Property values
                if(($PropertyName -eq $null) -and ($PropertyFrequency -eq $null)){$PropertyFrequency.Count++}   # If Property value is NULL, then add to count - still want to track this
                elseif($PropertyName -ceq $PropertyFrequency.$Property){$PropertyFrequency.Count++}             # Else If Property value is current value, then add to count
                else{
                    try{if($PropertyName.ToString("yyyy-MM") -ceq $PropertyFrequency.$Property){$PropertyFrequency.Count++}}
                    catch{# Nothing
                    }
                }
            }
        }
    }
    else{
        foreach($PropertyName in $Object.$Property){                                                            # For each value in Object
            if($Total -gt 0){Write-Progress -id 1 -Activity "Finding $Property Frequencies -> ( $([int]$ProgressCount) / $Total )" -Status "$(($ProgressCount++/$Total).ToString("P")) Complete"}
            foreach($PropertyFrequency in $PropertyFrequencies){                                                # Search through all existing Property values
                if(($PropertyName -eq $null) -and ($PropertyFrequency -eq $null)){$PropertyFrequency.Count++}   # If Property value is NULL, then add to count - still want to track this
                elseif($PropertyName -ceq $PropertyFrequency.$Property){$PropertyFrequency.Count++}             # Else If Property value is current value, then add to count
            }
        }
    }
    Write-Progress -id 1 -Completed -Activity "Complete"
    if($Total -gt 0){
        foreach($PropertyFrequency in $PropertyFrequencies){$PropertyFrequency.Frequency = ($PropertyFrequency.Count/$Total).ToString("P")}
    }
    return $PropertyFrequencies | select Count,$Property,Frequency | sort Count,$Property | Unique -AsString
}
function DisplayFrequencies($Property, $PropertyFrequencies){
    write-output "`n"
    $PropertyFrequencies | select Count,$Property,Frequency | sort Count,$Property,Frequency | unique -AsString | ft
    write-output "Total Number of Unique $($Property)(s): $(($PropertyFrequencies).count)"
}

function main{
    $quitProgram = $false
    While($quitProgram -eq $false){
        write-Host "`n Azure Active Directory USER Properties available to query:"
        foreach($Property in $AzureADUserProperties){Write-Host "`t$Property"}
        $Property = Read-Host "`nEnter one of the properties listed above or 'q' to quit"
        if($Property -eq "q"){$quitProgram = $true}
        else{
            $SmallerListOfProperties = @()
            $found = $false
            $index = 0
            foreach($P in $AzureADUserProperties){
                if($P -like "*$Property*"){
                    if($P -eq $Property){$found = $true}
                    else{$SmallerListOfProperties += New-Object -TypeName PSobject -Property @{Property=$P;Index=$index++}}
                }
            }
            if(($found -eq $false) -and ($SmallerListOfProperties -ne $null)){
                $SmallerListOfProperties | ft
                $Property = Read-Host "`nEnter one of the properties or index numbers listed above or 'q' to quit"
                if($Property -eq "q"){$quitProgram = $true}
                else{
                    foreach($Q in $SmallerListOfProperties){    
                        if(($Property -eq $Q.Index) -or ($Property -eq $Q.Property)){$Property = $Q.Property; $found = $true}
                    }
                }
            }
            if($found -eq $true){
                $Frequencies = Get-PropertyFrequencies $Property $ExistingUsers_AzureAD
                DisplayFrequencies $Property $Frequencies 
                Read-Host "`nPress Enter for Main Menu"
                clear
            }
            else{Write-Output "`nProperty '$Property' is not found in the list of properties available to query. `n" ; sleep 3.33}
        }
    }
}

$AzureADUserProperties = Get-AzureADUser -All $true | Select -First 1 | Get-Member | where{($_.MemberType -eq "Property") -and  ($_.Definition -notlike "*list*")} | select -ExpandProperty Name

$ExistingUsers_AzureAD = Get-ExistingUsers_AzureAD
if($ExistingUsers_AzureAD -eq $null){break}

Write-Host "This program will display various property frequencies from 'Azure AD Users'`n"

main
