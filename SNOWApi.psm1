##Get Information about Single CI
function get-SNOWWorkstation{
    [CmdletBinding(DefaultParameterSetName='Name')]
    Param(
	    [parameter(Mandatory=$true, ParameterSetName='Name')]
        [String]$ComputerName,
        [parameter(Mandatory=$true, ParameterSetName='SysID')]
        [String]$SysID,
        [ValidateSet('true','false','all')]
        [String]$Displayvalue ='true',
        [String]$proxyuri = $null
    )
    If(!$global:SNOWCredentials){
        $global:SNOWCredentials = Get-Credential
    }
    $uri = 'https://abraxas.service-now.com/api/now/table/u_cmdb_ci_workstation'
    $Body = @{
        'sysparm_query' = ''
        'sysparm_exclude_reference_link'='true'
    }
    
    switch($PSCmdlet.ParameterSetName){
        "Name"{$body.sysparm_query = 'name='+$ComputerName}
        "SysID"{$body.sysparm_query = 'sys_id='+$SysID}
    }


    $body.sysparm_display_value = $Displayvalue

    $result = Invoke-RestMethod -Uri $uri -Credential $global:SNOWCredentials -Body $Body -ContentType "application/json; charset=utf-8" -proxy $(if(([string]::IsNullOrEmpty($proxyuri))){$null }else{$proxyuri})
    return $result.result
}
##Get Inventory for Company specify name
function get-SNOWInventoyByCompany{
    Param(
	    [parameter(Mandatory=$true)]
	    [String]$CompanyID,
        [String]$CompanyID2,
        [Bool]$displayvalue,
        [String]$proxyuri = $null
    )

    If(!$global:SNOWCredentials){
        $global:SNOWCredentials = Get-Credential
    }
    
    $uri = 'https://abraxas.service-now.com/api/now/table/u_cmdb_ci_workstation'
    #construct filter
    $filter = 'companyLIKE'+$CompanyID
    if($CompanyID2){
        $filter = $filter+'^ORcompanyLIKE'+$CompanyID2
    }
    $filter = $filter+'^install_status!=7'

    $Body = @{ 
    'sysparm_query' = $filter
    'sysparm_limit'='1000'
    'sysparm_exclude_reference_link'='true'
    'sysparm_fields'='asset_tag,name,sys_id,u_type,model_id,company,assigned_to,install_status'
    } 

    if($displayvalue){
        $body.sysparm_display_value = 'true'
    }


    $result = (Invoke-RestMethod -Uri $uri -Credential $global:SNOWCredentials -Body $Body -ContentType "application/json; charset=utf-8"  -proxy $(if(([string]::IsNullOrEmpty($proxyuri))){$null }else{$proxyuri})).Result 
    return $result
}
##Get uncompleted Snow Tasks by short description
function get-SNOWTasks{
    [CmdletBinding(DefaultParameterSetName='Short_Description')]
    Param(
	    [parameter(Mandatory=$true, ParameterSetName='Short_Description')]
        [String]$Short_Description,
        [parameter(Mandatory=$true, ParameterSetName='Criteria')]
        [String]$Criteria,
        [ValidateSet('true','false','all')]
        [String]$Displayvalue ='true',
        [String]$proxyuri = $null
    )

    If(!$global:SNOWCredentials){
        $global:SNOWCredentials = Get-Credential
    }

    $uri = 'https://abraxas.service-now.com/api/now/table/sc_task'
    $Body = @{ 
    'sysparm_query'=''
    'sysparm_exclude_reference_link'='true'
    'sysparm_display_value'=''
    } 

    if($criteria){
        $Body.sysparm_query=$criteria
    }else{
        $Body.sysparm_query='state=1^short_descriptionCONTAINS'+$Short_Description
    }
    $Body.sysparm_display_value = $Displayvalue

    $result = Invoke-RestMethod -Uri $uri -Credential $SNOWCredentials -Body $Body -ContentType "application/json" -Proxy $(if(([string]::IsNullOrEmpty($proxyuri))){$null }else{$proxyuri})
    $tasks = $result.result
    return $tasks
}
##Resturns Snow Task with all properties as HashTabel | Requires TaskNumber
function get-SNOWTaskWithVar{
    Param(
	    [parameter(Mandatory=$true)]
	    [String]$TaskNumber,
        [String]$proxyuri = $null
    )

    If(!$global:SNOWCredentials){
        $global:SNOWCredentials = Get-Credential
    }

    $uri = 'https://abraxas.service-now.com/api/now/table/sc_task'
    $Body = @{ 
    'sysparm_query'='number='+$TaskNumber     
    'sysparm_limit'='1'
    'sysparm_exclude_reference_link'='true'
    'sysparm_display_value'='all'

    } 
    $result = Invoke-RestMethod -Uri $uri -Credential $SNOWCredentials -Body $Body -ContentType "application/json" -Proxy $(if(([string]::IsNullOrEmpty($proxyuri))){$null }else{$proxyuri})
    $Task = $result.result[0]

    $uri = 'https://abraxas.service-now.com/api/now/table/sc_item_option_mtom'
    $Body = @{ 
    'sysparm_query' = 'request_item='+$task.request_item.value
    'sysparm_exclude_reference_link'='true'
    'sysparm_display_value'='true'
    } 
    $result = Invoke-RestMethod -Uri $uri -Credential $SNOWCredentials -Body $Body -ContentType "application/json" -Proxy $(if(([string]::IsNullOrEmpty($proxyuri))){$null }else{$proxyuri})
    $ritmoptions = $result.result

    $uri = 'https://abraxas.service-now.com/api/now/table/sc_item_option'
    $task | Add-Member -MemberType NoteProperty -Name "OptionVariables" -Value @{} -Force
    foreach($option in $ritmoptions){
        $Body = @{ 
            'sysparm_query' = 'sys_id='+$option.sc_item_option 
            #'sysparm_limit'='10'
            'sysparm_exclude_reference_link'='true'
            'sysparm_display_value'='true'
        } 
        $result = Invoke-RestMethod -Uri $uri -Credential $SNOWCredentials -Body $Body -ContentType "application/json" -Proxy $(if(([string]::IsNullOrEmpty($proxyuri))){$null }else{$proxyuri})
        $task.OptionVariables["$($result.result[0].item_option_new)"] = $result.result[0].value
    }
    if($Task.OptionVariables.'Auswahl Person'){
        $uri = 'https://abraxas.service-now.com/api/now/table/sys_user'
        $Body = @{ 
        'sysparm_query'='sys_id='+$Task.OptionVariables.'Auswahl Person'
        'sysparm_limit'='1'
        } 
        $result = Invoke-RestMethod -Uri $uri -Credential $SNOWCredentials -Body $body -ContentType "application/json" -Proxy $(if(([string]::IsNullOrEmpty($proxyuri))){$null }else{$proxyuri})
        if($result.result[0].source -like "ldap*"){
            $Task.OptionVariables["OU"] = ($result.result[0].source).TrimStart("ldap:")
        }
    }
    return $Task
   }
##Complete Task
function set-SNOWTaskComplete{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(
	    [parameter(Mandatory=$true)]
	    $SNOWTask,
        [String]$Comment = "Closed by SNOW Worker Script",
        [String]$proxyuri = $null
    )
    $uri = 'https://abraxas.service-now.com/api/now/table/sc_task/'+$SNOWTask.sys_id.value
    $Body = @{ 
    'sysparm_exclude_reference_link'='true'
    'sysparm_display_value'='true'
    'sysparm_input_display_value'='true'
    'state'='3'
    'close_notes'=$Comment
    } 
    $Body = $Body | ConvertTo-Json 
    if ($PSCmdlet.ShouldProcess($($body+" "+$uri),"SNOW Update")) {
        $enc = [System.Text.Encoding]::UTF8
        $body= $enc.GetBytes($Body)
        try{
            $result = Invoke-RestMethod -Uri $uri -Credential $global:SNOWCredentials -Body $Body -ContentType "application/json; charset=utf-8" -Method Put -proxy $(if(([string]::IsNullOrEmpty($proxyuri))){$null }else{$proxyuri}) -ErrorAction stop
            Write-EventLog -LogName "SNOW Interaction Worker" -Source "User Expiration Script" -EventId 2 -EntryType Information -Message "Successfully Closed Task $($SNOWTASK.number.display_value) associated with User $($SNOWTask.OptionVariables.PersonName) Expiration date set to: $($SNOWTask.OptionVariables.Austrittsdatum)"

       }catch{
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-EventLog -LogName "SNOW Interaction Worker" -Source "User Expiration Script" -EventId 2 -EntryType Error -Message "Failed to Close Task $($SNOWTASK.number.display_value) associated with User $user"
            return "failed"
       }
       return $result.result
    }
}
##Set Various Properties on SNOW Computer CI must specify SysID
function set-SNOWWorkstation{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(
	    [parameter(Mandatory=$true)]
	    [String]$SysID,
        [ValidateSet('Installed','ToClarify','InStock','Retired')]
        [String]$Status,
        [String]$Company,
        [String]$Location,
        [hashtable]$BodyAdd,
        [String]$proxyuri = $null
    )

    If(!$global:SNOWCredentials){
    $global:SNOWCredentials = Get-Credential
    }

    $uri = 'https://abraxas.service-now.com/api/now/table/u_cmdb_ci_workstation/'+$SysID


     $Body = @{ 
    'sysparm_exclude_reference_link'='true'
    'sysparm_display_value'='true'
    'sysparm_input_display_value'='true'
    
    } 

    if($bodyadd){
        $Body += $bodyadd

    }

    if($status){
        $body | Add-Member -MemberType NoteProperty -Name "install_status" -Value ""
        Switch ($status){
            "Installed"{$Body.install_status = 1}
            "ToClarify"{$Body.install_status = 15}
            "InStock"{$Body.install_status = 6}
            "Retired"{$Body.install_status = 7}
            default{return "error"}  
        }
    }
    if($company -and $Company -ne ""){
        
        $body | Add-Member -MemberType NoteProperty -Name "company" -Value ""
        $body.company = $Company.ToString()

    }

    if($location -and $Location -ne ""){
        
        $body | Add-Member -MemberType NoteProperty -Name "location" -Value ""
        $body.location = $Location.ToString()

    }
    

    $Body = $Body | ConvertTo-Json 

    
    if ($PSCmdlet.ShouldProcess($body,"SNOW Update")) {
       
        $enc = [System.Text.Encoding]::UTF8
        $body= $enc.GetBytes($Body)
        $result = Invoke-RestMethod -Uri $uri -Credential $global:SNOWCredentials -Body $Body -ContentType "application/json; charset=utf-8" -Method Put -proxy $(if(([string]::IsNullOrEmpty($proxyuri))){$null }else{$proxyuri})
       
        return $result.result
    }
}
##Get SNOW User
function get-SNOWUser{
    [CmdletBinding(DefaultParameterSetName='Name')]
    Param(
	    [parameter(Mandatory=$true, ParameterSetName='Name')]
        [String]$UserName,
        [parameter(Mandatory=$true, ParameterSetName='SysID')]
        [String]$SysID,
        [parameter(Mandatory=$true, ParameterSetName='Email')]
        [String]$Email,
        [ValidateSet('true','false','all')]
        [String]$Displayvalue ='true',
        [String]$proxyuri = $null
    )
    If(!$global:SNOWCredentials){
        $global:SNOWCredentials = Get-Credential
    }
    $uri = 'https://abraxas.service-now.com/api/now/table/sys_user'
    $Body = @{
        'sysparm_query' = ''
        'sysparm_exclude_reference_link'='true'
    }
    
    switch($PSCmdlet.ParameterSetName){
        "Name"{$body.sysparm_query = 'user_name='+$UserName}
        "SysID"{$body.sysparm_query = 'sys_id='+$SysID}
        "Email"{$body.sysparm_query = 'email='+$Email}
    }

    $body.sysparm_display_value = $Displayvalue

    $result = Invoke-RestMethod -Uri $uri -Credential $global:SNOWCredentials -Body $Body -ContentType "application/json; charset=utf-8" -proxy $(if(([string]::IsNullOrEmpty($proxyuri))){$null }else{$proxyuri})
    return $result.result
}

function get-SNOWTableEntry{
    Param(
	    [parameter(Mandatory=$true)]
	    [String]$Table,
        [parameter(Mandatory=$true)]
	    [String]$sys_id,
        [ValidateSet('true','false','all')]
        [String]$Displayvalue ='true',
        [String]$proxyuri = $null
    )

    If(!$global:SNOWCredentials){
        $global:SNOWCredentials = Get-Credential
    }

    $Body = @{
        'sysparm_query' = ''
        'sysparm_exclude_reference_link'='true'
        'sysparm_limit'='1'
    }

    $body.sysparm_display_value = $Displayvalue


    $uri = 'https://abraxas.service-now.com/api/now/table/'+$table
    $Body.sysparm_query = 'sys_id='+$sys_id
    $result = Invoke-RestMethod -Uri $uri -Credential $SNOWCredentials -Body $body -ContentType "application/json" -Proxy $(if(([string]::IsNullOrEmpty($proxyuri))){$null }else{$proxyuri})
    return $result.result
}

function set-SNOWTableEntry{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(
	    [parameter(Mandatory=$true)]
	    [String]$SysID,
        [parameter(Mandatory=$true)]
        [String]$Table,
        [parameter(Mandatory=$true)]
        [hashtable]$Bodyadd,
        [String]$proxyuri = $null
    )

    If(!$global:SNOWCredentials){
    $global:SNOWCredentials = Get-Credential
    }

    $uri = 'https://abraxas.service-now.com/api/now/table/'+$table+'/'+$SysID
    
    $Body = @{ 
    'sysparm_exclude_reference_link'='true'
    'sysparm_display_value'='true'
    'sysparm_input_display_value'='true'
    }

    $Body += $bodyadd    

    $Body = $Body | ConvertTo-Json 

    if ($PSCmdlet.ShouldProcess($body,"SNOW Update")) {
       
        $enc = [System.Text.Encoding]::UTF8
        $body= $enc.GetBytes($Body)
        $result = Invoke-RestMethod -Uri $uri -Credential $global:SNOWCredentials -Body $Body -ContentType "application/json; charset=utf-8" -Method Put -proxy $(if(([string]::IsNullOrEmpty($proxyuri))){$null }else{$proxyuri})
       
        return $result.result
    }
}
