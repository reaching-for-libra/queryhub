###########################################################################################################
# Global
###########################################################################################################


$script:timeoutMinutes = 0
$script:OracleScriptDirectory = ""
$script:TextEditorCommand = $null

$script:OracleTnsNamesCommand = $null

$script:oracleConnections = [System.Collections.ArrayList]::new()
$script:oracleDefaultConnectionName = $null
$script:lastOracleQueryRequest = ""
$script:lastOracleQueryResult = $null
$script:lastOracleQueryExplain = $null
$script:lastOracleQueryTime = $null
$script:lastOracleDbmsOutput = $null
$script:lastOracleQueryError = $null

$script:OracleTnsNames = @()



$script:sqlserverConnections = [System.Collections.ArrayList]::new()
$script:sqlServerDefaultConnectionName = $null
$script:lastSqlServerQueryRequest = ""
$script:lastSqlServerQueryResult = $null
#$script:lastSqlServerQueryExplain = $null
$script:lastSqlServerQueryTime = $null
$script:lastSqlServerPrintStatements = $null
$script:lastSqlServerQueryError = $null


#Add-Type -AssemblyName System.Windows.Forms
#Add-Type -AssemblyName System.Drawing
add-type -path "$(split-path -parent $MyInvocation.MyCommand.Definition)\Oracle.ManagedDataAccess.dll"


if ((get-typedata -typename QueryHub.Config)){
    remove-typedata -typename QueryHub.Config -confirm:$false
}

update-typedata -typename QueryHub.Config -membertype NoteProperty -membername TimeoutMinutes -value ([int]$null) -confirm:$false
update-typedata -typename QueryHub.Config -membertype NoteProperty -membername OracleScriptDirectory -value ([string]$null) -confirm:$false
update-typedata -typename QueryHub.Config -membertype NoteProperty -membername SqlServerScriptDirectory -value ([string]$null) -confirm:$false
update-typedata -typename QueryHub.Config -membertype NoteProperty -membername TextEditorCommand -value ([string]$null) -confirm:$false
update-typedata -typename QueryHub.Config -membertype NoteProperty -membername OracleTnsNamesCommand -value ([string]$null) -confirm:$false

function ShowAsyncWaitMessage{

    param(
        [parameter(position=0,mandatory=$false,ValueFromPipeline=$true)][string]$Message = "Waiting",
        [parameter(position=1,mandatory=$true,ValueFromPipeline=$false)][string]$ConditionVariablename,
        [parameter(position=2,mandatory=$false,ValueFromPipeline=$false)][string[]]$ConditionMethodPropertyChain = @(),
        [parameter(position=3,mandatory=$false,ValueFromPipeline=$false)][object]$ConditionMatch = $true
    )

    $start = get-date
    $check = $start


    if ($host.name -eq 'consolehost'){
        $setControlC = $true
        $controlCSave = [console]::treatcontrolcasinput
    }else{
        $setControlC = $false
    }

    if ($setcontrolc){
        [console]::TreatControlCAsInput = $true
    }

    $completed = $true
    $progressbarcounter = 0

    $lasttimespancheck = (new-timespan $check (get-date))

    if ($host.name -eq 'consolehost'){

        $messageArray = ($message -split "`r`n") -split "`n"
        $messageOutput = ""
        
        for ($x=0;$x -lt $messagearray.count -and $x -lt 12;$x++){
            if ($x -eq 11){
                $messageoutput += "..."
                break
            }

            if ($messagearray[$x].trimend().length -gt ($host.ui.rawui.buffersize.width - 4)){
                $messageoutput += $messagearray[$x].trimend().substring(0, ($host.ui.rawui.buffersize.width - 4 - 3)) + "..."
            }else{
                $messageoutput += ($messagearray[$x].trimend() + (" " * ($host.ui.rawui.buffersize.width - $messagearray[$x].trimend().length - 4)))
            }
        }
    }else{
        $messageoutput = $message
    }

    while ($true){

        if ($setcontrolc -and [console]::KeyAvailable) {
            $key = [system.console]::readkey($true)
            if (($key.modifiers -band [consolemodifiers]"control") -and ($key.key -eq "C")) {
                $completed = $false
                break
            }
        }

        #get the scope of the variable based on whether it was called from this module or not
        if ((get-pscallstack)[1].scriptname -eq (get-pscallstack)[0].scriptname){
            $variablescope = 1
        }else{
            $variablescope = 2
        }
        
        #get the variable from the appropriate scope
        $checkVar = (get-variable -name $conditionVariablename -scope $variablescope).value

        #build the variable's property and method naming
        foreach ($methodproperty in $conditionmethodpropertychain){
            if ($methodproperty.trim() -match '^(?<name>[^(]{1,})(?<method>\(\)$){0,1}'){
                if ($matches.method){
                    $checkvar = $checkvar."$methodproperty"()
                }else{
                    $checkvar = $checkvar."$methodproperty"
                }
            }else{
                throw 'invalid condition method/property'
            }
        }
        
        #if the condition is met, exit 
        if ($checkvar -eq $conditionMatch){
            break
        }

        $timespan = (new-timespan $check (get-date))

        if ($timespan.totalmilliseconds - $lastTimeSpanCheck.totalmilliseconds -lt 5){
            continue
        }

        $lasttimespancheck = $timespan

        if ($progressbarcounter -ge 100){
            $directionup = $false
        }elseif($progressbarcounter -le 0){
            $directionup = $true
        }

        if ($directionup){
            $progressbarcounter++
        }else{
            $progressbarcounter--
        }

        $timespanstring = "$($timespan.seconds) seconds"

        if ($timespan.minutes -gt 0){
            $timespanstring = "$($timespan.minutes) minutes, " + $timespanstring
        }
        if ($timespan.hours -gt 0){
            $timespanstring = "$($timespan.hours) hours, " + $timespanstring
        }
        if ($timespan.days -gt 0){
            $timespanstring = "$($timespan.days) days, " + $timespanstring
        }


        write-progress -activity " " -status $messageoutput -PercentComplete $progressbarcounter -currentoperation "$timespanstring elapsed" 
    }
    
    write-progress -Activity " " -status $message -completed

    if ($setcontrolc){
        [console]::TreatControlCAsInput = $controlcsave
    }

#    $pos = [System.Management.Automation.Host.Coordinates]::new(0,$host.ui.rawui.cursorposition.y)
#    $newBuffer = $host.ui.rawui.newbuffercellarray((" " * $host.ui.rawui.buffersize.width),$host.ui.rawui.foregroundcolor,$host.ui.rawui.backgroundcolor)
#
#    $host.ui.rawui.SetBufferContents($pos,$newbuffer)

    $completed
    return 
}


#adapted from script: http://www.indented.co.uk/2015/06/03/dynamic-parameters/
function NewDynamicParameter {

  [CmdLetBinding()]
  param(
    [Parameter(ValueFromPipeline=$true,HelpMessage='Dictionary to add created dynamic parameter')] [pscustomobject] $DynamicParameters = $null,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$Name,
    [Object]$DefaultValue, 
    [Type]$Type = "Object",
    [Switch]$Mandatory,
    [Int32]$Position = -2147483648,
    [Switch]$ValueFromPipeline,
    [Switch]$ValueFromPipelineByPropertyName,
    [String[]]$ParameterSetNames = @("__AllParameterSets"),
    [Switch]$ValidateNotNullOrEmpty,
    [RegEx]$ValidatePattern,
    [Text.RegularExpressions.RegexOptions]$ValidatePatternOptions = [Text.RegularExpressions.RegexOptions]::IgnoreCase,
    [Object[]]$ValidateRange,
    [ScriptBlock]$ValidateScript,
    [Object[]]$ValidateSet,
    [Boolean]$ValidateSetIgnoreCase = $true
  )
  
  $AttributeCollection = New-Object 'Collections.ObjectModel.Collection[System.Attribute]'

foreach ($parametersetname in $parametersetnames){
  $ParameterAttribute = New-Object Management.Automation.ParameterAttribute
  $ParameterAttribute.Mandatory = $Mandatory
  $ParameterAttribute.Position = $Position
  $ParameterAttribute.ValueFromPipeline = $ValueFromPipeline
  $ParameterAttribute.ValueFromPipelineByPropertyName = $ValueFromPipelineByPropertyName
  $ParameterAttribute.parametersetname = $parametersetname

  $AttributeCollection.add($ParameterAttribute)
}

  if ($psboundparameters.ContainsKey('ValidateNotNullOrEmpty')) {
    $AttributeCollection.add((New-Object Management.Automation.ValidateNotNullOrEmptyAttribute))
  }
  if ($psboundparameters.ContainsKey('ValidatePattern') -and $validatepattern) {
    $ValidatePatternAttribute = New-Object Management.Automation.ValidatePatternAttribute($ValidatePattern.ToString())
    $ValidatePatternAttribute.Options = $ValidatePatternOptions

    $AttributeCollection.add($ValidatePatternAttribute)
  }
  if ($psboundparameters.ContainsKey('ValidateRange') -and $validrange) {
    $AttributeCollection.add((New-Object Management.Automation.ValidateRangeAttribute($ValidateRange)))
  }
  if ($psboundparameters.ContainsKey('ValidateScript') -and $validatescript) {
    $AttributeCollection.add((New-Object Management.Automation.ValidateScriptAttribute($ValidateScript)))
  }
  if ($psboundparameters.ContainsKey('ValidateSet') -and $validateset) {
    $ValidateSetAttribute = New-Object Management.Automation.ValidateSetAttribute($ValidateSet)
    $ValidateSetAttribute.IgnoreCase = $ValidateSetIgnoreCase

    $AttributeCollection.add($ValidateSetAttribute)
  }

  $Parameter = New-Object Management.Automation.RuntimeDefinedParameter($Name, $Type, $AttributeCollection)

    if (-not $DynamicParameters){
        
        $DynamicParameters = new-object PSObject
        $DynamicParameters | add-member -membertype NoteProperty -name "ParameterDictionary" -value ([System.Management.Automation.RuntimeDefinedParameterDictionary]::new())
        $DynamicParameters | add-member -membertype NoteProperty -name "DefaultValues" -value ([hashtable]::new())
    }

    $DynamicParameters.parameterdictionary.Add($Name, $Parameter)

  if ($psboundparameters.ContainsKey('DefaultValue')) {
    $DynamicParameters.defaultvalues.add($Name,$defaultvalue)
  }else{
    if ($type.name -eq 'SwitchParameter'){
      $DynamicParameters.defaultvalues.add($Name,$false)
    }else{
      $DynamicParameters.defaultvalues.add($Name,$null)
    }
  }

    $DynamicParameters
}

function SetDynamicParameterValues{
    param(
        [parameter(mandatory=$true)]$params,
        [parameter(mandatory=$true)]$defaultparams
    )

        foreach ($parameter in $defaultparams.defaultvalues.getenumerator()){
            if ($params[$parameter.key] -eq $null){
                $params[$parameter.key] = $parameter.value
            }
        }
        $params
}

function SetDynamicParameterVariables{
    param(
        [parameter(mandatory=$true)]$params
    )

    if ((get-pscallstack)[1].scriptname -eq (get-pscallstack)[0].scriptname){
        $variablescope = 1
    }else{
        $variablescope = 2
    }

    foreach ($parameter in $params.getenumerator()){
        set-variable -name $parameter.key -value $parameter.value -scope $variablescope -confirm:$false #scope 1 means immediate parent
    }
}


function getconsolecredential{
    [OutputType('System.Management.Automation.PSCredential')]
    param(
        [parameter(mandatory=$false)] $forreal = $true #this is a hack for allowing parameters to default a value to the return of this call, but preventing it when it's not needed - see invoke-sqlserverquery()
    )

    if (-not $forreal){
        return
    }

#    $login = New-Object System.Management.Automation.Host.FieldDescription "Login"
#    $login.Label = $null
#
#    $password = New-Object System.Management.Automation.Host.FieldDescription "Password"
#    $password.setparametertype([System.Security.SecureString])
#    $password.Label = $null
#
#    $fields = [System.Management.Automation.Host.FieldDescription[]]($login,$password)
#    $result = $Host.UI.Prompt($null,$null, $fields)
#    New-Object System.Management.Automation.PSCredential ($result["Login"],$result["Password"])

    $login = read-host Login
    $pass = read-host Password -assecurestring

    New-Object System.Management.Automation.PSCredential ($login,$pass)
}

function replacevariable{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param
    (
        [parameter(position=0,mandatory=$true,ValueFromPipeline=$false)] $variable,
        [parameter(position=1,mandatory=$false,ValueFromPipeline=$false)] $defaultValue = $null
    )

    $variable = "$variable[$defaultValue]"

    $returnObject = New-Object System.Management.Automation.PSObject

    $field = New-Object System.Management.Automation.Host.FieldDescription $variable
    if ($variable -match 'password'){
        $field.setparametertype([System.Security.SecureString])
    }
    $field.Label = $null
    $field.DefaultValue = $defaultValue
    $fields = [System.Management.Automation.Host.FieldDescription[]]($field)
    $result = $Host.UI.Prompt($null,$null, $fields)


    if ($result[$variable] -eq ""){

        Add-Member -InputObject $returnObject -MemberType NoteProperty -Name "IsCancelled" -Value $true
        Add-Member -InputObject $returnObject -MemberType NoteProperty -Name "Value" -Value $null
    }elseif ($result[$variable] -eq ""){

        Add-Member -InputObject $returnObject -MemberType NoteProperty -Name "IsCancelled" -Value $false
        Add-Member -InputObject $returnObject -MemberType NoteProperty -Name "Value" -Value ''

    }else{
        if ($result[$variable] -is [System.Security.SecureString]){
            $resultvalue = (New-Object System.Management.Automation.PSCredential 'N/A', $result[$variable]).GetNetworkCredential().Password
        }else{
            $resultValue = $result[$variable]
        }

        if ($resultValue -eq ""){
            $resultValue = $defaultValue
        }

        Add-Member -InputObject $returnObject -MemberType NoteProperty -Name "IsCancelled" -Value $false
        Add-Member -InputObject $returnObject -MemberType NoteProperty -Name "Value" -Value $resultValue
    }

    $returnobject

}

function Set-QueryHubParams{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param
    (
        [parameter(position=0,mandatory=$true,ValueFromPipeline=$true)] [validatescript({$_.psobject.typenames -contains 'QueryHub.Config'})] $Parameters
    )

    $script:timeoutMinutes = $parameters.timeoutMinutes
    $script:OracleScriptDirectory = $parameters.OraclescriptDirectory
    $script:sqlServerScriptDirectory = $parameters.sqlserverscriptDirectory
    $script:TextEditorCommand = $parameters.TextEditorCommand
    $script:OracleTnsNamesCommand = $parameters.OracleTnsNamesCommand

    $script:oracleTnsNames = $null

    if ($script:OracleTnsNamesCommand){
        $script:oracleTnsNames = [scriptblock]::create($script:OracleTnsNamesCommand).invoke()
    }
}


function Get-QueryHubParams{
    [outputtype('QueryHub.Config')]

    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    $result = [pscustomobject][ordered] @{
        TimeoutMinutes = $script:timeoutMinutes
        OracleScriptDirectory = $script:Oraclescriptdirectory
        SqlServerScriptDirectory = $script:SqlServerScriptDirectory
        TextEditorCommand = $script:TextEditorCommand
        OracleTnsNamesCommand = $script:OracleTnsNamesCommand
    }

    $result.psobject.typenames.insert(0,'QueryHub.Config')
    $result
}

###########################################################################################################
# Oracle
###########################################################################################################



function Get-LastOracleQueryRequest{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    $script:lastOracleQueryRequest
}

function Get-LastOracleQueryResult{
    [outputtype('QueryHub.OLast')]
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    if ($script:lastOracleQueryResult){
        $script:lastOracleQueryResult.psobject.typenames.insert(0,'QueryHub.OLast')
    }
    $script:lastOracleQueryResult
}

function Get-LastOracleQueryExplain{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    $script:lastOracleQueryExplain
}
    
function Get-LastOracleQueryTime{
    [outputtype('Timespan')]
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    $script:lastOracleQueryTime
}

function Get-LastOracleDbmsOutput{
    [outputtype('string[]')]
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    $script:lastOracledbmsoutput -split "`n"
}
function Get-LastOracleQueryError{
    [outputtype('String')]
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    $script:lastOracleQueryError
}

function Get-OracleConnectionSession{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -position 0 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections)  -default (get-oracledefaultqueryhubconnection)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
        setdynamicparametervariables $psboundparameters

        $script:oracleconnections | where {$_.name -eq $connectionname}
    }

    process{
        setdynamicparametervariables $psboundparameters
    }

    end{
    }
}

function Get-OracleConnectionNames{
    param(
    )

    $script:Oracletnsnames | select -expand tnsname

}

function Add-OracleQueryHubConnection{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -mandatory -position 0 -Type string -Name ServiceName -validateset (get-oracleconnectionnames) |
            newdynamicparameter -position 1 -Name Credentials -validatescript { $_ -is [pscredential] -and $_ -ne $null} |
            newdynamicparameter -position 2 -Type string -Name ConnectionName -defaultvalue "" |
            newdynamicparameter -position 3 -Type boolean -Name UseTns -defaultvalue $true |
            newdynamicparameter -position 4 -Type string -Name HostName -defaultvalue "" |
            newdynamicparameter -position 5 -Type int -Name Port -defaultvalue 0
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        if (-not $credentials){
            $credentials = getconsolecredential
        }
        
        if ($connectionname -eq ""){
            $connectionname = $servicename
        }

        if ($script:oracleconnections.name -contains $connectionname){
            throw 'Connection already exists'
        }

        $connection = [pscustomobject] @{
            Name = $connectionname
            Connection = [Oracle.ManagedDataAccess.Client.OracleConnection]::new()
            Sid = $null
            OraclePipeline = $null
        }

        if ($UseTns){
            $tns = $script:oracletnsnames | where {$_.tnsname -eq $servicename}
            $hostname = $tns.host
            $port = $tns.port
            $servicename = $tns.servicename
#            $connection.connection.connectionstring = "Data Source=$servicename;User Id=$($credentials.username);Password=$($credentials.getnetworkcredential().password);"
        }

#        }else{
        $connection.connection.connectionstring = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=$hostname)(PORT=$port)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=$servicename)));User Id=$($credentials.username);Password=$($credentials.getnetworkcredential().password);Min Pool Size=2;Max Pool Size=5"
#        $connection.connection.connectionstring = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=$hostname)(PORT=$port)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=$servicename)));User Id=$($credentials.username);Password=$($credentials.getnetworkcredential().password);pooling=false"
#        }

        try{
            $connection.connection.open()
        }catch{
            throw $_
            return
        }

#        $info = $connection.connection.getsessioninfo()
#        $info.timezone = 'US/Eastern'
#        $connection.connection.setsessioninfo($info)

        $connection.OraclePipeline = [powershell]::Create()
        $connection.OraclePipeline | Add-Member -MemberType NoteProperty -Name 'AsyncResult' -Value $null

        [void]$script:oracleconnections.add($connection)



#        if (-not $script:oracledefaultconnectionname){
            Set-OracleDefaultQueryHubConnection $connectionName
#        }

        try{
            $sid = "select sys_context('userenv','sid') sid from dual" | invoke-oraclequery -connectionname $connectionname | select -expand sid
            $connection.sid = $sid
#            $script:oracleconnections|where name -eq $connectionname
        }catch{
        }
    }
    end{
    }


}

function Invoke-ClearOracleQueryHubConnectionPool{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -mandatory -position 0 -valuefrompipeline -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }
    process{
        $connection = $script:oracleconnections | where {$_.name -eq $connectionname}

        if ($connection){
            [System.Data.SqlClient.SqlConnection]::ClearPool($connection.connection)
            $connection.name
        }
    }
    end{
    }
}

function Remove-OracleQueryHubConnection{


    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -mandatory -position 0 -valuefrompipeline -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
        $connections = @()
    }

    process{
        setdynamicparametervariables $psboundparameters

        $connections +=  $script:oracleconnections | where {$_.name -eq $connectionname}

    }
    end{


        foreach ($connection in $connections){

            $connection.name

            if (-not $connection){
                throw 'Connection not found'
            }

            try{
                $connection.OraclePipeline.stop()
            }catch{
                write-warning "Error closing oraclepipeline"
            }

            try{
                $connection.OraclePipeline.dispose()
            }catch{
                write-warning "Error disposing oraclepipeline"
            }


            try{
                $connection.connection.close()
            }catch{
                write-warning "Error closing database connection"
            }
            try{
                [Oracle.ManagedDataAccess.Client.OracleConnection]::clearpool($connection.connection)
            }catch{
                write-warning "Error removing database from pool"
            }

            try{
                $connection.connection.dispose()
            }catch{
                write-warning "Error disposing database connection"
            }

            if ($script:oracledefaultconnectionname -eq $connection.name){
                $script:oracledefaultconnectionname = $null
            }

            $script:oracleconnections.remove($connection)

        }
    }
}

function Get-OracleQueryHubConnections{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

	$script:oracleconnections | select -expand name
	
}


function Set-OracleDefaultQueryHubConnection{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -mandatory -position 0 -valuefrompipeline -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        $script:oracledefaultconnectionname = $connectionName
    }
    end {}

}

function Get-OracleDefaultQueryHubConnection{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    if (-not $script:oracledefaultconnectionname){
        throw "Default oracle connection is Not Set"
    }

    $script:oracledefaultconnectionname 

}

function Get-OracleScripts{
#    [CmdletBinding(SupportsShouldProcess=$true)]
#    param(
#    )
    begin{
    }
    process{
        get-childitem $script:oraclescriptdirectory | where {-not $_.psiscontainer} # | select -expand name
    }
    end{
    }
}

function Invoke-OracleScript{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -mandatory -position 0 -Type string -Name ScriptName -validateset (get-childitem -path $script:Oraclescriptdirectory -ea 0 | where {-not $_.psiscontainer}| select -expand name) |
            newdynamicparameter -position 1 -Type hashtable -Name BindReplacements |
            newdynamicparameter -position 2 -Type hashtable -Name StringReplacements |
            newdynamicparameter -position 3 -Type int -Name First |
            newdynamicparameter -position 4 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection) |
            newdynamicparameter -position 5 -Type switch -Name View |
            newdynamicparameter -position 6 -Type switch -Name GetPath |
            newdynamicparameter -position 7 -Type switch -Name Edit |
            newdynamicparameter -position 8 -Type switch -Name NoOutput 

        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{

        setdynamicparametervariables $psboundparameters

        if ($view){
            get-content "$($script:oraclescriptdirectory)\$scriptName" #-raw
        }elseif ($edit){
            if ($TextEditorCommand){
               [scriptblock]::create("$texteditorCommand $($script:oraclescriptdirectory)\$scriptName").invoke()
            }else{
                throw "Text editor command not set"
            }
        }elseif ($getpath){
            "$($script:oraclescriptdirectory)\$scriptName"
        }else{
            get-content "$($script:oraclescriptdirectory)\$scriptName" -raw | invoke-oraclequery -stringreplacements $stringreplacements -bindreplacements $bindreplacements -first $first -connectionname $connectionname -nooutput:($nooutput)
        }
    }

    end{}

    
}

function Get-OracleQueryHubReader{
   param(
       [oracle.manageddataaccess.client.oraclecommand]$Command,
       [string]$Refcursor = $null,
       [ref]$RecordsAffected,
       [ref]$ResultRows
   )

    if ($refcursor){
        $recordsAffected.value = $command.executenonquery()
        try{
            $reader = ([Oracle.ManagedDataAccess.Types.OracleRefCursor]$command.parameters[$refcursor].value).getdatareader()

        #ignore errors - if there was an error in the script being run, that will be captured. however, we don't want this error to propogate
        }catch{
        }
    }else{
        $reader = $command.executereader()
        $recordsAffected.value = $reader.recordsaffected
    }

    if (-not $reader.hasrows){
        $reader.dispose()
        return
    }

    $fieldnames = @()
    $fieldtypes = @()

    for ($x = 0; $x -lt $reader.fieldcount;$x++){

        $name = $reader.GetName($x)
        $fieldname = $name

        Try{
            $fieldType = $reader.GetFieldType($x)
        }catch{
            $fieldType = "asdf".GetType()
            $fieldName = $Name + "[error]"
        }

        $duplicatecount = 0

        while ($true){
            if ($fieldnames.contains($fieldname)){
                $duplicatecount++
                $fieldname = $name + '['+$duplicatecount+']'
            }else{
                break
            }
        }

        $fieldnames += $fieldname
        $fieldtypes += $fieldtype
    }

    $table = @()

    while ($reader.read()){

        $values = new-object object[]  $reader.fieldcount
#        $readfieldcount = $reader.getvalues($values)
        $readfieldcount = $reader.getproviderspecificvalues($values)
        $recordhash = New-Object System.Collections.Specialized.OrderedDictionary

        for ($fieldindex = 0; $fieldindex -lt $fieldnames.count; $fieldindex++){
            try{
                if ($values[$fieldindex] -is [Oracle.ManagedDataAccess.Types.OracleDecimal] -and -not $values[$fieldindex].isnull){
                    $recordhash.add($fieldnames[$fieldindex],[Oracle.ManagedDataAccess.Types.OracleDecimal]::setprecision($reader.getoracledecimal($fieldindex),28).value)
                }else{
                    $recordhash.add($fieldnames[$fieldindex],$values[$fieldindex].value)
                }
            }catch{
                $recordhash.add($fieldnames[$fieldindex],$_.exception)
            }
        }

        $newrecord = [pscustomobject] $recordhash 
        $table += $newrecord
    }

#    #override the tostring so that null values are hidden (instead of showing null)
#    foreach ($prop in $table|gm|where {$_.membertype -eq 'noteproperty'}){
#        $table."$($prop.name)" | add-member -membertype scriptmethod -name tostring -value {if ($this.isnull){""}else{$this.value.tostring()}} -force
#    }
#        $reader
        $reader.dispose()
        $resultrows.value = $table
#    }
#
#                        foreach ($record in $readerresult){
#
#
#                            $values = new-object object[] $fields.count
#                            $readfieldcount = $record.getvalues($values)
#
#                            $recordhash = New-Object System.Collections.Specialized.OrderedDictionary
#                            for ($fieldindex = 0; $fieldindex -lt $fields.count; $fieldindex++){
#                                $recordhash.add($fields[$fieldindex],$values[$fieldindex])
#                            }
#
#                            $newrecord = [pscustomobject] $recordhash 
#
#                            $table += $newrecord
#                        }
}


function Invoke-OracleQuery{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -position 0 -mandatory -valuefrompipeline -Type string -Name Script | 
            newdynamicparameter -position 1 -Type hashtable -Name BindReplacements | 
            newdynamicparameter -position 2 -Type hashtable -Name StringReplacements | 
            newdynamicparameter -position 3 -Type int -Name First |
            newdynamicparameter -position 4 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection) |
            newdynamicparameter -position 5 -Type switch -Name Count |
#            newdynamicparameter -position 6 -Type switch -Name KeepSemicolonEnding |
            newdynamicparameter -position 7 -Type switch -Name SkipStringReplacements |
            newdynamicparameter -position 8 -Type switch -Name NoSaveResults |
            newdynamicparameter -position 9 -Type switch -Name Quiet |
            newdynamicparameter -position 10 -Type switch -Name NoOutput |
            newdynamicparameter -position 11 -Type switch -Name GetExplainPlan
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        $connectionobject = $script:oracleconnections | where {$_.name -eq $connectionname} 

        if ($connectionobject.oraclepipeline.asyncresult -and -not $connectionobject.oraclepipeline.asyncresult.iscompleted){
            throw ("connection is busy")
        }

        $script = $script.trim()

        $queries = $script -split '(?m)^\s{0,}/\s{0,}$'


        foreach ($query in $queries){
            if ($query.trim() -match '^\s{0,}$' -or $query.trim() -match '^print [^ ;]{1,};'){
                continue
            }
            $query += "/"

            $errored = $null

            $bindVariables = @()
            $bindlines = 0

            while ($true){
                $matched = $query | select-string '(?m)^ {0,}variable [^;]*;'  -allmatches | select -expand matches 

                if (-not $matched){
                    break;
                }

                if ($details = ($matched[0] | select -expand value) -match '(?<=^ {0,})variable {1,}(?<bindName>[^ ;]+) (?<datatype>[^ ;]{1,})( {1,}default {1,}(''{0,1})(?<defaultValue>[^'']{0,})''{0,1}){0,1};'){

                    $bindvariable = new-object PSObject
                    $bindvariable | add-member -membertype NoteProperty -name "Name" -value $matches.bindName
                    $bindvariable | add-member -membertype NoteProperty -name "DataType" -value $matches.datatype

                    

                    if ($bindReplacements -and ("$($matches.bindName)" -in $bindreplacements.keys)){
                        
                        $bindvalue = $bindReplacements["$($matches.bindName)"]
                        if (-not $bindvalue -or $bindvalue -eq ''){
                            $bindvalue = [system.dbnull]::value
                        }

                        $bindvariable | add-member -membertype NoteProperty -name "DataValue" -value $bindvalue

                        $execBindDefault = $query | select-string "(?m)^ {0,}exec {1,}:$($bindvariable.name) {0,}:= {0,}('{0,1})(?<defaultValue>[^']{0,})'{0,1} {0,};" -allmatches | select -expand matches

                        if ($execbindDefault){
                            $query = $query.substring(0,$execbinddefault[0].index) + '--' + $execbinddefault[0].value + $query.substring($execbinddefault[0].index + $execbinddefault[0].length)
                        }
                    }elseif ($matches.datatype -eq "refcursor"){
                        $bindvariable | add-member -membertype NoteProperty -name "DataValue" -value $null
                    }else{

                        $binddefaultvalue = $matches.defaultvalue

                        if (-not $binddefaultvalue -or $binddefaultvalue -eq ''){
                            $binddefaultvalue = [system.dbnull]::value
                        }

                        $execBindDefault = $query | select-string "(?m)^ {0,}exec {1,}:$($bindvariable.name) {0,}:= {0,}('{0,1})(?<defaultValue>[^']{0,})'{0,1} {0,};" -allmatches | select -expand matches

                        if ($execbindDefault){
                            if ($execBindDetails = ($execBindDefault[0] | select -expand value) -match "(?<=^ {0,})exec {1,}:$($bindvariable.name) {0,}:= {0,}('(?<defaultValue>[^']{0,})'|(?<defaultValue>[^;]{1,}));" ){
                                $binddefaultvalue = $matches.defaultvalue

                                if (-not $binddefaultvalue -or $binddefaultvalue -eq ''){
                                    $binddefaultvalue = [system.dbnull]::value
                                }
                            }

                            $query = $query.substring(0,$execbinddefault[0].index) + '--' + $execbinddefault[0].value + $query.substring($execbinddefault[0].index + $execbinddefault[0].length)

                        }

                        $bindvariable | add-member -membertype NoteProperty -name "DataValue" -value (replacevariable "$($bindvariable.name)" $binddefaultvalue).value
                    }

                    $bindVariables += $bindvariable

                }

                $query = $query.substring(0,$matched[0].index) + '--' + $matched[0].value + $query.substring($matched[0].index + $matched[0].length)
            }

            if (-not $SkipStringReplacements) {

                if ($StringReplacements){
                    foreach ($repl in $StringReplacements.keys){
                        $query = $query.replace($repl,$stringreplacements.item($repl))
                    }
                }
                

                $skipIndex = -1

                while ($true){
                    $matched = $query | select-string '&[^ ''&]+'  -allmatches | select -expand matches | where index -gt $skipIndex

                    if (-not $matched){
                        break;
                    }

                    $skipIndex = $matched[0].index
                    $varname = $matched[0] | select -expand value;

                    if ($varname -match '(?<=\[)(?<defaultVar>[^)]+)(?=\])'){
                        $varDefault = $matches.defaultVar
                        $varName = $varName.replace("[$vardefault]","")
                    }else{
                        $vardefault = $null
                    }

                    $var = replacevariable $varname $varDefault
                    
                    if ($var.iscancelled){
    #                    continue
                        break;
                    }

                    $replacementText = $var.value
                    $replacementText = "$([regex]::split($replacementText, '(.{10})') | where {$_} | join -joinon "'||'")"

                    $skipIndex = $skipIndex + $var.value.length

                    $query = $query.replace(($matched[0] | select -expand value),$replacementText)
                }
            }


            $query = $query.trim()

            if ($query.endswith("/")){
                $query = $query.substring(0,$query.length-1)
            }

#            if (-not $keepsemicolonending -and $query.endswith(";")){
#                $query = $query.substring(0,$query.length-1)
#            }

             if ($First){
                $query = "select * from (`n$query`n) where rownum <= $First"
             }

             if ($count){
                $query = "select count(1) ""count(1)"" from (`n$query`n)"
             }

            if ($query -match "(?<!--.*)dbms_output\.put_line"){
                $search = [regex]::new("(?<!--.*)dbms_output\.put_line", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                $query = $search.Replace($query, "dbms_output.enable(buffer_size => null);dbms_output.put_line", 1)
            }

            try{

                $command = [oracle.manageddataaccess.client.oraclecommand]::new($query)
                $command.connection = $connectionobject.connection
                $command.commandtype = [system.data.commandtype]::text
                $command.commandtimeout = 60 * $script:timeoutminutes
                $command.bindbyname = $true

                $refcursor = $null
                $cancelled = $false

                foreach ($bindvariable in $bindvariables){
                    $param = [Oracle.ManagedDataAccess.Client.OracleParameter]::new()
                    $param.parametername = $bindvariable.name

                    switch -regex ($bindvariable.datatype){
                        "^refcursor$"{
                            $param.oracledbtype = [Oracle.ManagedDataAccess.Client.OracleDbType]::refcursor
                            $param.direction = [system.data.parameterdirection]::returnvalue
                            $refcursor = $bindvariable.name
                            break
                        }"^number$"{
                            $param.oracledbtype = [Oracle.ManagedDataAccess.Client.OracleDbType]::double
                            $param.direction = [system.data.parameterdirection]::inputoutput
                            $param.value = $bindvariable.datavalue
                            $param.parametername = $bindvariable.name
                            break
                        }"^varchar2"{
                            $param.oracledbtype = [Oracle.ManagedDataAccess.Client.OracleDbType]::varchar2
                            $param.direction = [system.data.parameterdirection]::inputoutput
                            $param.value = $bindvariable.datavalue
                            $param.parametername = $bindvariable.name
                            break
                        }"^clob$"{
                            $param.oracledbtype = [Oracle.ManagedDataAccess.Client.OracleDbType]::clob
                            $param.direction = [system.data.parameterdirection]::inputoutput
                            $param.value = $bindvariable.datavalue
                            $param.parametername = $bindvariable.name
                            break
                        }"^date$"{
                            $param.oracledbtype = [Oracle.ManagedDataAccess.Client.OracleDbType]::date
                            $param.direction = [system.data.parameterdirection]::inputoutput
                            $param.value = [datetime]$bindvariable.datavalue
                            $param.parametername = $bindvariable.name
                            break
                        }"^timestamp$"{
                            $param.oracledbtype = [Oracle.ManagedDataAccess.Client.OracleDbType]::timestamp
                            $param.direction = [system.data.parameterdirection]::inputoutput
                            $param.value = [datetime]$bindvariable.datavalue
                            $param.parametername = $bindvariable.name
                            break
                        }"^xmltype$"{
                            $param.oracledbtype = [Oracle.ManagedDataAccess.Client.OracleDbType]::xmltype
                            $param.direction = [system.data.parameterdirection]::inputoutput
                            $refcursor = $bindvariable.name
                            break
                        }default{
                            $param.oracledbtype = [Oracle.ManagedDataAccess.Client.OracleDbType]::object
                            $param.direction = [system.data.parameterdirection]::inputoutput
                            $param.value = $bindvariable.datavalue
                            $param.parametername = $bindvariable.name
                            break
                        }
                    }

                    [void]$command.parameters.add($param)
                }

                try{



                    try{
                        $recordsaffected = $null
                        $readerresult = $null


                        [void]$connectionobject.oraclePipeline.commands.clear()
                        [void]$connectionobject.oraclePipeline.streams.clearstreams()
                        [void]$connectionobject.oraclePipeline.Addcommand('Get-OracleQueryHubReader')
                        [void]$connectionobject.oraclePipeline.addparameter("Command",$command)
                        [void]$connectionobject.oraclePipeline.addparameter("Refcursor",$refcursor)
                        [void]$connectionobject.oraclePipeline.addparameter("RecordsAffected",[ref]$recordsaffected)
                        [void]$connectionobject.oraclePipeline.addparameter("ResultRows",[ref]$readerresult)

                        $connectionobject.oraclePipeline.AsyncResult = $connectionobject.oraclePipeline.BeginInvoke()

                        $queryStart = get-date

                        if ($quiet){

                            if ($host.name -eq 'consolehost'){
                                $setControlC = $true
                                $controlCSave = [console]::treatcontrolcasinput
                            }else{
                                $setControlC = $false
                            }

                            if ($setcontrolc){
                                [console]::TreatControlCAsInput = $true
                            }
                            while (-not $connectionobject.oraclepipeline.asyncresult.iscompleted){
                                if ($setcontrolc -and [console]::KeyAvailable) {
                                    $key = [system.console]::readkey($true)
                                    if (($key.modifiers -band [consolemodifiers]"control") -and ($key.key -eq "C")) {
                                        $completed = $false
                                        break
                                    }
                                }
                            }

                            if ($setcontrolc){
                                [console]::TreatControlCAsInput = $false
                                [console]::TreatControlCAsInput = $controlcsave
                            }

                        }else{
                            $completed = showasyncwaitmessage -message $command.commandtext -conditionvariablename "connectionobject" -conditionmethodpropertychain @('oraclePipeline','asyncresult','iscompleted') 
                        }

                        if (-not $completed){
                            $cancelled = $true
                            $command.cancel()
                            throw "Query was Cancelled"
                        }


                        $queryend = (new-timespan $querystart (get-date))

#                        $readerresult =  $connectionobject.oraclePipeline.EndInvoke($connectionobject.oraclePipeline.AsyncResult)
                        $connectionobject.oraclePipeline.EndInvoke($connectionobject.oraclePipeline.AsyncResult)

                        if ($connectionobject.oraclepipeline.streams.error){
                            throw $connectionobject.oraclepipeline.streams.error
                        }


                         #cleanup code.
                    }catch{
                        $errored = $_.exception.message | select-string '(?ms)(?<=Exception calling "(ExecuteReader|ExecuteNonQuery)" with "0" argument\(s\): ").*(?=" You cannot call a method on a null-valued expression.$)'
                        if ($errored){
                            $errored = $errored | select -expand matches | select -expand value
                        }else{
                            $errored = $_.exception.message
                        }
                    }finally{
                        [console]::TreatControlCAsInput = $false
                    }

                    $table = $readerresult

                    if (-not $cancelled -and -not $errored -and -not $nosaveresults -and $getExplainPlan){
                        $explainplan = get-oracleexplainplan -connectionname $connectionname
                    }

                    if (-not $cancelled -and -not $nosaveresults){
                        $dbmsoutput = ((get-oracledbmsoutput $connectionname -erroraction silentlycontinue) -join "`n")
                    }

                    if ($table.count -eq 0){
                        if ($errored){
                            $table = [pscustomobject] @{Error=$errored}
                        }elseif($recordsaffected -ge 0){
                            $table = [pscustomobject] @{RecordsAffected=$recordsaffected}
                        }elseif ($dbmsoutput){
                            $table = [pscustomobject] @{dbms_output=$dbmsoutput}
                        }else{
                            $null
                        }
                    }

                    if (-not $nosaveresults -and -not $cancelled){
                        $script:lastoraclequeryrequest = $query
                        $script:lastoraclequeryresult = $table
                        $script:lastoraclequeryexplain = $explainplan
                        $script:lastoraclequerytime = $queryend
                        $script:lastoracledbmsoutput = $dbmsoutput
                        $script:lastoraclequeryerror = $errored


                        if (get-typedata -typename QueryHub.OLast){
                            remove-typedata -typename QueryHub.OLast -confirm:$false
                        }

                        if ($script:lastoraclequeryresult){
                            foreach ($prop in ($script:lastoraclequeryresult|gm|where {$_.membertype -eq 'noteproperty'})){
                                update-typedata -typename QueryHub.OLast -membertype NoteProperty -membername $prop.name -value $null -confirm:$false
                            }
                            
                        }
                    }

                    if (-not $nooutput){
                        $table
                    }
                }finally{
                }
                
            }finally{
                if ($command){
                    $command.dispose()
                }
            }

        }
    }

    end{
    }
}

function Get-OracleDbmsOutput{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -position 0 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        $connection = $script:oracleconnections | where {$_.name -eq $connectionname} | select -expand connection

        $returnValue = @()

        #anonymous pl/sql block to get multiples lines of text per fetch
        $anonymous_block = "begin dbms_output.get_lines(:1, :2); end;"

        #used to indicate number of lines to get during each fetch
        $numToFetch = 8

        #used to determine number of rows fetched in anonymous pl/sql block
        $numLinesFetched = 0

        #create parameter objects for the anonymous pl/sql block
        $p_lines = [Oracle.ManagedDataAccess.Client.OracleParameter]::new("", [Oracle.ManagedDataAccess.Client.OracleDbType]::varchar2, $numtofetch, "", [system.data.parameterdirection]::output)

            $p_lines.CollectionType = [Oracle.ManagedDataAccess.Client.OracleCollectionType]::PLSQLAssociativeArray
            $p_lines.ArrayBindSize = new-object int[] ($numtofetch )

            #set the bind size value for each element
            For ($i = 0; $i -lt $numtofetch;$i++){
                $p_lines.ArrayBindSize[$i] = 32000
            }

            #this is an input output parameter...
            #on input it holds the number of lines requested to be fetched from the buffer
            #on output it holds the number of lines actually fetched from the buffer
            $p_numlines = [Oracle.ManagedDataAccess.Client.OracleParameter]::new("", [Oracle.ManagedDataAccess.Client.OracleDbType]::decimal, "", [system.data.parameterdirection]::inputoutput)

                #set the number of lines to fetch
                $p_numlines.Value = $numtofetch

                #set up command object and execute anonymous pl/sql block
                $cmd =  $connection.CreateCommand()
                    $cmd.CommandText = $anonymous_block
                    $cmd.commandtype = [system.data.commandtype]::text
                    [void]$cmd.Parameters.Add($p_lines)
                    [void]$cmd.Parameters.Add($p_numlines)
                    [void]$cmd.ExecuteNonQuery()

                    #get the number of lines that were fetched (0 = no more lines in buffer)
                    $numLinesFetched = [Int]::Parse(([Oracle.ManagedDataAccess.Types.oracledecimal]$p_numlines.Value).ToString())
                    $outlines = [Oracle.ManagedDataAccess.Types.oraclestring[]]$p_lines.Value

                    #as long as lines were fetched from the buffer...
                    While ($numLinesFetched -gt 0){
                        #write the text returned for each element in the pl/sql
                        #associative array to the console window
                        For ($i = 0; $i -lt $numlinesfetched;$i++){
                            if ($outlines[$i].isnull){
                                $returnvalue += ""
                            }else{
                                $returnvalue += $outlines[$i].value
                            }
                        }

                        #re-execute the command to fetch more lines (if any remain)
                        [void]$cmd.ExecuteNonQuery()

                        #get the number of lines that were fetched (0 = no more lines in buffer)
                        $numLinesFetched = [Int]::Parse(([Oracle.ManagedDataAccess.Types.oracledecimal]$p_numlines.Value).ToString())
                        $outlines = [Oracle.ManagedDataAccess.Types.oraclestring[]]$p_lines.Value
                    }



        $returnValue
    }
    end{
    }
}


function Invoke-OracleQueryFromFile{
    [CmdletBinding()] 
    param( 
    )

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -mandatory -position 0 -valuefrompipeline -Name File -validatescript {test-path $_ -PathType Leaf} |
            newdynamicparameter -position 1 -Type hashtable -Name BindReplacements |
            newdynamicparameter -position 2 -Type hashtable -Name StringReplacements |
            newdynamicparameter -position 3 -Type int -Name First |
            newdynamicparameter -position 4 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection) |
            newdynamicparameter -position 5 -Type switch -Name NoOutput |
            newdynamicparameter -position 6 -Type switch -Name NoSaveResults |
            newdynamicparameter -position 7 -Type switch -Name Quiet
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        cat $file -raw | invoke-oraclequery -connectionname $connectionName -bindreplacements $bindreplacements -stringreplacements $stringreplacements -first $first -nooutput:$nooutput -nosaveresults:$nosaveresults -quiet:$quiet
    }

    end{
    }
}


function Get-OracleErrors{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -position 0 -mandatory -valuefrompipeline -Type string -Name NameLike |
            newdynamicparameter -position 1 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection) |
            newdynamicparameter -position 2 -Type switch -Name Full
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        if ($full){
            "select * from all_errors where lower(name) like lower('$namelike') order by owner,name,type,sequence,line,position" | invoke-oraclequery -connectionname $connectionname
        } else {
            "select name,line,position,text from all_errors where lower(name) like lower('$namelike') order by owner,name,type,sequence,line,position" | invoke-oraclequery -connectionname $connectionname | ft -auto -wrap
        }
    }

    end{}
}

function Invoke-OraclePackageBuild{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )
	 
    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -position 0 -mandatory -valuefrompipeline -Type string -Name FilePath -validatescript {test-path $_ -PathType Leaf} |
            newdynamicparameter -position 1 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection) |
            newdynamicparameter -position 3 -Type switch -Name SkipStringReplacements
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        $file = (resolve-path $filepath).providerpath 
        $fileContents = cat $file -raw 
        $fileContents = $fileContents.trim()

        
        if ($skipstringreplacements){
#            $fileContents | invoke-oraclequery -connectionname $connectionname -keepsemicolonending -skipstringreplacements
            $fileContents | invoke-oraclequery -connectionname $connectionname -skipstringreplacements
        }else{
#            $fileContents | invoke-oraclequery -connectionname $connectionname -keepsemicolonending
            $fileContents | invoke-oraclequery -connectionname $connectionname 
        }

        if ($fileContents -match 'create {1,}(or {1,}replace {1,}){0,1}(force ){0,1}(?<type>trigger|package|function|procedure|view|type) {1,}(?<body>body {1,})*(?<packagename>[^ (]+)' -and $matches.packagename){
            if ($matches.body){
                $clause = "and type = 'PACKAGE BODY'"
            } else {
                $clause = "and type = '$($matches.type.toupper())'"
            }

            $package = $matches.packagename
            if ($package.contains(".")){
                $package = $package.substring($package.indexof(".") + 1)
            }

            #this so unused variables can be checked
            $alter = "alter session set PLSCOPE_SETTINGS='IDENTIFIERS:ALL'" | invoke-oraclequery -connectionname $connectionname

#write-host "select name,line,position,text from all_errors where lower(name) = lower('$package') $clause order by owner,name,type,sequence,line,position"
            $check = "select name,line,position,text from all_errors where lower(name) = lower('$package') $clause order by owner,name,type,sequence,line,position" | invoke-oraclequery -connectionname $connectionname
            
            if ($check) {
                $check 
            } else {
                $variableCheck = @"
                    select 
                        a.object_name name,
                        a.line,
                        a.col position,
                        'Unused variable: ' || a.name text
                    from user_identifiers a 
                    where a.object_type = 'PACKAGE BODY' 
                    and lower(a.object_name) = lower('$package')
                    and a.type in ('VARIABLE','DECLARATION') 
                    and a.usage = 'DECLARATION' 
                    and (
                        select count(1)
                        from user_identifiers b 
                        where b.object_type = a.object_type 
                        and b.object_name = a.object_name 
                        and b.signature = a.signature 
                        and b.usage <> 'DECLARATION'
                    ) = 0
--                    ) + (
--                        select count(1)
--                        from user_identifiers c 
--                        where c.object_type = a.object_type 
--                        and c.object_name = a.object_name 
--                        and c.signature = a.signature 
--                        and c.usage <> 'DECLARATION' 
--                        and c.line <> a.line 
--                        and c.col <> a.col
--                    ) < 2
                    order by object_name,line,col,name
"@ | invoke-oraclequery -connectionname $connectionname
                if ($variablecheck) {
                    $variablecheck 
                }
            }

        }else{
            write-warning 'Couldn''t determine package name. No code was compiled.'
        }
    }

    end{}
}

function Get-OracleDependencies(){

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -position 0 -valuefrompipeline -Type string -Name NameRegex |
            newdynamicparameter -position 1 -Type string -Name ReferencedNameRegex |
            newdynamicparameter -position 2 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection) 
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        if ($nameregex){
            
            $nameclause = "where regexp_like(name, '$nameRegex','i')"
        }
        if ($referencednameregex){

            if ($nameregex){
                $referencednameclause = "and "
            }else{
                $referencednameclause = "where "
            }

            $referencednameclause += "regexp_like(referenced_name, '$referencednameRegex','i')"
        }


        $query = "select * from all_dependencies $nameclause $referencednameclause order by name,type,referenced_name,referenced_type" 

        $result = $query | invoke-oraclequery -connectionname $connectionname

        $result 
    }

    end{}
}

function Get-OracleTableNames{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -position 0 -mandatory -valuefrompipeline -Type string -Name TableLike |
            newdynamicparameter -position 1 -Type char -Name EscapeCharacter -defaultvalue '\' |
            newdynamicparameter -position 2 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        $query = "select distinct owner || '.' || table_name table_name from all_tab_cols where table_name like '$($tablelike.toupper())' escape '$escapecharacter' order by table_name asc" 

        $result = $query | invoke-oraclequery -connectionname $connectionname

        $result | % {$_.table_name}
    }

    end{}
}

function Get-OracleTableSelectString{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -position 0 -mandatory -valuefrompipeline -Type string -Name Table |
            newdynamicparameter -position 1 -Type 'string[]' -Name OmitColumns -defaultvalue @() |
            newdynamicparameter -position 2 -Type 'string' -Name TableAlias |
            newdynamicparameter -position 3 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection) 
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters


        $omitSql = ""
        
        if ($omitcolumns.count -gt 0){
            $omitSql = "and column_name not in ('fakenomatch'"
            foreach ($column in $omitcolumns){
                $omitSql += ",'$($column.toupper())'"
            }
            $omitSql += ')'
        }

        $prepend = ""
        if ($tableAlias){
            $prepend = $tableAlias + '.'
        }

        $result = "select '$prepend' || column_name column_name from all_tab_cols where table_name = '$($table.toupper())' $omitSql order by column_name asc"  | invoke-oraclequery -connectionname $connectionname

        $result.column_name -join ","
    }
    end {}

}

function Get-OracleTableSchema{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -position 0 -valuefrompipeline -Type string -Name Table |
            newdynamicparameter -position 1 -Type string -Name Column |
            newdynamicparameter -position 2 -Type char -Name EscapeCharacter -defaultvalue '\' |
            newdynamicparameter -position 3 -Type string -Name Sort -defaultvalue 'Name' -validateset @('Name','Id') |
            newdynamicparameter -position 4 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection) |
            newdynamicparameter -position 5 -Type switch -Name TableIsLike |
            newdynamicparameter -position 6 -Type switch -Name ColumnIsLike
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        $tableQuery = "1=1"
        $columnQuery = "1=1"

        if ($table) {
            if ($tableislike){
                $tableQuery = "table_name like '$($table.toupper())' escape '$escapecharacter'"
            }else{
                $tableQuery = "table_name = '$($table.toupper())'" 
            }
        }

        if ($column) {
            if ($columnislike){
                $columnQuery = "column_name like '$($column.toupper())' escape '$escapecharacter'"
            }else{
                $columnQuery = "column_name = '$($column.toupper())'" 
            }
        }

        $result = "select table_name,column_id,column_name,data_type,data_length,data_precision,nullable,char_length from all_tab_cols where $tablequery and $columnQuery and hidden_column = 'NO' order by table_name asc, column_$sort asc"  | invoke-oraclequery -connectionname $connectionname

        $result
    }

    end{}
}

function Get-OracleCodeNames{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -position 0 -valuefrompipeline -Type string -Name NameLike |
            newdynamicparameter -position 1 -Type string -Name TextLike |
            newdynamicparameter -position 2 -Type char -Name EscapeCharacter -defaultvalue '\' |
            newdynamicparameter -position 3 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection) 
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters


        $nameQuery = $null
        $textQuery = $null

        if ($namelike){
            $nameQuery = "lower(name) like '$($namelike.tolower())' escape '$escapecharacter'"
        }

        if ($textlike){
            $textQuery = "lower(text) like '$($textlike.tolower())' escape '$escapecharacter'"
        }

        if (-not $namequery -and -not $textquery){
            "either name or text needs to be provided"
            return
        }

        if (-not $namequery){
            $textQuery = "where $textquery"
        }else{
            $nameQuery = "where $nameQuery"

            if ($textquery){
                $textQuery = "and $textquery"
            }
        }

        $wherequery = "$namequery $textQuery"


        $result = "select distinct owner,name,type from dba_source $whereQuery order by name asc,type asc,owner asc"  | invoke-oraclequery -connectionname $connectionname

        $result
    }
    end{}
}

function Test-OracleConnection{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -position 0 -valuefrompipeline -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        try{
            $result = "select 'asdf' from dual" | invoke-oraclequery -connectionname $connectionName
        }catch{
            if ($_.exception.message -match 'connection is busy'){
                $busy = $true
            }
        }

        $new = new-object PSCustomObject
        $new | add-member -membertype NoteProperty -name "ConnectionName" -value $connectionName

        if ($busy){
            $new | add-member -membertype NoteProperty -name "Result" -value "Busy"
        } elseif (-not $result -or ($result | gm | where {$_.name -match '^(error)$'})){
            $new | add-member -membertype NoteProperty -name "Result" -value $false
        } else {
            $new | add-member -membertype NoteProperty -name "Result" -value $true
        }

        $new
        return
    }
}

function Get-OracleCodeDefinition{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -position 0 -mandatory -valuefrompipeline -Type string -Name name |
            newdynamicparameter -position 1 -Type string -Name Type -validateset @('PackageBody','Package','Trigger','Procedure','Function','JavaSource','Library','TypeBody','Type') -defaultvalue PackageBody |
            newdynamicparameter -position 2 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        $typestring = $type

         if ($type.tolower() -eq 'packagebody'){
            $typestring = 'PACKAGE BODY'
        } elseif ($type.tolower() -eq 'typebody') {
            $typestring = 'TYPE BODY'
        } elseif ($type.tolower() -eq 'javasource') {
            $typestring = 'JAVA SOURCE'
        }

        $result = "select text from dba_source where type = '$($typestring.toupper())' and lower(name) = '$($name.tolower())' order by line asc" | invoke-oraclequery -connectionname $connectionname 
        
        if ($result){
            $result | select -expand text | % {$_ -replace "`n","`r"}
        }
    }

    end{}

}

function Get-OracleSession{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -position 0 -valuefrompipeline -Type string -Name SidSerial |
            newdynamicparameter -position 1 -Type string -Name OsUser -defaultvalue $env:username |
            newdynamicparameter -position 2 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection) |
            newdynamicparameter -position 3 -Type switch -Name IncludeSqlStatements |
            newdynamicparameter -position 4 -Type switch -Name Kill
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters


        $where = "where 1 = 1 "

        if ($sidserial){
            $where += "and b.sid || '.' || b.serial# = '$sidserial' "
        } 
        if ($osUser -and $osuser -ne '*'){
            $where += "and upper(b.osuser) = '$($osuser.toupper())' "
        }


        $results = "select b.sid || '.' || b.serial# ""SID.SERIAL"", a.spid, b.machine, b.username, b.osuser, b.program, b.logon_time, b.status, b.sql_id, b.event, c.sql_fulltext, (select listagg(d.name || ' = ' || nvl(d.value_string,anydata.accesstimestamp(d.value_anydata)), chr(10)) within group (order by d.position) from v`$sql_bind_capture d where d.sql_id = b.sql_id) bind_variables from v`$session b inner join v`$process a on a.addr = b.paddr inner join v`$sql c on c.sql_id = b.sql_id $where order by b.username asc" | invoke-oraclequery -connectionname $connectionName

        if (-not $IncludeSqlStatements){
            foreach ($result in $results){
                $result.sql_fulltext = $null
                $result.bind_variables = $null
            }
        }

        if ($kill){
            foreach ($result in $results){
                $result
                try{
                    stop-oraclesession -sidserial $result."sid.serial" -connectionname $connectionName
                }catch{
                    write-error $_
                }
            }
        }else{
            $results
        }
    }

    end{}
}

function Get-OracleSessionSqlHistory{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -position 0 -valuefrompipeline -Type string -Name SessionId -default ((get-oracleconnectionsession).sid)|
            newdynamicparameter -position 1 -valuefrompipeline -Type string -Name Regex -default '.*' |
            newdynamicparameter -position 2 -valuefrompipeline -Type int -Name Count -default 25 |
            newdynamicparameter -position 3 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection) 
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters



        "select b.sql_id, a.sql_exec_start, b.sql_fulltext from v`$active_session_history a inner join v`$sql b on b.sql_id = a.sql_id where a.session_id = $sessionid and regexp_like(b.sql_fulltext,'$($regex -replace "'","''")','i') order by a.sql_exec_start desc" | invoke-oraclequery -connectionname $connectionname -first $count | select * -uniq


    }

    end{}
}

function Get-OracleSqlHistory{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -valuefrompipeline -Type string -Name WhereSqlIs |
            newdynamicparameter -position 0 -Type string -Name OsUser -defaultvalue $env:username |
            newdynamicparameter -position 1 -Type int -Name Count -defaultvalue 25 |
            newdynamicparameter -position 2 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        $where = ""

        if ($wheresqlis){
            $results = @"
    variable rcursor refcursor;
    variable sql clob;
    declare
    begin
        open :rcursor for select c.last_load_time, c.first_load_time, c.sql_id, c.sql_fulltext from v`$sql c where dbms_lob.compare(upper(c.sql_fulltext), upper(:sql)) = 0 order by last_load_time desc;
    end;
    /
"@ | invoke-oraclequery -connectionname $connectionName -bind @{'sql'=$wheresqlis}
        }else{
            $results = "select c.last_load_time, c.first_load_time, c.sql_id, c.sql_fulltext from v`$sql c order by last_load_time desc" | invoke-oraclequery -connectionname $connectionName -first $count
        }


        $results



    }

    end{}
}

function Stop-OracleSession{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -mandatory -position 0 -valuefrompipeline -Type string -Name SidSerial |
            newdynamicparameter -position 1 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        "alter system kill session '$($sidserial.replace(".",","))'" | invoke-oraclequery  -connectionname $connectionname
    }

    end{}
}


function Get-OracleDefinition{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -position 0 -mandatory -valuefrompipeline -Type string -Name Name |
            newdynamicparameter -position 1 -Type string -Name Type -defaultvalue View -validateset @('AQ_QUEUE','REF_CONSTRAINT','AQ_QUEUE_TABLE','REFRESH_GROUP','AQ_TRANSFORM','RESOURCE_COST','ASSOCIATION','RLS_CONTEXT','AUDIT','RLS_GROUP','AUDIT_OBJ','RLS_POLICY','CLUSTER','RMGR_CONSUMER_GROUP','COMMENT','RMGR_INTITIAL_CONSUMER_GROUP','CONSTRAINT','RMGR_PLAN','CONTEXT','RMGR_PLAN_DIRECTIVE','DATABASE_EXPORT','ROLE','DB_LINK','ROLE_GRANT','DEFAULT_ROLE','ROLLBACK_SEGMENT','DIMENSION','SCHEMA_EXPORT','DIRECTORY','SEQUENCE','FGA_POLICY','SYNONYM','FUNCTION','SYSTEM_GRANT','INDEX_STATISTICS','TABLE','INDEX','TABLE_DATA','INDEXTYPE','TABLE_EXPORT','JAVA_SOURCE','TABLE_STATISTICS','JOB','TABLESPACE','LIBRARY','TABLESPACE_QUOTA','MATERIALIZED_VIEW','TRANSPORTABLE_EXPORT','MATERIALIZED_VIEW_LOG','TRIGGER','OBJECT_GRANT','TRUSTED_DB_LINK','OPERATOR','TYPE','PACKAGE','TYPE_BODY','PACKAGE_SPEC','TYPE_SPEC','PACKAGE_BODY','USER','PROCEDURE','VIEW','PROFILE','XMLSCHEMA','PROXY') |
            newdynamicparameter -position 2 -Type string -Name Owner -defaultvalue Apps |
            newdynamicparameter -position 3 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        $result = "select dbms_metadata.get_ddl('$($type.toupper())','$($name.toupper())','$($owner.toupper())') asdf from dual" | invoke-oraclequery -connectionname $connectionname

        if ($result -and $result.asdf){
            $result | select -expand asdf
        } else {
            $result
        }
    }

    end{}
}

function Get-OracleExplainPlan{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -position 0 -valuefrompipeline -Type string -Name SqlId |
            newdynamicparameter -position 1 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection) |
            newdynamicparameter -position 2 -Type switch -Name Raw 
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        if ($sqlid){
            $result = "select * from table(dbms_xplan.display_cursor('$sqlid'))" | invoke-oraclequery -connectionname $connectionname -nosaveresults
        }else{
            $result = "select * from table(dbms_xplan.display_cursor())" | invoke-oraclequery -connectionname $connectionname -nosaveresults
        }

        if (-not ($result -and $result.plan_table_output)){
            $result
            return
        }

        $output = $result.plan_table_output

        if ($raw){
            $output
            return
        }



        $sqlid =  ($output | select-string '(?<=^sql_id\s{1,})[^,]{1,}' | select -expand matches | select -expand value)
        $grep = $output | select-string '(?<=^\|\*{0,1}\s{0,})(?<id>\d{1,})\s{0,}\|(?<operation>[^|]{0,})\|(?<name>[^|]{0,})\|(?<rows>[^|]{0,})\|(?<bytes>[^|]{0,})\|(?<cost>[^|]{0,})\|(?<time>[^|]{0,})' | select -expand matches 

        $result = @()

        foreach ($row in $grep){
            $new = new-object PSCustomObject
            $new | add-member -membertype NoteProperty -name "SqlId" -value ($sqlId.trim())
            $new | add-member -membertype NoteProperty -name "SortIndex" -value ($null)
            $new | add-member -membertype NoteProperty -name "Id" -value ($row.groups[1].value.trim())
            $new | add-member -membertype NoteProperty -name "Operation" -value ($row.groups[2].value)
            $new | add-member -membertype NoteProperty -name "Name" -value ($row.groups[3].value.trim())
            $new | add-member -membertype NoteProperty -name "Rows" -value ($row.groups[4].value.trim())
            $new | add-member -membertype NoteProperty -name "Bytes" -value ($row.groups[5].value.trim())
            $new | add-member -membertype NoteProperty -name "Cost" -value ($row.groups[6].value.trim())
            $new | add-member -membertype NoteProperty -name "Time" -value ($row.groups[7].value.trim())
            $new | add-member -membertype NoteProperty -name "PredicateInformation" -value ($null)
            $result += $new
        }

        $predicateString = ($output -join "`n") | select-string '(?msi)(?<=predicate\D{1,})\d.*' | select -expand matches | select -expand value

        foreach ($line in ($predicateString -split "`n")){

            if ($line.trim() -match '(?<id>^\d{1,})\s-\s(?<info>.*)'){
                $lastId = $matches.id
                $record = $result | where {$_.id -eq $lastid}
                $result[$result.indexof($record)].PredicateInformation = $matches.info.trim()
            }else{
                $result[$result.indexof($record)].PredicateInformation += " $($line.trim())"
            }
        }

        $sortedresult = $result | sort @{expression={( $_.operation | grep '^\s{0,}' | select -expand matches | select -expand value).length};Descending=$true}, @{expression={[int]$_.id};Descending=$false}

        $x = 0
        foreach ($record in $sortedresult){
            $record.operation = $record.operation.trim()
            $record.sortindex = $x
            $x++
        }

        $sortedResult



    }

    end{}
}

function Get-OracleIndex{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -position 0 -valuefrompipeline -Type string -Name TableName |
            newdynamicparameter -position 1 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        "select * from all_ind_columns where upper(table_name) = '$($tablename.toupper())' order by index_name, column_position" | invoke-oraclequery -connectionname $connectionname
    }

    end{}
}

function Get-OracleFndTableKeys{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -position 0 -valuefrompipeline -Type string -Name TableName |
            newdynamicparameter -position 1 -Type string -Name ConnectionName -validateset (Get-OracleQueryHubConnections) -defaultvalue (Get-OracleDefaultQueryHubConnection)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        @"
select * 
from (
    select 
        'Primary' type,
        fpk.primary_key_name name,
        fc.column_name,
        fpkc.primary_key_sequence sequence,
        null referenced_table,
        null referenced_columns
    from fnd_tables ft
    inner join fnd_primary_keys fpk on fpk.table_id = ft.table_id
    inner join fnd_primary_key_columns fpkc on fpkc.primary_key_id = fpk.primary_key_id
    inner join fnd_columns fc on fc.column_id = fpkc.column_id
    where table_name = '$($tablename.toupper())'
    union all
    select 
        'Foreign' type,
        ffk.foreign_key_name name,
        fc.column_name,
        ffkc.foreign_key_sequence sequence,
        ft_ref.table_name referenced_table,
        (
            select listagg(fc_ref.column_name,',') within group (order by fc_ref.column_name)
            from fnd_primary_key_columns fpkc_ref 
            inner join fnd_columns fc_ref on fc_ref.column_id = fpkc_ref.column_id
            where fpkc_ref.primary_key_id = ffk.primary_key_id
        ) referenced_columns
    from fnd_tables ft
    inner join fnd_foreign_keys ffk on ffk.table_id = ft.table_id
    inner join fnd_foreign_key_columns ffkc on ffkc.foreign_key_id = ffk.foreign_key_id
    inner join fnd_columns fc on fc.column_id = ffkc.column_id
    inner join fnd_primary_keys fpk_ref on fpk_ref.primary_key_id = ffk.primary_key_id
    inner join fnd_tables ft_ref on ft_ref.table_id = ffk.primary_key_table_id
    where ft.table_name = '$($tablename.toupper())'
)
order by type desc,sequence asc,column_name asc
"@ | invoke-oraclequery -connectionname $connectionname
    }

    end{}
}


###########################################################################################################
# SqlServer
###########################################################################################################


function Get-LastSqlServerQueryRequest{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    $script:lastSqlServerQueryRequest
}

function Get-LastSqlServerQueryResult{
    [outputtype('QueryHub.SLast')]
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    if ($script:lastsqlserverQueryResult){
        $script:lastsqlserverQueryResult.psobject.typenames.insert(0,'QueryHub.SLast')
    }
    $script:lastsqlserverQueryResult
}
function Get-LastSqlServerQueryTime{
    [outputtype('Timespan')]
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    $script:lastSqlServerQueryTime
}

function Get-LastSqlServerPrintStatements{
    [outputtype('Timespan')]
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    $script:lastSqlServerPrintStatements
}

function Get-LastSqlServerQueryError{
    [outputtype('String')]
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    $script:lastSqlServerQueryError
}
    
function Remove-SqlServerQueryHubConnection{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -mandatory -position 0 -valuefrompipeline -Type string -Name ConnectionName -validateset (Get-SqlServerQueryHubConnections)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
        $connections = @()
    }

    process{
        setdynamicparametervariables $psboundparameters

        $connections +=  $script:sqlserverconnections | where {$_.name -eq $connectionname}

    }

    end{


        foreach ($connection in $connections){

            $connection.name

            if (-not $connection){
                throw 'Connection not found'
            }

            try{
                $connection.connection.close()
            }catch{
                write-warning "Error closing"
            }
            try{
                $connection.connection.dispose()
            }catch{
                write-warning "Error disposing"
            }

            if ($script:sqlserverdefaultconnectionname -eq $connection.name){
                $script:sqlserverdefaultconnectionname = $null
            }

            $script:sqlserverconnections.remove($connection)
        }
    }
}

function Add-SqlServerQueryHubConnection{

    [CmdletBinding(SupportsShouldProcess=$true)]
    [CmdletBinding(DefaultParameterSetName = "IntegratedSecurity")] 
    param(
    )

    dynamicparam {
        $dynamicparams = 
            newdynamicparameter -mandatory -parametersetname ('IntegratedSecurity','SqlLogin') -position 0 -Type string -Name Database |
            newdynamicparameter -mandatory -parametersetname ('IntegratedSecurity','SqlLogin') -position 1 -Type string -Name Host |
            newdynamicparameter -parametersetname ('IntegratedSecurity','SqlLogin') -position 2 -Type string -Name ConnectionName |
            newdynamicparameter -parametersetname 'SqlLogin' -position 3 -type switch -Name SqlLogin |
            newdynamicparameter -mandatory -parametersetname 'SqlLogin' -position 4 -Name Credentials -validatescript { $_ -is [pscredential] -and $_ -ne $null} |
            newdynamicparameter -parametersetname ('IntegratedSecurity','SqlLogin') -position 5 -Type int -Name Port -defaultvalue 1433 
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        if ($sqllogin -and -not $credentials){
            $credentials = getconsolecredential
        }
        
        if (-not $connectionname -or $connectionname -eq ""){
            $connectionname = $database
        }

        if ($port){
            $port = ",$port"
        }

        if ($script:sqlserverconnections.name -contains $connectionname){
            throw 'Connection already exists'
        }

        $connection = [pscustomobject] @{
            Name = $connectionname
            Connection = [System.Data.SqlClient.SqlConnection]::new()
            SqlServerPipeline = $null
            PrintStatements = [system.text.stringbuilder]::new()
        }

        if (-not $sqllogin){
            $connection.connection.connectionstring = "Server=$host$port;Database=$database;Integrated Security=True"
        }else{
            $connection.connection.connectionstring = "Server=$host$port;Database=$database;User Id=$($credentials.username);Password=$($credentials.getnetworkcredential().password)"
        }

        try{
            $connection.connection.open()
        }catch{
            return
        }

        $connection.SqlServerPipeline = [powershell]::Create()
        $connection.SqlServerPipeline | Add-Member -MemberType NoteProperty -Name 'AsyncResult' -Value $null

        [void]$script:sqlserverconnections.add($connection)



        if (-not $script:SqlServerdefaultconnectionname){
            Set-SqlServerDefaultQueryHubConnection $connectionName
        }

        $connection.connection.add_infomessage([System.Data.SqlClient.SqlInfoMessageEventHandler]{
            param(
                $src,
                $e
            )

            $connection = [system.data.sqlclient.sqlconnection]$src
            $connectionObject = $script:sqlserverconnections | where {$connection -eq $connection}
            [void]$connectionobject.PrintStatements.appendline($e.message)
        })

        $connection.connection.fireinfomessageeventonusererrors = $true

    }
    end{
    }

}

function Get-SqlServerQueryHubConnections{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    $script:sqlserverconnections | select -expand name
}



function Set-SqlServerDefaultQueryHubConnection{
    [CmdletBinding(SupportsShouldProcess=$true)]

    param()

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -mandatory -position 0 -valuefrompipeline -Type string -Name ConnectionName -validateset (Get-SqlServerQueryHubConnections)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        $script:sqlserverdefaultconnectionname = $connectionName
    }
    end {}



}

function Get-SqlServerDefaultQueryHubConnection{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    $script:SqlServerdefaultconnectionname 

}

function Get-SqlServerScripts{
#    [CmdletBinding(SupportsShouldProcess=$true)]
#    param(
#    )
    begin{
    }
    process{
        get-childitem $script:SqlServerscriptdirectory | where {-not $_.psiscontainer} # | select -expand name
    }
    end{
    }
}

function Invoke-SqlServerScript{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -mandatory -position 0 -Type string -Name ScriptName -validateset (get-childitem -path $script:SqlServerscriptdirectory -ea 0 | where {-not $_.psiscontainer}| select -expand name) |
            newdynamicparameter -position 1 -Type hashtable -Name BindReplacements |
            newdynamicparameter -position 2 -Type hashtable -Name StringReplacements |
            newdynamicparameter -position 3 -Type int -Name First |
            newdynamicparameter -position 4 -Type string -Name ConnectionName -validateset (Get-SqlServerQueryHubConnections) -defaultvalue (Get-SqlServerDefaultQueryHubConnection) |
            newdynamicparameter -position 5 -Type switch -Name View |
            newdynamicparameter -position 6 -Type switch -Name GetPath |
            newdynamicparameter -position 7 -Type switch -Name Edit |
            newdynamicparameter -position 8 -Type switch -Name NoOutput 

        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{

        setdynamicparametervariables $psboundparameters

        if ($view){
            get-content "$($script:SqlServerscriptdirectory)\$scriptName" -raw
        }elseif ($edit){
            if ($TextEditorCommand){
               [scriptblock]::create("$texteditorCommand $($script:SqlServerscriptdirectory)\$scriptName").invoke()
            }else{
                throw "Text editor command not set"
            }
        }elseif ($getpath){
            "$($script:SqlServerscriptdirectory)\$scriptName"
        }else{
            get-content "$($script:SqlServerscriptdirectory)\$scriptName" -raw | invoke-SqlServerquery -stringreplacements $stringreplacements -bindreplacements $bindreplacements -first $first -connectionname $connectionname -nooutput:($nooutput)
        }
    }

    end{}

    
}

function Get-SqlServerQueryHubReader{
   param(
       [system.data.sqlclient.sqlcommand]$Command,
       [ref]$RecordsAffected
   )

    $reader = $command.executereader()
    $recordsAffected.value = $reader.recordsaffected
#    $reader
#    $reader.dispose()

    if (-not $reader.hasrows){
        $reader.dispose()
        return
    }

    $fieldnames = @()
    $fieldtypes = @()

    for ($x = 0; $x -lt $reader.fieldcount;$x++){

        $name = $reader.GetName($x)
        if (-not $name){
            $name = '[unnamed]'
        }
        $fieldname = $name

        Try{
            $fieldType = $reader.GetFieldType($x)
        }catch{
            $fieldType = "asdf".GetType()
            $fieldName = $Name + "[error]"
        }

        $duplicatecount = 0

        while ($true){
            if ($fieldnames.contains($fieldname)){
                $duplicatecount++
                $fieldname = $name + '['+$duplicatecount+']'
            }else{
                break
            }
        }

        $fieldnames += $fieldname
        $fieldtypes += $fieldtype
    }

    $table = @()

    while ($reader.read()){

        $values = new-object object[]  $reader.fieldcount
        $readfieldcount = $reader.getproviderspecificvalues($values)
        $recordhash = New-Object System.Collections.Specialized.OrderedDictionary

        for ($fieldindex = 0; $fieldindex -lt $fieldnames.count; $fieldindex++){
            try{
                $recordhash.add($fieldnames[$fieldindex],$values[$fieldindex])
            }catch{
                $recordhash.add($fieldnames[$fieldindex],$_.exception)
            }
        }

        $newrecord = [pscustomobject] $recordhash 
        $table += $newrecord
    }

    #override the tostring so that null values are hidden (instead of showing null)
    foreach ($prop in $table|gm|where {$_.membertype -eq 'noteproperty'}){
        $table."$($prop.name)" | add-member -membertype scriptmethod -name tostring -value {if ($this.isnull){""}else{$this.value.tostring()}} -force
    }
    $table
        $reader.dispose()


}

function Invoke-SqlServerQuery{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -position 0 -mandatory -valuefrompipeline -Type string -Name Script | 
            newdynamicparameter -position 1 -Type hashtable -Name BindReplacements | 
            newdynamicparameter -position 2 -Type hashtable -Name StringReplacements | 
            newdynamicparameter -position 3 -Type int -Name First |
            newdynamicparameter -position 4 -Type string -Name ConnectionName -validateset (Get-SqlServerQueryHubConnections) -defaultvalue (Get-SqlServerDefaultQueryHubConnection) |
            newdynamicparameter -position 5 -Type switch -Name Count |
#            newdynamicparameter -position 6 -Type switch -Name KeepSemicolonEnding |
            newdynamicparameter -position 7 -Type switch -Name SkipStringReplacements |
            newdynamicparameter -position 8 -Type switch -Name NoSaveResults |
            newdynamicparameter -position 9 -Type switch -Name Quiet |
            newdynamicparameter -position 10 -Type switch -Name NoOutput
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        $connectionobject = $script:SqlServerconnections | where {$_.name -eq $connectionname} 

        if ($connectionobject.SqlServerpipeline.asyncresult -and -not $connectionobject.SqlServerpipeline.asyncresult.iscompleted){
            throw ("connection is busy")
        }

        $script = $script.trim()

        $queries = $script -split '(?m)^\s{0,}/\s{0,}$'


        foreach ($query in $queries){
            if ($query.trim() -match '^\s{0,}$'){
                continue
            }
            $query += "/"

            $errored = $null

            $bindVariables = @()
            $bindlines = 0

            while ($true){
                $matched = $query | select-string '(?m)^ {0,}variable [^;]*;'  -allmatches | select -expand matches 

                if (-not $matched){
                    break;
                }

                if ($details = ($matched[0] | select -expand value) -match '(?<=^ {0,})variable {1,}(?<bindName>[^ ;]+) (?<datatype>[^ ;]{1,})( {1,}default {1,}(''{0,1})(?<defaultValue>[^'']{1,})''{0,1}){0,1};'){
                    $bindvariable = new-object PSObject
                    $bindvariable | add-member -membertype NoteProperty -name "Name" -value $matches.bindName
                    $bindvariable | add-member -membertype NoteProperty -name "DataType" -value $matches.datatype

                    if ($bindReplacements -and (":$($matches.bindName)" -in $bindreplacements.keys)){
                        $bindvariable | add-member -membertype NoteProperty -name "DataValue" -value $bindReplacements[":$($matches.bindName)"]
                    }elseif ($matches.datatype -eq "refcursor"){
                        $bindvariable | add-member -membertype NoteProperty -name "DataValue" -value $null
                    }else{
                        $bindvariable | add-member -membertype NoteProperty -name "DataValue" -value (replacevariable ":$($matches.bindname)" $matches.defaultvalue).value
                    }

                    $bindVariables += $bindvariable

                }

                $query = $query.substring(0,$matched[0].index) + '--' + $matched[0].value + $query.substring($matched[0].index + $matched[0].length)
            }

            if (-not $SkipStringReplacements) {

                if ($StringReplacements){
                    foreach ($repl in $StringReplacements.keys){
                        $query = $query.replace($repl,$stringreplacements.item($repl))
                    }
                }
                

                $skipIndex = -1

                while ($true){
                    $matched = $query | select-string '&[^ ''&]+'  -allmatches | select -expand matches | where index -gt $skipIndex

                    if (-not $matched){
                        break;
                    }

                    $skipIndex = $matched[0].index
                    $varname = $matched[0] | select -expand value;

                    if ($varname -match '(?<=\[)(?<defaultVar>[^)]+)(?=\])'){
                        $varDefault = $matches.defaultVar
                        $varName = $varName.replace("[$vardefault]","")
                    }else{
                        $vardefault = $null
                    }

                    $var = replacevariable $varname $varDefault
                    
                    if ($var.iscancelled){
    #                    continue
                        break;
                    }

                    $replacementText = $var.value
                    $replacementText = "$([regex]::split($replacementText, '(.{10})') | where {$_} | join -joinon "'||'")"

                    $skipIndex = $skipIndex + $var.value.length

                    $query = $query.replace(($matched[0] | select -expand value),$replacementText)
                }
            }


            $query = $query.trim()

            if ($query.endswith("/")){
                $query = $query.substring(0,$query.length-1)
            }

#            if (-not $keepsemicolonending -and $query.endswith(";")){
#                $query = $query.substring(0,$query.length-1)
#            }

            if ($First){
                $lines = $query -split "`n"
                $newlines = @() 
                foreach ($line in $lines){
                    $newlines += $line -replace '--.*'
                }

                $query = $newlines -join "`n"

                $selectPosition = $query.tolower().indexof("select")
                if ($selectPosition -ge 0){
                    $query = $query.substring(0,$selectPosition) + "select top $First" + $query.substring($selectPosition + "select".length)
                }

            }

             if ($count){
                $query = "select count(1) [count(1)] from (`n$query`n) count_x_"
             }


            try{

                $command = [System.Data.SqlClient.SqlCommand]::new($query)
                $command.connection = $connectionobject.connection
                $command.commandtype = [system.data.commandtype]::text
                $command.commandtimeout = 60 * $script:timeoutminutes
#                $command.bindbyname = $true

                $cancelled = $false

                foreach ($bindvariable in $bindvariables){
                    $param = [System.Data.SqlClient.SqlParameter]::new()
                    $param.parametername = $bindvariable.name

                    switch -regex ($bindvariable.datatype) {
                        "^number$"{
                            $param.dbtype = [System.Data.SqlDbType]::decimal
                            $param.direction = [system.data.parameterdirection]::inputoutput
                            $param.value = $bindvariable.datavalue
                            break
                        }"^(varchar2|varchar|clob)$"{
                            $param.dbtype = [System.Data.SqlDbType]::varchar
                            $param.direction = [system.data.parameterdirection]::inputoutput
                            $param.value = $bindvariable.datavalue
                            break
                        }"^(date|datetime)$"{
                            $param.dbtype = [System.Data.SqlDbType]::datetime
                            $param.direction = [system.data.parameterdirection]::inputoutput
                            $param.value = $bindvariable.datavalue
                            break
                        }default{
                            $param.dbtype = [System.Data.SqlDbType]::variant
                            $param.direction = [system.data.parameterdirection]::inputoutput
                            $param.value = $bindvariable.datavalue
                            break
                        }
                    }

                    [void]$command.parameters.add($param)
                }

                try{

                    try{
                        $recordsaffected = $null
                        [void]$connectionobject.printstatements.clear()
                        [void]$connectionobject.SqlServerPipeline.commands.clear()
                        [void]$connectionobject.SqlServerPipeline.streams.clearstreams()
                        [void]$connectionobject.SqlServerPipeline.Addcommand('Get-SqlServerQueryHubReader')
                        [void]$connectionobject.SqlServerPipeline.addparameter("Command",$command)
                        [void]$connectionobject.SqlServerPipeline.addparameter("RecordsAffected",[ref]$recordsaffected)

                        $connectionobject.SqlServerPipeline.AsyncResult = $connectionobject.SqlServerPipeline.BeginInvoke()

                        $queryStart = get-date

                        $completed = showasyncwaitmessage -message $command.commandtext -conditionvariablename "connectionobject" -conditionmethodpropertychain @('sqlServerPipeline','asyncresult','iscompleted') 

                        if (-not $completed){
                            $cancelled = $true
                            $command.cancel()
                            throw "Query was Cancelled"
                        }

                        $queryend = (new-timespan $querystart (get-date))

                        $readerresult =  $connectionobject.SqlServerPipeline.EndInvoke($connectionobject.SqlServerPipeline.AsyncResult)

                        if ($connectionobject.SqlServerpipeline.streams.error){
                            throw $connectionobject.SqlServerpipeline.streams.error
                        }


                         #cleanup code.
                    }catch{
                        $errored = $_.exception.message | select-string '(?ms)(?<=Exception calling "(ExecuteReader|ExecuteNonQuery)" with "0" argument\(s\): ").*(?=" You cannot call a method on a null-valued expression.$)'
                        if ($errored){
                            $errored = $errored | select -expand matches | select -expand value
                        }else{
                            $errored = $_.exception.message
                        }
                    }finally{
                        [console]::TreatControlCAsInput = $false
                    }

                    $table = $readerresult

                    if (-not $cancelled -and -not $nosaveresults){
                        $printstatements = get-sqlserverprintstatements $connectionname
                    }

                    if ($table.count -eq 0){
                        if ($errored){
                            $table = [pscustomobject] @{Error=$errored}
                        }elseif($recordsaffected -ge 0){
                            $table = [pscustomobject] @{RecordsAffected=$recordsaffected}
                        }elseif ($printstatements){
                            $table = [pscustomobject] @{PrintStatements=$printstatements}
                        }else{
                            $null
                        }
                    }

                    if (-not $nosaveresults -and -not $cancelled){
                        $script:lastSqlServerqueryrequest = $query
                        $script:lastSqlServerqueryresult = $table
#                        $script:lastSqlServerqueryexplain = get-SqlServerexplainplan -connectionname $connectionname
                        $script:lastSqlServerquerytime = $queryend
                        $script:lastSqlServerPrintStatements = $printstatements
                        $script:lastoraclequeryerror = $errored


                        if (get-typedata -typename QueryHub.SLast){
                            remove-typedata -typename QueryHub.SLast -confirm:$false
                        }

                        if ($script:lastSqlServerqueryresult){
                            foreach ($prop in ($script:lastSqlServerqueryresult|gm|where {$_.membertype -eq 'noteproperty'})){
                                update-typedata -typename QueryHub.SLast -membertype NoteProperty -membername $prop.name -value $null -confirm:$false
                            }
                        }
                    }

                    if (-not $nooutput){
                        $table
                    }
                }finally{
                }
                
            }finally{
                if ($command){
                    $command.dispose()
                }
            }

        }
    }

    end{
    }
}

function Get-SqlServerPrintStatements{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -position 0 -Type string -Name ConnectionName -validateset (Get-SqlServerQueryHubConnections) -defaultvalue (Get-SqlServerDefaultQueryHubConnection)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        $connectionobject = $script:sqlserverconnections | where {$_.name -eq $connectionname}
        $results = $connectionobject.printstatements.tostring() -replace "$([environment]::newline)$"
        [void]$connectionobject.printstatements.clear()

        $results

    }
    end{
    }
}

function Invoke-SqlServerQueryFromFile{
    [CmdletBinding()] 
    param( 
    )

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -mandatory -position 0 -valuefrompipeline -Name File -validatescript {test-path $_ -PathType Leaf} |
            newdynamicparameter -position 1 -Type string -Name ConnectionName -validateset (Get-SqlServerQueryHubConnections) -defaultvalue (Get-SqlServerDefaultQueryHubConnection) 
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        cat $file -raw | invoke-SqlServerquery -connectionname $connectionName 
    }

    end{
    }
}


function Get-SqlServerTableNames{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -position 0 -Type string -Name TableLike -defaultvalue "%" | 
            newdynamicparameter -position 1 -Type string -Name ConnectionName -validateset (Get-SqlServerQueryHubConnections) -defaultvalue (Get-SqlServerDefaultQueryHubConnection)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        $query = "select distinct table_schema + '.' + table_name table_name from information_schema.tables where table_name like '$($tablelike.toupper())'  order by table_name asc" 

        $result = $query | invoke-SqlServerquery -connectionname $connectionname
        $result


    }

    end{}
}

function Get-SqlServerTableSchema{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -position 0 -valuefrompipeline -Type string -Name Table | 
            newdynamicparameter -position 1 -Type string -Name Column | 
            newdynamicparameter -position 2 -Type string -Name ConnectionName -validateset (Get-SqlServerQueryHubConnections) -defaultvalue (Get-SqlServerDefaultQueryHubConnection) |
            newdynamicparameter -position 3 -Type switch -Name TableIsLike |
            newdynamicparameter -position 4 -Type switch -Name ColumnIsLike
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        $tableQuery = "1=1"
        $columnQuery = "1=1"

        if ($table) {
            if ($tableislike){
                $tableQuery = "table_name like '$($table.toupper())' "
            }else{
                $tableQuery = "table_name = '$($table.toupper())'" 
            }
        }

        if ($column) {
            if ($columnislike){
                $columnQuery = "column_name like '$($column.toupper())' "
            }else{
                $columnQuery = "column_name = '$($column.toupper())'" 
            }
        }

        $result = "select * from information_schema.columns where $tablequery and $columnQuery order by table_name asc, column_name asc"  | invoke-SqlServerquery -connectionname $connectionname

        $result | select table_schema,table_name,column_name,data_type,character_maximum_length,numeric_precision,numeric_scale,is_nullable 
    }

    end{}
}

function Get-SqlServerCodeNames{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -position 0 -valuefrompipeline -Type string -Name Name | 
            newdynamicparameter -position 1 -Type string -Name Text | 
            newdynamicparameter -position 2 -Type string -Name ConnectionName -validateset (Get-SqlServerQueryHubConnections) -defaultvalue (Get-SqlServerDefaultQueryHubConnection) |
            newdynamicparameter -position 3 -Type switch -Name NameIsLike |
            newdynamicparameter -position 4 -Type switch -Name TextIsLike
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters


        $nameQuery = "1=1"
        $textQuery = "1=1"

        if ($Name) {
            if ($nameislike){
                $nameQuery = "lower(routine_name) like '$($name.tolower())'"
            }else{
                $nameQuery = "lower(routine_name) = '$($name.tolower())'" 
            }
        }

        if ($text) {
            if ($textislike){
                $textQuery = "lower(routine_definition) like '$($text.tolower())'"
            }else{
                $textQuery = "lower(routine_definition) = '$($text.tolower())'" 
            }
        }

     

        $result = "select routine_schema, routine_name, routine_type from information_schema.routines where $namequery and $textquery order by routine_name,routine_type asc, routine_schema asc"   | invoke-SqlServerquery -connectionname $connectionname

        $result 
    }

    end{}
}

function Get-SqlServerCodeDefinition{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -position 0 -mandatory -valuefrompipeline -Type string -Name Name | 
            newdynamicparameter -position 1 -Type string -Name Type -validateset @('Function','Procedure') -defaultvalue 'Function' | 
            newdynamicparameter -position 2 -Type string -Name ConnectionName -validateset (Get-SqlServerQueryHubConnections) -defaultvalue (Get-SqlServerDefaultQueryHubConnection)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        $typestring = $type


        $result = "select routine_definition from information_schema.routines where routine_type = '$($typestring.toupper())' and lower(routine_name) = '$($name.tolower())' " | invoke-SqlServerquery -connectionname $connectionname 
        
        $result | select -expand routine_definition | select -expand value #| % {$_ -replace "`n","`r"}
    }

    end{}

}

function Test-SqlServerConnection{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -position 0 -valuefrompipeline -Type string -Name ConnectionName -validateset (Get-SqlServerQueryHubConnections) -defaultvalue (Get-SqlServerDefaultQueryHubConnection)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        try{
            $result = "select 'asdf'" | invoke-SqlServerquery -connectionname $connectionName
        }catch{
            if ($_.exception.message -match 'connection is busy'){
            }else{
                $busy = $true
            }
        }

       
        $new = new-object PSCustomObject
        $new | add-member -membertype NoteProperty -name "ConnectionName" -value $connectionName

        if ($busy){
            $new | add-member -membertype NoteProperty -name "Result" -value "Busy"
        } elseif (-not $result -or ($result | gm | where {$_.name -match '^(error)$'})){
            $new | add-member -membertype NoteProperty -name "Result" -value $false
        } else {
            $new | add-member -membertype NoteProperty -name "Result" -value $true
        }

        $new
        return
 
    }

    end{}
}

function Get-SqlServerSession{

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
    )

    dynamicparam{
        $dynamicparams = 
            newdynamicparameter -position 0 -Type string -Name ConnectionName -validateset (Get-SqlServerQueryHubConnections) -defaultvalue (Get-SqlServerDefaultQueryHubConnection)
        
        $dynamicparams.parameterdictionary
        $psboundparameters = SetDynamicParameterValues $psboundparameters $dynamicparams
    }

    begin{
    }

    process{
        setdynamicparametervariables $psboundparameters

        $result = "sp_who" | invoke-SqlServerquery -connectionname $connectionName

        $result | % {$_.cmd = $_.cmd.tostring().trim(); $_.dbname = $_.dbname.tostring().trim(); $_.hostname = $_.hostname.tostring().trim(); $_.loginame = $_.loginame.tostring().trim();  $_.status = $_.status.tostring().trim();}

        $result
    }

    end{}
}


#TODO: use smo library?
#
#function Get-SqlServerDefinition{
#
#    [CmdletBinding(SupportsShouldProcess=$true)]
#    param(
#        [parameter(position=0,mandatory=$true,ValueFromPipeline=$true)] [string]$Name,
#        [parameter(position=1,mandatory=$true,ValueFromPipeline=$true)] [ValidateSet('AQ_QUEUE','REF_CONSTRAINT','AQ_QUEUE_TABLE','REFRESH_GROUP','AQ_TRANSFORM','RESOURCE_COST','ASSOCIATION','RLS_CONTEXT','AUDIT','RLS_GROUP','AUDIT_OBJ','RLS_POLICY','CLUSTER','RMGR_CONSUMER_GROUP','COMMENT','RMGR_INTITIAL_CONSUMER_GROUP','CONSTRAINT','RMGR_PLAN','CONTEXT','RMGR_PLAN_DIRECTIVE','DATABASE_EXPORT','ROLE','DB_LINK','ROLE_GRANT','DEFAULT_ROLE','ROLLBACK_SEGMENT','DIMENSION','SCHEMA_EXPORT','DIRECTORY','SEQUENCE','FGA_POLICY','SYNONYM','FUNCTION','SYSTEM_GRANT','INDEX_STATISTICS','TABLE','INDEX','TABLE_DATA','INDEXTYPE','TABLE_EXPORT','JAVA_SOURCE','TABLE_STATISTICS','JOB','TABLESPACE','LIBRARY','TABLESPACE_QUOTA','MATERIALIZED_VIEW','TRANSPORTABLE_EXPORT','MATERIALIZED_VIEW_LOG','TRIGGER','OBJECT_GRANT','TRUSTED_DB_LINK','OPERATOR','TYPE','PACKAGE','TYPE_BODY','PACKAGE_SPEC','TYPE_SPEC','PACKAGE_BODY','USER','PROCEDURE','VIEW','PROFILE','XMLSCHEMA','PROXY')] [string]$Type,
#        [parameter(position=2,mandatory=$false,ValueFromPipeline=$false)] [validatescript({ $_ -ne $null})] [string]$ConnectionName = $script:SqlServerdefaultConnectionName
#    )
#
#	$result = "select dbms_metadata.get_ddl('$($type.toupper())','$($name.toupper())') asdf from dual" | invoke-SqlServerquery -connectionname $connectionname
#
#	if ($result -and $result.asdf){
#		$result | select -expand asdf
#	} else {
#		$result
#	}
#}






$config = get-childitem "$(split-path -parent $MyInvocation.MyCommand.Definition)\QueryHub.Config" -ea silentlycontinue
if ($config){
    $parameters = import-csv -delim "," $config
    $parameters.psobject.typenames.add("QueryHub.Config")
    $parameters | Set-QueryHubParams
}

Export-ModuleMember -Function  *-*

