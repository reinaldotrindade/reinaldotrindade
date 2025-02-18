<# 
.SYNOPSIS
Script para a instalação remota de patches Microsoft através da interface WMI do agente do SCCM.

.DESCRIPTION 
Facilita a administração das atualizações de segurança através do agente do SCCM de modo remoto.
Com o script é possível consultar os patches pendentes, executar a instalação, reiniciar as máquinas, executar as ações do agente, corrigir alguns erros.

.PARAMETER Server
Especifica o computador remoto

.PARAMETER ServerList
Especifica um arquivo com a lista dos computadores remotos 

.PARAMETER ServerOU
Especifica uma OU no Active Directory com os computadores remotos

.PARAMETER ServerGroup
Especifica um Grupo no Active Directory com os computadores remotos

.PARAMETER Install
Envia o comando para instalar os updates pendentes no computadore remoto

.PARAMETER Reboot
Reinicia o computador remoto caso ele se encontre com boot pendente

.PARAMETER ForceReboot
Reinicia o computador remoto independente do status de boot pendente

.PARAMETER ExecuteAllActions
Envia o comando para execução de todas as ações do SCCM client agent

.PARAMETER ExecuteUpdatesActions
Envia o comando para execução das ações de Software Update do SCCM client agent

.PARAMETER ExportCSV
Especifica o arquivo de saída em formato .csv

.PARAMETER Details
Especifica se serão exibidos as informações em modo detalhado

.PARAMETER FixUpdateError
Tenta reparar o serviço de do SCCM e WU

.PARAMETER CheckHotFix
Verifica se um determinado KB está instalado

.INPUTS 
Não suporta nenhum tipo de input via pipe.

.OUTPUTS 
Tabela com resulta da consulta/ações executadas.

.EXAMPLE
Realiza a consulta dos patches pendentes para os servidores ctx8001 e ctx8002
SCCM-Updates.ps1 -Server ctx8001,ctx8002

.EXAMPLE
Realiza a consulta dos patches pendentes para os servidores ctx8001 e ctx8002 com detalhes de cada patche pendente
SCCM-Updates.ps1 -Server ctx8001,ctx8002 -Details

.EXAMPLE
Realiza a consulta dos patches pendentes para os servidores listas no arquivo servidores.txt (um por linha)
SCCM-Updates.ps1 -ServerList C:\Scripts\Lista\servidores.txt

.EXAMPLE
Executa a instalação dos patches pendentes para os servidores ctx8001 e ctx8002
SCCM-Updates.ps1 -Server ctx8001,ctx8002 -Install

.EXAMPLE
Executa a instalação dos patches pendentes para os servidores ctx8001 e ctx8002 e 
reinicia os servidores casa estejam com status "REBOOT PENDENTE" no momento da execuação
SCCM-Updates.ps1 -Server ctx8001,ctx8002 -Install -Reboot

.EXAMPLE
Executa a todas as ações do agente do SCCM para os servidores ctx8001 e ctx8002
SCCM-Updates.ps1 -Server ctx8001,ctx8002 -ExecuteAllActions

.EXAMPLE
Verifica se o patche KB4048957 está instalado nos servidores ctx8001 e ctx8002
SCCM-Updates.ps1 -Server ctx8001,ctx8002 -CheckHotFix KB4048957

.EXAMPLE
Exporta os resultados exibidos para o arquivo relatorio.csv em formato CSV
SCCM-Updates.ps1 -Server ctx8001,ctx8002 -ExportCSV c:\temp\relatorio.csv

#>

[CmdletBinding(DefaultParameterSetName="Query")]

param(
    [parameter(Mandatory=$false,ValueFromPipeline=$false,Position=1)][ValidateNotNullOrEmpty()]
        [string[]]$Server,
    
    [parameter(Mandatory=$false,ValueFromPipeline=$false)][ValidateNotNullOrEmpty()][ValidateScript({Test-Path $_})]
		[string]$ServerList,

    [parameter(Mandatory=$false,ValueFromPipeline=$false)][ValidateNotNullOrEmpty()]
		[string]$ServerOU,
    
    [parameter(Mandatory=$false,ValueFromPipeline=$false)][ValidateNotNullOrEmpty()]
		[string]$ServerGroup,

    [parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName = "Query")][ValidateNotNullOrEmpty()]    
        [switch]$Details,

    [parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName = "Install")][ValidateNotNullOrEmpty()]
		[switch]$Install,
    
    [parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName = "Install")][ValidateNotNullOrEmpty()]
        [switch]$Reboot,

    [parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName = "Install")][ValidateNotNullOrEmpty()]
    [parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName = "Query")][ValidateNotNullOrEmpty()]
        [switch]$FixUpdateError,

    [parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName = "ForceReboot")][ValidateNotNullOrEmpty()]  
        [switch]$ForceReboot,
    
    [parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName = "ActionsAll")][ValidateNotNullOrEmpty()]
        [switch]$ExecuteAllActions,
    
    [parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName = "ActionsUpdates")][ValidateNotNullOrEmpty()]
        [switch]$ExecuteUpdatesActions,

    [parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName = "ActionsHwInventory")][ValidateNotNullOrEmpty()]
        [switch]$ExecuteHwInventoryAction,

    [parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName = "ActionsFullHwInventory")][ValidateNotNullOrEmpty()]
        [switch]$ExecuteFullHwInventoryAction,
    
    [parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName = "Check")][ValidateNotNullOrEmpty()]
        [string]$CheckHotFix,

    [parameter(Mandatory=$false,ValueFromPipeline=$false)][ValidateNotNullOrEmpty()]
        [string]$ExportCSV
)

begin {

    $version = "1.09"
    $computers = @()
    $report = @()
    $k = 0
    $linhaTabela = "--------------------------------------------------------------------"

    if($Server){    
        foreach($s in $Server){ 
            $computers += $s.Trim() 
        }
    }
    if($ServerList){
        $computers += (Get-Content $ServerList).Trim()
    }
    if($ServerOU){
        $computers += @(Get-ADComputer -SearchBase $ServerOU -Filter * | select name).name
    }
    if($ServerGroup){
        $computers += @(Get-ADGroupMember -Identity $ServerGroup | Where-Object objectClass -eq "computer" | select name).name
    }
    if($computers.Count -eq 0){
       $computers += "localhost"
    }


    #Definição das constantes --------------------------------------------------------------------------------------------------------
    Set-Variable STATUS_SUCCESS_INSTALL -Option Constant -Value "INSTALANDO"
    Set-Variable STATUS_SUCCESS_ACTIONS -Option Constant -Value "ACTIONS OK"
    Set-Variable STATUS_FAIL_INSTALL -Option Constant -Value "FALHA AO INSTALAR"
    Set-Variable STATUS_FAIL_ACTIONS -Option Constant -Value "FALHA AO EXECUTAR"
    Set-Variable STATUS_NO_UPDATES -Option Constant -Value "SEM UPDATES"
    Set-Variable STATUS_OFFILINE -Option Constant -Value "OFFLINE"
    Set-Variable STATUS_FAIL_GET_UPDATES -Option Constant -Value "FALHA AO CONSULTAR"
    Set-Variable STATUS_PENDING_UPDATES -Option Constant -Value "UPDATES PENDENTES"
    Set-Variable STATUS_PENDING_REBOOT -Option Constant -Value "REBOOT PENDENTE"
    Set-Variable STATUS_UPDATES_INSTALADOS -Option Constant -Value "UPDATES INSTALADOS"
    Set-Variable STATUS_UPDATES_ERROR -Option Constant -Value "UPDATE ERROR"
    Set-Variable STATUS_UPDATES_ERROR_FIX -Option Constant -Value "FIX UPDATE ERROR"
    Set-Variable CHECK_HOTFIX_OK -Option Constant -Value "INSTALADO"
    Set-Variable CHECK_HOTFIX_NOT -Option Constant -Value " --------- "
    Set-Variable REBOOT_OK -Option Constant -Value "OK"
    Set-Variable REBOOT_FAIL -Option Constant -Value "FALHA"

    Set-Variable PADNum -Option Constant -Value 3 
    Set-Variable PADComputer -Option Constant -Value 18
    Set-Variable PADUpdates -Option Constant -Value 10
    Set-Variable PADStatus -Option Constant -Value 20
    Set-Variable PADReboot -Option Constant -Value 7


    #Criando Hash Table das Actions do SCCM ----------------------------------------------------------------------------------------------
    $Actions = @{}
    $Actions.Add("Application Deployment Evaluation Cycle", "{00000000-0000-0000-0000-000000000121}")
    $Actions.Add("Discovery Data Collection Cycle", "{00000000-0000-0000-0000-000000000003}")
    $Actions.Add("File Collection Cycle", "{00000000-0000-0000-0000-000000000010}")
    $Actions.Add("Hardware Inventory Cycle", "{00000000-0000-0000-0000-000000000001}")
    $Actions.Add("Machine Policy Retrieval Cycle", "{00000000-0000-0000-0000-000000000021}")
    $Actions.Add("Machine Policy Evaluation Cycle", "{00000000-0000-0000-0000-000000000022}")
    $Actions.Add("Software Inventory Cycle", "{00000000-0000-0000-0000-000000000002}")
    $Actions.Add("Software Metering Usage Report Cycle", "{00000000-0000-0000-0000-000000000031}")
    $Actions.Add("Software Update Deployment Evaluation Cycle", "{00000000-0000-0000-0000-000000000108}")
    $Actions.Add("Software Update Sotre", "{00000000-0000-0000-0000-000000000114}")
    $Actions.Add("Software Update Scan Cycle", "{00000000-0000-0000-0000-000000000113}")
    $Actions.Add("State Message Refresh", "{00000000-0000-0000-0000-000000000111}")
    #$Actions.Add("User Policy Retrieval Cycle", "{00000000-0000-0000-0000-000000000026}")
    #$Actions.Add("User Policy Evaluation Cycle", "{00000000-0000-0000-0000-000000000027}")
    $Actions.Add("Windows Installers Source List Update Cycle", "{00000000-0000-0000-0000-000000000032}")

    #Criando Hash Table dos EvalutionState Codes do SCCM ----------------------------------------------------------------------------------------------
    $Evaluation = @{}
    $Evaluation.Add(0,"JobStateNone")
    $Evaluation.Add(1,"JobStateAvailable")
    $Evaluation.Add(2,"JobStateSubmitted")
    $Evaluation.Add(3,"JobStateDetecting")
    $Evaluation.Add(4,"JobStatePreDownload")
    $Evaluation.Add(5,"JobStateDownloading")
    $Evaluation.Add(6,"JobStateWaitInstall")
    $Evaluation.Add(7,"JobStateInstalling")
    $Evaluation.Add(8,"JobStatePendingSoftReboot")
    $Evaluation.Add(9,"JobStatePendingHardReboot")
    $Evaluation.Add(10,"JobStateWaitReboot")
    $Evaluation.Add(11,"JobStateVerifying")
    $Evaluation.Add(12,"JobStateInstallComplete")
    $Evaluation.Add(13,"JobStateError")
    $Evaluation.Add(14,"JobStateWaitServiceWindow")
    $Evaluation.Add(15,"JobStateWaitUserLogon")
    $Evaluation.Add(16,"JobStateWaitUserLogoff")
    $Evaluation.Add(17,"JobStateWaitJobUserLogon")
    $Evaluation.Add(18,"JobStateWaitUserReconnect")
    $Evaluation.Add(19,"JobStatePendingUserLogoff")
    $Evaluation.Add(20,"JobStatePendingUpdate")
    $Evaluation.Add(21,"JobStateWaitingRetry")
    $Evaluation.Add(22,"JobStateWaitPresModeOff")
    $Evaluation.Add(23,"JobStateWaitForOrchestration")


    #Declarção das funções ----------------------------------------------------------------------------------------------

    Function CreateRow(){
        return $obj = "" | select Num,Computer,Updates,Status,Reboot
    }

    Function PrintRow(){
        Write-Host -NoNewline "|"
        Write-Host -NoNewline $row.Num.PadRight($PADNum-1)
        Write-Host -NoNewline "| "
        Write-Host -NoNewline $row.Computer.PadRight($PADComputer)
        Write-Host -NoNewline "| "
        Write-Host -NoNewline $row.Updates.PadRight($PADUpdates)
        Write-Host -NoNewline "| " 

        if($row.Status -eq $STATUS_SUCCESS_INSTALL -or $row.Status -eq $STATUS_SUCCESS_ACTIONS -or $row.Status -eq $REBOOT_OK -or $row.Status -eq $STATUS_UPDATES_ERROR_FIX -or $row.Status -eq $CHECK_HOTFIX_OK){
    
            Write-Host -NoNewline $row.Status.PadRight($PADStatus) -ForegroundColor Green
    
        }elseif($row.Status -eq $STATUS_FAIL_INSTALL -or $row.Status -eq $STATUS_FAIL_ACTIONS -or $row.Status -eq $STATUS_FAIL_GET_UPDATES -or $row.Status -eq $STATUS_OFFILINE -or $row.Status -eq $STATUS_UPDATES_ERROR){
    
            Write-Host -NoNewline $row.Status.PadRight($PADStatus) -ForegroundColor Red
    
        }elseif($row.Status -eq $STATUS_PENDING_UPDATES -or $row.Status -eq $STATUS_PENDING_REBOOT -or $Row.Status -eq $CHECK_HOTFIX_NOT){
    
            Write-Host -NoNewline $row.Status.PadRight($PADStatus) -ForegroundColor Yellow
    
        }else{
    
            Write-Host -NoNewline $row.Status.PadRight($PADStatus)
        }
    
        Write-Host -NoNewline "| "
    
        if($row.Reboot -eq $REBOOT_OK){
        
            Write-Host -NoNewline $row.Reboot.PadRight($PADReboot) -ForegroundColor Green
    
        }elseif($row.Reboot -eq $REBOOT_FAIL){
    
            Write-Host -NoNewline $row.Reboot.PadRight($PADReboot) -ForegroundColor Red
    
        }else{
    
            Write-Host -NoNewline $row.Reboot.PadRight($PADReboot)
    
        }
    
        Write-Host "|"
    }

    Function GetUpdates($ComputerName){
        try{
            $Script:UpdatesPendentes = @()
            $Script:UpdatesPendentes = @(Get-WmiObject -ComputerName $ComputerName -Class CCM_SoftwareUpdate -Filter ComplianceState=0 -Namespace root\CCM\ClientSDK -ErrorAction SilentlyContinue| Foreach-Object {[WMI]$_.__PATH})
            return $true
        }catch{
            return $false
        }
    }

    Function GetRebootStatus($ComputerName){
        try{
            $objWMI = Invoke-WmiMethod -ComputerName $ComputerName -Class CCM_ClientUtilities -Namespace root\CCM\ClientSDK -Name DetermineIfRebootPending -ErrorAction SilentlyContinue
            return $objWMI.RebootPending
        }catch{
            return $false
        }
    }

    Function InstallUpdates($ComputerName,$Updates){
        try{
            $objWMI = Invoke-WmiMethod -ComputerName $ComputerName -Class CCM_SoftwareUpdatesManager -Name InstallUpdates -ArgumentList (,$Updates) -Namespace root\ccm\ClientSDK -ErrorAction SilentlyContinue
            if($objWMI.ReturnValue -eq 0){
                return $STATUS_SUCCESS_INSTALL     
            }else{
                return $STATUS_FAIL_INSTALL
            }
        }catch{
            return $STATUS_FAIL_INSTALL
        }
    }

    Function RebootComputer($ComputerName, $Force){    
        try{
            if($Force){
                $objWMI = (Get-WMIObject Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction SilentlyContinue).Win32Shutdown(6)
            }else{
                $objWMI = Invoke-WmiMethod -ComputerName $ComputerName -Class CCM_ClientUtilities -Namespace root\CCM\ClientSDK -Name RestartComputer -ErrorAction SilentlyContinue
            }
        
            if($objWMI.ReturnValue -eq 0){
                return $REBOOT_OK
            }else{
                return $REBOOT_FAIL
            }                
        }catch{
            return $REBOOT_FAIL
        }
    }

    Function FixError($ComputerName){    
        try{
            Get-Service -ComputerName $ComputerName -Name wuauserv -ErrorAction SilentlyContinue | Stop-Service -EA SilentlyContinue -WA SilentlyContinue
            Get-Service -ComputerName $ComputerName -Name cryptSvc -ErrorAction SilentlyContinue | Stop-Service -EA SilentlyContinue -WA SilentlyContinue
            Get-Service -ComputerName $ComputerName -Name bits -ErrorAction SilentlyContinue | Stop-Service -EA SilentlyContinue -WA SilentlyContinue
            Get-Service -ComputerName $ComputerName -Name msiserver -ErrorAction SilentlyContinue | Stop-Service -EA SilentlyContinue -WA SilentlyContinue
            Get-Service -ComputerName $ComputerName -Name CcmExec -ErrorAction SilentlyContinue | Stop-Service -EA SilentlyContinue -WA SilentlyContinue
            Remove-Item \\$ComputerName\c$\Windows\SoftwareDistribution -Force -Recurse -ErrorAction SilentlyContinue
            Remove-Item \\$ComputerName\c$\Windows\ccmcache -Force -Recurse -ErrorAction SilentlyContinue
            Get-Service -ComputerName $ComputerName -Name wuauserv -ErrorAction SilentlyContinue | Start-Service -EA SilentlyContinue -WA SilentlyContinue
            Get-Service -ComputerName $ComputerName -Name cryptSvc -ErrorAction SilentlyContinue | Start-Service -EA SilentlyContinue -WA SilentlyContinue
            Get-Service -ComputerName $ComputerName -Name bits -ErrorAction SilentlyContinue | Start-Service -EA SilentlyContinue -WA SilentlyContinue
            Get-Service -ComputerName $ComputerName -Name msiserver -ErrorAction SilentlyContinue | Start-Service -EA SilentlyContinue -WA SilentlyContinue
            Get-Service -ComputerName $ComputerName -Name CcmExec -ErrorAction SilentlyContinue | Start-Service -EA SilentlyContinue -WA SilentlyContinue
            return $true
        }catch{
            return $false
        }
    }

    Function GetHotfix($ComputerName, $KB){    
        try{
            if(Get-HotFix -ComputerName $ComputerName -Id $KB -ErrorAction SilentlyContinue){
                return $CHECK_HOTFIX_OK
            }
            else{
                return $CHECK_HOTFIX_NOT
            }
        }
        catch{
            return $STATUS_OFFILINE
        }
    }

    Function StartAction($computerName,$Action){
        Invoke-WmiMethod -ComputerName $ComputerName -Class SMS_Client -Name TriggerSchedule $Action -Namespace root\ccm -ErrorAction SilentlyContinue
    }

    Function StartAllActions($ComputerName){
    
        try{
            foreach ($ActionValue in $Actions.Values){
                StartAction $ComputerName $ActionValue
            }
            return $true
        }catch{
            return $false
        }

    }

    Function StartUpdatesActions($ComputerName){
        try{
            StartAction $ComputerName $Actions.'Software Update Deployment Evaluation Cycle'
            StartAction $ComputerName $Actions.'Software Update Scan Cycle'
            return $true
        }catch{
            return $false
        }
    }

    Function StartHwInventoryAction($ComputerName){
        try{
            StartAction $ComputerName $Actions.'Hardware Inventory Cycle'
            return $true
        }catch{
            return $false
        }
    }

    Function ForceFullInventory($computerName){
        try{
            Get-WmiObject -ComputerName $ComputerName -Namespace root\ccm\invagt -Query 'select * from inventoryActionStatus where InventoryActionID="{00000000-0000-0000-0000-000000000001}"' | Remove-WmiObject -Confirm:$false
            StartAction $ComputerName $Actions.'Hardware Inventory Cycle'
            return $true
        }
        catch{
            return $false
        }
    }

    Function VerifyConnection($ComputerName){

        try{

            #Testa conectividade via ping
            if(Test-Connection -computername $ComputerName -count 1 -quiet){
           
                #Testa conectividade via WMI
                if(Get-WmiObject win32_service -ComputerName $ComputerName -ErrorAction SilentlyContinue){
                    
                    #Verifica se o serviço CcmExec está no ar
                    if((Get-Service -name CcmExec -ComputerName $ComputerName -ErrorAction SilentlyContinue).Status -eq "Running"){
                    
                        return $true

                    }
                }
            }

            return $false

        }catch{
            return $false
        }    
    }

}

process{
    
    #Gerando cabeçalho
    $row = CreateRow
    $row.Num = " # "
    $row.Computer = "COMPUTER"
    $row.Updates = "UPDATES"
    $row.Status = "STATUS"
    $row.Reboot = "REBOOT"

    $title = "SCCM Update Tool v$version"
    $runDate = Get-Date -Format G
    $header = $title + $runDate.PadLeft($linhaTabela.Length-$title.Length)

    write-host "`n`n`n`n`n`n`n"
    Write-host $header
    Write-Host $linhaTabela
    PrintRow
    Write-Host $linhaTabela


    #Código Main ------------------------------------------------------------------------------------------------
    foreach ($computer in $computers){    
        
        $k++
        $num = "{0,$PADNum :D$PADNum}" -f $k
        $UpdatesPendentes = @()

        $row = CreateRow
        $row.Num = "$num"
        $row.Computer = $computer
        $row.Updates = " "
        $row.Status = " "
        $row.Reboot = " "
    
        Write-Progress -Activity "SCCM Update Tool v$version" -Status "Processando $k/$($computers.Count) - $computer" -PercentComplete (($k / $computers.Count)  * 100)

        If (VerifyConnection $computer){

            if($PSCmdlet.ParameterSetName -ieq "ForceReboot"){

                $row.Reboot = RebootComputer $computer $ForceReboot
            }
            elseif($PSCmdlet.ParameterSetName -ilike "Actions*"){
            
                switch ($PSCmdlet.ParameterSetName){

                    'ActionsAll' {$status = StartAllActions $computer}
                
                    'ActionsUpdates' {$status = StartUpdatesActions $computer}

                    'ActionsHwInventory' {$status = StartHwInventoryAction $computer}
                
                    'ActionsFullHwInventory' {$status = ForceFullInventory $computer}
                }
            
                if($status){
                    $row.Status = $STATUS_SUCCESS_ACTIONS

                }
                else{
                    $row.Status = $STATUS_FAIL_ACTIONS
                }
            }
            elseif($PSCmdlet.ParameterSetName -ieq "Check"){

                $row.Updates = $CheckHotFix
                $row.Status = GetHotfix $computer $CheckHotFix

            }
            elseif($PSCmdlet.ParameterSetName -ieq "Query" -or $PSCmdlet.ParameterSetName -ieq "Install"){
            
                $Error.Clear()
                if(!(GetUpdates $computer)){
                    
                    $row.Status = $STATUS_FAIL_GET_UPDATES

                }
                else{
                    
                    $row.Updates = $Script:UpdatesPendentes.Count.ToString()

                    if($Script:UpdatesPendentes.Count){
    
                        #Verificando se a máquina possui reboot pendente
                        if(GetRebootStatus $computer){
                            
                            $row.Status = $STATUS_PENDING_REBOOT
                
                            if($Reboot){ 
                                $row.Reboot = RebootComputer $computer $ForceReboot
                            }

                        }
                        else{          
                            
                            $UpdateError = $false
                            $UpdateInstalado = $true

                            foreach ($update in $script:UpdatesPendentes){
                                
                                # Verificar se o patches está com status instalado
                                if($Evaluation.[int]$update.EvaluationState -ne 'JobStatePendingSoftReboot' -and
                                   $Evaluation.[int]$update.EvaluationState -ne 'JobStateInstallComplete')
                                {
                                    
                                    $UpdateInstalado = $false
                                     
                                    #Verifica se o status é JobStateError
                                    if($update.EvaluationState -eq 13){
                                        $UpdateError = $true
                                    }
                                   
                                }
                            }

                            if($UpdateError){
                                
                                $row.Status = $STATUS_UPDATES_ERROR
                                
                                if($FixUpdateError){
                                    
                                    if(FixError -ComputerName $computer){
                                        $row.Status = $STATUS_UPDATES_ERROR_FIX        
                                    }
                                }
                            }
                            elseif(!$UpdateInstalado){
                                
                                $row.Status = $STATUS_PENDING_UPDATES
                               
                                if($Install){                            
                                    
                                    $row.Status = InstallUpdates $computer $Script:UpdatesPendentes
                                }
                            }
                            else{
                                $row.Status = $STATUS_UPDATES_INSTALADOS
                            }
                        }

                    }else{                       
                        $row.Status = $STATUS_NO_UPDATES
                    }
                }
            }
        }
        else {
            $row.Status = $STATUS_OFFILINE
        }
        
        PrintRow

        #Exibe detalhes dos KB listados
        if($Details){                   

            if($Script:UpdatesPendentes.Count){
                
                Write-Host $linhaTabela

                $Script:UpdatesPendentes | 
                Select-Object @{LABEL='KB';EXPRESSION={$_.ArticleID}}, 
                              @{LABEL='CCMStatus';EXPRESSION={$key = [int]$_.EvaluationState; $Evaluation.$key}}, 
                              @{LABEL='PercentComplete';EXPRESSION={if($_.EvaluationState -eq 7){($_.PercentComplete).toString()+"%"}else{""}}}, 
                              @{LABEL='Size';EXPRESSION={([System.Math]::Round($_.ContentSize / 1KB)).tostring()+"MB"}},
                              @{LABEL='Description';EXPRESSION={$_.name}} |
                              Sort-Object CCMStatus,KB | Format-Table -AutoSize  
                
                if($computers.IndexOf($computer) -ne ($computers.Count-1)){
                    Write-Host $linhaTabela
                }
            }
        }

        if($ExportCSV){
            $report += $row
        }
    }
}

end{
    
    Write-Host $linhaTabela
    
    #Export para arquivo CSV
    if($ExportCSV){
        $report | Export-csv -NoTypeInformation -UseCulture $ExportCSV
    }
}
