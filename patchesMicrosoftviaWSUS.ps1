# Importar o módulo necessário
Import-Module -Name PSWindowsUpdate
 
# Definir a lista de servidores
$servidores = @("servidor1", "servidor2", "servidor3")
 
# Definir o servidor WSUS
$wsusServer = "nome-do-servidor-wsus"
 
# Definir o porto do servidor WSUS
$wsusPort = 8530
 
# Definir o nome do grupo de computadores no WSUS
$wsusGroupName = "Nome-do-Grupo"
 
# Loop para cada servidor
foreach ($servidor in $servidores) {
  # Conectar ao servidor WSUS
  $wsus = Get-WSUS Server -Name $wsusServer -Port $wsusPort
 
  # Obter a lista de atualizações disponíveis para o servidor
  $atualizacoes = Get-WsusUpdate -WsusServer $wsus -ComputerTargetGroups $wsusGroupName -ComputerName $servidor
 
  # Instalar as atualizações
  foreach ($atualizacao in $atualizacoes) {
    Write-Host "Instalando atualização $($atualizacao.Title) no servidor $servidor"
    Install-WsusUpdate -WsusServer $wsus -Update $atualizacao
  }
 
  # Reiniciar o servidor se necessário
  if ((Get-WsusUpdate -WsusServer $wsus -ComputerTargetGroups $wsusGroupName -ComputerName $servidor | Where-Object {$_.IsInstalled -eq $false}).Count -gt 0) {
    Write-Host "Reiniciando o servidor $servidor"
    Restart-Computer -ComputerName $servidor -Force
  }
}
