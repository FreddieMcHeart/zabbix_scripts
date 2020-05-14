Param ( $Srv )
$svc = Get-Service zabbix* 
$chek = Get-NetFirewallRule -DisplayName "Zabbix*"

function Get-Conf {
    $conf = "C:\zabbix\zabbix_agentd.conf"
    $HostName = hostname
 
        New-Item $conf -ItemType File 
        Add-Content $conf "LogFile=c:\zabbix\zabbix_agentd.log"
        Add-Content $conf "Server=$Srv"
        Add-Content $conf "ServerActive=$Srv"
        Add-Content $conf "Hostname=$HostName"
}

function Dell-Path {
    $string = ps zabbix* | ls | Convert-Path 
    $zabPath = "C:\zabbix"    
        Stop-Service $svc
        Pause
        Write-Host "Zabbix Agent`s service was stoped"
        if ( $string.Replace( "\zabbix_agentd.exe" , "" ) -eq $zabPath )
            { 
                Remove-Item -path $string.Replace( "\zabbix_agentd.exe" , "" ) -Exclude *.txt
                Pause
            }
        else 
            {
                Remove-Item $string.Replace( "\zabbix_agentd.exe" , "" )
                Pause 
            }               
         
 }

function Get-Agent {
    Get-Conf
    Write-Host "Conf File was created"
    $download_url = "https://pendalf.help-it.biz/s/2Qghr2dZr4GjZnx/download"
    $local_path = "C:\zabbix\zabbix_agentd.exe" 
    $WebClient = New-Object System.Net.WebClient
    $WebClient.DownloadFile($download_url, $local_path)
    Pause
    Write-Host "Agent was downloaded" 
    
    if ( !$chek ) 
        { 
            New-NetFirewallRule -DisplayName "Zabbix" -Direction Inbound -LocalPort 10050 -Protocol TCP -Action Allow
            Pause
            Write-Host "New Firewall Rule was added"
        }

    Start-Service zabbix*
    
    if ( $svc ) 
        {
            Write-Host "Zabbix Agent`s service was started"
        }
    else
        {
            C:\zabbix\zabbix_agentd.exe --config C:\zabbix\zabbix_agentd.conf --install   
            Pause
            Write-Host "Zabbix Agent`s service was installed"
            Start-Service zabbix*
            Write-Host "Zabbix Agent`s service was started"
            Pause
        }
}

if ( !$svc ) 
   {
       New-Item C:\zabbix -ItemType Directory
       Write-Host "Directory Added"
       Pause
       Get-Agent
       Pause
   }
else 
   {
       Dell-Path
       Write-Host "Old agent was delleted"
       Pause
       New-Item C:\zabbix -ItemType Directory
       Write-Host "Directory Added"
       Pause
       Get-Agent
       Pause
   } 