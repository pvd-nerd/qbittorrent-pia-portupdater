$qbit_ini = "C:\Users\Administrator\AppData\Roaming\qBittorrent\qBittorrent.ini"
$qbit_port = Select-String -Path $qbit_ini 'Session\\Port=\d*' | Select-Object -ExpandProperty Line
$qbit_port = $qbit_port -creplace '^[^0-9]*'
$qbit_process = Get-Process qbittorrent -ErrorAction SilentlyContinue

$pia_port = & "C:\Program Files\Private Internet Access\piactl.exe" get portforward
$pia_port = $pia_port -as [int]

if ($qbit_port -ne $pia_port) { 
    if ($qbit_process) {
        Write-Output "PIA port updated, need to update qBittorrent!"
        Write-Output "Stopping process and waiting 120 seconds..."
        Stop-Process -Name "qbittorrent" 
        Start-Sleep -Seconds 30

        Write-Output "Updating qBittorrent config..."
        (Get-Content $qbit_ini) -replace "Session\\Port=\d*", "Session\Port=$pia_port" | Set-Content -Path $qbit_ini
        Start-Sleep -Seconds 1

        Write-Output "Starting process..."
        Start-Process -FilePath "C:\Program Files\qBittorrent\qbittorrent.exe"
        Start-Sleep -Seconds 5

        $qbit_process = Get-Process qbittorrent -ErrorAction SilentlyContinue
            if ($qbit_process) {
                $qbit_process.CloseMainWindow()
                Write-Output "qBittorrent started! Exiting..."
            }
            else {
                Write-Output "qBittorrent not started yet. Waiting..."
                Start-Sleep -Seconds 10
                $qbit_process = Get-Process qbittorrent -ErrorAction SilentlyContinue
                $qbit_process.CloseMainWindow()
                Write-Output "qBittorrent started! Exiting..."
            }
    }
}
else {
    Write-Output "Ports match, nothing to do..."
}
