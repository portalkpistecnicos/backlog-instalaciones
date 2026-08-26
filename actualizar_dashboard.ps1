<#
  Actualiza el dashboard Backlog Instalaciones a partir del CSV mas reciente
  en la carpeta de origen, guarda un respaldo historico (HTML del dia + fila
  de KPIs) y sube los cambios a GitHub Pages.

  No editar los valores a mano en index.html: este script reemplaza solo el
  bloque de datos marcado con /*DATA_START*/ ... /*DATA_END*/ dentro del
  archivo, dejando intacto el resto del HTML/CSS/JS.
#>

$ErrorActionPreference = 'Stop'

$ProjectDir  = "G:\Mi unidad\Respaldo\Documents\KPI's\Backlog Instala"
$SourceDir   = "G:\Mi unidad\Respaldo\Documents\KPI's\Backlog Instala\BBDD"
$IndexFile   = Join-Path $ProjectDir "index.html"
$HistDir     = Join-Path $ProjectDir "historico"
$HistHtmlDir = Join-Path $HistDir "html"
$KpiCsv      = Join-Path $HistDir "kpis_historico.csv"
$LogDir      = Join-Path $ProjectDir "logs"
$LogFile     = Join-Path $LogDir ("actualizar_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
$Utf8NoBom   = New-Object System.Text.UTF8Encoding($false)
$Inv         = [System.Globalization.CultureInfo]::InvariantCulture

New-Item -ItemType Directory -Force -Path $HistHtmlDir | Out-Null
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

function Write-Log {
    param([string]$Message)
    $line = "{0}  {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
    Write-Host $line
}

try {
    Write-Log "===== Inicio actualizacion Backlog Instalaciones ====="

    # 1. Ubicar el CSV fuente mas reciente (soporta el cambio de nombre mensual del archivo)
    $csvFile = Get-ChildItem -Path $SourceDir -Filter "base-backlog-instalaciones_*.csv" -ErrorAction Stop |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $csvFile) { throw "No se encontro ningun archivo base-backlog-instalaciones_*.csv en $SourceDir" }
    Write-Log "Archivo fuente: $($csvFile.FullName)"

    $data = Import-Csv -Path $csvFile.FullName -Delimiter ';' -Encoding UTF8
    Write-Log "Filas leidas: $($data.Count)"

    # 2. Determinar snapshot mas reciente y rango de fechas del archivo
    $allDates = $data | ForEach-Object { [datetime]::ParseExact($_.partition_date, 'dd-MM-yyyy', $null) }
    $maxDate = ($allDates | Measure-Object -Maximum).Maximum
    $minDate = ($allDates | Measure-Object -Minimum).Minimum
    $maxDateStr = $maxDate.ToString('dd-MM-yyyy')
    $snapshotCount = ($data | Select-Object -ExpandProperty partition_date -Unique).Count
    Write-Log "Snapshot mas reciente: $maxDateStr | Rango: $($minDate.ToString('dd-MM-yyyy')) a $maxDateStr | Snapshots: $snapshotCount"

    # 3. Snapshot de hoy, excluyendo filas marcadas para exclusion de backlog
    $hoy = $data | Where-Object { $_.partition_date -eq $maxDateStr -and $_.filtro_backlog_exclusion -eq '0' }

    function Get-AgencyStats {
        param([string]$AgenciaRaw)
        $sub = $hoy | Where-Object { $_.agen_descripcion -eq $AgenciaRaw }
        $proc = $sub | Where-Object { $_.espe_nombre -eq 'En Proceso' }
        $term = ($sub | Where-Object { $_.espe_nombre -eq 'Terminada' }).Count
        $canc = ($sub | Where-Object { $_.espe_nombre -eq 'Cancelada' }).Count
        $con = 0
        $venc = 0
        $sinAtender1Dia = 0
        foreach ($r in $proc) {
            if ($r.dia_especifico -ne '') {
                $con++
                $f = [datetime]::ParseExact($r.dia_especifico, 'dd-MM-yyyy', $null)
                if ($f -lt $maxDate) { $venc++ }
            }
            if ($r.peti_fecha_ingreso -ne '') {
                $fi = [datetime]::ParseExact($r.peti_fecha_ingreso, 'dd-MM-yyyy', $null)
                $diasAbierta = (New-TimeSpan -Start $fi -End $maxDate).Days
                if ($diasAbierta -gt 1) { $sinAtender1Dia++ }
            }
        }
        [PSCustomObject]@{
            enProceso      = $proc.Count
            terminada      = $term
            cancelada      = $canc
            agendaCon      = $con
            agendaVencida  = $venc
            sinAtender1Dia = $sinAtender1Dia
        }
    }

    $agencyMap = [ordered]@{
        "Viña del Mar" = "VINA DEL MAR"
        "Valparaíso"   = "VALPARAISO"
        "San Antonio"  = "SAN ANTONIO"
    }

    $agencies = @()
    foreach ($displayName in $agencyMap.Keys) {
        $s = Get-AgencyStats -AgenciaRaw $agencyMap[$displayName]
        $agencies += [PSCustomObject]@{
            name           = $displayName
            enProceso      = $s.enProceso
            terminada      = $s.terminada
            cancelada      = $s.cancelada
            agendaCon      = $s.agendaCon
            agendaVencida  = $s.agendaVencida
            sinAtender1Dia = $s.sinAtender1Dia
        }
        Write-Log ("Agencia {0}: EnProceso={1} Terminada={2} Cancelada={3} AgendaVencida={4}/{5} SinAtender1Dia={6}" -f $displayName, $s.enProceso, $s.terminada, $s.cancelada, $s.agendaVencida, $s.agendaCon, $s.sinAtender1Dia)
    }

    # 4. Buckets: todo el archivo cargado (todas las fechas), mismo filtro
    $allFiltered = $data | Where-Object { $_.filtro_backlog_exclusion -eq '0' -and $_.vpi_cod_bucket -ne '' }
    $bucketGroups = $allFiltered | Group-Object vpi_cod_bucket
    $buckets = [ordered]@{}
    foreach ($g in $bucketGroups) {
        $c = ($g.Group | Where-Object { $_.espe_nombre -eq 'Cancelada' }).Count
        $p = ($g.Group | Where-Object { $_.espe_nombre -eq 'En Proceso' }).Count
        $t = ($g.Group | Where-Object { $_.espe_nombre -eq 'Terminada' }).Count
        $buckets[$g.Name] = [PSCustomObject]@{ c = $c; p = $p; t = $t }
    }
    Write-Log "Buckets encontrados: $($buckets.Count)"

    # 4b. Consolidado por agencia: mismo acumulado del mes (todas las fechas), agrupado por agen_descripcion
    $allFilteredAllAgencies = $data | Where-Object { $_.filtro_backlog_exclusion -eq '0' }
    $agencyConsolidated = @()
    foreach ($displayName in $agencyMap.Keys) {
        $sub = $allFilteredAllAgencies | Where-Object { $_.agen_descripcion -eq $agencyMap[$displayName] }
        $c = ($sub | Where-Object { $_.espe_nombre -eq 'Cancelada' }).Count
        $p = ($sub | Where-Object { $_.espe_nombre -eq 'En Proceso' }).Count
        $t = ($sub | Where-Object { $_.espe_nombre -eq 'Terminada' }).Count
        $agencyConsolidated += [PSCustomObject]@{ name = $displayName; c = $c; p = $p; t = $t }
        Write-Log ("Consolidado mes {0}: Cancelada={1} EnProceso={2} Terminada={3}" -f $displayName, $c, $p, $t)
    }

    # 5. Armar el bloque de datos y convertirlo a JSON
    $dashboardData = [ordered]@{
        sourceFile          = $csvFile.Name
        snapshotDate        = $maxDateStr
        snapshotCount       = $snapshotCount
        periodLabel         = "$($minDate.ToString('dd-MM')) al $maxDateStr"
        generatedAt         = (Get-Date -Format 'dd-MM-yyyy HH:mm')
        agencies            = $agencies
        agencyConsolidated  = $agencyConsolidated
        buckets             = $buckets
    }
    $json = $dashboardData | ConvertTo-Json -Depth 6 -Compress
    Write-Log "JSON generado ($($json.Length) caracteres)"

    # 6. Reemplazar SOLO el bloque entre marcadores dentro de index.html
    $htmlBytes = [System.IO.File]::ReadAllBytes($IndexFile)
    $html = [System.Text.Encoding]::UTF8.GetString($htmlBytes)

    $pattern = '(?<=/\*DATA_START\*/)[\s\S]*?(?=/\*DATA_END\*/)'
    if ($html -notmatch $pattern) { throw "No se encontraron los marcadores DATA_START/DATA_END en index.html" }
    $evaluator = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $json }
    $newHtml = [System.Text.RegularExpressions.Regex]::Replace($html, $pattern, $evaluator, 'Singleline')

    [System.IO.File]::WriteAllText($IndexFile, $newHtml, $Utf8NoBom)
    Write-Log "index.html actualizado"

    # 7. Copia historica del HTML del dia
    $histHtmlFile = Join-Path $HistHtmlDir ("{0}.html" -f $maxDate.ToString('yyyy-MM-dd'))
    Copy-Item -Path $IndexFile -Destination $histHtmlFile -Force
    Write-Log "Copia historica guardada: $histHtmlFile"

    # 8. Fila en el historico de KPIs (reemplaza la fila si ya existia una para esa fecha)
    $zoneEnProceso = ($agencies | Measure-Object -Property enProceso -Sum).Sum
    $zoneTerminada = ($agencies | Measure-Object -Property terminada -Sum).Sum
    $zoneCancelada = ($agencies | Measure-Object -Property cancelada -Sum).Sum
    $zoneRate = ($zoneCancelada + $zoneTerminada) / 6
    $zoneBacklog = if ($zoneRate -gt 0) { [math]::Round($zoneEnProceso / $zoneRate, 2) } else { 0 }

    if (-not (Test-Path $KpiCsv)) {
        $header = "fecha;en_proceso_zona;terminada_zona;cancelada_zona;backlog_zona;" +
            (($agencies | ForEach-Object { "en_proceso_$($_.name);terminada_$($_.name);cancelada_$($_.name);backlog_$($_.name)" }) -join ';')
        [System.IO.File]::WriteAllText($KpiCsv, $header + "`r`n", $Utf8NoBom)
    }

    $agencyParts = foreach ($a in $agencies) {
        $rate = ($a.cancelada + $a.terminada) / 6
        $backlog = if ($rate -gt 0) { [math]::Round($a.enProceso / $rate, 2) } else { 0 }
        "$($a.enProceso);$($a.terminada);$($a.cancelada);$($backlog.ToString($Inv))"
    }
    $row = "$maxDateStr;$zoneEnProceso;$zoneTerminada;$zoneCancelada;$($zoneBacklog.ToString($Inv));" + ($agencyParts -join ';')

    $existingRows = @()
    if (Test-Path $KpiCsv) { $existingRows = Get-Content -Path $KpiCsv -Encoding UTF8 }
    $datePrefix = "^$([regex]::Escape($maxDateStr));"
    $alreadyHasRowForDate = @($existingRows | Where-Object { $_ -match $datePrefix }).Count -gt 0
    if (-not $alreadyHasRowForDate) {
        Add-Content -Path $KpiCsv -Value $row -Encoding UTF8
        Write-Log "Fila agregada al historico de KPIs: $maxDateStr"
    } else {
        $newRows = @($existingRows | Where-Object { $_ -notmatch $datePrefix }) + $row
        [System.IO.File]::WriteAllLines($KpiCsv, $newRows, $Utf8NoBom)
        Write-Log "Fila de $maxDateStr en historico de KPIs actualizada (ya existia)"
    }

    # 9. Commit + push a GitHub (GitHub Pages se actualiza solo)
    # git escribe su progreso normal por stderr; con $ErrorActionPreference='Stop'
    # PowerShell 5.1 convertiria esas lineas en errores terminantes si se capturan
    # con 2>&1, asi que se baja a 'Continue' solo para estos comandos nativos.
    Push-Location $ProjectDir
    $prevEAP = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        git add index.html historico .gitignore actualizar_dashboard.ps1 actualizar_dashboard.bat *>$null
        $changes = git status --porcelain
        if ($changes) {
            git -c user.email="ximeloquilla@gmail.com" -c user.name="xarancibias" commit -m "Actualizacion automatica $maxDateStr" *>$null
            # Redirigido via cmd.exe (no PowerShell) para evitar que 5.1 envuelva
            # cada linea de stderr de git como un ErrorRecord ruidoso en el log.
            $pushLogTmp = Join-Path $LogDir "_push.tmp"
            cmd.exe /c "git push origin master > `"$pushLogTmp`" 2>&1"
            $pushExit = $LASTEXITCODE
            $pushOutput = (Get-Content -Path $pushLogTmp -Raw -ErrorAction SilentlyContinue)
            Remove-Item -Path $pushLogTmp -ErrorAction SilentlyContinue
            Write-Log ("git push (exit $pushExit): " + ($pushOutput -replace '[\r\n]+', ' '))
            if ($pushExit -ne 0) { throw "git push fallo con codigo $pushExit" }
            Write-Log "Cambios enviados a GitHub"
        } else {
            Write-Log "Sin cambios respecto al ultimo commit; no se hizo push"
        }
    } finally {
        Pop-Location
        $ErrorActionPreference = $prevEAP
    }

    Write-Log "===== Actualizacion completada OK ====="
    exit 0
}
catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    Write-Log $_.ScriptStackTrace
    exit 1
}
