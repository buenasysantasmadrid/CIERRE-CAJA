# Al abrir Claude Code en este proyecto:
# 1) trae lo último de GitHub (el usuario trabaja desde 2 ordenadores)
# 2) compara el Apps Script publicado en Google con los .gs de esta carpeta
# Devuelve JSON: systemMessage (lo ve el usuario) y additionalContext (lo ve Claude).

$ErrorActionPreference = 'Continue'
$env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')
$raiz = Split-Path -Parent $PSScriptRoot
Set-Location $raiz
$avisos = @()

# --- 1) git pull ---
$pull = (git pull --ff-only 2>&1 | Out-String).Trim()
if ($LASTEXITCODE -ne 0) {
  $avisos += "git pull FALLÓ (hay que resolverlo antes de tocar nada): $pull"
} elseif ($pull -match 'Already up to date|Ya está actualizado') {
  $avisos += 'git: al día con GitHub.'
} else {
  $avisos += "git: bajados cambios del otro ordenador. $pull"
}

# --- 2) Apps Script en Google vs local ---
if (-not (Get-Command clasp -ErrorAction SilentlyContinue)) {
  $avisos += 'clasp no está instalado en este ordenador: no se pudo comparar con Google.'
} elseif (-not (Test-Path '.clasp.json')) {
  $avisos += 'Falta .clasp.json: no se pudo comparar con Google.'
} else {
  $id = (Get-Content '.clasp.json' -Raw | ConvertFrom-Json).scriptId
  $tmp = Join-Path $env:TEMP ('clasp-remoto-' + [guid]::NewGuid())
  New-Item -ItemType Directory $tmp | Out-Null
  Push-Location $tmp
  $clon = (clasp clone $id 2>&1 | Out-String)
  Pop-Location
  if (-not (Get-ChildItem $tmp -Filter *.js -ErrorAction SilentlyContinue)) {
    $avisos += "No se pudo descargar el Apps Script de Google (¿falta clasp login?): $($clon.Trim())"
  } else {
    $norm = { param($p) ((Get-Content $p -Raw -Encoding UTF8) -replace "`r", '').TrimEnd() }
    $locales = @{}
    Get-ChildItem $raiz -Filter *.gs | ForEach-Object { $locales[$_.BaseName.ToLower()] = $_.FullName }
    $distintos = @()
    Get-ChildItem $tmp -Filter *.js | ForEach-Object {
      $clave = $_.BaseName.ToLower()
      if (-not $locales.ContainsKey($clave)) { $distintos += "$($_.BaseName) (solo está en Google)" }
      elseif ((& $norm $_.FullName) -ne (& $norm $locales[$clave])) { $distintos += "$($_.BaseName) (distinto)" }
      $locales.Remove($clave)
    }
    foreach ($k in $locales.Keys) { $distintos += "$k (solo está en local)" }
    if ($distintos.Count -eq 0) { $avisos += 'Apps Script: Google coincide con los .gs locales.' }
    else { $avisos += 'Apps Script DIFIERE de Google: ' + ($distintos -join ', ') + ". Copia de Google en: $tmp" }
  }
}

$texto = $avisos -join "`n"
@{ systemMessage = $texto; hookSpecificOutput = @{ hookEventName = 'SessionStart'; additionalContext = $texto } } | ConvertTo-Json -Compress
