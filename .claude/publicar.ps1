# Sube los .gs a Google y publica una nueva versión en la implementación de la
# app (misma URL). Uso: powershell -ExecutionPolicy Bypass -File .claude\publicar.ps1 "descripción"
param([string]$Descripcion = 'Actualización')

$env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')
Set-Location (Split-Path -Parent $PSScriptRoot)
$IMPLEMENTACION = 'AKfycbyrfOE73JqgjTXAEaLetRz8aJg1Ldb-Du4dsqGgW34H-zgwLSJMrAkSKApiTCoz6t7m_A'

# La app corre como la cuenta que publica (executeAs USER_DEPLOYING): tiene que
# ser lfisbein, la dueña de las planillas. Con otra cuenta, Contabilidad y
# Albaranes (protegidas para lfisbein) dejan de guardarse.
$cuenta = (node -e "const c=JSON.parse(require('fs').readFileSync(require('os').homedir()+'/.clasprc.json','utf8'));const t=c.tokens?Object.values(c.tokens)[0]:c.token;fetch('https://oauth2.googleapis.com/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:new URLSearchParams({client_id:t.client_id,client_secret:t.client_secret,refresh_token:t.refresh_token,grant_type:'refresh_token'})}).then(r=>r.json()).then(j=>fetch('https://www.googleapis.com/oauth2/v3/userinfo',{headers:{Authorization:'Bearer '+j.access_token}})).then(r=>r.json()).then(u=>console.log(u.email||''))" | Out-String).Trim()
if ($cuenta -ne 'lfisbein@gmail.com') {
  Write-Host ("clasp está con la cuenta '" + $cuenta + "', no con lfisbein@gmail.com. No se publica. Hay que hacer: clasp login (y elegir lfisbein).") -ForegroundColor Red
  exit 1
}

# clasp push BORRA en Google lo que no esté aquí: comprobar antes que no falte nada.
$id = (Get-Content '.clasp.json' -Raw | ConvertFrom-Json).scriptId
$tmp = Join-Path $env:TEMP ('clasp-remoto-' + [guid]::NewGuid())
New-Item -ItemType Directory $tmp | Out-Null
Push-Location $tmp; clasp clone $id | Out-Null; Pop-Location
$remotos = Get-ChildItem $tmp -Filter *.js
if (-not $remotos) { Write-Host 'No se pudo leer el proyecto de Google. No se sube nada.' -ForegroundColor Red; exit 1 }
$faltan = $remotos | Where-Object { -not (Test-Path ($_.BaseName + '.gs')) }
if ($faltan) {
  Write-Host ('Faltan en este ordenador: ' + (($faltan | ForEach-Object BaseName) -join ', ') + '. No se sube nada.') -ForegroundColor Red
  exit 1
}

clasp push --force
if ($LASTEXITCODE -ne 0) { Write-Host 'Falló la subida. No se publica.' -ForegroundColor Red; exit 1 }
clasp deploy -i $IMPLEMENTACION -d $Descripcion
if ($LASTEXITCODE -ne 0) { Write-Host 'Falló la publicación.' -ForegroundColor Red; exit 1 }
Write-Host 'Listo: subido y publicado.' -ForegroundColor Green
