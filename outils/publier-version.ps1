# ═══════════════════════════════════════════════════════════════════════════
#  Publication d'une version de SageBridge sur GitHub (dépôt public)
# ═══════════════════════════════════════════════════════════════════════════
#  1. lit la version de bin\SageBridge.exe (AssemblyVersion de SageBridge.cs) ;
#  2. utilise le compte GitHub déjà enregistré dans Git sur ce PC
#     (Gestionnaire d'identifiants Windows) — aucun mot de passe n'est stocké ici ;
#  3. crée le dépôt public s'il n'existe pas, pousse le code source (branche main) ;
#  4. crée la release vX.Y.Z avec SageBridge-Setup.exe, son empreinte .sha256
#     et la documentation PDF, puis la rend publique.
#  Les SageBridge installés la détectent au démarrage ou dans l'heure.
# ═══════════════════════════════════════════════════════════════════════════

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[Net.ServicePointManager]::Expect100Continue = $false
function Remove-UntaggedTags {
    # Étiquettes « untagged-… » laissées par GitHub quand une version est publiée sans étiquette.
    foreach ($t in @(Api 'GET' "https://api.github.com/repos/$Repo/tags?per_page=100")) {
        if ($t.name -like 'untagged-*') {
            Api 'DELETE' "https://api.github.com/repos/$Repo/git/refs/tags/$($t.name)" | Out-Null
            Write-Host "   Étiquette parasite $($t.name) supprimée."
        }
    }
}

try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
# Sans BOM : sinon .NET écrit l'en-tête UTF-8 au début de l'entrée de git (« missing protocol field »).
try { [Console]::InputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }

$Repo = 'Gnaris/SageBridge'
$Root = Split-Path -Parent $PSScriptRoot
$LogFile = Join-Path $Root 'publier-version.log'
Start-Transcript -Path $LogFile -Force | Out-Null

function Step($t) { Write-Host ''; Write-Host "▶ $t" -ForegroundColor Cyan }

function Invoke-Git {
    $old = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    $out = & $script:Git @args 2>&1 | ForEach-Object { "$_" }
    $script:GitCode = $LASTEXITCODE
    $ErrorActionPreference = $old
    if ($out) { $out | ForEach-Object { Write-Host "   $_" } }
}

function Api($Method, $Uri, $Body) {
    $p = @{ Method = $Method; Uri = $Uri; Headers = $script:H }
    if ($Body) { $p.Body = [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 5)); $p.ContentType = 'application/json; charset=utf-8' }
    # « return » énumère les tableaux JSON (Invoke-RestMethod de PowerShell 5 ne le fait pas).
    $r = Invoke-RestMethod @p
    return $r
}

function StatusOf($err) {
    try { return [int]$err.Exception.Response.StatusCode } catch { return 0 }
}

try {
    $exe = Join-Path $Root 'bin\SageBridge.exe'
    $setup = Join-Path $Root 'bin\SageBridge-Setup.exe'
    $sha = "$setup.sha256"
    $pdf = Join-Path $Root 'docs\SageBridge - Documentation.pdf'
    foreach ($f in $exe, $setup, $sha) { if (-not (Test-Path $f)) { throw "Fichier manquant : $f — lancez d'abord build.cmd." } }
    $v = [Reflection.AssemblyName]::GetAssemblyName($exe).Version
    $ver = '{0}.{1}.{2}' -f $v.Major, $v.Minor, $v.Build
    $tag = "v$ver"
    Write-Host "Publication de SageBridge $ver sur https://github.com/$Repo" -ForegroundColor Green

    Step 'Recherche de Git'
    $script:Git = (Get-Command git.exe -ErrorAction SilentlyContinue | Select-Object -First 1).Source
    if (-not $script:Git) {
        foreach ($p in "$env:ProgramFiles\Git\cmd\git.exe", "${env:ProgramFiles(x86)}\Git\cmd\git.exe", "$env:LOCALAPPDATA\Programs\Git\cmd\git.exe") {
            if (Test-Path $p) { $script:Git = $p; break }
        }
    }
    if (-not $script:Git) { throw 'Git est introuvable sur ce PC (https://git-scm.com).' }
    Write-Host "   $script:Git"

    Step 'Compte GitHub enregistré sur ce PC'
    # Écriture directe sur l'entrée standard de git (un tube PowerShell ajouterait un BOM).
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = $script:Git; $psi.Arguments = 'credential fill'
    $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true
    $psi.RedirectStandardInput = $true; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
    $old = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    $helpers = & $script:Git config --show-origin --get-all credential.helper 2>&1 | ForEach-Object { "$_" }
    $ErrorActionPreference = $old
    Write-Host "   Gestionnaire d'identifiants Git : $($helpers -join ' | ')"
    Write-Host "   (une fenêtre de connexion GitHub peut s'ouvrir : connectez-vous avec le compte Gnaris)"
    $proc = [Diagnostics.Process]::Start($psi)
    $bytes = [Text.Encoding]::ASCII.GetBytes("protocol=https`nhost=github.com`n`n")
    $proc.StandardInput.BaseStream.Write($bytes, 0, $bytes.Length)
    $proc.StandardInput.Close()
    $credText = $proc.StandardOutput.ReadToEnd()
    $credErr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit(300000) | Out-Null
    if ($credErr) { Write-Host "   git : $($credErr.Trim())" }
    $cred = $credText -split "`r?`n"
    $token = (($cred | Where-Object { $_ -like 'password=*' } | Select-Object -First 1) -replace '^password=', '')
    if (-not $token) { throw "Aucun compte GitHub n'est enregistré dans Git sur ce PC (faites un premier « git push » sur un dépôt pour vous connecter)." }
    $script:H = @{ Authorization = "token $token"; 'User-Agent' = 'SageBridge-publication'; Accept = 'application/vnd.github+json' }
    $me = Api 'GET' 'https://api.github.com/user'
    $owner, $name = $Repo.Split('/')
    Write-Host "   Connecté en tant que $($me.login)"
    if ($me.login -ne $owner) { throw "Le compte GitHub enregistré est « $($me.login) », pas « $owner »." }

    Step "Dépôt $Repo"
    $exists = $true
    try { Api 'GET' "https://api.github.com/repos/$Repo" | Out-Null }
    catch { if ((StatusOf $_) -eq 404) { $exists = $false } else { throw } }
    if ($exists) { Write-Host '   Le dépôt existe déjà.' }
    else {
        Api 'POST' 'https://api.github.com/user/repos' @{
            name = $name; private = $false; has_wiki = $false; has_projects = $false
            description = "Création de factures Sage 50 Gestion commerciale par l'interface COM — installateur Windows et mises à jour automatiques"
            homepage = "https://github.com/$Repo/releases/latest"
        } | Out-Null
        Write-Host '   Dépôt public créé.'
    }

    # Version déjà publiée mais sans étiquette (publication interrompue) : on pose l'étiquette et on s'arrête.
    $orphan = @(Api 'GET' "https://api.github.com/repos/$Repo/releases?per_page=50") | Where-Object { -not $_.draft -and $_.name -eq "SageBridge $ver" -and $_.tag_name -ne $tag } | Select-Object -First 1
    if ($orphan) {
        Api 'PATCH' "https://api.github.com/repos/$Repo/releases/$($orphan.id)" @{ tag_name = $tag; target_commitish = 'main'; make_latest = 'true' } | Out-Null
        Write-Host "   Étiquette $tag posée sur la version déjà publiée."
        Remove-UntaggedTags
        Write-Host "✔ SageBridge $ver publié : https://github.com/$Repo/releases/tag/$tag" -ForegroundColor Green
        Write-Host 'PUBLICATION OK'
        return
    }

    $already = $true
    try { Api 'GET' "https://api.github.com/repos/$Repo/releases/tags/$tag" | Out-Null }
    catch { if ((StatusOf $_) -eq 404) { $already = $false } else { throw } }
    if ($already) { Remove-UntaggedTags; throw "La version $ver est déjà publiée. Augmentez AssemblyVersion dans SageBridge.cs, ajoutez la section « ## $ver » au CHANGELOG.md, relancez build.cmd puis ce script." }

    Step 'Envoi du code source (API GitHub, branche principale)'
    # Envoi fichier par fichier par l'API REST (plus fiable qu'un git push sur certaines connexions).
    $sha1 = [Security.Cryptography.SHA1]::Create()
    $files = Get-ChildItem -Path $Root -Recurse -File -Force | Where-Object {
        $_.FullName -notmatch '\\(bin|obj|\.git|\.vs)\\' -and $_.Extension -ne '.log'
    }
    $sent = 0
    foreach ($f in $files) {
        $rel = $f.FullName.Substring($Root.Length + 1).Replace('\', '/')
        $bytes = [IO.File]::ReadAllBytes($f.FullName)
        # Empreinte « blob » de Git : permet de ne renvoyer que les fichiers modifiés.
        $head = [Text.Encoding]::ASCII.GetBytes("blob $($bytes.Length)`0")
        $blob = New-Object byte[] ($head.Length + $bytes.Length)
        [Array]::Copy($head, $blob, $head.Length); [Array]::Copy($bytes, 0, $blob, $head.Length, $bytes.Length)
        $local = -join ($sha1.ComputeHash($blob) | ForEach-Object { $_.ToString('x2') })
        $url = "https://api.github.com/repos/$Repo/contents/" + [Uri]::EscapeUriString($rel)
        $remote = $null
        try { $remote = (Api 'GET' $url).sha } catch { if ((StatusOf $_) -ne 404) { throw } }
        if ($remote -eq $local) { continue }
        $body = @{ message = "SageBridge $ver — $rel"; content = [Convert]::ToBase64String($bytes) }
        if ($remote) { $body.sha = $remote }
        Api 'PUT' $url $body | Out-Null
        Write-Host "   $rel"
        $sent++
    }
    Write-Host "   $sent fichier(s) envoyé(s)."

    Step "Release $tag"
    $notes = "SageBridge $ver"
    $changelog = Join-Path $Root 'CHANGELOG.md'
    if (Test-Path $changelog) {
        $text = [IO.File]::ReadAllText($changelog, [Text.Encoding]::UTF8)
        $m = [regex]::Match($text, "(?ms)^## \[?$([regex]::Escape($ver))\]?[^\r\n]*\r?\n(.*?)(?=^## |\z)")
        if ($m.Success) { $notes = $m.Groups[1].Value.Trim() }
    }
    $notes += "`n`n---`n**Installation** : téléchargez ``SageBridge-Setup.exe`` ci-dessous et lancez-le (Windows peut afficher « Windows a protégé votre ordinateur » : *Informations complémentaires* → *Exécuter quand même*). Les SageBridge déjà installés se mettent à jour tout seuls."
    # Reprend un brouillon laissé par une publication interrompue, sinon en crée un.
    $rel = @(Api 'GET' "https://api.github.com/repos/$Repo/releases?per_page=50") | Where-Object { $_.draft -and $_.tag_name -eq $tag } | Select-Object -First 1
    if ($rel) {
        Write-Host "   Reprise du brouillon existant."
        Api 'PATCH' "https://api.github.com/repos/$Repo/releases/$($rel.id)" @{ tag_name = $tag; target_commitish = 'main'; name = "SageBridge $ver"; body = $notes } | Out-Null
    } else {
        $rel = Api 'POST' "https://api.github.com/repos/$Repo/releases" @{
            tag_name = $tag; target_commitish = 'main'; name = "SageBridge $ver"; body = $notes; draft = $true; prerelease = $false
        }
    }
    $upload = $rel.upload_url -replace '\{.*\}$', ''
    $assets = @(@($setup, 'SageBridge-Setup.exe'), @($sha, 'SageBridge-Setup.exe.sha256'))
    if (Test-Path $pdf) { $assets += , @($pdf, 'SageBridge-Documentation.pdf') }
    $existing = @(Api 'GET' "https://api.github.com/repos/$Repo/releases/$($rel.id)/assets")
    foreach ($a in $assets) {
        foreach ($old in ($existing | Where-Object { $_.name -eq $a[1] })) {
            Api 'DELETE' "https://api.github.com/repos/$Repo/releases/assets/$($old.id)" | Out-Null
        }
        $size = [Math]::Round((Get-Item $a[0]).Length / 1KB)
        for ($try = 1; $try -le 4; $try++) {
            Write-Host "   Envoi de $($a[1]) ($size Ko)… essai $try"
            try {
                Invoke-RestMethod -Method Post -Uri ("${upload}?name=" + [Uri]::EscapeDataString($a[1])) -Headers $script:H `
                    -ContentType 'application/octet-stream' -InFile $a[0] -TimeoutSec 600 | Out-Null
                $ok = $true; break
            }
            catch {
                $ok = $false
                Write-Host "   échec : $($_.Exception.Message)"
                # Un envoi interrompu peut laisser un fichier incomplet : on le supprime avant de réessayer.
                foreach ($bad in (@(Api 'GET' "https://api.github.com/repos/$Repo/releases/$($rel.id)/assets") | Where-Object { $_.name -eq $a[1] })) {
                    Api 'DELETE' "https://api.github.com/repos/$Repo/releases/assets/$($bad.id)" | Out-Null
                }
                Start-Sleep -Seconds 5
            }
        }
        if (-not $ok) { throw "Impossible d'envoyer $($a[1]) : relancez publier-version.cmd (le brouillon sera repris)." }
    }
    Api 'PATCH' "https://api.github.com/repos/$Repo/releases/$($rel.id)" @{ tag_name = $tag; target_commitish = 'main'; draft = $false; make_latest = 'true' } | Out-Null

    Remove-UntaggedTags
    Write-Host ''
    Write-Host "✔ SageBridge $ver publié : https://github.com/$Repo/releases/tag/$tag" -ForegroundColor Green
    Write-Host 'PUBLICATION OK'
}
catch {
    Write-Host ''
    Write-Host "✖ Échec : $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'PUBLICATION ECHEC'
}
finally {
    Stop-Transcript | Out-Null
}
