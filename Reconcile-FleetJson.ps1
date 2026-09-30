<#
.SYNOPSIS
    Reconcile-FleetJson.ps1 - Reconcilie les JSON du parc apres durcissement ACL.

.DESCRIPTION
    A executer periodiquement sur le SERVEUR (tache planifiee), sous un compte
    ayant Modify (donc suppression/deplacement) sur la racine du share + pending\
    + archive\. Deux regles :

      1. PROMOTION. Un poste REIMAGE porte un nouveau SID AD : il n'est plus
         proprietaire de son ancien <PC>.json (cree par l'ancien SID) et, sous
         l'ACL durcie, ne peut plus l'ecraser. Son Collector (>= 2.4.9) depose
         alors le rapport dans pending\<PC>.json. Ici on detecte le binome
         "root du meme nom PERIME (en arret de cycle) + pending FRAIS (actif)" et
         on PROMEUT : suppression de l'ancien root, puis RENOMMAGE du pending vers
         la racine. Le renommage preserve le proprietaire (nouveau SID) -> le poste
         reecrase son root normalement aux cycles suivants.

      2. MENAGE. Un <PC>.json de la racine PERIME depuis > ArchiveDays ET SANS
         pending associe (= poste decommissionne / definitivement parti) est
         deplace dans archive\, puis purge apres ArchivePurgeDays supplementaires.

    IMPORTANT : le seuil "root perime" (StaleHours) ne declenche RIEN tout seul.
    Il n'est evalue que DANS le binome de la regle 1, c'est-a-dire uniquement pour
    un <PC> qui a AUSSI un pending frais du meme nom. Un poste juste offline (sans
    pending) n'est jamais touche par la regle 1 ; il ne part en archive (regle 2)
    qu'apres ArchiveDays (30j par defaut).

    Idempotent, et -WhatIf montre ce qui SERAIT fait sans rien modifier (a utiliser
    pour les essais en bac a sable).

.PARAMETER SharePath
    Racine des JSON. Sur le serveur, prefere le chemin LOCAL (ex: 'D:\PCPulse').
.PARAMETER StaleHours
    Age min du root pour le juger "en arret de cycle" DANS un binome (defaut 3 =
    ~2 cycles horaires + delai anti-collision + marge).
.PARAMETER PendingFreshHours
    Age max du pending pour le juger "frais/actif" (defaut 2).
.PARAMETER ArchiveDays
    Age du root SANS pending avant archivage (defaut 30).
.PARAMETER ArchivePurgeDays
    Retention dans archive\ (a partir de la date d'archivage) avant suppression (defaut 30).

.EXAMPLE
    .\Reconcile-FleetJson.ps1 -SharePath 'D:\PCPulse' -WhatIf   # essai a blanc
.EXAMPLE
    .\Reconcile-FleetJson.ps1 -SharePath 'D:\PCPulse'          # execution reelle
#>

#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$SharePath        = '\\SERVER\PCPulse$',
    [int]   $StaleHours       = 3,
    [int]   $PendingFreshHours = 2,
    [int]   $ArchiveDays      = 30,
    [int]   $ArchivePurgeDays = 30
)

$ErrorActionPreference = 'Stop'
$now        = Get-Date
$rootDir    = $SharePath
$pendingDir = Join-Path $SharePath 'pending'
$archiveDir = Join-Path $SharePath 'archive'

if (-not (Test-Path -LiteralPath $rootDir)) { Write-Error "SharePath introuvable : $rootDir"; exit 1 }

# Journal dans archive\ (le compte de reconcile y a Modify ; jamais purge car .log)
function Write-RecLog {
    param([string]$Message, [string]$Level = 'INFO')
    $line = ('{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message)
    Write-Host $line
    try {
        if (-not (Test-Path -LiteralPath $archiveDir)) { $null = New-Item -ItemType Directory -Path $archiveDir -Force -ErrorAction SilentlyContinue }
        Add-Content -LiteralPath (Join-Path $archiveDir 'reconcile.log') -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {}
}

$promoted = 0; $archived = 0; $purged = 0

# ============================================================
# REGLE 1 : PROMOTION (pending frais + root du meme nom perime)
# ============================================================
if (Test-Path -LiteralPath $pendingDir) {
    foreach ($pend in @(Get-ChildItem -LiteralPath $pendingDir -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
        if ($pend.Name -like '*.tmp') { continue }

        # Le pending doit etre FRAIS (le poste reimage reporte activement).
        if (($now - $pend.LastWriteTime).TotalHours -ge $PendingFreshHours) { continue }

        $rootFile   = Join-Path $rootDir $pend.Name
        $rootExists = Test-Path -LiteralPath $rootFile
        $doPromote  = $false
        if (-not $rootExists) {
            # Plus de root (deja archive/supprime) : on promeut directement.
            $doPromote = $true
        } else {
            $rootAgeH = ($now - (Get-Item -LiteralPath $rootFile).LastWriteTime).TotalHours
            # Root en arret de cycle depuis >= StaleHours ALORS que le pending est
            # frais -> binome "meme nom, l'un mort l'autre vivant" = reimage confirme.
            if ($rootAgeH -ge $StaleHours) { $doPromote = $true }
        }

        if ($doPromote) {
            if ($PSCmdlet.ShouldProcess($pend.Name, "Promouvoir pending -> racine")) {
                try {
                    if ($rootExists) { Remove-Item -LiteralPath $rootFile -Force }
                    # MOVE (rename) : preserve le proprietaire du pending (nouveau SID).
                    Move-Item -LiteralPath $pend.FullName -Destination $rootFile -Force
                    Write-RecLog "PROMU : $($pend.Name) (pending -> racine)" 'OK'
                    $promoted++
                } catch { Write-RecLog "Echec promotion $($pend.Name) : $_" 'ERREUR' }
            } else { $promoted++ }   # comptabilise le WOULD en -WhatIf
        }
    }
}

# ============================================================
# REGLE 2 : MENAGE (root perime > ArchiveDays ET SANS pending)
# ============================================================
foreach ($root in @(Get-ChildItem -LiteralPath $rootDir -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
    if ($root.Name -like '*.tmp') { continue }
    # Un root qui a un pending releve de la regle 1 (jamais archive).
    if (Test-Path -LiteralPath (Join-Path $pendingDir $root.Name)) { continue }

    $ageDays = ($now - $root.LastWriteTime).TotalDays
    if ($ageDays -ge $ArchiveDays) {
        if ($PSCmdlet.ShouldProcess($root.Name, "Archiver (perime $([int]$ageDays)j)")) {
            try {
                if (-not (Test-Path -LiteralPath $archiveDir)) { $null = New-Item -ItemType Directory -Path $archiveDir -Force }
                $dest = Join-Path $archiveDir $root.Name
                if (Test-Path -LiteralPath $dest) { Remove-Item -LiteralPath $dest -Force }
                Move-Item -LiteralPath $root.FullName -Destination $dest -Force
                # Horodater a MAINTENANT : le compte de retention (ArchivePurgeDays)
                # part de la date d'ARCHIVAGE, pas du dernier report (deja ancien).
                (Get-Item -LiteralPath $dest).LastWriteTime = $now
                Write-RecLog "ARCHIVE : $($root.Name) (perime $([int]$ageDays)j)" 'OK'
                $archived++
            } catch { Write-RecLog "Echec archivage $($root.Name) : $_" 'ERREUR' }
        } else { $archived++ }
    }
}

# ============================================================
# PURGE archive\ (au-dela de ArchivePurgeDays a partir de l'archivage)
# ============================================================
if (Test-Path -LiteralPath $archiveDir) {
    foreach ($old in @(Get-ChildItem -LiteralPath $archiveDir -Filter '*.json' -File -ErrorAction SilentlyContinue |
                        Where-Object { ($now - $_.LastWriteTime).TotalDays -ge $ArchivePurgeDays })) {
        if ($PSCmdlet.ShouldProcess($old.Name, "Purger archive")) {
            try { Remove-Item -LiteralPath $old.FullName -Force; Write-RecLog "PURGE archive : $($old.Name)"; $purged++ }
            catch { Write-RecLog "Echec purge $($old.Name) : $_" 'ERREUR' }
        } else { $purged++ }
    }
}

Write-RecLog ("Termine : {0} promu(s), {1} archive(s), {2} purge(s)." -f $promoted, $archived, $purged) 'OK'
