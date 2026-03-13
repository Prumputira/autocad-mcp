# AutoCAD MCP Server — LISP auto-load configurator
# Adds mcp_dispatch.lsp to AutoCAD's Startup Suite (same as APPLOAD)
# and adds lisp-code/ to support paths and trusted paths.
# Called by setup.bat. Expects ACAD_DIR environment variable.
#
# To undo: run with -Uninstall flag.

param([switch]$Uninstall)

$lispDir = $env:ACAD_DIR + '\lisp-code'
$lispFile = $lispDir + '\mcp_dispatch.lsp'
if (-not $Uninstall -and -not (Test-Path $lispFile)) {
  Write-Host '  [FEL] mcp_dispatch.lsp hittades inte i:' $lispDir
  exit 1
}

$acadRoot = 'HKCU:\SOFTWARE\Autodesk\AutoCAD'
if (-not (Test-Path $acadRoot)) {
  Write-Host '  [INFO] Ingen AutoCAD-installation hittades i registret'
  exit 0
}

$found = $false
$lispDirLower = $lispDir.ToLower()
$lispFileLower = $lispFile.ToLower()
$errors = 0

foreach ($ver in (Get-ChildItem $acadRoot -ErrorAction SilentlyContinue)) {
  foreach ($prod in (Get-ChildItem $ver.PSPath -ErrorAction SilentlyContinue)) {
    if ($prod.PSChildName -notmatch '^ACAD-') { continue }

    $profRoot = Join-Path $prod.PSPath 'Profiles'
    if (-not (Test-Path $profRoot)) { continue }

    foreach ($profile in (Get-ChildItem $profRoot -ErrorAction SilentlyContinue)) {
      $genKey = Join-Path $profile.PSPath 'General'
      if (-not (Test-Path $genKey)) { continue }

      $acadPaths = (Get-ItemProperty $genKey -Name 'ACAD' -ErrorAction SilentlyContinue).ACAD
      if (-not $acadPaths) { continue }

      # Strip trailing semicolons
      $acadPaths = $acadPaths.TrimEnd(';')

      if ($Uninstall) {
        # --- Remove lisp-code from support paths ---
        $pathList = $acadPaths -split ';'
        $filtered = $pathList | Where-Object { $_.ToLower().Trim() -ne $lispDirLower -and $_.Trim() -ne '' }
        $newPaths = ($filtered -join ';')
        if ($newPaths -ne $acadPaths) {
          Set-ItemProperty $genKey -Name 'ACAD' -Value $newPaths
          if ($?) { Write-Host "  [OK] Tog bort lisp-code fran sokvagar for profil: $($profile.PSChildName)" }
          else { Write-Host "  [FEL] Kunde inte uppdatera sokvagar" -ForegroundColor Red; $errors++ }
        }

        # --- Remove lisp-code from trusted paths ---
        $varKey = Join-Path $profile.PSPath 'Variables'
        if (Test-Path $varKey) {
          $trusted = (Get-ItemProperty $varKey -Name 'TRUSTEDPATHS' -ErrorAction SilentlyContinue).TRUSTEDPATHS
          if ($trusted) {
            $trusted = $trusted.TrimEnd(';')
            $tList = $trusted -split ';'
            $tFiltered = $tList | Where-Object { $_.ToLower().Trim() -ne $lispDirLower -and $_.Trim() -ne '' }
            $newTrusted = ($tFiltered -join ';')
            if ($newTrusted -ne $trusted) {
              Set-ItemProperty $varKey -Name 'TRUSTEDPATHS' -Value $newTrusted
              if ($?) { Write-Host "  [OK] Tog bort lisp-code fran betrodda sokvagar" }
              else { Write-Host "  [FEL] Kunde inte uppdatera betrodda sokvagar" -ForegroundColor Red; $errors++ }
            }
          }
        }

        # --- Remove mcp_dispatch.lsp from Startup Suite ---
        $startupKey = Join-Path $profile.PSPath 'Dialogs\Appload\Startup'
        if (Test-Path $startupKey) {
          $props = Get-ItemProperty $startupKey -ErrorAction SilentlyContinue
          $numStartup = $props.NumStartup
          if ($numStartup -gt 0) {
            $kept = @()
            for ($i = 1; $i -le $numStartup; $i++) {
              $val = $props."${i}Startup"
              if ($val -and $val.ToLower() -notmatch 'mcp_dispatch\.lsp') {
                $kept += $val
              }
            }
            # Clear old entries
            for ($i = 1; $i -le $numStartup; $i++) {
              Remove-ItemProperty $startupKey -Name "${i}Startup" -ErrorAction SilentlyContinue
            }
            # Write back kept entries
            Set-ItemProperty $startupKey -Name 'NumStartup' -Value $kept.Count
            for ($i = 0; $i -lt $kept.Count; $i++) {
              $idx = $i + 1
              New-ItemProperty $startupKey -Name "${idx}Startup" -Value $kept[$i] -PropertyType String -Force | Out-Null
            }
            if ($kept.Count -lt $numStartup) {
              Write-Host "  [OK] Tog bort mcp_dispatch.lsp fran Startup Suite for profil: $($profile.PSChildName)"
            }
          }
        }
      } else {
        # --- Add lisp-code to support paths ---
        $pathList = $acadPaths -split ';'
        $alreadyInPath = $pathList | Where-Object { $_.ToLower().Trim() -eq $lispDirLower }
        if (-not $alreadyInPath) {
          $newPaths = $lispDir + ';' + $acadPaths
          Set-ItemProperty $genKey -Name 'ACAD' -Value $newPaths
          if ($?) { Write-Host "  [OK] Lade till lisp-code i sokvagar for profil: $($profile.PSChildName)" }
          else { Write-Host "  [FEL] Kunde inte uppdatera sokvagar" -ForegroundColor Red; $errors++ }
        } else {
          Write-Host "  [OK] lisp-code redan i sokvagar for profil: $($profile.PSChildName)"
        }

        # --- Add lisp-code to trusted paths ---
        $varKey = Join-Path $profile.PSPath 'Variables'
        if (Test-Path $varKey) {
          $trusted = (Get-ItemProperty $varKey -Name 'TRUSTEDPATHS' -ErrorAction SilentlyContinue).TRUSTEDPATHS
          if ($trusted) {
            $trusted = $trusted.TrimEnd(';')
            $tList = $trusted -split ';'
            $alreadyTrusted = $tList | Where-Object { $_.ToLower().Trim() -eq $lispDirLower }
            if (-not $alreadyTrusted) {
              Set-ItemProperty $varKey -Name 'TRUSTEDPATHS' -Value ($trusted + ';' + $lispDir)
              if ($?) { Write-Host "  [OK] Lade till lisp-code som betrodd sokvag" }
              else { Write-Host "  [FEL] Kunde inte uppdatera betrodda sokvagar" -ForegroundColor Red; $errors++ }
            } else {
              Write-Host "  [OK] lisp-code redan betrodd"
            }
          } else {
            New-ItemProperty $varKey -Name 'TRUSTEDPATHS' -Value $lispDir -PropertyType String -Force | Out-Null
            if ($?) { Write-Host "  [OK] Skapade betrodd sokvag for lisp-code" }
            else { Write-Host "  [FEL] Kunde inte skapa betrodd sokvag" -ForegroundColor Red; $errors++ }
          }
        }

        # --- Add mcp_dispatch.lsp to Startup Suite ---
        $startupKey = Join-Path $profile.PSPath 'Dialogs\Appload\Startup'
        if (-not (Test-Path $startupKey)) {
          New-Item $startupKey -Force | Out-Null
          Set-ItemProperty $startupKey -Name 'NumStartup' -Value 0
        }
        $props = Get-ItemProperty $startupKey -ErrorAction SilentlyContinue
        $numStartup = [int]($props.NumStartup)
        # Check if already registered
        $alreadyRegistered = $false
        for ($i = 1; $i -le $numStartup; $i++) {
          $val = $props."${i}Startup"
          if ($val -and $val.ToLower() -match 'mcp_dispatch\.lsp') {
            $alreadyRegistered = $true
            break
          }
        }
        if (-not $alreadyRegistered) {
          $newIdx = $numStartup + 1
          Set-ItemProperty $startupKey -Name 'NumStartup' -Value $newIdx
          New-ItemProperty $startupKey -Name "${newIdx}Startup" -Value $lispFile -PropertyType String -Force | Out-Null
          if ($?) { Write-Host "  [OK] Lade till mcp_dispatch.lsp i Startup Suite for profil: $($profile.PSChildName)" }
          else { Write-Host "  [FEL] Kunde inte uppdatera Startup Suite" -ForegroundColor Red; $errors++ }
        } else {
          Write-Host "  [OK] mcp_dispatch.lsp redan i Startup Suite for profil: $($profile.PSChildName)"
        }
      }

      $found = $true
    }
  }
}

if (-not $found) {
  Write-Host '  [INFO] Inga AutoCAD-profiler hittades. Ladda mcp_dispatch.lsp manuellt via APPLOAD.'
}

if ($errors -gt 0) {
  Write-Host "  [VARNING] $errors registerskrivningar misslyckades." -ForegroundColor Yellow
  exit 1
}
