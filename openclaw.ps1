#!/usr/bin/env pwsh
$basedir=Split-Path $MyInvocation.MyCommand.Definition -Parent

$exe=""
$pathsep=":"
$env_node_path=$env:NODE_PATH
$new_node_path="C:\Users\MINI-PC\AppData\Local\pnpm\global\5\.pnpm\openclaw@2026.3.13_@napi-rs_e5f2bb93bf52304d09fd1d7f29f705c6\node_modules\openclaw\node_modules;C:\Users\MINI-PC\AppData\Local\pnpm\global\5\.pnpm\openclaw@2026.3.13_@napi-rs_e5f2bb93bf52304d09fd1d7f29f705c6\node_modules;C:\Users\MINI-PC\AppData\Local\pnpm\global\5\.pnpm\node_modules"
if ($PSVersionTable.PSVersion -lt "6.0" -or $IsWindows) {
  # Fix case when both the Windows and Linux builds of Node
  # are installed in the same directory
  $exe=".exe"
  $pathsep=";"
} else {
  $new_node_path="/mnt/c/Users/MINI-PC/AppData/Local/pnpm/global/5/.pnpm/openclaw@2026.3.13_@napi-rs_e5f2bb93bf52304d09fd1d7f29f705c6/node_modules/openclaw/node_modules:/mnt/c/Users/MINI-PC/AppData/Local/pnpm/global/5/.pnpm/openclaw@2026.3.13_@napi-rs_e5f2bb93bf52304d09fd1d7f29f705c6/node_modules:/mnt/c/Users/MINI-PC/AppData/Local/pnpm/global/5/.pnpm/node_modules"
}
if ([string]::IsNullOrEmpty($env_node_path)) {
  $env:NODE_PATH=$new_node_path
} else {
  $env:NODE_PATH="$new_node_path$pathsep$env_node_path"
}

$ret=0
if (Test-Path "$basedir/node$exe") {
  # Support pipeline input
  if ($MyInvocation.ExpectingInput) {
    $input | & "$basedir/node$exe"  "$basedir/global/5/.pnpm/openclaw@2026.3.13_@napi-rs_e5f2bb93bf52304d09fd1d7f29f705c6/node_modules/openclaw/openclaw.mjs" $args
  } else {
    & "$basedir/node$exe"  "$basedir/global/5/.pnpm/openclaw@2026.3.13_@napi-rs_e5f2bb93bf52304d09fd1d7f29f705c6/node_modules/openclaw/openclaw.mjs" $args
  }
  $ret=$LASTEXITCODE
} else {
  # Support pipeline input
  if ($MyInvocation.ExpectingInput) {
    $input | & "node$exe"  "$basedir/global/5/.pnpm/openclaw@2026.3.13_@napi-rs_e5f2bb93bf52304d09fd1d7f29f705c6/node_modules/openclaw/openclaw.mjs" $args
  } else {
    & "node$exe"  "$basedir/global/5/.pnpm/openclaw@2026.3.13_@napi-rs_e5f2bb93bf52304d09fd1d7f29f705c6/node_modules/openclaw/openclaw.mjs" $args
  }
  $ret=$LASTEXITCODE
}
$env:NODE_PATH=$env_node_path
exit $ret
