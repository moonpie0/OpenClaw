<#
.SYNOPSIS
    Completely uninstall OpenClaw and related components
.DESCRIPTION
    Stops and removes the OpenClaw Gateway service, uninstalls the global npm package,
    and deletes the user configuration and pnpm cache directories.
.NOTES
    It is recommended to run this script as an administrator.
#>

Write-Host "Stopping OpenClaw Gateway service..." -ForegroundColor Cyan
openclaw gateway stop

Write-Host "Uninstalling OpenClaw Gateway service..." -ForegroundColor Cyan
openclaw gateway uninstall

Write-Host "Uninstalling global npm package openclaw..." -ForegroundColor Cyan
npm uninstall -g openclaw

Write-Host "Removing user configuration directory .openclaw..." -ForegroundColor Cyan
Remove-Item -Recurse -Force "$env:USERPROFILE\.openclaw" -ErrorAction SilentlyContinue

Write-Host "Removing pnpm related directories..." -ForegroundColor Cyan
Remove-Item -Recurse -Force "$env:USERPROFILE\AppData\Local\pnpm" -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force "$env:USERPROFILE\AppData\Local\pnpm-cache" -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force "$env:USERPROFILE\AppData\Local\pnpm-state" -ErrorAction SilentlyContinue

Write-Host "Uninstall completed!" -ForegroundColor Green