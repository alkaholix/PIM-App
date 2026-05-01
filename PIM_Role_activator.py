#!/usr/bin/env python3
"""
PIM Role Activator - A GUI application to activate Azure AD roles via PIM
without needing to access the Azure web portal.

All PowerShell scripts are embedded directly in this application.

Usage:
    python pim_activator.py

Requirements:
    - Python 3.8+
    - PowerShell 7.0+
    - Microsoft.Graph PowerShell module (auto-installed by scripts)
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import subprocess
import tempfile
import os
import sys
import threading
import secrets
import string
import json

# Role mappings: (Display Name, PowerShell Script Content)
# Each script is embedded directly to avoid file path issues
# Note: Using r""" raw strings to avoid Python escape sequence warnings
ROLE_SCRIPTS = {
    "Authentication Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Authentication Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Authentication Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Cloud Device Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Cloud Device Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Cloud Device Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Exchange Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Exchange Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Exchange Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Fabric Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Fabric Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Fabric Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Global Reader": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Global Reader'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Global Reader'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Groups Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Groups Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Groups Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Helpdesk Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Helpdesk Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Helpdesk Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Intune Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Intune Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Intune Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Microsoft 365 Backup Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Microsoft 365 Backup Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Microsoft 365 Backup Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Microsoft 365 Migration Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Microsoft 365 Migration Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Microsoft 365 Migration Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Microsoft Entra Joined Device Local Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Microsoft Entra Joined Device Local Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Microsoft Entra Joined Device Local Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Office Apps Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Office Apps Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Office Apps Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Power Platform Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Power Platform Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Power Platform Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Privileged Authentication Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Privileged Authentication Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Privileged Authentication Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Reports Reader": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Reports Reader'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Reports Reader'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Security Reader": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Security Reader'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Security Reader'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "SharePoint Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'SharePoint Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'SharePoint Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "SharePoint Embedded Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'SharePoint Embedded Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'SharePoint Embedded Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Teams Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Teams Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Teams Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Teams Communications Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Teams Communications Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Teams Communications Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Teams Devices Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Teams Devices Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Teams Devices Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "Teams Telephony Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'Teams Telephony Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'Teams Telephony Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "User Administrator": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][ValidatePattern('^PT(\d+H)?(\d+M)?(\d+S)?$')][string]$Duration,
  [Parameter(Mandatory)][string]$Justification
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop
}
function Connect-GraphForPIM {
  $scopes = @('User.Read','RoleEligibilitySchedule.Read.Directory','RoleAssignmentSchedule.ReadWrite.Directory')
  if (-not (Get-MgContext)) { Connect-MgGraph -Scopes $scopes | Out-Null }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
function Get-CurrentUserId {
  $ctx = Get-MgContext
  (Get-MgUser -UserId $ctx.Account).Id
}
function Get-EligibleRoleRecord {
  param([Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDisplayName)
  $eligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty RoleDefinition -Filter "principalId eq '$PrincipalId'"
  $match = $eligible | Where-Object { $_.RoleDefinition.DisplayName -ieq $RoleDisplayName }
  return $match
}
function New-ActivationRequest {
  param(
    [Parameter(Mandatory)][string]$PrincipalId, [Parameter(Mandatory)][string]$RoleDefinitionId,
    [Parameter(Mandatory)][string]$DirectoryScopeId, [Parameter(Mandatory)][string]$Justification,
    [Parameter(Mandatory)][string]$Duration, [string]$TicketNumber, [string]$TicketSystem
  )
  $body = @{
    Action='selfActivate', PrincipalId=$PrincipalId, RoleDefinitionId=$RoleDefinitionId,
    DirectoryScopeId=$DirectoryScopeId, Justification=$Justification,
    ScheduleInfo=@{ StartDateTime=(Get-Date); Expiration=@{ Type='AfterDuration'; Duration=$Duration } }
  }
  if ($TicketNumber -or $TicketSystem) { $body.TicketInfo = @{ TicketNumber=$TicketNumber; TicketSystem=$TicketSystem } }
  New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $body
}
$RoleDisplayName = 'User Administrator'
try {
  Ensure-GraphModules
  Connect-GraphForPIM
  $me = Get-CurrentUserId
  Write-Host "Activating role: 'User Administrator'" -ForegroundColor Cyan
  $rec = Get-EligibleRoleRecord -PrincipalId $me -RoleDisplayName $RoleDisplayName
  if (-not $rec) { throw "You are not eligible for role '$RoleDisplayName'." }
  $res = New-ActivationRequest -PrincipalId $me -RoleDefinitionId $rec.RoleDefinitionId -DirectoryScopeId $rec.DirectoryScopeId -Justification $Justification -Duration $Duration
  $obj = [pscustomobject]@{ Role=$RoleDisplayName; RequestId=$res.Id; Status=$res.Status; Action=$res.Action; Scope=$rec.DirectoryScopeId }
  $obj | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',
}


USERS_SCRIPTS = {
    "search": r'''#requires -Version 7.0
[CmdletBinding()]
param([Parameter(Mandatory)][string]$Query)
function Ensure-GraphModules {
  $mods = @('Microsoft.Graph.Users')
  foreach ($m in $mods) {
    if (-not (Get-Module -ListAvailable $m)) {
      Write-Host "Installing $m..." -ForegroundColor Yellow
      Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $m -ErrorAction Stop
  }
}
function Connect-GraphForAdmin {
  param([string[]]$Scopes)
  $ctx = Get-MgContext
  if (-not $ctx -or ($Scopes | Where-Object { $_ -notin $ctx.Scopes })) {
    Connect-MgGraph -Scopes $Scopes | Out-Null
  }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
try {
  Ensure-GraphModules
  Connect-GraphForAdmin -Scopes @('User.Read.All')
  $users = Get-MgUser -Filter "startsWith(userPrincipalName,'$Query') or startsWith(displayName,'$Query')" `
    -Property Id,DisplayName,UserPrincipalName,AccountEnabled,JobTitle,Department `
    -Top 20 -ConsistencyLevel eventual -CountVariable ignored
  if (-not $users) { Write-Host "No users found matching '$Query'" -ForegroundColor Yellow; exit 0 }
  $results = $users | ForEach-Object {
    [pscustomobject]@{
      DisplayName = $_.DisplayName
      UPN         = $_.UserPrincipalName
      Enabled     = $_.AccountEnabled
      Title       = $_.JobTitle
      Department  = $_.Department
    }
  }
  $results | Format-Table -AutoSize
  Write-Host "##JSON_START##"
  $results | ConvertTo-Json -Compress
  Write-Host "##JSON_END##"
}
catch { Write-Error $_; exit 1 }''',

    "toggle": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][string]$UPN,
  [Parameter(Mandatory)][ValidateSet('Enable','Disable')][string]$Action
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph.Users)) {
    Write-Host "Installing Microsoft.Graph.Users..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Users -ErrorAction Stop
}
function Connect-GraphForAdmin {
  param([string[]]$Scopes)
  $ctx = Get-MgContext
  if (-not $ctx -or ($Scopes | Where-Object { $_ -notin $ctx.Scopes })) {
    Connect-MgGraph -Scopes $Scopes | Out-Null
  }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
try {
  Ensure-GraphModules
  Connect-GraphForAdmin -Scopes @('User.ReadWrite.All')
  $user   = Get-MgUser -UserId $UPN -Property Id,DisplayName,AccountEnabled
  $enable = ($Action -eq 'Enable')
  Update-MgUser -UserId $user.Id -AccountEnabled:$enable
  $verb  = if ($enable) { 'ENABLED' } else { 'DISABLED' }
  $color = if ($enable) { 'Green'  } else { 'Yellow'  }
  Write-Host "Account $verb`: $($user.DisplayName) ($UPN)" -ForegroundColor $color
}
catch { Write-Error $_; exit 1 }''',

    "reset_password": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][string]$UPN,
  [Parameter(Mandatory)][string]$TempPassword
)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph.Users)) {
    Write-Host "Installing Microsoft.Graph.Users..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Users -ErrorAction Stop
}
function Connect-GraphForAdmin {
  param([string[]]$Scopes)
  $ctx = Get-MgContext
  if (-not $ctx -or ($Scopes | Where-Object { $_ -notin $ctx.Scopes })) {
    Connect-MgGraph -Scopes $Scopes | Out-Null
  }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
try {
  Ensure-GraphModules
  Connect-GraphForAdmin -Scopes @('User.ReadWrite.All')
  $params = @{
    PasswordProfile = @{
      Password                      = $TempPassword
      ForceChangePasswordNextSignIn = $true
    }
  }
  Update-MgUser -UserId $UPN -BodyParameter $params
  Write-Host "Password reset for: $UPN" -ForegroundColor Green
  Write-Host "Temporary password: $TempPassword" -ForegroundColor Cyan
  Write-Host "User must change password at next sign-in." -ForegroundColor Cyan
}
catch { Write-Error $_; exit 1 }''',

    "mfa_methods": r'''#requires -Version 7.0
[CmdletBinding()]
param([Parameter(Mandatory)][string]$UPN)
function Ensure-GraphModules {
  $mods = @('Microsoft.Graph.Users','Microsoft.Graph.Identity.SignIns')
  foreach ($m in $mods) {
    if (-not (Get-Module -ListAvailable $m)) {
      Write-Host "Installing $m..." -ForegroundColor Yellow
      Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $m -ErrorAction Stop
  }
}
function Connect-GraphForAdmin {
  param([string[]]$Scopes)
  $ctx = Get-MgContext
  if (-not $ctx -or ($Scopes | Where-Object { $_ -notin $ctx.Scopes })) {
    Connect-MgGraph -Scopes $Scopes | Out-Null
  }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
try {
  Ensure-GraphModules
  Connect-GraphForAdmin -Scopes @('UserAuthenticationMethod.Read.All')
  $userId = (Get-MgUser -UserId $UPN -Property Id).Id
  Write-Host "MFA authentication methods for: $UPN" -ForegroundColor Cyan
  Write-Host ""
  $methods = @()
  try { Get-MgUserAuthenticationFido2Method -UserId $userId | ForEach-Object {
    $methods += [pscustomobject]@{ Type='FIDO2 Security Key'; Detail=$_.DisplayName; Registered=$_.CreatedDateTime }
  }} catch {}
  try { Get-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $userId | ForEach-Object {
    $methods += [pscustomobject]@{ Type='Microsoft Authenticator'; Detail=$_.DisplayName; Registered=$_.CreatedDateTime }
  }} catch {}
  try { Get-MgUserAuthenticationPhoneMethod -UserId $userId | ForEach-Object {
    $methods += [pscustomobject]@{ Type="Phone ($($_.PhoneType))"; Detail=$_.PhoneNumber; Registered='N/A' }
  }} catch {}
  try { Get-MgUserAuthenticationEmailMethod -UserId $userId | ForEach-Object {
    $methods += [pscustomobject]@{ Type='Email OTP'; Detail=$_.EmailAddress; Registered='N/A' }
  }} catch {}
  if ($methods.Count -eq 0) { Write-Host "No MFA methods registered for this user." -ForegroundColor Yellow }
  else {
    $methods | Format-Table -AutoSize
    Write-Host "Total methods registered: $($methods.Count)" -ForegroundColor Cyan
  }
}
catch { Write-Error $_; exit 1 }''',

    "group_memberships": r'''#requires -Version 7.0
[CmdletBinding()]
param([Parameter(Mandatory)][string]$UPN)
function Ensure-GraphModules {
  $mods = @('Microsoft.Graph.Users','Microsoft.Graph.Groups')
  foreach ($m in $mods) {
    if (-not (Get-Module -ListAvailable $m)) {
      Write-Host "Installing $m..." -ForegroundColor Yellow
      Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $m -ErrorAction Stop
  }
}
function Connect-GraphForAdmin {
  param([string[]]$Scopes)
  $ctx = Get-MgContext
  if (-not $ctx -or ($Scopes | Where-Object { $_ -notin $ctx.Scopes })) {
    Connect-MgGraph -Scopes $Scopes | Out-Null
  }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
try {
  Ensure-GraphModules
  Connect-GraphForAdmin -Scopes @('User.Read.All','GroupMember.Read.All')
  $userId = (Get-MgUser -UserId $UPN -Property Id).Id
  Write-Host "Group memberships for: $UPN" -ForegroundColor Cyan
  Write-Host ""
  $memberships = Get-MgUserMemberOf -UserId $userId -All |
    Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.group' }
  if (-not $memberships) { Write-Host "User is not a member of any groups." -ForegroundColor Yellow }
  else {
    $memberships | ForEach-Object {
      $gt = $_.AdditionalProperties['groupTypes']
      [pscustomobject]@{
        DisplayName = $_.AdditionalProperties['displayName']
        Type        = if ($gt -contains 'Unified') {'M365'} elseif ($gt -contains 'DynamicMembership') {'Dynamic'} else {'Security'}
        Id          = $_.Id
      }
    } | Sort-Object DisplayName | Format-Table -AutoSize
    Write-Host "Total groups: $($memberships.Count)" -ForegroundColor Cyan
  }
}
catch { Write-Error $_; exit 1 }''',
}


GROUPS_SCRIPTS = {
    "search": r'''#requires -Version 7.0
[CmdletBinding()]
param([Parameter(Mandatory)][string]$Query)
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph.Groups)) {
    Write-Host "Installing Microsoft.Graph.Groups..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph.Groups -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Groups -ErrorAction Stop
}
function Connect-GraphForAdmin {
  param([string[]]$Scopes)
  $ctx = Get-MgContext
  if (-not $ctx -or ($Scopes | Where-Object { $_ -notin $ctx.Scopes })) {
    Connect-MgGraph -Scopes $Scopes | Out-Null
  }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
try {
  Ensure-GraphModules
  Connect-GraphForAdmin -Scopes @('Group.Read.All')
  $groups = Get-MgGroup -Filter "startsWith(displayName,'$Query')" `
    -Property Id,DisplayName,Description,GroupTypes,CreatedDateTime `
    -Top 20 -ConsistencyLevel eventual -CountVariable ignored
  if (-not $groups) { Write-Host "No groups found matching '$Query'" -ForegroundColor Yellow; exit 0 }
  $results = $groups | ForEach-Object {
    $gt = $_.GroupTypes
    [pscustomobject]@{
      DisplayName = $_.DisplayName
      Type        = if ($gt -contains 'Unified') {'M365'} elseif ($gt -contains 'DynamicMembership') {'Dynamic'} else {'Security'}
      Id          = $_.Id
    }
  }
  $results | Format-Table -AutoSize
  Write-Host "##JSON_START##"
  $results | ConvertTo-Json -Compress
  Write-Host "##JSON_END##"
}
catch { Write-Error $_; exit 1 }''',

    "get_members": r'''#requires -Version 7.0
[CmdletBinding()]
param([Parameter(Mandatory)][string]$GroupId)
function Ensure-GraphModules {
  $mods = @('Microsoft.Graph.Groups','Microsoft.Graph.Users')
  foreach ($m in $mods) {
    if (-not (Get-Module -ListAvailable $m)) {
      Write-Host "Installing $m..." -ForegroundColor Yellow
      Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $m -ErrorAction Stop
  }
}
function Connect-GraphForAdmin {
  param([string[]]$Scopes)
  $ctx = Get-MgContext
  if (-not $ctx -or ($Scopes | Where-Object { $_ -notin $ctx.Scopes })) {
    Connect-MgGraph -Scopes $Scopes | Out-Null
  }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
try {
  Ensure-GraphModules
  Connect-GraphForAdmin -Scopes @('GroupMember.Read.All')
  $group   = Get-MgGroup -GroupId $GroupId -Property DisplayName
  Write-Host "Members of: $($group.DisplayName)" -ForegroundColor Cyan
  Write-Host ""
  $members = Get-MgGroupMember -GroupId $GroupId -All
  if (-not $members) { Write-Host "Group has no members." -ForegroundColor Yellow }
  else {
    $members | ForEach-Object {
      $type = ($_.'@odata.type' -replace '#microsoft.graph.','').ToUpper()
      [pscustomobject]@{
        Type        = $type
        DisplayName = $_.AdditionalProperties['displayName']
        UPN         = $_.AdditionalProperties['userPrincipalName']
      }
    } | Sort-Object Type,DisplayName | Format-Table -AutoSize
    Write-Host "Total members: $($members.Count)" -ForegroundColor Cyan
  }
}
catch { Write-Error $_; exit 1 }''',

    "add_member": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][string]$GroupId,
  [Parameter(Mandatory)][string]$MemberUPN
)
function Ensure-GraphModules {
  $mods = @('Microsoft.Graph.Groups','Microsoft.Graph.Users')
  foreach ($m in $mods) {
    if (-not (Get-Module -ListAvailable $m)) {
      Write-Host "Installing $m..." -ForegroundColor Yellow
      Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $m -ErrorAction Stop
  }
}
function Connect-GraphForAdmin {
  param([string[]]$Scopes)
  $ctx = Get-MgContext
  if (-not $ctx -or ($Scopes | Where-Object { $_ -notin $ctx.Scopes })) {
    Connect-MgGraph -Scopes $Scopes | Out-Null
  }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
try {
  Ensure-GraphModules
  Connect-GraphForAdmin -Scopes @('GroupMember.ReadWrite.All','User.Read.All')
  $userId = (Get-MgUser -UserId $MemberUPN -Property Id).Id
  $group  = Get-MgGroup -GroupId $GroupId -Property DisplayName
  New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $userId
  Write-Host "Added $MemberUPN to group '$($group.DisplayName)'" -ForegroundColor Green
}
catch { Write-Error $_; exit 1 }''',

    "remove_member": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][string]$GroupId,
  [Parameter(Mandatory)][string]$MemberUPN
)
function Ensure-GraphModules {
  $mods = @('Microsoft.Graph.Groups','Microsoft.Graph.Users')
  foreach ($m in $mods) {
    if (-not (Get-Module -ListAvailable $m)) {
      Write-Host "Installing $m..." -ForegroundColor Yellow
      Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $m -ErrorAction Stop
  }
}
function Connect-GraphForAdmin {
  param([string[]]$Scopes)
  $ctx = Get-MgContext
  if (-not $ctx -or ($Scopes | Where-Object { $_ -notin $ctx.Scopes })) {
    Connect-MgGraph -Scopes $Scopes | Out-Null
  }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
try {
  Ensure-GraphModules
  Connect-GraphForAdmin -Scopes @('GroupMember.ReadWrite.All','User.Read.All')
  $userId = (Get-MgUser -UserId $MemberUPN -Property Id).Id
  $group  = Get-MgGroup -GroupId $GroupId -Property DisplayName
  Remove-MgGroupMemberByRef -GroupId $GroupId -DirectoryObjectId $userId
  Write-Host "Removed $MemberUPN from group '$($group.DisplayName)'" -ForegroundColor Yellow
}
catch { Write-Error $_; exit 1 }''',
}


LICENSES_SCRIPTS = {
    "tenant_inventory": r'''#requires -Version 7.0
[CmdletBinding()]
param()
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph.Identity.DirectoryManagement)) {
    Write-Host "Installing Microsoft.Graph.Identity.DirectoryManagement..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
}
function Connect-GraphForAdmin {
  param([string[]]$Scopes)
  $ctx = Get-MgContext
  if (-not $ctx -or ($Scopes | Where-Object { $_ -notin $ctx.Scopes })) {
    Connect-MgGraph -Scopes $Scopes | Out-Null
  }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
try {
  Ensure-GraphModules
  Connect-GraphForAdmin -Scopes @('Directory.Read.All')
  Write-Host "Tenant License Inventory" -ForegroundColor Cyan
  Write-Host ""
  $skus    = Get-MgSubscribedSku -All
  $results = $skus | ForEach-Object {
    [pscustomobject]@{
      'SKU Name'  = $_.SkuPartNumber
      'Total'     = $_.PrepaidUnits.Enabled
      'Assigned'  = $_.ConsumedUnits
      'Available' = $_.PrepaidUnits.Enabled - $_.ConsumedUnits
      'SKU ID'    = $_.SkuId
    }
  } | Sort-Object 'SKU Name'
  $results | Format-Table -AutoSize
  Write-Host "##JSON_START##"
  $results | ConvertTo-Json -Compress
  Write-Host "##JSON_END##"
}
catch { Write-Error $_; exit 1 }''',

    "user_licenses": r'''#requires -Version 7.0
[CmdletBinding()]
param([Parameter(Mandatory)][string]$UPN)
function Ensure-GraphModules {
  $mods = @('Microsoft.Graph.Users','Microsoft.Graph.Identity.DirectoryManagement')
  foreach ($m in $mods) {
    if (-not (Get-Module -ListAvailable $m)) {
      Write-Host "Installing $m..." -ForegroundColor Yellow
      Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $m -ErrorAction Stop
  }
}
function Connect-GraphForAdmin {
  param([string[]]$Scopes)
  $ctx = Get-MgContext
  if (-not $ctx -or ($Scopes | Where-Object { $_ -notin $ctx.Scopes })) {
    Connect-MgGraph -Scopes $Scopes | Out-Null
  }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
try {
  Ensure-GraphModules
  Connect-GraphForAdmin -Scopes @('User.Read.All','Directory.Read.All')
  $skuMap = @{}
  Get-MgSubscribedSku -All | ForEach-Object { $skuMap[$_.SkuId] = $_.SkuPartNumber }
  $user = Get-MgUser -UserId $UPN -Property Id,DisplayName,AssignedLicenses
  Write-Host "Licenses for: $($user.DisplayName) ($UPN)" -ForegroundColor Cyan
  Write-Host ""
  if (-not $user.AssignedLicenses -or $user.AssignedLicenses.Count -eq 0) {
    Write-Host "No licenses assigned." -ForegroundColor Yellow
  } else {
    $user.AssignedLicenses | ForEach-Object {
      [pscustomobject]@{
        License          = if ($skuMap[$_.SkuId]) { $skuMap[$_.SkuId] } else { $_.SkuId }
        'SKU ID'         = $_.SkuId
        'Disabled Plans' = $_.DisabledPlans.Count
      }
    } | Format-Table -AutoSize
  }
}
catch { Write-Error $_; exit 1 }''',

    "assign_license": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][string]$UPN,
  [Parameter(Mandatory)][string]$SkuId
)
function Ensure-GraphModules {
  $mods = @('Microsoft.Graph.Users','Microsoft.Graph.Identity.DirectoryManagement')
  foreach ($m in $mods) {
    if (-not (Get-Module -ListAvailable $m)) {
      Write-Host "Installing $m..." -ForegroundColor Yellow
      Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $m -ErrorAction Stop
  }
}
function Connect-GraphForAdmin {
  param([string[]]$Scopes)
  $ctx = Get-MgContext
  if (-not $ctx -or ($Scopes | Where-Object { $_ -notin $ctx.Scopes })) {
    Connect-MgGraph -Scopes $Scopes | Out-Null
  }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
try {
  Ensure-GraphModules
  Connect-GraphForAdmin -Scopes @('User.ReadWrite.All','Directory.ReadWrite.All')
  $addLicenses = @(@{ SkuId = $SkuId; DisabledPlans = @() })
  Set-MgUserLicense -UserId $UPN -AddLicenses $addLicenses -RemoveLicenses @()
  $skuName = (Get-MgSubscribedSku -All | Where-Object SkuId -eq $SkuId).SkuPartNumber
  Write-Host "Assigned license '$skuName' to $UPN" -ForegroundColor Green
}
catch { Write-Error $_; exit 1 }''',

    "remove_license": r'''#requires -Version 7.0
[CmdletBinding()]
param(
  [Parameter(Mandatory)][string]$UPN,
  [Parameter(Mandatory)][string]$SkuId
)
function Ensure-GraphModules {
  $mods = @('Microsoft.Graph.Users','Microsoft.Graph.Identity.DirectoryManagement')
  foreach ($m in $mods) {
    if (-not (Get-Module -ListAvailable $m)) {
      Write-Host "Installing $m..." -ForegroundColor Yellow
      Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $m -ErrorAction Stop
  }
}
function Connect-GraphForAdmin {
  param([string[]]$Scopes)
  $ctx = Get-MgContext
  if (-not $ctx -or ($Scopes | Where-Object { $_ -notin $ctx.Scopes })) {
    Connect-MgGraph -Scopes $Scopes | Out-Null
  }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
try {
  Ensure-GraphModules
  Connect-GraphForAdmin -Scopes @('User.ReadWrite.All','Directory.ReadWrite.All')
  $skuName = (Get-MgSubscribedSku -All | Where-Object SkuId -eq $SkuId).SkuPartNumber
  Set-MgUserLicense -UserId $UPN -AddLicenses @() -RemoveLicenses @($SkuId)
  Write-Host "Removed license '$skuName' from $UPN" -ForegroundColor Yellow
}
catch { Write-Error $_; exit 1 }''',
}


TENANT_SCRIPTS = {
    "overview": r'''#requires -Version 7.0
[CmdletBinding()]
param()
function Ensure-GraphModules {
  $mods = @('Microsoft.Graph.Identity.DirectoryManagement','Microsoft.Graph.Users','Microsoft.Graph.Groups')
  foreach ($m in $mods) {
    if (-not (Get-Module -ListAvailable $m)) {
      Write-Host "Installing $m..." -ForegroundColor Yellow
      Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $m -ErrorAction Stop
  }
}
function Connect-GraphForAdmin {
  param([string[]]$Scopes)
  $ctx = Get-MgContext
  if (-not $ctx -or ($Scopes | Where-Object { $_ -notin $ctx.Scopes })) {
    Connect-MgGraph -Scopes $Scopes | Out-Null
  }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
try {
  Ensure-GraphModules
  Connect-GraphForAdmin -Scopes @('Organization.Read.All','Directory.Read.All')
  $org = Get-MgOrganization | Select-Object -First 1
  Write-Host "TENANT OVERVIEW" -ForegroundColor Cyan
  Write-Host ("=" * 54) -ForegroundColor Cyan
  Write-Host ""
  [pscustomobject]@{
    'Tenant Name'  = $org.DisplayName
    'Tenant ID'    = $org.Id
    'Country'      = $org.CountryLetterCode
    'Created'      = $org.CreatedDateTime
    'Tech Contact' = ($org.TechnicalNotificationMails -join ', ')
  } | Format-List
  Write-Host ("─" * 54) -ForegroundColor Cyan
  Write-Host "Verified Domains" -ForegroundColor Cyan
  Write-Host ""
  $org.VerifiedDomains | ForEach-Object {
    [pscustomobject]@{ Domain=$_.Name; Default=$_.IsDefault; Type=$_.Type }
  } | Format-Table -AutoSize
  Write-Host ("─" * 54) -ForegroundColor Cyan
  Write-Host "Counts" -ForegroundColor Cyan
  Write-Host ""
  try {
    $uc = (Invoke-MgGraphRequest -Method GET -Uri '/v1.0/users/$count' -Headers @{'ConsistencyLevel'='eventual'})
    $gc = (Invoke-MgGraphRequest -Method GET -Uri '/v1.0/groups/$count' -Headers @{'ConsistencyLevel'='eventual'})
    Write-Host ("  Total Users  : {0}" -f $uc) -ForegroundColor Green
    Write-Host ("  Total Groups : {0}" -f $gc) -ForegroundColor Green
  } catch { Write-Host "  (Could not retrieve counts — check Directory.Read.All consent)" -ForegroundColor Yellow }
  Write-Host ""
  Write-Host ("─" * 54) -ForegroundColor Cyan
  Write-Host "License Summary" -ForegroundColor Cyan
  Write-Host ""
  Get-MgSubscribedSku -All | ForEach-Object {
    [pscustomobject]@{ SKU=$_.SkuPartNumber; Total=$_.PrepaidUnits.Enabled; Used=$_.ConsumedUnits; Available=($_.PrepaidUnits.Enabled - $_.ConsumedUnits) }
  } | Format-Table -AutoSize
}
catch { Write-Error $_; exit 1 }''',

    "service_health": r'''#requires -Version 7.0
[CmdletBinding()]
param()
function Ensure-GraphModules {
  if (-not (Get-Module -ListAvailable Microsoft.Graph.Reports)) {
    Write-Host "Installing Microsoft.Graph.Reports..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph.Reports -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module Microsoft.Graph.Reports -ErrorAction Stop
}
function Connect-GraphForAdmin {
  param([string[]]$Scopes)
  $ctx = Get-MgContext
  if (-not $ctx -or ($Scopes | Where-Object { $_ -notin $ctx.Scopes })) {
    Connect-MgGraph -Scopes $Scopes | Out-Null
  }
  if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) { Select-MgProfile -Name 'v1.0' }
}
try {
  Ensure-GraphModules
  Connect-GraphForAdmin -Scopes @('ServiceHealth.Read.All')
  Write-Host "M365 SERVICE HEALTH" -ForegroundColor Cyan
  Write-Host ("=" * 54) -ForegroundColor Cyan
  Write-Host ""
  $health = Invoke-MgGraphRequest -Method GET -Uri '/v1.0/admin/serviceAnnouncement/healthOverviews'
  $health.value | Sort-Object service | ForEach-Object {
    $status = $_.status
    $color  = switch ($status) {
      'serviceOperational'  { 'Green'  }
      'investigating'       { 'Yellow' }
      'serviceInterruption' { 'Red'    }
      'serviceDegradation'  { 'Yellow' }
      'serviceRestored'     { 'Cyan'   }
      default               { 'White'  }
    }
    $icon = if ($status -eq 'serviceOperational') { 'OK ' } else { '!! ' }
    Write-Host ("  {0}  {1,-45} {2}" -f $icon, $_.service, $status) -ForegroundColor $color
  }
  Write-Host ""
  Write-Host ("─" * 54) -ForegroundColor Yellow
  Write-Host "Active Incidents" -ForegroundColor Yellow
  Write-Host ""
  $issues = Invoke-MgGraphRequest -Method GET -Uri '/v1.0/admin/serviceAnnouncement/issues?$filter=isResolved eq false'
  if (-not $issues.value -or $issues.value.Count -eq 0) {
    Write-Host "  No active incidents." -ForegroundColor Green
  } else {
    foreach ($i in $issues.value) {
      Write-Host "  [$($i.id)] $($i.title)" -ForegroundColor Red
      Write-Host "    Service : $($i.service)"      -ForegroundColor Gray
      Write-Host "    Status  : $($i.status)"       -ForegroundColor Gray
      Write-Host "    Started : $($i.startDateTime)" -ForegroundColor Gray
      Write-Host ""
    }
  }
}
catch { Write-Error $_; exit 1 }''',
}


# Duration options (ISO 8601 duration format)
DURATIONS = [
    ("30 minutes", "PT30M"),
    ("1 hour", "PT1H"),
    ("2 hours", "PT2H"),
    ("4 hours", "PT4H"),
    ("8 hours", "PT8H"),
]

# Justification templates
JUSTIFICATION_TEMPLATES = [
    ("Emergency access needed for incident response", "Emergency access needed for incident response"),
    ("Performing scheduled maintenance", "Performing scheduled maintenance"),
    ("User account management task", "User account management task"),
    ("Device management and troubleshooting", "Device management and troubleshooting"),
    ("Security audit or compliance review", "Security audit or compliance review"),
    ("Teams configuration and administration", "Teams configuration and administration"),
    ("SharePoint site administration", "SharePoint site administration"),
    ("Exchange mailbox management", "Exchange mailbox management"),
    ("Intune device policy deployment", "Intune device policy deployment"),
    ("Power Platform admin task", "Power Platform admin task"),
    ("Custom justification...", ""),
]


COLORS = {
    'bg':           '#0b0f1a',
    'surface':      '#111827',
    'surface2':     '#1c2336',
    'border':       '#2d3748',
    'accent':       '#3b82f6',
    'accent_dim':   '#1e40af',
    'btn':          '#2563eb',
    'btn_hover':    '#3b82f6',
    'btn_active':   '#1d4ed8',
    'btn_disable':  '#1a2540',
    'text':         '#cbd5e1',
    'text_muted':   '#4b5563',
    'text_bright':  '#f1f5f9',
    'heading':      '#93c5fd',
    'success':      '#22c55e',
    'error':        '#f87171',
    'warning':      '#fbbf24',
    'info':         '#38bdf8',
    'terminal':     '#060912',
}


class PIMActivatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Tenant Admin Suite")
        self.root.geometry("1040x760")
        self.root.resizable(True, True)
        self.root.minsize(1040, 760)
        self.root.configure(bg=COLORS['bg'])

        self.root.option_add('*TCombobox*Listbox.background', COLORS['surface2'])
        self.root.option_add('*TCombobox*Listbox.foreground', COLORS['text'])
        self.root.option_add('*TCombobox*Listbox.selectBackground', COLORS['accent_dim'])
        self.root.option_add('*TCombobox*Listbox.selectForeground', COLORS['text_bright'])
        self.root.option_add('*TCombobox*Listbox.font', ('Segoe UI', 10))

        # JSON parsing state (for structured PS output)
        self._json_collecting = False
        self._json_buf = []
        self._json_callback = None

        # Per-section selection state
        self.users_selected_upn = ""
        self.groups_selected_id = ""
        self.groups_selected_name = ""
        self.licenses_sku_map = {}

        self._setup_styles()
        self._create_widgets()

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Custom.TCombobox',
            fieldbackground=COLORS['surface2'], background=COLORS['surface2'],
            foreground=COLORS['text'], selectbackground=COLORS['accent_dim'],
            selectforeground=COLORS['text_bright'], arrowcolor=COLORS['accent'],
            bordercolor=COLORS['border'], lightcolor=COLORS['border'],
            darkcolor=COLORS['border'], padding=6,
        )
        style.map('Custom.TCombobox',
            fieldbackground=[('readonly', COLORS['surface2']), ('disabled', COLORS['surface'])],
            foreground=[('readonly', COLORS['text']), ('disabled', COLORS['text_muted'])],
            background=[('active', COLORS['surface2']), ('readonly', COLORS['surface2'])],
            bordercolor=[('focus', COLORS['accent'])],
        )

    def _create_widgets(self):
        outer = tk.Frame(self.root, bg=COLORS['bg'])
        outer.pack(fill=tk.BOTH, expand=True)

        self._create_header(outer)
        self._create_nav(outer)

        # Fixed-height content pane — holds whichever section is active
        self._content_pane = tk.Frame(outer, bg=COLORS['bg'])
        self._content_pane.configure(height=262)
        self._content_pane.pack(fill=tk.X)
        self._content_pane.pack_propagate(False)

        self._section_frames = {}
        for key in ('pim', 'users', 'groups', 'licenses', 'tenant'):
            self._section_frames[key] = tk.Frame(self._content_pane, bg=COLORS['bg'])

        self._create_pim_section(self._section_frames['pim'])
        self._create_users_section(self._section_frames['users'])
        self._create_groups_section(self._section_frames['groups'])
        self._create_licenses_section(self._section_frames['licenses'])
        self._create_tenant_section(self._section_frames['tenant'])

        self._create_terminal(outer)
        self._show_section('pim')

    # ── Header ────────────────────────────────────────────────────────────

    def _create_header(self, outer):
        hdr = tk.Canvas(outer, height=80, bg=COLORS['surface'], highlightthickness=0)
        hdr.pack(fill=tk.X)
        hdr.create_rectangle(0, 0, 5, 80, fill=COLORS['accent'], outline='')
        hdr.create_text(22, 27, anchor='w', text="Tenant Admin Suite",
                        font=('Segoe UI', 17, 'bold'), fill=COLORS['text_bright'])
        hdr.create_text(22, 55, anchor='w',
                        text="Azure AD  ·  Microsoft 365  ·  Privileged Identity Management",
                        font=('Segoe UI', 9), fill=COLORS['text_muted'])
        tk.Frame(outer, height=1, bg=COLORS['accent_dim']).pack(fill=tk.X)

    # ── Nav bar ───────────────────────────────────────────────────────────

    def _create_nav(self, outer):
        nav = tk.Frame(outer, bg=COLORS['surface'], height=44)
        nav.pack(fill=tk.X)
        nav.pack_propagate(False)
        self._nav_frame = nav

        self._nav_buttons = {}
        for key, label in [('pim', 'PIM Activation'), ('users', 'Users'),
                            ('groups', 'Groups'), ('licenses', 'Licenses'), ('tenant', 'Tenant')]:
            btn = tk.Button(nav, text=label, font=('Segoe UI', 10),
                            bg=COLORS['surface'], fg=COLORS['text_muted'],
                            relief=tk.FLAT, bd=0, padx=18, pady=0, cursor='hand2',
                            command=lambda k=key: self._show_section(k))
            btn.pack(side=tk.LEFT, fill=tk.Y)
            btn.bind('<Enter>', lambda e, b=btn: b.config(fg=COLORS['text_bright'])
                     if b.cget('bg') != COLORS['surface2'] else None)
            btn.bind('<Leave>', lambda e, b=btn: b.config(fg=COLORS['text_muted'])
                     if b.cget('bg') != COLORS['surface2'] else None)
            self._nav_buttons[key] = btn

        self._nav_indicator = tk.Frame(nav, height=3, bg=COLORS['accent'])
        tk.Frame(outer, height=1, bg=COLORS['border']).pack(fill=tk.X)

    def _show_section(self, key):
        for frame in self._section_frames.values():
            frame.pack_forget()
        self._section_frames[key].pack(fill=tk.BOTH, expand=True)

        for k, btn in self._nav_buttons.items():
            btn.config(bg=COLORS['surface2'] if k == key else COLORS['surface'],
                       fg=COLORS['text_bright'] if k == key else COLORS['text_muted'])

        self._active_section = key
        btn = self._nav_buttons[key]
        self._nav_frame.update_idletasks()
        self._nav_indicator.place(x=btn.winfo_x(), y=41, width=btn.winfo_width(), height=3)

    # ── PIM Section ───────────────────────────────────────────────────────

    def _create_pim_section(self, parent):
        body = tk.Frame(parent, bg=COLORS['bg'], padx=26, pady=10)
        body.pack(fill=tk.BOTH, expand=True)

        self._section_label(body, "SELECT ROLE")
        self.role_var = tk.StringVar()
        self.role_combo = ttk.Combobox(body, textvariable=self.role_var,
            values=sorted(ROLE_SCRIPTS.keys()), state='readonly',
            style='Custom.TCombobox', font=('Segoe UI', 10))
        self.role_combo.pack(fill=tk.X, pady=(0, 4))
        self.role_combo.bind('<<ComboboxSelected>>', self._pim_on_role_selected)

        row = tk.Frame(body, bg=COLORS['bg'])
        row.pack(fill=tk.X, pady=(4, 0))

        dur_col = tk.Frame(row, bg=COLORS['bg'])
        dur_col.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 14))
        self._section_label(dur_col, "DURATION")
        self.duration_var = tk.StringVar(value="4 hours (PT4H)")
        self.duration_combo = ttk.Combobox(dur_col, textvariable=self.duration_var,
            values=[f"{lbl} ({val})" for lbl, val in DURATIONS],
            state='readonly', style='Custom.TCombobox', font=('Segoe UI', 10))
        self.duration_combo.pack(fill=tk.X)

        just_col = tk.Frame(row, bg=COLORS['bg'])
        just_col.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._section_label(just_col, "JUSTIFICATION")
        self.justification_var = tk.StringVar()
        self.justification_combo = ttk.Combobox(just_col, textvariable=self.justification_var,
            values=[t[0] for t in JUSTIFICATION_TEMPLATES],
            state='readonly', style='Custom.TCombobox', font=('Segoe UI', 10))
        self.justification_combo.pack(fill=tk.X)
        self.justification_combo.bind('<<ComboboxSelected>>', self._pim_on_justification_selected)

        cf = tk.Frame(body, bg=COLORS['bg'])
        cf.pack(fill=tk.X, pady=(6, 0))
        tk.Label(cf, text="Custom justification:", bg=COLORS['bg'],
                 fg=COLORS['text_muted'], font=('Segoe UI', 8)).pack(anchor=tk.W)
        self.custom_justification_entry = tk.Entry(cf, font=('Segoe UI', 10),
            bg=COLORS['surface2'], fg=COLORS['text'], insertbackground=COLORS['accent'],
            relief=tk.FLAT, bd=0, highlightthickness=1, highlightcolor=COLORS['accent'],
            highlightbackground=COLORS['border'], state='disabled',
            disabledbackground=COLORS['surface'], disabledforeground=COLORS['text_muted'])
        self.custom_justification_entry.pack(fill=tk.X, ipady=6, pady=(3, 0))

        self.activate_btn = self._make_btn(body, "⚡  ACTIVATE ROLE", self._pim_activate,
                                           font_size=11, bold=True)
        self.activate_btn.pack(fill=tk.X, pady=(12, 2))

    # ── Users Section ─────────────────────────────────────────────────────

    def _create_users_section(self, parent):
        body = tk.Frame(parent, bg=COLORS['bg'], padx=18, pady=10)
        body.pack(fill=tk.BOTH, expand=True)
        row = tk.Frame(body, bg=COLORS['bg'])
        row.pack(fill=tk.BOTH, expand=True)

        left = tk.Frame(row, bg=COLORS['bg'], width=340)
        left.pack(side=tk.LEFT, fill=tk.Y)
        left.pack_propagate(False)

        self._section_label(left, "SEARCH USER")
        sf = tk.Frame(left, bg=COLORS['bg'])
        sf.pack(fill=tk.X, pady=(0, 6))
        self.users_search_var = tk.StringVar()
        ue = self._make_entry(sf, self.users_search_var)
        ue.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=6, padx=(0, 6))
        ue.bind('<Return>', self._users_search)
        self._make_btn(sf, "Search", self._users_search, width=7).pack(side=tk.LEFT)

        self._section_label(left, "ACTIONS")
        ar1 = tk.Frame(left, bg=COLORS['bg'])
        ar1.pack(fill=tk.X, pady=(0, 4))
        self.users_enable_btn = self._make_btn(ar1, "Enable Account", self._users_enable)
        self.users_enable_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self.users_disable_btn = self._make_btn(ar1, "Disable Account", self._users_disable,
                                                color=COLORS['warning'])
        self.users_disable_btn.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ar2 = tk.Frame(left, bg=COLORS['bg'])
        ar2.pack(fill=tk.X, pady=(0, 4))
        self._make_btn(ar2, "View MFA Methods", self._users_mfa_methods).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self._make_btn(ar2, "View Groups", self._users_group_memberships).pack(
            side=tk.LEFT, fill=tk.X, expand=True)

        self._make_btn(left, "🔑  Reset Password", self._users_reset_password,
                       color='#7c2d12').pack(fill=tk.X)

        right = tk.Frame(row, bg=COLORS['bg'])
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(16, 0))
        self._section_label(right, "RESULTS")
        self.users_listbox = self._make_listbox(right)
        self.users_listbox.pack(fill=tk.BOTH, expand=True)
        self.users_listbox.bind('<<ListboxSelect>>', self._users_on_select)
        self._users_selected_var = tk.StringVar(value="No user selected")
        tk.Label(right, textvariable=self._users_selected_var, bg=COLORS['bg'],
                 fg=COLORS['heading'], font=('Segoe UI', 9, 'italic')).pack(anchor=tk.W, pady=(4, 0))
        self._users_data = []

    # ── Groups Section ────────────────────────────────────────────────────

    def _create_groups_section(self, parent):
        body = tk.Frame(parent, bg=COLORS['bg'], padx=18, pady=10)
        body.pack(fill=tk.BOTH, expand=True)
        row = tk.Frame(body, bg=COLORS['bg'])
        row.pack(fill=tk.BOTH, expand=True)

        left = tk.Frame(row, bg=COLORS['bg'], width=340)
        left.pack(side=tk.LEFT, fill=tk.Y)
        left.pack_propagate(False)

        self._section_label(left, "SEARCH GROUP")
        sf = tk.Frame(left, bg=COLORS['bg'])
        sf.pack(fill=tk.X, pady=(0, 6))
        self.groups_search_var = tk.StringVar()
        ge = self._make_entry(sf, self.groups_search_var)
        ge.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=6, padx=(0, 6))
        ge.bind('<Return>', self._groups_search)
        self._make_btn(sf, "Search", self._groups_search, width=7).pack(side=tk.LEFT)

        self._section_label(left, "MEMBER OPERATIONS")
        ml = tk.Frame(left, bg=COLORS['bg'])
        ml.pack(fill=tk.X, pady=(0, 4))
        tk.Label(ml, text="Member UPN:", bg=COLORS['bg'], fg=COLORS['text_muted'],
                 font=('Segoe UI', 9)).pack(anchor=tk.W)
        self.groups_member_var = tk.StringVar()
        self._make_entry(ml, self.groups_member_var).pack(fill=tk.X, ipady=6, pady=(2, 0))

        mar = tk.Frame(left, bg=COLORS['bg'])
        mar.pack(fill=tk.X, pady=(4, 0))
        self._make_btn(mar, "Add Member", self._groups_add_member).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self._make_btn(mar, "Remove Member", self._groups_remove_member,
                       color=COLORS['warning']).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._make_btn(left, "List Members", self._groups_get_members).pack(fill=tk.X, pady=(6, 0))

        right = tk.Frame(row, bg=COLORS['bg'])
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(16, 0))
        self._section_label(right, "GROUPS")
        self.groups_listbox = self._make_listbox(right)
        self.groups_listbox.pack(fill=tk.BOTH, expand=True)
        self.groups_listbox.bind('<<ListboxSelect>>', self._groups_on_select)
        self._groups_id_var = tk.StringVar(value="No group selected")
        tk.Label(right, textvariable=self._groups_id_var, bg=COLORS['bg'],
                 fg=COLORS['text_muted'], font=('Segoe UI', 8),
                 wraplength=280).pack(anchor=tk.W, pady=(4, 0))
        self._groups_data = []

    # ── Licenses Section ──────────────────────────────────────────────────

    def _create_licenses_section(self, parent):
        body = tk.Frame(parent, bg=COLORS['bg'], padx=18, pady=10)
        body.pack(fill=tk.BOTH, expand=True)
        row = tk.Frame(body, bg=COLORS['bg'])
        row.pack(fill=tk.BOTH, expand=True)

        left = tk.Frame(row, bg=COLORS['bg'], width=320)
        left.pack(side=tk.LEFT, fill=tk.Y)
        left.pack_propagate(False)

        self._section_label(left, "USER")
        self.licenses_upn_var = tk.StringVar()
        self._make_entry(left, self.licenses_upn_var,
                         placeholder="user@domain.com").pack(fill=tk.X, ipady=6, pady=(0, 6))
        lr1 = tk.Frame(left, bg=COLORS['bg'])
        lr1.pack(fill=tk.X, pady=(0, 6))
        self._make_btn(lr1, "View User Licenses", self._licenses_view_user).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self._make_btn(lr1, "Tenant Inventory", self._licenses_tenant_inventory).pack(
            side=tk.LEFT, fill=tk.X, expand=True)

        self._section_label(left, "ASSIGNMENT")
        sl = tk.Frame(left, bg=COLORS['bg'])
        sl.pack(fill=tk.X, pady=(0, 4))
        tk.Label(sl, text="SKU ID:", bg=COLORS['bg'], fg=COLORS['text_muted'],
                 font=('Segoe UI', 9)).pack(anchor=tk.W)
        self.licenses_sku_var = tk.StringVar()
        self._make_entry(sl, self.licenses_sku_var).pack(fill=tk.X, ipady=5, pady=(2, 0))
        lr2 = tk.Frame(left, bg=COLORS['bg'])
        lr2.pack(fill=tk.X, pady=(6, 0))
        self._make_btn(lr2, "Assign License", self._licenses_assign).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self._make_btn(lr2, "Remove License", self._licenses_remove,
                       color=COLORS['warning']).pack(side=tk.LEFT, fill=tk.X, expand=True)

        right = tk.Frame(row, bg=COLORS['bg'])
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(16, 0))
        self._section_label(right, "AVAILABLE SKUs  — run Tenant Inventory to populate")
        self.licenses_sku_listbox = self._make_listbox(right)
        self.licenses_sku_listbox.pack(fill=tk.BOTH, expand=True)
        self.licenses_sku_listbox.bind('<<ListboxSelect>>', self._licenses_sku_on_select)

    # ── Tenant Section ────────────────────────────────────────────────────

    def _create_tenant_section(self, parent):
        body = tk.Frame(parent, bg=COLORS['bg'], padx=18, pady=14)
        body.pack(fill=tk.BOTH, expand=True)
        self._section_label(body, "TENANT OPERATIONS")
        btn_row = tk.Frame(body, bg=COLORS['bg'])
        btn_row.pack(fill=tk.X, pady=(0, 14))
        self._make_btn(btn_row, "🏢  Tenant Overview", self._tenant_overview).pack(
            side=tk.LEFT, padx=(0, 10), ipadx=14, ipady=8)
        self._make_btn(btn_row, "🔔  Service Health", self._tenant_service_health).pack(
            side=tk.LEFT, ipadx=14, ipady=8)
        tk.Label(body,
                 text=("Tenant Overview — org name, tenant ID, verified domains, user/group counts, "
                       "and license inventory.\n\n"
                       "Service Health — current status of all M365 services and any active incidents."),
                 bg=COLORS['bg'], fg=COLORS['text_muted'], font=('Segoe UI', 9),
                 wraplength=700, justify=tk.LEFT).pack(anchor=tk.W)

    # ── Shared Terminal ───────────────────────────────────────────────────

    def _create_terminal(self, outer):
        hdr = tk.Frame(outer, bg=COLORS['bg'], padx=26)
        hdr.pack(fill=tk.X)
        lf = tk.Frame(hdr, bg=COLORS['bg'])
        lf.pack(fill=tk.X, pady=(10, 5))
        tk.Frame(lf, width=3, bg=COLORS['accent']).pack(side=tk.LEFT, fill=tk.Y, padx=(0, 8))
        tk.Label(lf, text="OUTPUT", bg=COLORS['bg'], fg=COLORS['heading'],
                 font=('Segoe UI', 8, 'bold')).pack(side=tk.LEFT, pady=1)
        tk.Button(lf, text="Clear", font=('Segoe UI', 8), bg=COLORS['surface2'],
                  fg=COLORS['text_muted'], relief=tk.FLAT, bd=0, cursor='hand2',
                  padx=8, pady=2, command=self.clear_output).pack(side=tk.RIGHT)

        wrap = tk.Frame(outer, bg=COLORS['bg'], padx=26)
        wrap.pack(fill=tk.BOTH, expand=True, pady=(0, 6))
        term_wrap = tk.Frame(wrap, bg=COLORS['border'])
        term_wrap.pack(fill=tk.BOTH, expand=True)

        font_face = 'Cascadia Code' if self._font_exists('Cascadia Code') else 'Consolas'
        self.output_text = scrolledtext.ScrolledText(
            term_wrap, height=10, font=(font_face, 9),
            bg=COLORS['terminal'], fg=COLORS['text'], insertbackground=COLORS['accent'],
            selectbackground=COLORS['accent_dim'], selectforeground=COLORS['text_bright'],
            relief=tk.FLAT, bd=10, wrap=tk.WORD, state='disabled')
        self.output_text.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        for tag, color in [('success', COLORS['success']), ('error', COLORS['error']),
                           ('info', COLORS['info']), ('warning', COLORS['warning']),
                           ('muted', COLORS['text_muted']), ('bright', COLORS['text_bright'])]:
            self.output_text.tag_configure(tag, foreground=color)

        self.status_var = tk.StringVar(value="Ready")
        tk.Frame(outer, height=1, bg=COLORS['border']).pack(fill=tk.X, side=tk.BOTTOM)
        self.status_bar = tk.Label(outer, textvariable=self.status_var,
            bg=COLORS['surface'], fg=COLORS['text_muted'], font=('Segoe UI', 9),
            anchor=tk.W, padx=16, pady=6)
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    # ── Widget helpers ────────────────────────────────────────────────────

    def _section_label(self, parent, text):
        f = tk.Frame(parent, bg=COLORS['bg'])
        f.pack(fill=tk.X, pady=(10, 4))
        tk.Frame(f, width=3, bg=COLORS['accent']).pack(side=tk.LEFT, fill=tk.Y, padx=(0, 8))
        tk.Label(f, text=text, bg=COLORS['bg'], fg=COLORS['heading'],
                 font=('Segoe UI', 8, 'bold')).pack(side=tk.LEFT, pady=1)

    def _make_entry(self, parent, textvariable, placeholder=None):
        e = tk.Entry(parent, textvariable=textvariable, font=('Segoe UI', 10),
                     bg=COLORS['surface2'], fg=COLORS['text'],
                     insertbackground=COLORS['accent'], relief=tk.FLAT, bd=0,
                     highlightthickness=1, highlightcolor=COLORS['accent'],
                     highlightbackground=COLORS['border'])
        if placeholder and not textvariable.get():
            e.insert(0, placeholder)
            e.config(fg=COLORS['text_muted'])
            def _on_focus_in(ev, entry=e, ph=placeholder, tv=textvariable):
                if entry.get() == ph:
                    entry.delete(0, tk.END)
                    entry.config(fg=COLORS['text'])
            def _on_focus_out(ev, entry=e, ph=placeholder, tv=textvariable):
                if not entry.get():
                    entry.insert(0, ph)
                    entry.config(fg=COLORS['text_muted'])
            e.bind('<FocusIn>', _on_focus_in)
            e.bind('<FocusOut>', _on_focus_out)
        return e

    def _make_listbox(self, parent):
        return tk.Listbox(parent, bg=COLORS['surface2'], fg=COLORS['text'],
                          selectbackground=COLORS['accent_dim'],
                          selectforeground=COLORS['text_bright'],
                          font=('Segoe UI', 9), relief=tk.FLAT, bd=0,
                          highlightthickness=1, highlightcolor=COLORS['border'],
                          highlightbackground=COLORS['border'], activestyle='none',
                          cursor='hand2')

    def _make_btn(self, parent, text, command, width=None, color=None,
                  font_size=10, bold=False):
        bg = color or COLORS['btn']
        hover = COLORS['btn_hover'] if color is None else color
        font_weight = 'bold' if bold else 'normal'
        kw = {'width': width} if width else {}
        btn = tk.Button(parent, text=text, command=command,
                        font=('Segoe UI', font_size, font_weight),
                        bg=bg, fg=COLORS['text_bright'],
                        activebackground=COLORS['btn_active'],
                        activeforeground=COLORS['text_bright'],
                        relief=tk.FLAT, bd=0, cursor='hand2',
                        padx=10, pady=8, **kw)
        btn.bind('<Enter>', lambda e, b=btn, h=hover:
                 b.config(bg=h) if str(b['state']) == 'normal' else None)
        btn.bind('<Leave>', lambda e, b=btn, n=bg:
                 b.config(bg=n) if str(b['state']) == 'normal' else None)
        return btn

    @staticmethod
    def _font_exists(name):
        try:
            import tkinter.font as tkfont
            return name in tkfont.families()
        except Exception:
            return False

    def _set_status(self, text, kind='normal'):
        self.status_var.set(text)
        self.status_bar.configure(fg={
            'success': COLORS['success'], 'error': COLORS['error'],
            'warning': COLORS['warning'], 'busy': COLORS['info'],
        }.get(kind, COLORS['text_muted']))

    def _require_selection(self, value, noun="item"):
        if not value:
            messagebox.showwarning("Nothing Selected", f"Please select a {noun} first.")
            return False
        return True

    # ── Centralised PS runner ─────────────────────────────────────────────

    def _run_ps(self, script_content, ps_args,
                on_success="Operation completed.", on_failure="Operation failed.",
                lock_widget=None, lock_text="Running…",
                unlock_text=None, unlock_bg=None, json_callback=None):

        self._json_collecting = False
        self._json_buf = []
        self._json_callback = json_callback

        orig_text = lock_widget.cget('text') if lock_widget else None
        orig_bg   = unlock_bg or COLORS['btn']
        unlock    = unlock_text or orig_text or ""

        if lock_widget:
            lock_widget.config(state='disabled', bg=COLORS['btn_disable'], text=lock_text)
        self._set_status("Running…", 'busy')
        self.clear_output()

        def _worker():
            tmp_path = None
            try:
                with tempfile.NamedTemporaryFile(
                        mode='w', suffix='.ps1', delete=False, encoding='utf-8') as tmp:
                    tmp.write(script_content)
                    tmp_path = tmp.name

                ps_path = tmp_path.replace('\\', '/')
                cmd = ["pwsh", "-NoProfile", "-NonInteractive",
                       "-ExecutionPolicy", "Bypass", "-File", ps_path
                       ] + [str(a) for a in ps_args]

                self.root.after(0, self.append_output,
                                f"Executing: pwsh ... {' '.join(str(a) for a in ps_args)}\n", 'muted')
                self.root.after(0, self.append_output, "─" * 54 + "\n", 'muted')

                process = subprocess.Popen(cmd, stdout=subprocess.PIPE,
                                           stderr=subprocess.STDOUT, text=True, bufsize=1)
                for line in process.stdout:
                    self.root.after(0, self.append_output, line)
                rc = process.wait()

                self.root.after(0, self.append_output, "─" * 54 + "\n", 'muted')
                if rc == 0:
                    self.root.after(0, self._set_status, on_success, 'success')
                    self.root.after(0, self.append_output, f"\n✅ {on_success}\n", 'success')
                else:
                    self.root.after(0, self._set_status, on_failure, 'error')
                    self.root.after(0, self.append_output,
                                    f"\n❌ {on_failure} (exit {rc})\n", 'error')

            except FileNotFoundError:
                self.root.after(0, self.append_output,
                                "❌ PowerShell 7 (pwsh) not found.\n"
                                "Install from: github.com/PowerShell/PowerShell\n", 'error')
                self.root.after(0, self._set_status, "Error: pwsh not found", 'error')
            except Exception as exc:
                self.root.after(0, self.append_output, f"❌ {exc}\n", 'error')
                self.root.after(0, self._set_status, "Error occurred", 'error')
            finally:
                if tmp_path:
                    try: os.unlink(tmp_path)
                    except Exception: pass
                if lock_widget:
                    self.root.after(0, lambda: lock_widget.config(
                        state='normal', bg=orig_bg, text=unlock))

        threading.Thread(target=_worker, daemon=True).start()

    # ── PIM logic ─────────────────────────────────────────────────────────

    def _pim_on_role_selected(self, event=None):
        self.clear_output()

    def _pim_on_justification_selected(self, event=None):
        selected = self.justification_var.get()
        for label, _ in JUSTIFICATION_TEMPLATES:
            if selected == label and label.startswith("Custom"):
                self.custom_justification_entry.configure(state='normal')
                self.custom_justification_entry.focus()
                return
        self.custom_justification_entry.configure(state='disabled')
        self.custom_justification_entry.delete(0, tk.END)

    # Legacy aliases
    on_role_selected = _pim_on_role_selected
    on_justification_selected = _pim_on_justification_selected

    def _pim_get_justification(self):
        selected = self.justification_var.get()
        for label, _ in JUSTIFICATION_TEMPLATES:
            if selected == label and label.startswith("Custom"):
                custom_text = self.custom_justification_entry.get().strip()
                if custom_text:
                    return custom_text
                messagebox.showwarning("Input Required", "Please enter a custom justification.")
                return None
        for label, template in JUSTIFICATION_TEMPLATES:
            if label == selected:
                return template
        return selected if selected else None

    def _pim_get_duration(self):
        selected = self.duration_var.get()
        for label, value in DURATIONS:
            if selected.startswith(label):
                return value
        return "PT4H"

    def _pim_activate(self):
        role = self.role_var.get()
        if not role:
            messagebox.showwarning("No Role Selected", "Please select a role to activate.")
            return
        justification = self._pim_get_justification()
        if not justification:
            return
        self._run_ps(
            ROLE_SCRIPTS[role],
            ["-Duration", self._pim_get_duration(), "-Justification", justification],
            on_success="Role activation completed successfully.",
            on_failure="Role activation failed.",
            lock_widget=self.activate_btn,
            lock_text="⏳  ACTIVATING…",
            unlock_text="⚡  ACTIVATE ROLE",
            unlock_bg=COLORS['btn'],
        )

    activate_role = _pim_activate

    # ── Users logic ───────────────────────────────────────────────────────

    def _users_search(self, event=None):
        query = self.users_search_var.get().strip()
        if not query:
            messagebox.showwarning("Input Required", "Enter a name or UPN to search.")
            return
        self.users_listbox.delete(0, tk.END)
        self._users_data = []
        self.users_selected_upn = ""
        self._users_selected_var.set("No user selected")

        def on_json(data):
            self._users_data = data if isinstance(data, list) else [data]
            self.root.after(0, self._users_populate_listbox)

        self._run_ps(USERS_SCRIPTS['search'], ["-Query", query],
                     on_success="User search complete.", on_failure="User search failed.",
                     json_callback=on_json)

    def _users_populate_listbox(self):
        self.users_listbox.delete(0, tk.END)
        for u in self._users_data:
            line = f"{u.get('DisplayName','?')}  |  {u.get('UPN','?')}"
            if not u.get('Enabled', True):
                line += "  [DISABLED]"
            self.users_listbox.insert(tk.END, line)

    def _users_on_select(self, event=None):
        sel = self.users_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx < len(self._users_data):
            u = self._users_data[idx]
            self.users_selected_upn = u.get('UPN', '')
        else:
            parts = self.users_listbox.get(idx).split('|')
            self.users_selected_upn = parts[1].strip().split()[0] if len(parts) >= 2 else ''
        self._users_selected_var.set(f"Selected: {self.users_selected_upn}")

    def _users_require_selection(self):
        return self._require_selection(self.users_selected_upn, "user")

    def _users_enable(self):
        if not self._users_require_selection(): return
        if not messagebox.askyesno("Confirm", f"Enable account for:\n{self.users_selected_upn}?"): return
        self._run_ps(USERS_SCRIPTS['toggle'],
                     ["-UPN", self.users_selected_upn, "-Action", "Enable"],
                     on_success=f"Account enabled: {self.users_selected_upn}",
                     on_failure="Failed to enable account.",
                     lock_widget=self.users_enable_btn, lock_text="Enabling…",
                     unlock_text="Enable Account", unlock_bg=COLORS['btn'])

    def _users_disable(self):
        if not self._users_require_selection(): return
        if not messagebox.askyesno("Confirm — Destructive",
                f"DISABLE account for:\n{self.users_selected_upn}\n\n"
                "This blocks sign-in immediately."): return
        self._run_ps(USERS_SCRIPTS['toggle'],
                     ["-UPN", self.users_selected_upn, "-Action", "Disable"],
                     on_success=f"Account disabled: {self.users_selected_upn}",
                     on_failure="Failed to disable account.",
                     lock_widget=self.users_disable_btn, lock_text="Disabling…",
                     unlock_text="Disable Account", unlock_bg=COLORS['warning'])

    def _users_reset_password(self):
        if not self._users_require_selection(): return
        alpha = string.ascii_letters + string.digits + "!@#$%"
        pw = (secrets.choice(string.ascii_uppercase) + secrets.choice(string.ascii_lowercase) +
              secrets.choice(string.digits) + secrets.choice("!@#$%") +
              ''.join(secrets.choice(alpha) for _ in range(12)))
        if not messagebox.askyesno("Confirm Password Reset",
                f"Reset password for:\n{self.users_selected_upn}\n\n"
                f"New temporary password:\n{pw}\n\n"
                "User must change it on next sign-in. Proceed?"): return
        self._run_ps(USERS_SCRIPTS['reset_password'],
                     ["-UPN", self.users_selected_upn, "-TempPassword", pw],
                     on_success=f"Password reset for {self.users_selected_upn}",
                     on_failure="Password reset failed.")

    def _users_mfa_methods(self):
        if not self._users_require_selection(): return
        self._run_ps(USERS_SCRIPTS['mfa_methods'], ["-UPN", self.users_selected_upn],
                     on_success="MFA methods retrieved.", on_failure="Failed to retrieve MFA methods.")

    def _users_group_memberships(self):
        if not self._users_require_selection(): return
        self._run_ps(USERS_SCRIPTS['group_memberships'], ["-UPN", self.users_selected_upn],
                     on_success="Group memberships retrieved.",
                     on_failure="Failed to retrieve group memberships.")

    # ── Groups logic ──────────────────────────────────────────────────────

    def _groups_search(self, event=None):
        query = self.groups_search_var.get().strip()
        if not query:
            messagebox.showwarning("Input Required", "Enter a group name to search.")
            return
        self.groups_listbox.delete(0, tk.END)
        self._groups_data = []
        self.groups_selected_id = ""
        self.groups_selected_name = ""
        self._groups_id_var.set("No group selected")

        def on_json(data):
            self._groups_data = data if isinstance(data, list) else [data]
            self.root.after(0, self._groups_populate_listbox)

        self._run_ps(GROUPS_SCRIPTS['search'], ["-Query", query],
                     on_success="Group search complete.", on_failure="Group search failed.",
                     json_callback=on_json)

    def _groups_populate_listbox(self):
        self.groups_listbox.delete(0, tk.END)
        for g in self._groups_data:
            self.groups_listbox.insert(tk.END, g.get('DisplayName', g.get('Id', '?')))

    def _groups_on_select(self, event=None):
        sel = self.groups_listbox.curselection()
        if not sel: return
        idx = sel[0]
        if idx < len(self._groups_data):
            g = self._groups_data[idx]
            self.groups_selected_id = g.get('Id', '')
            self.groups_selected_name = g.get('DisplayName', '')
            self._groups_id_var.set(f"ID: {self.groups_selected_id}")

    def _groups_require_selection(self):
        return self._require_selection(self.groups_selected_id, "group")

    def _groups_get_members(self):
        if not self._groups_require_selection(): return
        self._run_ps(GROUPS_SCRIPTS['get_members'], ["-GroupId", self.groups_selected_id],
                     on_success=f"Members of '{self.groups_selected_name}' retrieved.",
                     on_failure="Failed to retrieve members.")

    def _groups_add_member(self):
        if not self._groups_require_selection(): return
        member = self.groups_member_var.get().strip()
        if not member:
            messagebox.showwarning("Input Required", "Enter a member UPN."); return
        if not messagebox.askyesno("Confirm",
                f"Add {member} to '{self.groups_selected_name}'?"): return
        self._run_ps(GROUPS_SCRIPTS['add_member'],
                     ["-GroupId", self.groups_selected_id, "-MemberUPN", member],
                     on_success=f"Added {member} to '{self.groups_selected_name}'.",
                     on_failure="Failed to add member.")

    def _groups_remove_member(self):
        if not self._groups_require_selection(): return
        member = self.groups_member_var.get().strip()
        if not member:
            messagebox.showwarning("Input Required", "Enter a member UPN."); return
        if not messagebox.askyesno("Confirm — Destructive",
                f"Remove {member} from '{self.groups_selected_name}'?"): return
        self._run_ps(GROUPS_SCRIPTS['remove_member'],
                     ["-GroupId", self.groups_selected_id, "-MemberUPN", member],
                     on_success=f"Removed {member} from '{self.groups_selected_name}'.",
                     on_failure="Failed to remove member.")

    # ── Licenses logic ────────────────────────────────────────────────────

    def _licenses_require_upn(self):
        return self._require_selection(self.licenses_upn_var.get().strip(), "user UPN")

    def _licenses_require_sku(self):
        return self._require_selection(self.licenses_sku_var.get().strip(), "SKU ID")

    def _licenses_tenant_inventory(self):
        def on_json(data):
            rows = data if isinstance(data, list) else [data]
            self.root.after(0, self._licenses_populate_sku_listbox, rows)

        self._run_ps(LICENSES_SCRIPTS['tenant_inventory'], [],
                     on_success="License inventory retrieved.",
                     on_failure="Failed to retrieve license inventory.",
                     json_callback=on_json)

    def _licenses_populate_sku_listbox(self, rows):
        self.licenses_sku_listbox.delete(0, tk.END)
        self.licenses_sku_map = {}
        for r in rows:
            name   = r.get('SKU Name', r.get('SkuPartNumber', '?'))
            sku_id = r.get('SKU ID', r.get('SkuId', ''))
            avail  = r.get('Available', '?')
            display = f"{name}  ({avail} available)"
            self.licenses_sku_map[display] = sku_id
            self.licenses_sku_listbox.insert(tk.END, display)

    def _licenses_sku_on_select(self, event=None):
        sel = self.licenses_sku_listbox.curselection()
        if not sel: return
        sku_id = self.licenses_sku_map.get(self.licenses_sku_listbox.get(sel[0]), '')
        if sku_id:
            self.licenses_sku_var.set(sku_id)

    def _licenses_view_user(self):
        if not self._licenses_require_upn(): return
        self._run_ps(LICENSES_SCRIPTS['user_licenses'],
                     ["-UPN", self.licenses_upn_var.get().strip()],
                     on_success="User licenses retrieved.",
                     on_failure="Failed to retrieve user licenses.")

    def _licenses_assign(self):
        if not self._licenses_require_upn(): return
        if not self._licenses_require_sku(): return
        upn, sku = self.licenses_upn_var.get().strip(), self.licenses_sku_var.get().strip()
        if not messagebox.askyesno("Confirm", f"Assign SKU:\n{sku}\nto user:\n{upn}?"): return
        self._run_ps(LICENSES_SCRIPTS['assign_license'], ["-UPN", upn, "-SkuId", sku],
                     on_success=f"License assigned to {upn}.",
                     on_failure="License assignment failed.")

    def _licenses_remove(self):
        if not self._licenses_require_upn(): return
        if not self._licenses_require_sku(): return
        upn, sku = self.licenses_upn_var.get().strip(), self.licenses_sku_var.get().strip()
        if not messagebox.askyesno("Confirm — Destructive",
                f"Remove SKU:\n{sku}\nfrom user:\n{upn}?"): return
        self._run_ps(LICENSES_SCRIPTS['remove_license'], ["-UPN", upn, "-SkuId", sku],
                     on_success=f"License removed from {upn}.",
                     on_failure="License removal failed.")

    # ── Tenant logic ──────────────────────────────────────────────────────

    def _tenant_overview(self):
        self._run_ps(TENANT_SCRIPTS['overview'], [],
                     on_success="Tenant overview loaded.",
                     on_failure="Failed to load tenant overview.")

    def _tenant_service_health(self):
        self._run_ps(TENANT_SCRIPTS['service_health'], [],
                     on_success="Service health loaded.",
                     on_failure="Failed to load service health.")

    # ── Shared terminal ───────────────────────────────────────────────────

    def clear_output(self):
        self.output_text.configure(state='normal')
        self.output_text.delete(1.0, tk.END)
        self.output_text.configure(state='disabled')

    def append_output(self, text, tag=None):
        # Intercept JSON markers — silent, never shown in terminal
        stripped = text.strip()
        if stripped == "##JSON_START##":
            self._json_collecting = True
            self._json_buf = []
            return
        if stripped == "##JSON_END##":
            self._json_collecting = False
            if self._json_callback:
                try:
                    self._json_callback(json.loads(''.join(self._json_buf)))
                except Exception:
                    pass
            return
        if self._json_collecting:
            self._json_buf.append(text.strip())
            return

        self.output_text.configure(state='normal')
        if tag is None:
            low = text.lower()
            if any(k in low for k in ('✅', 'completed successfully', ' enabled', 'assigned', 'added', 'reset')):
                tag = 'success'
            elif any(k in low for k in ('❌', 'error', 'failed', 'exception', 'denied')):
                tag = 'error'
            elif any(k in low for k in ('warning', 'warn', ' disabled', 'removed')):
                tag = 'warning'
            elif text.startswith('Executing:') or '─' in text:
                tag = 'muted'
            elif any(k in low for k in ('activating', 'connecting', 'installing', 'fetching')):
                tag = 'info'
        if tag:
            self.output_text.insert(tk.END, text, tag)
        else:
            self.output_text.insert(tk.END, text)
        self.output_text.see(tk.END)
        self.output_text.configure(state='disabled')


def main():
    root = tk.Tk()
    app = PIMActivatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
