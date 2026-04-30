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
import re

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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
    Action='selfActivate'
    PrincipalId=$PrincipalId
    RoleDefinitionId=$RoleDefinitionId
    DirectoryScopeId=$DirectoryScopeId
    Justification=$Justification
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
  $obj | Format-List
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
        self.root.title("PIM Role Activator")
        self.root.geometry("740x660")
        self.root.resizable(False, False)
        self.root.configure(bg=COLORS['bg'])

        # Style the Combobox dropdown listbox
        self.root.option_add('*TCombobox*Listbox.background', COLORS['surface2'])
        self.root.option_add('*TCombobox*Listbox.foreground', COLORS['text'])
        self.root.option_add('*TCombobox*Listbox.selectBackground', COLORS['accent_dim'])
        self.root.option_add('*TCombobox*Listbox.selectForeground', COLORS['text_bright'])
        self.root.option_add('*TCombobox*Listbox.font', ('Segoe UI', 10))

        self._setup_styles()
        self._create_widgets()

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')

        style.configure('Custom.TCombobox',
            fieldbackground=COLORS['surface2'],
            background=COLORS['surface2'],
            foreground=COLORS['text'],
            selectbackground=COLORS['accent_dim'],
            selectforeground=COLORS['text_bright'],
            arrowcolor=COLORS['accent'],
            bordercolor=COLORS['border'],
            lightcolor=COLORS['border'],
            darkcolor=COLORS['border'],
            padding=6,
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

        # ── Header ─────────────────────────────────────────────────────────
        hdr = tk.Canvas(outer, height=80, bg=COLORS['surface'], highlightthickness=0)
        hdr.pack(fill=tk.X)
        hdr.create_rectangle(0, 0, 5, 80, fill=COLORS['accent'], outline='')
        hdr.create_text(22, 27, anchor='w',
                        text="PIM Role Activator",
                        font=('Segoe UI', 17, 'bold'),
                        fill=COLORS['text_bright'])
        hdr.create_text(22, 55, anchor='w',
                        text="Azure AD  ·  Privileged Identity Management",
                        font=('Segoe UI', 9),
                        fill=COLORS['text_muted'])

        tk.Frame(outer, height=1, bg=COLORS['accent_dim']).pack(fill=tk.X)

        # ── Body ───────────────────────────────────────────────────────────
        body = tk.Frame(outer, bg=COLORS['bg'], padx=26, pady=14)
        body.pack(fill=tk.BOTH, expand=True)

        # Role
        self._section_label(body, "SELECT ROLE")
        self.role_var = tk.StringVar()
        self.role_combo = ttk.Combobox(
            body,
            textvariable=self.role_var,
            values=sorted(ROLE_SCRIPTS.keys()),
            state='readonly',
            style='Custom.TCombobox',
            font=('Segoe UI', 10),
        )
        self.role_combo.pack(fill=tk.X, pady=(0, 4))
        self.role_combo.bind('<<ComboboxSelected>>', self.on_role_selected)

        # Duration + Justification side by side
        row = tk.Frame(body, bg=COLORS['bg'])
        row.pack(fill=tk.X, pady=(6, 0))

        dur_col = tk.Frame(row, bg=COLORS['bg'])
        dur_col.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 14))
        self._section_label(dur_col, "DURATION")
        self.duration_var = tk.StringVar(value="4 hours (PT4H)")
        self.duration_combo = ttk.Combobox(
            dur_col,
            textvariable=self.duration_var,
            values=[f"{lbl} ({val})" for lbl, val in DURATIONS],
            state='readonly',
            style='Custom.TCombobox',
            font=('Segoe UI', 10),
        )
        self.duration_combo.pack(fill=tk.X)

        just_col = tk.Frame(row, bg=COLORS['bg'])
        just_col.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._section_label(just_col, "JUSTIFICATION")
        self.justification_var = tk.StringVar()
        self.justification_combo = ttk.Combobox(
            just_col,
            textvariable=self.justification_var,
            values=[t[0] for t in JUSTIFICATION_TEMPLATES],
            state='readonly',
            style='Custom.TCombobox',
            font=('Segoe UI', 10),
        )
        self.justification_combo.pack(fill=tk.X)
        self.justification_combo.bind('<<ComboboxSelected>>', self.on_justification_selected)

        # Custom justification entry
        cf = tk.Frame(body, bg=COLORS['bg'])
        cf.pack(fill=tk.X, pady=(8, 0))
        tk.Label(cf, text="Custom justification:",
                 bg=COLORS['bg'], fg=COLORS['text_muted'],
                 font=('Segoe UI', 8)).pack(anchor=tk.W)
        self.custom_justification_entry = tk.Entry(
            cf,
            font=('Segoe UI', 10),
            bg=COLORS['surface2'],
            fg=COLORS['text'],
            insertbackground=COLORS['accent'],
            relief=tk.FLAT,
            bd=0,
            highlightthickness=1,
            highlightcolor=COLORS['accent'],
            highlightbackground=COLORS['border'],
            state='disabled',
            disabledbackground=COLORS['surface'],
            disabledforeground=COLORS['text_muted'],
        )
        self.custom_justification_entry.pack(fill=tk.X, ipady=7, pady=(4, 0))

        # Activate button
        self.activate_btn = tk.Button(
            body,
            text="⚡  ACTIVATE ROLE",
            font=('Segoe UI', 11, 'bold'),
            bg=COLORS['btn'],
            fg=COLORS['text_bright'],
            activebackground=COLORS['btn_active'],
            activeforeground=COLORS['text_bright'],
            relief=tk.FLAT,
            bd=0,
            cursor='hand2',
            pady=12,
            command=self.activate_role,
        )
        self.activate_btn.pack(fill=tk.X, pady=(20, 6))
        self.activate_btn.bind('<Enter>', self._btn_enter)
        self.activate_btn.bind('<Leave>', self._btn_leave)

        # Output terminal
        self._section_label(body, "OUTPUT")
        term_wrap = tk.Frame(body, bg=COLORS['border'], bd=0)
        term_wrap.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        font_face = 'Cascadia Code' if self._font_exists('Cascadia Code') else 'Consolas'
        self.output_text = scrolledtext.ScrolledText(
            term_wrap,
            height=10,
            font=(font_face, 9),
            bg=COLORS['terminal'],
            fg=COLORS['text'],
            insertbackground=COLORS['accent'],
            selectbackground=COLORS['accent_dim'],
            selectforeground=COLORS['text_bright'],
            relief=tk.FLAT,
            bd=10,
            wrap=tk.WORD,
            state='disabled',
        )
        self.output_text.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        self.output_text.tag_configure('success', foreground=COLORS['success'])
        self.output_text.tag_configure('error',   foreground=COLORS['error'])
        self.output_text.tag_configure('info',    foreground=COLORS['info'])
        self.output_text.tag_configure('warning', foreground=COLORS['warning'])
        self.output_text.tag_configure('muted',   foreground=COLORS['text_muted'])
        self.output_text.tag_configure('bright',  foreground=COLORS['text_bright'])

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        tk.Frame(outer, height=1, bg=COLORS['border']).pack(fill=tk.X, side=tk.BOTTOM)
        self.status_bar = tk.Label(
            outer,
            textvariable=self.status_var,
            bg=COLORS['surface'],
            fg=COLORS['text_muted'],
            font=('Segoe UI', 9),
            anchor=tk.W,
            padx=16,
            pady=6,
        )
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    # ── Helpers ────────────────────────────────────────────────────────────

    def _section_label(self, parent, text):
        f = tk.Frame(parent, bg=COLORS['bg'])
        f.pack(fill=tk.X, pady=(12, 5))
        tk.Frame(f, width=3, bg=COLORS['accent']).pack(side=tk.LEFT, fill=tk.Y, padx=(0, 8))
        tk.Label(f, text=text, bg=COLORS['bg'], fg=COLORS['heading'],
                 font=('Segoe UI', 8, 'bold')).pack(side=tk.LEFT, pady=1)

    @staticmethod
    def _font_exists(name):
        try:
            import tkinter.font as tkfont
            return name in tkfont.families()
        except Exception:
            return False

    def _btn_enter(self, _event=None):
        if str(self.activate_btn['state']) == 'normal':
            self.activate_btn.config(bg=COLORS['btn_hover'])

    def _btn_leave(self, _event=None):
        if str(self.activate_btn['state']) == 'normal':
            self.activate_btn.config(bg=COLORS['btn'])

    def _set_status(self, text, kind='normal'):
        self.status_var.set(text)
        colours = {
            'success': COLORS['success'],
            'error':   COLORS['error'],
            'warning': COLORS['warning'],
            'busy':    COLORS['info'],
            'normal':  COLORS['text_muted'],
        }
        self.status_bar.configure(fg=colours.get(kind, COLORS['text_muted']))

    # ── Logic (unchanged interface) ────────────────────────────────────────

    def on_role_selected(self, event=None):
        self.clear_output()

    def on_justification_selected(self, event=None):
        selected = self.justification_var.get()
        for label, _ in JUSTIFICATION_TEMPLATES:
            if selected == label and label.startswith("Custom"):
                self.custom_justification_entry.configure(state='normal')
                self.custom_justification_entry.focus()
                return
        self.custom_justification_entry.configure(state='disabled')
        self.custom_justification_entry.delete(0, tk.END)

    def get_script(self):
        selected_role = self.role_var.get()
        if not selected_role:
            return None
        return ROLE_SCRIPTS.get(selected_role)

    def get_justification(self):
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

    def get_duration(self):
        selected = self.duration_var.get()
        for label, value in DURATIONS:
            if selected.startswith(label):
                return value
        return "PT4H"

    def clear_output(self):
        self.output_text.configure(state='normal')
        self.output_text.delete(1.0, tk.END)
        self.output_text.configure(state='disabled')

    ANSI_ESCAPE_RE = re.compile(r'\x1B\[[0-?]*[ -/]*[@-~]')

    def append_output(self, text, tag=None):
        text = self.ANSI_ESCAPE_RE.sub('', text)
        self.output_text.configure(state='normal')
        if tag is None:
            low = text.lower()
            if any(k in low for k in ('✅', 'completed successfully')):
                tag = 'success'
            elif any(k in low for k in ('❌', 'error', 'failed', 'exception')):
                tag = 'error'
            elif 'warning' in low:
                tag = 'warning'
            elif text.startswith('Executing:') or '─' in text:
                tag = 'muted'
            elif any(k in low for k in ('activating', 'connecting', 'installing')):
                tag = 'info'
        if tag:
            self.output_text.insert(tk.END, text, tag)
        else:
            self.output_text.insert(tk.END, text)
        self.output_text.see(tk.END)
        self.output_text.configure(state='disabled')

    def activate_role(self):
        script_content = self.get_script()
        if not script_content:
            messagebox.showwarning("No Role Selected", "Please select a role to activate.")
            return

        justification = self.get_justification()
        if not justification:
            return

        duration = self.get_duration()

        self.activate_btn.configure(state='disabled', text='⏳  ACTIVATING…', bg=COLORS['btn_disable'])
        self._set_status("Activating role…", 'busy')
        self.clear_output()

        try:
            with tempfile.NamedTemporaryFile(mode='w', suffix='.ps1', delete=False, encoding='utf-8') as tmp:
                tmp.write(script_content)
                tmp_path = tmp.name

            ps_path = tmp_path.replace('\\', '/')

            try:
                cmd = [
                    "pwsh", "-NoProfile", "-NonInteractive",
                    "-ExecutionPolicy", "Bypass",
                    "-File", ps_path,
                    "-Duration", duration,
                    "-Justification", justification,
                ]

                self.append_output(
                    f"Executing: pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass "
                    f"-File {ps_path} -Duration {duration} -Justification {justification}\n\n",
                    'muted'
                )
                self.append_output("─" * 54 + "\n", 'muted')

                process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                )

                while True:
                    line = process.stdout.readline()
                    if not line and process.poll() is not None:
                        break
                    if line:
                        self.append_output(line)
                    self.root.update()

                return_code = process.wait()
                self.append_output("─" * 54 + "\n", 'muted')

                if return_code == 0:
                    self._set_status("Activation successful!", 'success')
                    self.append_output("\n✅ Role activation completed successfully!\n", 'success')
                else:
                    self._set_status(f"Activation failed (exit {return_code})", 'error')
                    self.append_output(
                        f"\n❌ Role activation failed with exit code {return_code}\n", 'error'
                    )

            finally:
                try:
                    os.unlink(tmp_path)
                except Exception:
                    pass

        except FileNotFoundError:
            error_msg = (
                "PowerShell 7 (pwsh) not found!\n\n"
                "Please ensure PowerShell 7 is installed.\n"
                "Download from: https://github.com/PowerShell/PowerShell/releases"
            )
            self.append_output(f"❌ Error: {error_msg}\n", 'error')
            messagebox.showerror("PowerShell Not Found", error_msg)
            self._set_status("Error: PowerShell 7 not found", 'error')

        except Exception as e:
            error_msg = f"An error occurred: {str(e)}"
            self.append_output(f"❌ Error: {error_msg}\n", 'error')
            messagebox.showerror("Error", error_msg)
            self._set_status("Error occurred", 'error')

        finally:
            self.activate_btn.configure(state='normal', text='⚡  ACTIVATE ROLE', bg=COLORS['btn'])


def main():
    root = tk.Tk()
    app = PIMActivatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
