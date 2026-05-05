$modulePath = Join-Path -Path $PSScriptRoot -ChildPath 'M365.Exchange.Tools\M365.Exchange.Tools.psd1'
Import-Module $modulePath -Force
Show-M365Prerequisites
Show-M365ExchangeMenu