#!%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe

# =======================================================================================
# Name: tado_smart_thermostat.ps1
#
# Description: Gets desired tado settings from a Google Sheet, gets current tado settings
# does some magic, then sets your tado to your desired settings.
#
# Change History:
# 2016-11-01: 1.0 :Kyle Prochaska: 	Initial Release.
# =======================================================================================

function script_setup {
	$ErrorActionPreference="SilentlyContinue"
	Stop-Transcript | Out-Null
	$ErrorActionPreference = "Continue"
	Start-Transcript -Path ./logs/tado.log -Append | Out-Null
	$script:Tado_Username = 'TADO_USERNAME'
	$script:Tado_Password = 'TADO_PASSWORD'
	$script:Tado_ApiUri = 'https://my.tado.com/api/v2'
	$script:Google_ApiKey = 'GOOGLE_APIKEY'
	$script:Google_SheetApiUri = 'GOOGLE_SHEETAPIURI'
}

function google_sheets_import_array {
	foreach ($Array in $script:Tado_GoogleSheet){
		if ($Array[0] -eq 'Tado_DesiredPower'){$script:Tado_DesiredPower = $Array[1]}
		if ($Array[0] -eq 'Tado_DesiredMode'){$script:Tado_DesiredMode = $Array[1]}
		if ($Array[0] -eq 'Tado_DesiredTemp'){[int]$script:Tado_DesiredTemp = $Array[1]}
		if ($Array[0] -eq 'Tado_DesiredFanSpeed'){$script:Tado_DesiredFanSpeed = $Array[1]}
		if ($Array[0] -eq 'Tado_DesiredLouverSwing'){$script:Tado_DesiredLouverSwing = $Array[1]}
	}
}

function google_sheets_api_v4_get {
	$LogStatusPrefix = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss : ")LoggedOutput | google_sheets_api_v4_get |"
	$script:Tado_GoogleSheetRequest = Invoke-WebRequest -Uri "$script:Google_SheetApiUri/values/B2%3AC6?majorDimension=ROWS&fields=values&key=$script:Google_ApiKey"
	$script:Tado_GoogleSheet = ($script:Tado_GoogleSheetRequest.Content | ConvertFrom-Json).values
	if (($script:Tado_GoogleSheetRequest).StatusDescription -eq 'OK'){
		Write-Host "$LogStatusPrefix SUCCESS"
		$script:Tado_GoogleSheet | Export-Clixml -Path ./cache/tado_cached_google_sheet.xml
		google_sheets_import_array
	}
	else {
		Write-Host "$LogStatusPrefix ERROR | Reading from cached file instead of Google Sheets API"
		$script:Tado_GoogleSheet = Import-Clixml -Path ./cache/tado_cached_google_sheet.xml
		google_sheets_import_array
	}
}

function tado_api_v2_get_me {
	$LogStatusPrefix = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss : ")LoggedOutput | tado_api_v2_get_me |"
	$script:Tado_MeRequest = Invoke-WebRequest -Uri "$script:Tado_ApiUri/me?username=$script:Tado_Username&password=$script:Tado_Password"
	if (($script:Tado_MeRequest).StatusDescription -eq 'OK'){
		Write-Host "$LogStatusPrefix SUCCESS"
		$script:Tado_MeRequest | Export-Clixml -Path ./cache/tado_cached_tado_me.xml
	}
	else {
		Write-Host "$LogStatusPrefix ERROR | Reading from cached file instead of Tado API"
		$script:Tado_MeRequest = Import-Clixml -Path ./cache/tado_cached_tado_me.xml
	}
	$script:Tado_Me = ($script:Tado_MeRequest).Content | ConvertFrom-Json
}

function tado_api_v2_get_homeinfo {
	$LogStatusPrefix = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss : ")LoggedOutput | tado_api_v2_get_homeinfo |"
	$script:Tado_HomeInfoRequest = Invoke-WebRequest -Uri "$script:Tado_ApiUri/homes/$script:Tado_HomeId/?username=$script:Tado_Username&password=$script:Tado_Password"
	if (($script:Tado_HomeInfoRequest).StatusDescription -eq 'OK'){
		Write-Host "$LogStatusPrefix SUCCESS"
		$script:Tado_HomeInfoRequest | Export-Clixml -Path ./cache/tado_cached_tado_homeinfo.xml
	}
	else {
		Write-Host "$LogStatusPrefix ERROR | Reading from cached file instead of Tado API"
		$script:Tado_HomeInfoRequest = Import-Clixml -Path ./cache/tado_cached_tado_homeinfo.xml
	}
	$script:Tado_HomeInfo = ($script:Tado_HomeInfoRequest).Content | ConvertFrom-Json
}

function tado_api_v2_get_state {
	$LogStatusPrefix = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss : ")LoggedOutput | tado_api_v2_get_state |"
	$script:Tado_StateRequest = Invoke-WebRequest -Uri "$script:Tado_ApiUri/homes/$script:Tado_HomeId/zones/1/state/?username=$script:Tado_Username&password=$script:Tado_Password"
	if (($script:Tado_StateRequest).StatusDescription -eq 'OK'){
		Write-Host "$LogStatusPrefix SUCCESS"
		$script:Tado_StateRequest | Export-Clixml -Path ./cache/tado_cached_tado_state.xml
	}
	else {
		Write-Host "$LogStatusPrefix ERROR | Reading from cached file instead of Tado API"
		$script:Tado_StateRequest = Import-Clixml -Path ./cache/tado_cached_tado_state.xml
	}
	$script:Tado_State = ($script:Tado_StateRequest).Content | ConvertFrom-Json
}

function tado_api_v2_get {
	tado_api_v2_get_me
	$script:Tado_HomeId = $script:Tado_Me.homes.id
	tado_api_v2_get_homeinfo
	tado_api_v2_get_state
#	$script:Tado_Capabilities = (Invoke-WebRequest -Uri "$script:Tado_ApiUri/homes/$script:Tado_HomeId/zones/1/capabilities/?username=$script:Tado_Username&password=$script:Tado_Password").Content | ConvertFrom-Json
#	$script:Tado_OutdoorWeather = (Invoke-WebRequest -Uri "$script:Tado_ApiUri/homes/$script:Tado_HomeId/weather/?username=$script:Tado_Username&password=$script:Tado_Password").Content | ConvertFrom-Json
	$script:Tado_HomeAway = $script:Tado_State.tadoMode
	$script:Tado_GetTemp = [System.Math]::Round($script:Tado_State.sensorDataPoints.insideTemperature.fahrenheit, 0)
	$script:Tado_CurrentHumidity = [System.Math]::Round($script:Tado_State.sensorDataPoints.humidity.percentage, 0)
}

function tado_off {
	$script:Tado_SetPower = 'OFF'
	$script:Tado_ApiPutJson = @{
		"type" = "MANUAL"
		"setting" = @{
			"type" = "AIR_CONDITIONING"
			"power" = "$script:Tado_SetPower"
		}
		"termination" = @{
			"type" = "MANUAL"
		}
	} | ConvertTo-Json
}

function tado_cool {
	if ($script:Tado_DesiredFanSpeed -eq 'AUTO'){$script:Tado_SmartFanSpeedEnabled = $True}
	if ($script:Tado_GetTemp -ge ($script:Tado_DesiredTemp +3)){
		$script:Tado_SetPower = 'ON'
		$script:Tado_SetMode = 'COOL'
		if ($script:Tado_SmartFanSpeedEnabled -eq $True){$script:Tado_SetFanSpeed = 'HIGH'}
		else {$script:Tado_SetFanSpeed = $script:Tado_DesiredFanSpeed}
		$script:Tado_SetTemp = $script:Tado_DesiredTemp
	}
	elseif ($script:Tado_GetTemp -eq ($script:Tado_DesiredTemp +2)){
		$script:Tado_SetPower = 'ON'
		$script:Tado_SetMode = 'COOL'
		if ($script:Tado_SmartFanSpeedEnabled -eq $True){$script:Tado_SetFanSpeed = 'MIDDLE'}
		else {$script:Tado_SetFanSpeed = $script:Tado_DesiredFanSpeed}
		$script:Tado_SetTemp = $script:Tado_DesiredTemp
	}
	elseif ($script:Tado_GetTemp -eq ($script:Tado_DesiredTemp +1)){
		$script:Tado_SetPower = 'ON'
		$script:Tado_SetMode = 'COOL'
		if ($script:Tado_SmartFanSpeedEnabled -eq $True){$script:Tado_SetFanSpeed = 'LOW'}
		else {$script:Tado_SetFanSpeed = $script:Tado_DesiredFanSpeed}
		$script:Tado_SetTemp = $script:Tado_DesiredTemp
	}
	elseif ($script:Tado_GetTemp -eq ($script:Tado_DesiredTemp)){
		$script:Tado_SetPower = 'ON'
		$script:Tado_SetMode = 'FAN'
		if ($script:Tado_SmartFanSpeedEnabled -eq $True){$script:Tado_SetFanSpeed = 'LOW'}
		else {$script:Tado_SetFanSpeed = $script:Tado_DesiredFanSpeed}
	}
	elseif ($script:Tado_GetTemp -le ($script:Tado_DesiredTemp -1)){
		$script:Tado_SetPower = 'OFF'
	}
	if ($script:Tado_SetPower -eq 'ON' -and $script:Tado_SetMode -eq 'COOL'){
		$script:Tado_ApiPutJson = @{
			"type" = "MANUAL"
			"setting" = @{
				"type" = "AIR_CONDITIONING"
				"power" = "$script:Tado_SetPower"
				"mode" = "$script:Tado_SetMode"
				"temperature" = @{
					"fahrenheit" = $script:Tado_SetTemp
				}
				"fanSpeed" = "$script:Tado_SetFanSpeed"
				"swing" = "$script:Tado_DesiredLouverSwing"
			}
			"termination" = @{
				"type" = "MANUAL"
			}
		} | ConvertTo-Json
	}
	elseif ($script:Tado_SetPower -eq 'ON' -and $script:Tado_SetMode -eq 'FAN'){
		$script:Tado_ApiPutJson = @{
			"type" = "MANUAL"
			"setting" = @{
				"type" = "AIR_CONDITIONING"
				"power" = "$script:Tado_SetPower"
				"mode" = "$script:Tado_SetMode"
				"fanSpeed" = "$script:Tado_SetFanSpeed"
				"swing" = "$script:Tado_DesiredLouverSwing"
			}
			"termination" = @{
				"type" = "MANUAL"
			}
		} | ConvertTo-Json
	}
	elseif ($script:Tado_SetPower -eq 'OFF'){
		tado_off
	}
}

function tado_heat {
	if ($script:Tado_DesiredFanSpeed -eq 'AUTO'){$script:Tado_SmartFanSpeedEnabled = $True}
	if ($script:Tado_GetTemp -lt ($script:Tado_DesiredTemp)){
		$script:Tado_SetPower = 'ON'
		$script:Tado_SetMode = 'HEAT'
		$script:Tado_SetTemp = $script:Tado_DesiredTemp
	}
	elseif ($script:Tado_GetTemp -eq ($script:Tado_DesiredTemp)){
		$script:Tado_SetPower = 'ON'
		$script:Tado_SetMode = 'FAN'
		$script:Tado_SetFanSpeed = 'LOW'
	}
	elseif ($script:Tado_GetTemp -le ($script:Tado_DesiredTemp +1)){
		$script:Tado_SetPower = 'OFF'
	}
	if ($script:Tado_SetPower -eq 'ON' -and $script:Tado_SetMode -eq 'HEAT'){
		$script:Tado_ApiPutJson = @{
			"type" = "MANUAL"
			"setting" = @{
				"type" = "AIR_CONDITIONING"
				"power" = "$script:Tado_SetPower"
				"mode" = "$script:Tado_SetMode"
				"temperature" = @{
					"fahrenheit" = $script:Tado_SetTemp
				}
			}
			"termination" = @{
				"type" = "MANUAL"
			}
		} | ConvertTo-Json
	}
	elseif ($script:Tado_SetPower -eq 'ON' -and $script:Tado_SetMode -eq 'FAN'){
		$script:Tado_ApiPutJson = @{
			"type" = "MANUAL"
			"setting" = @{
				"type" = "AIR_CONDITIONING"
				"power" = "$script:Tado_SetPower"
				"mode" = "$script:Tado_SetMode"
				"fanSpeed" = "$script:Tado_SetFanSpeed"
				"swing" = "$script:Tado_DesiredLouverSwing"
			}
			"termination" = @{
				"type" = "MANUAL"
			}
		} | ConvertTo-Json
	}
	elseif ($script:Tado_SetPower -eq 'OFF'){
		tado_off
	}
}

function tado_fan {
	if ($script:Tado_DesiredFanSpeed -eq 'AUTO'){$script:Tado_SmartFanSpeedEnabled = $True}
	if ($script:Tado_GetTemp -ge ($script:Tado_DesiredTemp + 3)){
		$script:Tado_SetPower = 'ON'
		$script:Tado_SetMode = 'FAN'
		if ($script:Tado_SmartFanSpeedEnabled -eq $True){$script:Tado_SetFanSpeed = 'HIGH'}
		else {$script:Tado_SetFanSpeed = $script:Tado_DesiredFanSpeed}
	}
	elseif ($script:Tado_GetTemp -eq ($script:Tado_DesiredTemp + 2)){
		$script:Tado_SetPower = 'ON'
		$script:Tado_SetMode = 'FAN'
		if ($script:Tado_SmartFanSpeedEnabled -eq $True){$script:Tado_SetFanSpeed = 'MIDDLE'}
		else {$script:Tado_SetFanSpeed = $script:Tado_DesiredFanSpeed}
	}
	elseif ($script:Tado_GetTemp -eq ($script:Tado_DesiredTemp + 1)){
		$script:Tado_SetPower = 'ON'
		$script:Tado_SetMode = 'FAN'
		if ($script:Tado_SmartFanSpeedEnabled -eq $True){$script:Tado_SetFanSpeed = 'LOW'}
		else {$script:Tado_SetFanSpeed = $script:Tado_DesiredFanSpeed}
	}
	elseif ($script:Tado_GetTemp -eq ($script:Tado_DesiredTemp)){
		$script:Tado_SetPower = 'ON'
		$script:Tado_SetMode = 'FAN'
		if ($script:Tado_SmartFanSpeedEnabled -eq $True){$script:Tado_SetFanSpeed = 'LOW'}
		else {$script:Tado_SetFanSpeed = $script:Tado_DesiredFanSpeed}
	}
	elseif ($script:Tado_GetTemp -le ($script:Tado_DesiredTemp - 1)){
		$script:Tado_SetPower = 'OFF'
	}
	if ($script:Tado_SetPower -eq 'ON' -and $script:Tado_SetMode -eq 'FAN'){
		$script:Tado_ApiPutJson = @{
			"type" = "MANUAL"
			"setting" = @{
				"type" = "AIR_CONDITIONING"
				"power" = "$script:Tado_SetPower"
				"mode" = "$script:Tado_SetMode"
				"fanSpeed" = "$script:Tado_SetFanSpeed"
				"swing" = "$script:Tado_DesiredLouverSwing"
			}
			"termination" = @{
				"type" = "MANUAL"
			}
		} | ConvertTo-Json
	}
	else {tado_off}
}

function tado_dry {
	if ($script:Tado_CurrentHumidity -gt 50){
		$script:Tado_SetPower = 'ON'
		$script:Tado_SetMode = 'DRY'
		$script:Tado_ApiPutJson = @{
			"type" = "MANUAL"
			"setting" = @{
				"type" = "AIR_CONDITIONING"
				"power" = "$script:Tado_SetPower"
				"mode" = "$script:Tado_SetMode"
			}
			"termination" = @{
				"type" = "MANUAL"
			}
		} | ConvertTo-Json
	}
	elseif ($script:Tado_CurrentHumidity -lt 49 -and $script:Tado_CurrentHumidity -gt 40){
		$script:Tado_SetPower = 'ON'
		$script:Tado_SetMode = 'FAN'
		$script:Tado_SetFanSpeed = 'LOW'
		
	}
	else {tado_off}
	if ($script:Tado_SetPower -eq 'ON' -and $script:Tado_SetMode -eq 'DRY'){
		$script:Tado_ApiPutJson = @{
			"type" = "MANUAL"
			"setting" = @{
				"type" = "AIR_CONDITIONING"
				"power" = "$script:Tado_SetPower"
				"mode" = "$script:Tado_SetMode"
			}
			"termination" = @{
				"type" = "MANUAL"
			}
		} | ConvertTo-Json
	}
	if ($script:Tado_SetPower -eq 'ON' -and $script:Tado_SetMode -eq 'FAN'){
		$script:Tado_ApiPutJson = @{
			"type" = "MANUAL"
			"setting" = @{
				"type" = "AIR_CONDITIONING"
				"power" = "$script:Tado_SetPower"
				"mode" = "$script:Tado_SetMode"
				"fanSpeed" = "$script:Tado_SetFanSpeed"
				"swing" = "$script:Tado_DesiredLouverSwing"
			}
			"termination" = @{
				"type" = "MANUAL"
			}
		} | ConvertTo-Json
	}
}

function tado_mode_determination {
	if ($script:Tado_DesiredPower -eq 'AUTO'){
		if ($script:Tado_HomeAway -eq 'HOME'){
			if ($script:Tado_DesiredMode -eq 'COOL'){tado_cool}
			elseif ($script:Tado_DesiredMode -eq 'FAN'){tado_fan}
			elseif ($script:Tado_DesiredMode -eq 'HEAT'){tado_heat}
			elseif ($script:Tado_DesiredMode -eq 'DRY'){tado_dry}
		}
		elseif ($script:Tado_HomeAway -eq 'AWAY'){
			tado_off
		}
	}
	elseif ($script:Tado_DesiredPower -eq 'ON'){
		if ($script:Tado_DesiredMode -eq 'COOL'){tado_cool}
		elseif ($script:Tado_DesiredMode -eq 'FAN'){tado_fan}
		elseif ($script:Tado_DesiredMode -eq 'HEAT'){tado_heat}
		elseif ($script:Tado_DesiredMode -eq 'DRY'){tado_dry}
	}
	elseif ($script:Tado_DesiredPower -eq 'OFF'){
		tado_off
	}
}

function tado_cooldown_timer_pre {
	$script:Tado_Cooldown = $False
	$script:Tado_CooldownTimer = Import-Clixml -Path ./assets/tado_cooldown_timer.xml
	if ($script:Tado_CooldownTimer -eq $null){$script:Tado_CooldownTimer = 0}
	if ($script:Tado_CooldownTimer -ge 1){
		$script:Tado_ProcessApiRequest = $False
		$script:Tado_Cooldown = $True
	}
	$script:Tado_CooldownTimer = $script:Tado_CooldownTimer - 1
	if ($script:Tado_CooldownTimer -lt 1){$script:Tado_CooldownTimer = 0}
}

function tado_cooldown_timer_post {
	$script:Tado_CooldownTimer | Export-Clixml -Path ./assets/tado_cooldown_timer.xml
}
	
function tado_api_v2_put {
	$script:Tado_ProcessApiRequest = $True
	$CurrentSettings = $script:Tado_State.setting
	$DesiredSettings = ($script:Tado_ApiPutJson | ConvertFrom-Json).setting
	if ($CurrentSettings.power -eq $DesiredSettings.power -and $CurrentSettings.mode -eq $DesiredSettings.mode -and $CurrentSettings.temperature -eq $DesiredSettings.temperature -and $CurrentSettings.fanSpeed -eq $DesiredSettings.fanSpeed -and $CurrentSettings.swing -eq $DesiredSettings.swing){
		$script:Tado_ProcessApiRequest = $False
	}
	$LogStatusPrefix = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss : ")LoggedOutput | tado_api_v2_put |"
	tado_cooldown_timer_pre
	if ($script:Tado_ProcessApiRequest -eq $True){
		$script:Tado_ApiPutOutput = Invoke-WebRequest -Uri "$script:Tado_ApiUri/homes/$script:Tado_HomeId/zones/1/overlay?username=$script:Tado_Username&password=$script:Tado_Password" -Method Put -Body $script:Tado_ApiPutJson
		if ($script:Tado_ApiPutOutput.StatusDescription -eq 'OK'){
			Write-Host "$LogStatusPrefix SUCCESS | power=$script:Tado_SetPower,mode=$script:Tado_SetMode,temperature=$script:Tado_SetTemp,fanSpeed=$script:Tado_SetFanSpeed,swing=$script:Tado_DesiredLouverSwing"
		}
		else {Write-Host "$LogStatusPrefix ERROR | power=$script:Tado_SetPower,mode=$script:Tado_SetMode,temperature=$script:Tado_SetTemp,fanSpeed=$script:Tado_SetFanSpeed,swing=$script:Tado_DesiredLouverSwing"}
		$script:Tado_CooldownTimer = $script:Tado_CooldownTimer + 5
	}
	elseif ($script:Tado_ProcessApiRequest -eq $FALSE -and $script:Tado_Cooldown -eq $True){
		Write-Host "$LogStatusPrefix OK | Temperature hysteresis cooldown"
	}
	else {
		Write-Host "$LogStatusPrefix OK | Update not required"
	}
	tado_cooldown_timer_post
}

script_setup
google_sheets_api_v4_get
tado_api_v2_get
tado_mode_determination
tado_api_v2_put

Stop-Transcript | Out-Null

