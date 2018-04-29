function Invoke-UserSimulator {
<#
.SYNOPSIS

Simulates common user behaviour on local and remote Windows hosts.
Authors:  Barrett Adams (@peewpw) and Chris Myers (@swizzlez_)

.DESCRIPTION

Performs different actions to simulate real user activity and is intended for use in a lab
environment. It will browse the internet using Internet Explorer, attempt to map non-existant
network shares, and open emails with Outlook, including embeded links and attachments.

.PARAMETER Standalone

Define if the script should run as a standalone script on the localhost or on remote systems.

.PARAMETER ConfigXML

The configuration xml file to use when running on remote hosts.

.PARAMETER IE

Run the Internet Explorer simulation.

.PARAMETER Shares

 Run the mapping shares simulation.
 
.PARAMETER Email

Run the opening email simulation.

.PARAMETER All

Run all script simulation functions (IE, Shares, Email).

.EXAMPLE

Import the script modules:
PS>Import-Module .\Invoke-UserSimulator.ps1

Run only the Internet Explorer function on the local host:
PS>Invoke-UserSimulator -StandAlone -IE

Configure remote hosts prior to running the script remotely:
PS>Invoke-ConfigureHosts -ConfigXML .\config.xml

Run all simulation functionality on remote hosts configured in the config.xml file:
PS>Invoke-UserSimulator -ConfigXML .\config.xml -All

#>
    [CmdletBinding()]
    Param(
		[Parameter(Mandatory=$False)]
		[switch]$StandAlone,

		[Parameter(Mandatory=$False)]
		[switch]$Email,

		[Parameter(Mandatory=$False)]
		[switch]$IE,

		[Parameter(Mandatory=$False)]
		[switch]$Shares,

		[Parameter(Mandatory=$False)]
		[switch]$All,

		[Parameter(Mandatory=$False)]
		[string]$ConfigXML
    )

    $RemoteScriptBlock = {
	# TOOL FUNCTIONS

        [CmdletBinding()]
        Param(
            [Parameter(Position = 0, Mandatory = $false)]
            [Int]$EmailInterval,
			
            [Parameter(Position = 1, Mandatory = $false)]
            [Int]$PageDuration,
			
            [Parameter(Position = 2, Mandatory = $false)]
            [Int]$LinkDepth,
            
            [Parameter(Position = 3, Mandatory = $false)]
            [Int]$MountInterval,
			
            [Parameter(Position = 4, Mandatory=$False)]
            [Int]$Email,
			
            [Parameter(Position = 5, Mandatory=$False)]
            [Int]$IE,
			
            [Parameter(Position = 6, Mandatory=$False)]
            [Int]$Shares,
			
            [Parameter(Position = 7, Mandatory=$False)]
            [Int]$All
        )

        # Creates an Outlook COM Object, then iterates through unread mail. Parses the mail
        # and opens links in IE. Downloads and executes any attachments in a Microsoft trusted
        # folder, resulting in automatic MACRO execution.
        $InvokeOpenEmail = {
            Param(
                [Parameter(Position = 0, Mandatory=$True)]
                [int]$Interval
            )

            While ($True) {

                Add-type -assembly "Microsoft.Office.Interop.Outlook"
                $Outlook = New-Object -comobject Outlook.Application
                $Outlook.visible = $True
                $namespace = $Outlook.GetNameSpace("MAPI")
                $namespace.Logon("","",$false,$true)
                $namespace.SendAndReceive($false)
                Start-Sleep -s 60
                $inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)
                $filepath = ("C:\Users\" + $env:username + "\AppData\Roaming\Microsoft\Templates")

                ForEach($mail in $inbox.Items.Restrict("[Unread] = True")) {
                    Write-Host "Found Emails"
                    If ($mail.Attachments) {
                        Write-Host "Found attachments!!"
                        ForEach ($attach in $mail.Attachments) {
                            Write-Host "in attachment loop..."
                            $path = (Join-Path $filepath $attach.filename)
                            $attach.saveasfile($path)
                            Invoke-Item $path
                        }
                    }
					
                    # URL Regex
                    [regex]$regex = "([a-zA-Z]{3,})://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)*?"
                    $URLs = echo $mail.Body | Select-String -Pattern $regex -AllMatches

                    ForEach ($object in $URLs.Matches) {
                        $ie = new-object -COM "InternetExplorer.Application"
                        $ie.visible=$true
                        echo "Parsed link:" $object.Value
                        $ie.navigate2($object.Value)
						$timeout = 0
						While ($ie.busy -And $timeout -lt 10) {
							Start-Sleep -milliseconds 1000
							$timeout += 1
						}
                        Start-Sleep -s 15
                        $ie.Quit()
                    }
					
                    # Set mail object as read
                    $mail.Unread = $false
                }
                Start-Sleep -s $Interval
            }
        }
        $InvokeMapShares = {
            Param(
                [Parameter(Position = 0, Mandatory=$True)]
                [int]$Interval
            )

            While ($True) {
                $randShare = -join ((65..90) + (97..122) | Get-Random -Count 10 | % {[char]$_})
                New-PSDrive -Name "K" -PSProvider FileSystem -Root "\\$randShare\sendCreds"
                Start-Sleep -s $Interval
            }
        }
        
        # Simulates a user browsing the Internet. Will open pseudo-random URLs
        $InvokeIETraffic = {
            Param(
                [Parameter(Position = 0, Mandatory=$True)]
                [int]$Interval,

                [Parameter(Position = 1, Mandatory=$True)]
                [int]$MaxLinkDepth
            )

            $URIs = "https://news.google.com","https://www.reddit.com","https://www.msn.com","http://www.cnn.com","http://www.bbc.com","http://www.uroulette.com"
            $depth = 0

            $ie = New-Object -ComObject "InternetExplorer.Application"
            $ie.visible=$true 
            While ($ie.Application -ne $null) {
                If ($depth -eq 0) {
                    $requestUri = get-random -input $URIs
                }
				
                $depth = $depth + 1

                $ie.navigate2($requestUri)
				$timeout = 0
                While ($ie.busy -And $timeout -lt 10) {
					Start-Sleep -milliseconds 1000
					$timeout += 1
				}

                $linklist = New-Object System.Collections.ArrayList($null)
                ForEach ($link in $ie.document.getElementsByTagName("a")) {
                    If ($link.href.length -gt 1) {
                        $linklist.add($link.href)
                        echo $link.href
                    }
                }

                $requestUri = get-random -input $linklist
                echo $requestUri
                echo $depth
                Start-Sleep -s $Interval
                If ($depth -eq $MaxLinkDepth) {
                    $depth = 0
                }
            }
        }
        
        $emailJob = 0
        $ieJob = 0
        $shareJob = 0
        
        If ($Email -or $All) { $emailJob = Start-Job -ScriptBlock $InvokeOpenEmail -ArgumentList @($EmailInterval) -Name 'usersimemail' }
        If ($IE -or $All) { $ieJob = Start-Job -ScriptBlock $InvokeIETraffic -ArgumentList @($PageDuration, $LinkDepth) -Name 'usersimie' }
        If ($Shares -or $All) { $shareJob = Start-Job -ScriptBlock $InvokeMapShares -ArgumentList @($MountInterval) -Name 'usersimshares' }

        # Start health check loop

        $StartTime = Get-Date
        $TimeOut = New-TimeSpan -Hours 1
        While ($True) {
            Start-Sleep -Seconds 60

            If (($All -or $Email) -and $emailJob.State -ne 'Running') {
                $emailJob = Start-Job -ScriptBlock $InvokeOpenEmail -ArgumentList @($EmailInterval) -Name 'usersimemail'
            }
            If (($All -or $IE) -and $ieJob.State -ne 'Running') {
                $ieJob = Start-Job -ScriptBlock $InvokeIETraffic -ArgumentList @($PageDuration, $LinkDepth) -Name 'usersimie'
            }
            If (($All -or $Shares) -and $shareJob.State -ne 'Running') {
                $shareJob = Start-Job -ScriptBlock $InvokeMapShares -ArgumentList @($MountInterval) -Name 'usersimshares'
            }
            
            If ((New-TimeSpan -Start $StartTime -End (Get-Date)) -gt $TimeOut) {
                    If ($All -or $Email) {
                        Stop-Job -Job $emailJob
                        Stop-Process -Name outlook
                        Stop-Process -Name werfault
                    }
                    If ($All -or $IE) {
                        Stop-Job -Job $ieJob
                        Stop-Process -Name iexplore
                        Stop-Process -Name werfault
                    }
                    If ($All -or $Shares) {
                        Stop-Job -Job $shareJob
                    }
                    $StartTime = Get-Date
            }
        }
    }
	
    If ($StandAlone) {
	# CLIENT BEHAVIOR
        If ($ConfigXML) {
            [xml]$XML = Get-Content $ConfigXML
            $EmailInterval = $XML.usersim.email.checkinInterval
            $PageDuration = $XML.usersim.web.pageDuration
            $LinkDepth = $XML.usersim.web.linkDepth
            $MountInterval = $XML.usersim.shares.mountInterval
        } Else {
            $EmailInterval = 300
            $PageDuration = 20
            $LinkDepth = 10
            $MountInterval = 30
        }

        # Make sure variables have values of the right type
		$myEmailInterval = if ($EmailInterval) {$EmailInterval} else {300}
		$myPageDuration = if ($PageDuration) {$PageDuration} else {20}
		$myLinkDepth = if ($LinkDepth) {$LinkDepth} else {10}
		$myMountInterval = if ($MountInterval) {$MountInterval} else {30}
		$myEmail = if ($Email) {1} else {0}
		$myIE = if ($IE) {1} else {0}
		$myShares = if ($Shares) {1} else {0}
		$myAll = if ($All) {1} else {0}
		
        Invoke-Command -ScriptBlock $RemoteScriptBlock -ArgumentList @($myEmailInterval, $myPageDuration, $myLinkDepth, $myMountInterval, $myEmail, $myIE, $myShares, $myAll)
    } Else {
	# SERVER BEHAVIOR
        If (!$ConfigXML) {
			Write-Host "Please provide a configuration file with '-ConfigXML' flag."
			Break
		}
		
		$ShareName = 'UserSim'
		
		Create-UserSimShare "C:\UserSim" $ShareName $RemoteScriptBlock
		
		[xml]$XML = Get-Content $ConfigXML
		$EmailInterval = $XML.usersim.email.checkinInterval
		$PageDuration = $XML.usersim.web.pageDuration
		$LinkDepth = $XML.usersim.web.linkDepth
		$MountInterval = $XML.usersim.shares.mountInterval
		$Server = $XML.usersim.serverIP
		
		$myEmailInterval = if ($EmailInterval) {$EmailInterval} else {300}
		$myPageDuration = if ($PageDuration) {$PageDuration} else {20}
		$myLinkDepth = if ($LinkDepth) {$LinkDepth} else {10}
		$myMountInterval = if ($MountInterval) {$MountInterval} else {30}
		$myEmail = if ($Email) {1} else {0}
		$myIE = if ($IE) {1} else {0}
		$myShares = if ($Shares) {1} else {0}
		$myAll = if ($All) {1} else {0}
		
		$TaskArgs = "-exec bypass -w hidden -c ipmo \\$Server\$ShareName\UserSim.ps1;Invoke-UserSim $myEmailInterval $myPageDuration $myLinkDepth $myMountInterval $myEmail $myIE $myShares $myAll"

		$XML.usersim.client | ForEach-Object { 

			$myHost = $_.host
			$username = $_.username
			$password = $_.password
			$domain = $_.domain
			$DomainUser = $domain+'\'+$username
		
			Write-Host "In foreach loop..."
			Write-Host "$myHost"
			cmdkey /delete:"$myHost"
			Start-RemoteUserSimTask $myHost $TaskArgs $DomainUser $password
			start-sleep -s 1
			cmdkey /add:$myHost /user:"$domain\$username" /pass:"$password"
			Start-Sleep -s 1
			mstsc.exe /v "$myHost"
			Start-Sleep -s 5
			cmdkey /delete:"$myHost"

			Write-Host "Starting usersim on '$myHost' with username '$username'"
		}
    }
}

function Create-UserSimShare {
    [CmdletBinding()]
    Param(
		[Parameter(Position = 0, Mandatory=$True)]
		[String]$SharePath,

		[Parameter(Position = 1, Mandatory=$True)]
		[String]$ShareName,

		[Parameter(Position = 2, Mandatory=$True)]
		[Management.Automation.ScriptBlock]$ScriptBlock
    )
	
	If (Test-Path $SharePath) {
		Remove-Item -Path $SharePath -Recurse -Force
        Start-Sleep -Milliseconds 400
	}
	New-Item $SharePath -Type Directory 

	$Shares=[WMICLASS]"Win32_Share"

	$old_shares = Get-WMIObject Win32_Share -Filter "name='$ShareName'"
	If ($old_shares) { 
		ForEach ($share in $old_shares) {
			$delete = $share.Delete()
		}
	}

	$Shares.Create($SharePath,$ShareName,0)
	$functionString = "function Invoke-UserSim {" + $ScriptBlock + "}"
    $functionString | Out-File $SharePath\UserSim.ps1
}

function Start-RemoteUserSimTask {
    [CmdletBinding()]
    Param(
		[Parameter(Position = 0, Mandatory=$True)]
		[String]$RemoteHost,
		
		[Parameter(Position = 1, Mandatory=$True)]
		[String]$TaskArguments,
		
		[Parameter(Position = 2, Mandatory=$True)]
		[String]$TaskUser,

		[Parameter(Position = 3, Mandatory=$True)]
		[String]$TaskPassword
    )
	
	$RemoteStart = {
		Param(
			[Parameter(Position = 0, Mandatory=$True)]
			[String]$TaskArguments,
			
			[Parameter(Position = 1, Mandatory=$True)]
			[String]$TaskUser,
			
			[Parameter(Position = 2, Mandatory=$True)]
			[String]$TaskPassword
		)
	
		$ExpTime = (Get-Date).AddMinutes(10).GetDateTimeFormats('s')[0]
		
		$ShedService = New-Object -comobject 'Schedule.Service'
		$ShedService.Connect('localhost')

		$Task = $ShedService.NewTask(0)
		$Task.RegistrationInfo.Description = 'Temporary User Sim Task'
		$Task.Settings.Enabled = $true
		$Task.Settings.AllowDemandStart = $true
		$Task.Settings.DeleteExpiredTaskAfter = 'PT5M'

		$trigger = $Task.Triggers.Create(11)
		$trigger.Enabled = $true
		$trigger.UserId = $TaskUser
		$trigger.StateChange = 3
		$trigger.EndBoundary = $ExpTime
		
		$trigger2 = $Task.Triggers.Create(9)
		$trigger2.Enabled = $true
		$trigger2.UserId = $TaskUser
		$trigger2.EndBoundary = $ExpTime

		$action = $Task.Actions.Create(0)
		$action.Path = "powershell"
		$action.Arguments = $TaskArguments

		$taskFolder = $ShedService.GetFolder("\")
		$taskFolder.RegisterTaskDefinition("UserSim", $Task , 6, $TaskUser, $TaskPassword, 3)
	}
	
	Invoke-Command -ScriptBlock $RemoteStart -ArgumentList @($TaskArguments, $TaskUser, $TaskPassword) -ComputerName $RemoteHost
}

function Invoke-ConfigureHosts {
<#
.SYNOPSIS

Configure remote hosts in preperation for Invoke-UserSimulator

.DESCRIPTION

Sets some registry keys to allow programatic access to Outlook and prevent the "welcome" window
in Internet Explorer. Also adds the user to run as to the "Remote Desktop Users" group on the
remote computer.

.PARAMETER ConfigXML

The configuration xml file to use for host configuration on remote hosts.

.EXAMPLE

Import the script modules:
PS>Import-Module .\Invoke-UserSimulator.ps1

Configure remote hosts prior to running the script remotely:
PS>Invoke-ConfigureHosts -ConfigXML .\config.xml

#>
    Param(
		[Parameter(Position = 0, Mandatory=$True)]
		[String]$ConfigXML
    )

    If (!$ConfigXML) {
			Write-Host "Please provide a configuration file with '-ConfigXML' flag."
			Break
	} else {
        [xml]$XML = Get-Content $ConfigXML

        $XML.usersim.client | ForEach-Object {

	        $myHost = $_.host
	        $username = $_.username
            $domain = $_.domain
            $domUser = "$domain\$username"

            $configBlock = {
                Param(
		            [Parameter(Position = 0, Mandatory=$True)]
		            [String]$domuser
                )
                net localgroup "Remote Desktop Users" $domuser /add

                $registryPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\16.0\Outlook\Security"
                New-Item -Path $registryPath -Force
                Set-ItemProperty -Path $registryPath -Name ObjectModelGuard -Value 2 -Type DWord

                $registryPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Outlook\Security"
                New-Item -Path $registryPath -Force
                Set-ItemProperty -Path $registryPath -Name ObjectModelGuard -Value 2 -Type DWord

                $registryPath = "HKLM:\Software\Policies\Microsoft\Internet Explorer\Main"
                New-Item -Path $registryPath  -Force
                Set-ItemProperty -Path $registryPath -Name DisableFirstRunCustomize -Value 1
            }
			
            Invoke-Command -ScriptBlock $configBlock -ArgumentList $domuser -ComputerName $myHost
        }
    }
}