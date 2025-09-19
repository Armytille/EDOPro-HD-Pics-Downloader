<#
.SYNOPSIS
    EDOPro HD Pics Downloader - PowerShell GUI Edition
.DESCRIPTION
    Downloads HD images of Yu-Gi-Oh! cards for EDOPro.
.NOTES
    Uses Windows Forms for GUI
#>

# Strict mode and error handling
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Load required assemblies - AVANT toute création d'objet Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Net.Http

# IMPORTANT: Ces appels doivent être faits AVANT toute création d'objet Windows Forms
[System.Windows.Forms.Application]::EnableVisualStyles()


# Script configuration
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ImgBaseUrl = "https://images.ygoprodeck.com/images/cards"
$DefaultPicsDir = Join-Path $ScriptDir "pics"
$FieldPicsDir = Join-Path $DefaultPicsDir "field"
$ApiUrl = "https://db.ygoprodeck.com/api/v7/cardinfo.php"
$MaxConcurrency = 20
$RetryCount = 3
$TimeoutSeconds = 30

# ==================== Utility Functions ====================

function Ensure-Directory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )
    
    if (-not (Test-Path -Path $Path -PathType Container)) {
        try {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
            Write-Verbose "Created directory: $Path"
        }
        catch {
            throw "Failed to create directory '$Path': $_"
        }
    }
}

function Initialize-SyncObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$Total
    )
    
    return [hashtable]::Synchronized(@{
        Total = $Total
        Logs = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        FinishedQueue = [System.Collections.Concurrent.ConcurrentQueue[PSCustomObject]]::new()
        Running = $true
    })
}

# ==================== GUI Creation ====================

# Main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Armytille's EDOPro HD Pics Downloader"
$form.Size = New-Object System.Drawing.Size(650, 550)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.MaximizeBox = $false

# Download button
$buttonStart = New-Object System.Windows.Forms.Button
$buttonStart.Text = "Download All Cards"
$buttonStart.Location = New-Object System.Drawing.Point(10, 15)
$buttonStart.Size = New-Object System.Drawing.Size(200, 30)
$buttonStart.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
$form.Controls.Add($buttonStart)

# Cancel button
$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Text = "Cancel"
$buttonCancel.Location = New-Object System.Drawing.Point(220, 15)
$buttonCancel.Size = New-Object System.Drawing.Size(100, 30)
$buttonCancel.Enabled = $false
$buttonCancel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
$form.Controls.Add($buttonCancel)

# Force overwrite checkbox
$checkForce = New-Object System.Windows.Forms.CheckBox
$checkForce.Text = "Force Overwrite Existing"
$checkForce.Location = New-Object System.Drawing.Point(330, 20)
$checkForce.Size = New-Object System.Drawing.Size(180, 20)
$checkForce.Checked = $false
$form.Controls.Add($checkForce)

# Progress bar
$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point(10, 60)
$ProgressBar.Size = New-Object System.Drawing.Size(620, 25)
$ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$form.Controls.Add($ProgressBar)

# Status label
$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Text = "Ready to download"
$StatusLabel.Location = New-Object System.Drawing.Point(10, 90)
$StatusLabel.Size = New-Object System.Drawing.Size(620, 20)
$StatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$form.Controls.Add($StatusLabel)

# Log box
$LogBox = New-Object System.Windows.Forms.TextBox
$LogBox.Multiline = $true
$LogBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$LogBox.ReadOnly = $true
$LogBox.Location = New-Object System.Drawing.Point(10, 120)
$LogBox.Size = New-Object System.Drawing.Size(620, 380)
$LogBox.Font = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Regular)
$LogBox.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$form.Controls.Add($LogBox)

# Script-level variables for job management
$script:sync = $null
$script:runspacePool = $null
$script:jobs = [System.Collections.Generic.List[hashtable]]::new()
$script:timer = $null
$script:isCancelled = $false
$script:allFinished = [System.Collections.Generic.List[object]]::new()

# ==================== Functions that need GUI elements ====================

function Get-IdsFromWeb {
    [CmdletBinding()]
    param()
    
    $LogBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] Retrieving card IDs from YGOProDeck API...`r`n")
    $LogBox.ScrollToCaret()
    
    try {
        # Create HTTP client for better control
        $httpClient = [System.Net.Http.HttpClient]::new()
        $httpClient.Timeout = [TimeSpan]::FromSeconds($TimeoutSeconds)
        
        try {

            # Make API request
            $response = $httpClient.GetAsync($ApiUrl).GetAwaiter().GetResult()
            
            if (-not $response.IsSuccessStatusCode) {
                throw "API returned status code: $($response.StatusCode)"
            }
            
            $jsonContent = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            $cards = $jsonContent | ConvertFrom-Json
            
            # Check for API error response
            if ($cards.PSObject.Properties.Name -contains 'error') {
                throw "API Error: $($cards.error)"
            }
            
            # Extract and validate IDs
            if (-not $cards.data) {
                throw "No card data received from API"
            }

            $script:cardInfo = [System.Collections.Concurrent.ConcurrentDictionary[int, object]]::new()

            $ids = @($cards.data | ForEach-Object {
                if ($null -ne $_.id) {
                    $script:cardInfo.TryAdd($_.id, $_) | Out-Null
                    $_.id.ToString()
                }
            } | Where-Object { $_ })
            
            if ($ids.Count -eq 0) {
                throw "No valid card IDs found in API response"
            }



            $LogBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] Successfully retrieved $($ids.Count) card IDs.`r`n")
            $LogBox.ScrollToCaret()
            
            return $ids
        }
        finally {
            $httpClient.Dispose()
        }
    }
    catch {
        $errorMsg = "Error retrieving card data: $_"
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "API Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $LogBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] $errorMsg`r`n")
        $LogBox.ScrollToCaret()
        return @()
    }
}

# ==================== Download Script Block ====================

$downloadScriptBlock = {
    param($id, $targetDir, $imgBase, $sync, $force)

    # Vérifie les infos de la carte
    if (-not $sync.CardInfo.ContainsKey($id)) {
        $sync.Logs.Enqueue("[$(Get-Date -Format 'HH:mm:ss')] Card info not found for $id")
        return
    }
    $cardData = $sync.CardInfo[$id]

    # Téléchargement image principale
    $url = "$imgBase/$id.jpg"
    $outfile = Join-Path $targetDir "$id.jpg"

    if ((Test-Path -Path $outfile -PathType Leaf) -and -not $force) {
        $sync.FinishedQueue.Enqueue([PSCustomObject]@{Status='Skipped'; Id=$id})
    }
    else {
        $retryCount = 0
        $maxRetries = 3
        $success = $false

        while ($retryCount -lt $maxRetries -and -not $success -and $sync.Running) {
            try {
                foreach ($imgObj in $cardData.card_images) {
                    $url = $imgObj.image_url
                    $outfile = Join-Path $targetDir "$($imgObj.id).jpg"
                    Invoke-WebRequest -Uri $url -OutFile $outfile -TimeoutSec 30
                }
                $success = $true
            }
            catch {
                $retryCount++
                if ($retryCount -eq $maxRetries) {
                    $sync.Logs.Enqueue("[$(Get-Date -Format 'HH:mm:ss')] Error downloading $($id): $($_.Exception.Message)")
                    $sync.FinishedQueue.Enqueue([PSCustomObject]@{Status='Error'; Id=$id})
                }
                else {
                    Start-Sleep -Milliseconds (500 * $retryCount)
                }
            }
        }

        if ($success) {
            $sync.FinishedQueue.Enqueue([PSCustomObject]@{Status='Success'; Id=$id})
        }
    }

    # --- Téléchargement cropped pour Field Spells ---
    if ($cardData.humanReadableCardType -eq "Field Spell" -and $cardData.card_images[0].image_url_cropped) {
        $FieldDir = Join-Path $targetDir "field"
        if (-not (Test-Path $FieldDir)) { New-Item -ItemType Directory -Path $FieldDir -Force | Out-Null }

        $croppedUrl = $cardData.card_images[0].image_url_cropped
        $croppedFile = Join-Path $FieldDir "$id.jpg"

        if ((-not (Test-Path $croppedFile)) -or $force) {
            try {
                Invoke-WebRequest -Uri $croppedUrl -OutFile $croppedFile -TimeoutSec 30 -ErrorAction Stop
            }
            catch {
                $cardName = if ($null -ne $cardData -and $cardData.name) { $cardData.name } else { "<unknown>" }
                $sync.Logs.Enqueue("[$(Get-Date -Format 'HH:mm:ss')] ERROR: Failed to download cropped image for Field Spell $id ($cardName) - $($_.Exception.Message)")
            }
        }
    }
}

# ==================== Event Handlers ====================

$buttonStart.Add_Click({
    # Check if download is already in progress
    if ($null -ne $script:sync -and $script:sync.Running) {
        $LogBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] Download already in progress.`r`n")
        $LogBox.ScrollToCaret()
        return
    }
    
    # Reset cancellation flag
    $script:isCancelled = $false
    
    # Ensure pics directory exists
    try {
        Ensure-Directory -Path $DefaultPicsDir
        Ensure-Directory -Path $FieldPicsDir
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to create directory: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Get card IDs from API
    $ids = Get-IdsFromWeb
    if ($ids.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No card IDs were retrieved from the API.", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # Initialize UI for download
    $LogBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] Starting download of $($ids.Count) images...`r`n")
    $LogBox.ScrollToCaret()
    $ProgressBar.Value = 0
    $ProgressBar.Maximum = 100
    $StatusLabel.Text = "Initializing download..."
    $buttonStart.Enabled = $false
    $buttonCancel.Enabled = $true
    
    $forceOverwrite = $checkForce.Checked
    
    # Initialize synchronization object
    $script:sync = Initialize-SyncObject -Total $ids.Count
    if ($null -ne $script:cardInfo) {
        $script:sync.CardInfo = $script:cardInfo
    } else {
        # Safety: create an empty ConcurrentDictionary to avoid null issues in workers
        $script:sync.CardInfo = [System.Collections.Concurrent.ConcurrentDictionary[int, object]]::new()
    }
    # Create runspace pool
    $script:runspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxConcurrency)
    $script:runspacePool.ApartmentState = [System.Threading.ApartmentState]::MTA
    $script:runspacePool.Open()
    
    # Clear previous jobs
    $script:jobs.Clear()
    
    # Create jobs for each card ID
    foreach ($id in $ids) {
        if ($script:isCancelled) {
            break
        }
        
        $ps = [powershell]::Create()
        $ps.RunspacePool = $script:runspacePool
        
        [void]$ps.AddScript($downloadScriptBlock)
        [void]$ps.AddArgument($id)
        [void]$ps.AddArgument($DefaultPicsDir)
        [void]$ps.AddArgument($ImgBaseUrl)
        [void]$ps.AddArgument($script:sync)
        [void]$ps.AddArgument($forceOverwrite)
        
        $handle = $ps.BeginInvoke()
        $script:jobs.Add(@{
            PowerShell = $ps
            Handle = $handle
            Id = $id
        })
    }
    
    # Start monitoring timer
    if ($null -ne $script:timer) {
        $script:timer.Dispose()
    }
    
    $script:timer = New-Object System.Windows.Forms.Timer
    $script:timer.Interval = 250
    $script:timer.Add_Tick({
        if ($null -eq $script:sync) {
            $script:timer.Stop()
            return
        }
        
        # Vider la queue et mettre dans une liste

        $item = $null
        while ($script:sync.FinishedQueue.TryDequeue([ref]$item)) {
            $script:allFinished.Add($item)
        }


        # Compteurs fiables avec Measure-Object
        $skipped   = ($script:allFinished | Where-Object { $_.Status -eq 'Skipped' } | Measure-Object).Count
        $errors    = ($script:allFinished | Where-Object { $_.Status -eq 'Error' } | Measure-Object).Count
        $processed = ($script:allFinished | Where-Object { $_.Status -eq 'Success' } | Measure-Object).Count
        $finished  = $processed + $skipped + $errors

        # Update progress
        $ProgressBar.Value = if ($script:sync.Total -gt 0) {
            [Math]::Min(100, [int](($finished / $script:sync.Total) * 100))
        } else { 0 }

        $StatusLabel.Text = "Processed: $finished/$($script:sync.Total) | Skipped: $skipped | Errors: $errors"

        # Process log messages
        $log = $null
        while ($script:sync.Logs.TryDequeue([ref]$log)) {
            $LogBox.AppendText("$log`r`n")
        }
        $LogBox.ScrollToCaret()
        
        # Check if download is complete
        if ($finished -ge $script:sync.Total -or $script:isCancelled) {
            $script:timer.Stop()
            
            # Update final status
            if ($script:isCancelled) {
                $StatusLabel.Text = "Download cancelled. Processed: $($finished)/$($script:sync.Total)"
                $LogBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] Download cancelled by user.`r`n")
            }
            else {
                $StatusLabel.Text = "Download completed. Errors: $($errors) | Skipped: $($skipped)"
                $LogBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] Download completed successfully.`r`n")
            }
            $LogBox.ScrollToCaret()
            
            # Save final counts before cleanup
            $finalErrors =  $errors  
            $finalSkipped = $skipped 
            $finalProcessed = $finished
            $finalTotal = $script:sync.Total

            # Clean up resources
            $script:sync.Running = $false
            
            # Wait for and dispose of all jobs
            foreach ($job in $script:jobs) {
                try {
                    if ($script:isCancelled) {
                        $job.PowerShell.Stop()
                    }
                    elseif ($job.Handle.IsCompleted -eq $false) {
                        $job.PowerShell.EndInvoke($job.Handle)
                    }
                }
                catch {
                    $LogBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] Error completing job for ID $($job.Id): $_`r`n")
                }
                finally {
                    $job.PowerShell.Dispose()
                }
            }
            
            # Clean up runspace pool
            if ($null -ne $script:runspacePool) {
                $script:runspacePool.Close()
                $script:runspacePool.Dispose()
                $script:runspacePool = $null
            }
            
            # Clean up timer
            $script:timer.Dispose()
            $script:timer = $null

            # Vider la liste accumulée
            $script:allFinished.Clear()

            # Reset variables
            $script:sync = $null
            $script:jobs.Clear()
            
            # Re-enable buttons
            $buttonStart.Enabled = $true
            $buttonCancel.Enabled = $false
            
            # Show completion message (only if not cancelled)
            if (-not $script:isCancelled) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Download completed!`n`nTotal: $finalTotal`nProcessed: $finalProcessed`nErrors: $finalErrors`nSkipped: $finalSkipped", 
                    "Download Complete", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
        }
    })
    
    $script:timer.Start()
})

$buttonCancel.Add_Click({
    if ($null -ne $script:sync -and $script:sync.Running) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to cancel the download?", 
            "Confirm Cancellation", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo, 
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $script:isCancelled = $true
            $script:sync.Running = $false
            $buttonCancel.Enabled = $false
            $StatusLabel.Text = "Cancelling download..."
            $LogBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] Cancellation requested...`r`n")
            $LogBox.ScrollToCaret()
        }
    }
})

# Form closing event
$form.Add_FormClosing({
    param($sender, $e)
    
    if ($null -ne $script:sync -and $script:sync.Running) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "A download is in progress. Are you sure you want to exit?", 
            "Confirm Exit", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo, 
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::No) {
            $e.Cancel = $true
            return
        }
        
        # Stop all operations
        $script:isCancelled = $true
        $script:sync.Running = $false
        
        # Clean up resources
        if ($null -ne $script:timer) {
            $script:timer.Stop()
            $script:timer.Dispose()
        }
        
        foreach ($job in $script:jobs) {
            try {
                $job.PowerShell.Stop()
                $job.PowerShell.Dispose()
            }
            catch { }
        }
        
        if ($null -ne $script:runspacePool) {
            $script:runspacePool.Close()
            $script:runspacePool.Dispose()
        }
    }
})

# ==================== Application Launch ====================

# Add initial log entry
$LogBox.AppendText("========================================`r`n")
$LogBox.AppendText("Armytille's EDOPro HD Pics Downloader`r`n")
$LogBox.AppendText("PowerShell Version: $($PSVersionTable.PSVersion)`r`n")
$LogBox.AppendText("========================================`r`n")
$LogBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] Application started. Ready to download.`r`n")

# Show form and activate
$form.Add_Shown({ $form.Activate() })

# Run the application
[System.Windows.Forms.Application]::Run($form)