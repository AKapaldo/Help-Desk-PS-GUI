<# 
.NAME
Help Desk GUI
.Synopsis
Graphical interface for help desk techs to use PowerShell.
.Description
Graphical interface for help desk techs to use PowerShell.
.Notes
None.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(725,485)
$Form.text                       = "Help Desk Ninja"
$Form.TopMost                    = $false
$Form.BackColor          		     = "#ffffff"

#----Add an icon -----
#$Icon                           = New-Object system.drawing.icon ("icon.ico")
#$Form.Icon               		   = $Icon


### -------- Top Row ---------------
$Title                           = New-Object system.Windows.Forms.Label
$Title.text                      = "Computer Name:"
$Title.AutoSize                  = $true
$Title.width                     = 25
$Title.height                    = 20
$Title.location                  = New-Object System.Drawing.Point(15,10)
$Title.Font                      = New-Object System.Drawing.Font('Microsoft Sans Serif',12)

$ComputerName                    = New-Object system.Windows.Forms.TextBox
$ComputerName.multiline          = $false
$ComputerName.width              = 220
$ComputerName.height             = 20
$ComputerName.CharacterCasing	   = 'Upper'
$ComputerName.BackColor			     = 'LemonChiffon'
$ComputerName.BorderStyle		     = 'FixedSingle'
$ComputerName.location           = New-Object System.Drawing.Point(155,10)
$ComputerName.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',12)
$ComputerName.Text			      	 = $env:COMPUTERNAME

$label_PingStatus                = New-Object system.Windows.Forms.Label
$label_PingStatus.text           = "Ping"
$label_PingStatus.AutoSize       = $true
$label_PingStatus.width          = 25
$label_PingStatus.height         = 20
$label_PingStatus.location       = New-Object System.Drawing.Point(540,15)
$label_PingStatus.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$label_PermissionStatus          = New-Object system.Windows.Forms.Label
$label_PermissionStatus.text     = "Permission"
$label_PermissionStatus.AutoSize = $true
$label_PermissionStatus.width    = 25
$label_PermissionStatus.height   = 20
$label_PermissionStatus.location = New-Object System.Drawing.Point(610,15)
$label_PermissionStatus.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$CheckButton                     = New-Object system.Windows.Forms.Button
$CheckButton.text                = "Check Connection"
$CheckButton.width               = 130
$CheckButton.height              = 30
$CheckButton.location            = New-Object System.Drawing.Point(400,10)
$CheckButton.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)


### --------------- Second Row ------------------
$Welcome                         = New-Object system.Windows.Forms.Label
$Welcome.text                    = "Welcome to the Help Desk Ninja. Choose an Option Below:"
$Welcome.AutoSize                = $true
$Welcome.width                   = 325
$Welcome.height                  = 10
$Welcome.location                = New-Object System.Drawing.Point(15,45)
$Welcome.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Service                         = New-Object system.Windows.Forms.Label
$Service.text                    = "Service:"
$Service.AutoSize                = $true
$Service.width                   = 25
$Service.height                  = 10
$Service.location                = New-Object System.Drawing.Point(15,45)
$Service.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ServBox                         = New-Object system.Windows.Forms.TextBox
$ServBox.multiline               = $false
$ServBox.width                   = 305
$ServBox.height                  = 20
$ServBox.location                = New-Object System.Drawing.Point(70,45)
$ServBox.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$StatusBox                       = New-Object system.Windows.Forms.RichTextBox
$StatusBox.multiline             = $true
$StatusBox.width                 = 700
$StatusBox.height                = 150
$StatusBox.location              = New-Object System.Drawing.Point(15,45)
$StatusBox.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

### -------------------- Third Row -----------------
$Status                          = New-Object system.Windows.Forms.Label
$Status.text                     = "Status:"
$Status.AutoSize                 = $true
$Status.width                    = 25
$Status.height                   = 10
$Status.location                 = New-Object System.Drawing.Point(15,75)
$Status.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 305
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(70,75)
$TextBox1.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

### ------------------- Fourth Row ----------------
$ServTitle                       = New-Object system.Windows.Forms.Label
$ServTitle.text                  = "Service Resets:"
$ServTitle.AutoSize              = $true
$ServTitle.width                 = 360
$ServTitle.height                = 10
$ServTitle.BorderStyle	         = 'None'
$ServTitle.location              = New-Object System.Drawing.Point(15,205)
$ServTitle.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

###------------------- Fifth Row ------------------
$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "WebClient Service Reset"
$Button1.width                   = 110
$Button1.height                  = 40
$Button1.location                = New-Object System.Drawing.Point(15,225)
$Button1.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = "DameWare Service Reset"
$Button2.width                   = 110
$Button2.height                  = 40
$Button2.location                = New-Object System.Drawing.Point(140,225)
$Button2.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Button3                         = New-Object system.Windows.Forms.Button
$Button3.text                    = "button"
$Button3.width                   = 110
$Button3.height                  = 40
$Button3.location                = New-Object System.Drawing.Point(265,225)
$Button3.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

###------------------- Sixth Row -------------------
$SWTitle                        = New-Object system.Windows.Forms.Label
$SWTitle.text                   = "Software Installs:"
$SWTitle.AutoSize               = $true
$SWTitle.width                  = 360
$SWTitle.height                 = 10
$SWTitle.BorderStyle	          = 'None'
$SWTitle.location               = New-Object System.Drawing.Point(15,275)
$SWTitle.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

###--------------------- Seventh Row ----------------------
$Button4                        = New-Object system.Windows.Forms.Button
$Button4.text                   = "Content Manager Install"
$Button4.width                  = 110
$Button4.height                 = 40
$Button4.location               = New-Object System.Drawing.Point(15,295)
$Button4.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

###------------------- Bottom Row ---------------------

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 460
$ProgressBar1.height             = 30
$ProgressBar1.location           = New-Object System.Drawing.Point(15,360)

$Close                           = New-Object system.Windows.Forms.Button
$Close.text                      = "Close"
$Close.width                     = 60
$Close.height                    = 30
$Close.location                  = New-Object System.Drawing.Point(650,440)
$Close.Font                      = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Close.ForeColor                 = [System.Drawing.ColorTranslator]::FromHtml("")
$Close.DialogResult         	   = [System.Windows.Forms.DialogResult]::Cancel


###------------- Build the Form ----------------
$Form.controls.AddRange(@($Button1,$Button2,$Button3,$Button4,$Close,$Title,$Status,$TextBox1,$ProgressBar1,$ServBox,$Service,$Welcome,$ComputerName,$ServTitle,$SWTitle,$label_PingStatus,$label_PermissionStatus, $CheckButton,$StatusBox))

#----------------Functions ---------------


function Add-TextBox{
		[CmdletBinding()]
		param ($text)
		$StatusBox.Text += "$text"
		$StatusBox.Text += "`n# # # # # # # # # #`n"
	}

function TestConnect {
$Computer = $ComputerName.text
if (Test-Connection $Computer -Count 1 -Quiet) {
			$label_PingStatus.Text = "Ping: OK";$label_PingStatus.ForeColor = "green"}
else {$label_PingStatus.Text = "Ping: FAIL";$label_PingStatus.ForeColor = "red"}
if (Test-Path "\\$Computer\c$"){
				$label_PermissionStatus.Text = "Permission: OK";$label_PermissionStatus.ForeColor = "green"}
else {$label_PermissionStatus.Text = "Permission: FAIL";$label_PermissionStatus.ForeColor = "red"}
}

###----- WebClient Service Reset ---------------
function WebClient { 
  $Status.visible = $true
  $TextBox1.Visible = $true
  $TextBox1.ForeColor = "#000000"
  $TextBox1.Text = 'Checking Service...'
  $Computer = $ComputerName.text
  IF (Get-Service -Name "WebClient" -erroraction silentlycontinue){
  $process = Get-Service -Name "WebClient" | Select -expand Status
  $Welcome.Visible = $false
  $Service.Visible = $true
  $ServBox.Visible = $true
  $Status.visible = $true
  $TextBox1.text = "$process on $computer"
  $ServBox.text = "WebClient"  }
  
  else {
    $TextBox1.Visible = $true
	$Welcome.Visible = $false
    $Service.Visible = $true
    $ServBox.Visible = $true
    $Status.visible = $true
    $TextBox1.Visible = $true
    $TextBox1.text = $process
    $ServBox.text = "WebClient"  
	$TextBox1.ForeColor = "#D0021B"
    $TextBox1.Text = 'Service Not Found'
  }
 
}

###------------- DameWare Reset ------------
function DameWare { 
  if ($ComputerName.text -ne "") {
  $Status.visible = $true
  $TextBox1.Visible = $true
  $TextBox1.ForeColor = "#000000"
  $TextBox1.Text = 'Checking Service...'
  $Computer = $ComputerName.text
		if (Get-Service -name DWMRCS -computername $Computer -erroraction SilentlyContinue) {
			$TextBox1.Visible = $true
			$Service.Visible = $true
			$ServBox.Visible = $true
			$ServBox.Text = "DameWare Mini Remote Control"
			$TextBox1.ForeColor = "#000000"
			$TextBox1.Text = "Resetting Dameware for $Computer ."
			Get-Service -name DWMRCS -computername $Computer | Restart-Service
			$TextBox1.Text = "Service Reset."
		}
		else {
		$Service.Visible = $true
		$ServBox.Visible = $true
		$ServBox.Text = "DameWare Mini Remote Control"
		$TextBox1.ForeColor = "#D0021B"
		$TextBox1.Text = "Dameware not found on $Computer ."
		}
  }
  else {
	$TextBox1.Visible = $true
	$Welcome.Visible = $false
    $Service.Visible = $true
    $ServBox.Visible = $true
    $Status.visible = $true
    $TextBox1.Visible = $true
    $TextBox1.text = $process
    $ServBox.text = "DameWare Mini Remote Control"  
	$TextBox1.ForeColor = "#D0021B"
    $TextBox1.Text = 'Service Not Found'
	}
	}
  
   
  
  	
	function TextTest
	{	
	$Computer = $ComputerName.text
	$Welcome.Visible = $false
    $Service.Visible = $false
    $ServBox.Visible = $false
    $Status.visible = $false
    $TextBox1.Visible = $false
	$StatusBox.Visible = $true
	$StatusBox.Text = "Checking Ping"
	$TextTest = Test-Connection $Computer -Count 3 | Out-String
	Add-TextBox $TextTest
	
	
	}
	
 



#---------------------------------------------------------[Opening State]--------------------------------------------------------

  $TextBox1.Visible = $false
  $Button1.Visible = $true
  $Close.Visible = $true
  $Title.Visible = $true
  $TextBox1.Visible = $false
  $ProgressBar1.Visible = $false
  $Status.Visible = $false
  $Service.Visible = $false
  $ServBox.Visible= $false
  $Welcome.Visible = $true
  $ComputerName.Visible = $true
  $StatusBox.Visible = $false


$Button1.Add_Click({ WebClient })
$Button2.Add_Click({ DameWare })
$Button3.Add_Click({ TextTest })
$CheckButton.Add_Click({ TestConnect })

[void]$Form.ShowDialog()
