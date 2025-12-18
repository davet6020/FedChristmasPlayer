Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---- CONFIG ----
$musicDir = "Songs"
$volume   = 40
$titleText  = "Christmas Music Player"
$titleColor = [System.Drawing.Color]::Green
$borderColor = [System.Drawing.Color]::Red
# ----------------

# ---- Media Player ----
$player = New-Object -ComObject WMPlayer.OCX
$player.settings.volume = $volume
$player.settings.playCount = 0

$playlist = $player.playlistCollection.newPlaylist("Loop")
Get-ChildItem $musicDir -Filter *.mp3 | ForEach-Object {
    $playlist.appendItem($player.newMedia($_.FullName))
}

$player.currentPlaylist = $playlist
$player.controls.play()

# ---- Form ----
$form = New-Object System.Windows.Forms.Form
$form.FormBorderStyle = "None"
$form.Size = New-Object System.Drawing.Size(420,220)
$form.StartPosition = "CenterScreen"
$form.TopMost = $true

# ---- Outer Border ----
$borderPanel = New-Object System.Windows.Forms.Panel
$borderPanel.Dock = "Fill"
$borderPanel.Padding = 4
$borderPanel.BackColor = $borderColor

# ---- Content Panel ----
$content = New-Object System.Windows.Forms.Panel
$content.Dock = "Fill"
$content.BackColor = [System.Drawing.Color]::White

# ---- Green Title Bar ----
$titleBar = New-Object System.Windows.Forms.Panel
$titleBar.Height = 30
$titleBar.Dock = "Top"
$titleBar.BackColor = $titleColor

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = $titleText
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$titleLabel.Dock = "Fill"
$titleLabel.TextAlign = "MiddleLeft"
$titleLabel.Padding = "10,0,0,0"

$titleBar.Controls.Add($titleLabel)

# Drag window by title bar
$titleBar.Add_MouseDown({
    $form.Capture = $false
    $msg = [System.Windows.Forms.Message]::Create($form.Handle, 0xA1, 2, 0)
    [System.Windows.Forms.Application]::Run($msg)
})

# ---- Body ----
$label = New-Object System.Windows.Forms.Label
$label.Text = "I cannot believe you actually inserted this USB from an unknown source into your computer! You must be really desperate for music. 

Fortunately for you, I am not malicious. Merry Christmas and enjoy the tunes!"
$label.Dock = "Top"
$label.Height = 100
$label.TextAlign = "MiddleCenter"

$button = New-Object System.Windows.Forms.Button
$button.Text = "Close"
$button.Width = 100
$button.Height = 30
$button.Top = 120
$button.Left = 160
$button.Add_Click({
    $player.controls.stop()
    $form.Close()
})

# ---- Assemble ----
$content.Controls.Add($button)
$content.Controls.Add($label)
$content.Controls.Add($titleBar)
$borderPanel.Controls.Add($content)
$form.Controls.Add($borderPanel)

$form.ShowDialog()
