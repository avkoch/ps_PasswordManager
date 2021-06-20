# (C) Hydeen
# .LINK
# https://www.codeproject.com/Articles/1107067/PowerShell-Application-Password-Manager


[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Security")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Windows.Forms.Application]::EnableVisualStyles()

$PathToStore = "$env:LOCALAPPDATA\Temp\PassKeeper.dat"

########################### FUNCTIONS ##########################################
function CreateColumns() {
    $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{
        HeaderText = "Name"
        Width      = 163+180
    }
    $grid.Columns.Add($column) | Out-Null
    $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{
        HeaderText = "Password"
        Width      = 180
        Visible    = $false
    }
    $grid.Columns.Add($column) | Out-Null
    $column = New-Object System.Windows.Forms.DataGridViewButtonColumn -Property @{
        #HeaderText = "View"
        Width      = 50
    }
    $grid.Columns.Add($column) | Out-Null
    $column = New-Object System.Windows.Forms.DataGridViewButtonColumn -Property @{
        #HeaderText = "Copy"
        Width      = 50
    }
    $grid.Columns.Add($column) | Out-Null
    $column = New-Object System.Windows.Forms.DataGridViewButtonColumn -Property @{
        #HeaderText = "Delete"
        Width      = 50
    }
    $grid.Columns.Add($column) | Out-Null
}

function Set-EncryptedString($String, $Passphrase, $salt = "SaltCrypto", $init = "PassKeeper", [switch]$arrayOutput) {
    # Info: More information and good documentation on these functions can be found in the link above
    $r = New-Object System.Security.Cryptography.RijndaelManaged
    $pass = [Text.Encoding]::UTF8.GetBytes($Passphrase)
    $salt = [Text.Encoding]::UTF8.GetBytes($salt)
    $r.Key = (New-Object Security.Cryptography.PasswordDeriveBytes $pass, $salt, "SHA1", 5).GetBytes(32)
    $r.IV = (New-Object Security.Cryptography.SHA1Managed).ComputeHash( [Text.Encoding]::UTF8.GetBytes($init) )[0..15]
    $c = $r.CreateEncryptor()
    $ms = New-Object IO.MemoryStream
    $cs = New-Object Security.Cryptography.CryptoStream $ms, $c, "Write"
    $sw = New-Object IO.StreamWriter $cs
    $sw.Write($String)
    $sw.Close()
    $cs.Close()
    $ms.Close()
    $r.Clear()
    [byte[]]$result = $ms.ToArray()
    return [Convert]::ToBase64String($result)
}

function Get-EncryptedString($Encrypted, $Passphrase, $salt = "SaltCrypto", $init = "PassKeeper") {
    if ($Encrypted -is [string]) {
        $Encrypted = [Convert]::FromBase64String($Encrypted)
    }

    $r = New-Object System.Security.Cryptography.RijndaelManaged
    $pass = [Text.Encoding]::UTF8.GetBytes($Passphrase)
    $salt = [Text.Encoding]::UTF8.GetBytes($salt)
    $r.Key = (New-Object Security.Cryptography.PasswordDeriveBytes $pass, $salt, "SHA1", 5).GetBytes(32) #256/8
    $r.IV = (New-Object Security.Cryptography.SHA1Managed).ComputeHash( [Text.Encoding]::UTF8.GetBytes($init) )[0..15]
    $d = $r.CreateDecryptor()
    $ms = New-Object IO.MemoryStream @(, $Encrypted)
    $cs = New-Object Security.Cryptography.CryptoStream $ms, $d, "Read"
    $sr = New-Object IO.StreamReader $cs
    $text = $sr.ReadToEnd()
    $sr.Close()
    $cs.Close()
    $ms.Close()
    $r.Clear()
    return $text
}

function Export() {
    # Info: Exports all relevent cells to a .dat file in the %temp% directory
    if (Test-Path $Script:PathToStore) {
        Remove-Item $Script:PathToStore -Force -Confirm:$false | Out-Null
    }
    for ($i = 0; $i -lt $grid.RowCount; $i++) {
        $item = $grid.Rows[$i].Cells.Value
        "$($item[0])→$($item[1])" | Out-File $Script:PathToStore -Append
    }
}

function Import() {
    # Info: Imports .dat file, if it exists to retreive old saved secrets
    if (Test-Path $Script:PathToStore) {
        $grid.RowCount = 0
        Get-Content $Script:PathToStore | ForEach-Object {
            $grid.Rows.Add("$($_.Split("→")[0])", "$($_.Split("→")[1])","View","Copy","Delete")
        }
    }
}

function HideConsole($show = $false) {
    # Info: Hides the console window
    # Params: show, if I would like to show the console again. Just add the parameter -show:$true
    if ($show) { $v = 5 } else { $v = 0 }
    Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, $v)
}
############################ BODY ##############################################
HideConsole

$form = New-Object System.Windows.Forms.Form -Property @{
    Icon            = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\certutil.exe")
    Text            = "Keep passwords safe & encrypted - Powershell"
    Size            = New-Object System.Drawing.Size (530, 345)
    FormBorderStyle = "FixedDialog"
    MaximizeBox     = $false
    KeyPreview      = $true
}

$grid = New-Object System.Windows.Forms.DataGridView -Property @{
    Size                        = New-Object System.Drawing.Size (300, 250)
    BackgroundColor             = "Gray"
    ColumnHeadersHeightSizeMode = "AutoSize"
    AutoSizeRowsMode            = "AllCells"
    CellBorderStyle             = "None"
    RowHeadersVisible           = $false
    ReadOnly                    = $true
    AllowUserToAddRows          = $False
    AllowUserToDeleteRows       = $False
    Dock                        = "Top"
    GridColor                   = "Black"
}
$form.Controls.Add($grid)

CreateColumns
Import

$labelMaster = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Size (3, 255)
    Size     = New-Object System.Drawing.Size (70, 20)
    Text     = "Master key"
}
$form.Controls.Add($labelMaster)

$textBoxMaster = New-Object System.Windows.Forms.TextBox -Property @{
    Location              = New-Object System.Drawing.Size (73, 253)
    Size                  = New-Object System.Drawing.Size (310, 20)
    UseSystemPasswordChar = $true
}
$form.Controls.Add($textBoxMaster)

$labelName = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Size (3, 282)
    Size     = New-Object System.Drawing.Size (40, 20)
    Text     = "Name"
}
$form.Controls.Add($labelName)

$textBoxName = New-Object System.Windows.Forms.TextBox -Property @{
    Location = New-Object System.Drawing.Size (43, 280)
    Size     = New-Object System.Drawing.Size (128, 20)
}
$form.Controls.Add($textBoxName)

$labelPass = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Size (175, 282)
    Size     = New-Object System.Drawing.Size (60, 20)
    Text     = "Password"
}
$form.Controls.Add($labelPass)

$textBoxPass = New-Object System.Windows.Forms.TextBox -Property @{
    Location              = New-Object System.Drawing.Size (235, 280)
    Size                  = New-Object System.Drawing.Size (148, 20)
    UseSystemPasswordChar = $true
}
$form.Controls.Add($textBoxPass)

$buttonAdd = New-Object System.Windows.Forms.Button -Property @{
    Location = New-Object System.Drawing.Size (390, 253)
    Size     = New-Object System.Drawing.Size (120, 50)
    Text     = "Add secret"
}
$form.Controls.Add($buttonAdd)

$timer = New-Object System.Windows.Forms.Timer -Property @{
    Interval = 3000
    Enabled  = $false
}

$timer.Add_Tick( {
        try {
            $grid[1, $Script:tempArray[2]].Value = $Script:tempArray[0]
        }
        catch { }
        $timer.Enabled = $false
        $grid.Columns[0].Width = 163+180
        $grid.Columns[1].Visible = $false
    })

$buttonAdd.Add_Click( {
        if (($textBoxName.Text.Length -gt 0) -and ($textBoxPass.Text.Length -gt 0) -and ($textBoxMaster.Text.Length -gt 0)) {
            $grid.Rows.Add(($textBoxName.Text), (Set-EncryptedString -Passphrase $textBoxMaster.Text -String ($textBoxPass.Text) -salt ($textBoxName.Text) -init "$($textBoxName.Text.Length)"),"View","Copy","Delete")
            $textBoxName.Text = ""
            $textBoxPass.Text = ""
            $grid.ClearSelection()
            Export
        }
    })

$grid.Add_CellContentClick( {
        $textBoxName.Text = ""
        if ($grid.CurrentCell.ColumnIndex -ge 2 -and (!$timer.Enabled)) {
            # If any of the ButtonColumns are pressed
            switch ($grid.CurrentCell.ColumnIndex) {
                2 {
                    # View column
                    try {
                        $decrypt = (Get-EncryptedString -Encrypted ($grid[1, $grid.CurrentCell.RowIndex].Value) -Passphrase $textBoxMaster.Text -salt ($grid[0, $grid.CurrentCell.RowIndex].Value) -init "$(($grid[0,$grid.CurrentCell.RowIndex].Value).Length)")
                        $Script:tempArray = @($grid[1, $grid.CurrentCell.RowIndex].Value, $decrypt, $grid.CurrentCell.RowIndex)
                        $grid.Columns[0].Width = 163
                        $grid.Columns[1].Visible = $true
                        $grid[1, $grid.CurrentCell.RowIndex].Value = $decrypt
                        $timer.Enabled = $true
                    }
                    catch { $textBoxName.Text = "Unable to decrypt" }
                }
                3 {
                    # Copy column
                    try {
                        $decrypt = (Get-EncryptedString -Encrypted ($grid[1, $grid.CurrentCell.RowIndex].Value) -Passphrase $textBoxMaster.Text -salt ($grid[0, $grid.CurrentCell.RowIndex].Value) -init "$(($grid[0,$grid.CurrentCell.RowIndex].Value).Length)")
                        Set-Clipboard -Value $decrypt
                    }
                    catch { $textBoxName.Text = "Unable to decrypt" }
                }
                4 {
                    # Delete column
                    $grid.Rows.Remove($grid.Rows[$grid.CurrentCell.RowIndex])
                    Export
                }
            }
        }
        $grid.ClearSelection()
    })

[System.Windows.Forms.Application]::Run($form)