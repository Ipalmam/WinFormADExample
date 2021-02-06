import-module activedirectory
function checkpath {##testing if file and folder exist to sve report in case needed, this function create it in case it does not
    param ()
    $folder = Test-Path C:\ITFormProcess
    if ($folder -eq $true) {
        Clear-content "C:\ITFormProcess\Report.txt" -Force
    }
    else {
        New-Item -ItemType Directory -Force -Path C:\ITFormProcess
        New-Item -Path C:\ITFormProcess -Name "Report.txt" -ItemType "file"
    }
}
function GetUserData {              ##This function validate data entry, if format is correct search for ADUser account information 
    param (                          
        $UserInput
    )
    $ITFormSummary
    $ITFormNotes
    if ($UserInput -match "\+") {     ## Validate data entry, is invalid explain user why, if data is correct saves ADUser account information on an array
        [System.Windows.MessageBox]::Show('Invalid character in input data')
        return
    }else {
        if ($UserInput -match '\,') {
            Write-Output "123"
            $ITFormNotes = get-aduser -filter 'Name -like $UserInput' -properties  * | select-object DisplayName, SamAccountName, Manager, Department, Title, Office
        }else {
            if ($UserInput -match '\@') {
                $ITFormNotes = get-aduser -filter 'UserPrincipalName -like $UserInput' -properties  * | select-object DisplayName, SamAccountName, Manager, Department, Title, Office
            }else {
                Write-Output "456"
            $ITFormNotes = get-aduser -filter 'SamAccountName -like $UserInput' -properties  * | select-object DisplayName, SamAccountName, Manager, Department, Title, Office
            }
        }
    }
    if ($null -eq $ITFormNotes) {      ##If ADUser account information was retrieved is validated to be right, if is not explains to user
        $UserWInput = '*' + $UserInput + '*'
        $ITFormNotes = get-aduser -filter 'UserPrincipalName -like $UserWInput' -properties  * | select-object DisplayName, SamAccountName
        Add-Content -Path "C:\ITFormProcess\Report.txt" $ITFormNotes
        $ITFormNotes = Get-Content -Path "C:\ITFormProcess\Report.txt"
        Clear-content "C:\ITFormProcess\Report.txt" -Force
        [System.Windows.MessageBox]::Show($ITFormNotes, "Users found with incomplete data provided")
        return
    }else {                 ##If ADUser data is correct format data as requested and save it in clipboard and text boxes
        Add-Content -Path "C:\ITFormProcess\Report.txt" $ITFormNotes
        $ITFormNotes = Get-Content -Path "C:\ITFormProcess\Report.txt"
        Clear-content "C:\ITFormProcess\Report.txt" -Force
        $Datas = @()
        $Datas = $ITFormNotes.Split('=')
        $Value = @()
        $ReportdData = @()
        $ChanArray = @()
        $FormatManager = @()
        if ($Datas.Length -eq 16) {     ##validate data size as Canonical names have different lenght depending of company names
            for ($i = 0; $i -lt $Datas.Length; $i++){
                switch ($i) {
                    1 { $Value = $Datas[$i].Split(';')
                        $Chan = $Value[0]
                        $ChanArray = $Chan.Split(' ')
                        [array]::Reverse($ChanArray)
                        $LastNAme = $ChanArray[1].Substring(0,$ChanArray[1].Length-1)
                        $ReportdData += "Name: " + $ChanArray[0] + " "+ $LastNAme }   ##Formaing Name as requested
                    2 { $Value = $Datas[$i].Split(';')
                        $ReportdData += "User ID: " + $Value[0] }
                    4 { $Value = $Datas[$i].Split('\')
                        $Manager = $Value[0] + $Value[1]
                        $Manager = $Manager.Substring(0,$Manager.Length-3)
                        $FormatManager = $Manager.Split(' ')
                        [array]::Reverse($FormatManager)
                        $ForMana = $FormatManager[1].Substring(0,$FormatManager[1].Length-1)
                        $ReportdData += "Manager: " + $FormatManager[0] + " " + $ForMana }
                    13 {$Depto = $Datas[$i]                                                           
                        $Depto = $Depto.Substring(0,$Depto.Length-7)
                        $ReportdData += "Department: " + $Depto }
                    14 {$JobP = $Datas[$i]
                        $JobP = $JobP.Substring(0,$JobP.Length-8)
                        $ReportdData += "Job Title: " + $JobP }
                    15 {$Office = $Datas[$i]
                        $Office = $Office.Substring(0,$Office.Length-1)
                        $ReportdData += "Location: " + $Office }        
                    Default {}
                }
            }
        }else {
            if ($Datas.Length -eq 15) {
                for ($i = 0; $i -lt $Datas.Length; $i++){
                    switch ($i) {
                        1 { $Value = $Datas[$i].Split(';')
                            $Chan = $Value[0]
                            $ChanArray = $Chan.Split(' ')
                            [array]::Reverse($ChanArray)
                            $LastNAme = $ChanArray[1].Substring(0,$ChanArray[1].Length-1)
                            $ReportdData += "Name: " + $ChanArray[0] + " "+ $LastNAme }   ##Formaing Name as requested
                        2 { $Value = $Datas[$i].Split(';')
                            $ReportdData += "User ID: " + $Value[0] }
                        4 { $Value = $Datas[$i].Split('\')
                            $Manager = $Value[0] + $Value[1]
                            $Manager = $Manager.Substring(0,$Manager.Length-3)
                            $FormatManager = $Manager.Split(' ')
                            [array]::Reverse($FormatManager)
                            $ForMana = $FormatManager[1].Substring(0,$FormatManager[1].Length-1)
                            $ReportdData += "Manager: " + $FormatManager[0] + " " + $ForMana }
                        12 {$Depto = $Datas[$i]
                            $Depto = $Depto.Substring(0,$Depto.Length-7)
                            $ReportdData += "Department: " + $Depto }
                        13 {$JobP = $Datas[$i]
                            $JobP = $JobP.Substring(0,$JobP.Length-8)
                            $ReportdData += "Job Title: " + $JobP }
                        14 {$Office = $Datas[$i]
                            $Office = $Office.Substring(0,$Office.Length-1)
                            $ReportdData += "Location: " + $Office }        
                        Default {}
                    }
                }
            }
        }
        
    }
    $ReportdData = $ReportdData + "Requested Access: " + "Request Details: " + "-----------------------------------------"##formating data and saving it in an array
    $tBSummaryEU.Text = "UA – Existing User – ", $ChanArray[0], $LastNAme                                                 ##formating data and saving it in a textbox
    $tBSummaryNE.Text = "UA – New User – ", $ChanArray[0], $LastNAme                                                      ##formating data and saving it in a textbox
    Set-Clipboard -Value $ReportdData                                                                                     ##Sagind data in clipboard
}
checkpath
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")  ##Loading Asamblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void] [System.Windows.Forms.Application]::EnableVisualStyles()
Add-Type -AssemblyName PresentationFramework
$UserInput = ""
$Form = New-Object system.Windows.Forms.Form ##Object creation Form
$Form.Size = New-Object System.Drawing.Size(420,600) ##Form cuztomization, fisrt line size
$form.MaximizeBox = $false 
$Form.StartPosition = "CenterScreen" 
$Form.FormBorderStyle = 'Fixed3D' 
$Form.Text = "IT Form Fill out Automation Tool"
$labelInputData = New-Object System.Windows.Forms.Label ##Create a label
$labelInputData.Location = New-Object System.Drawing.Point(10,20)
$labelInputData.Size = New-Object System.Drawing.Size(380,20)
$labelInputData.Text = 'Type user ID or email address'
$form.Controls.Add($labelInputData)
$tBInputData = New-Object System.Windows.Forms.TextBox ##Create a text box to input user information
$tBInputData.Location = New-Object System.Drawing.Point(10,40)
$tBInputData.Size = New-Object System.Drawing.Size(380,40)
$tBInputData.Font = New-Object System.Drawing.Font("Lucida Console",18,[System.Drawing.FontStyle]::Regular)
$tBInputData.Add_keyDown({
    if ($_.KeyCode -eq "Enter") {
        GetUserData($tBInputData.Text)
    }
})
$form.Controls.Add($tBInputData)
$labelInstruc = New-Object System.Windows.Forms.Label
$labelInstruc.Location = New-Object System.Drawing.Point(10,140)
$labelInstruc.Size = New-Object System.Drawing.Size(380,80)
$labelInstruc.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Regular)
$labelInstruc.Text = 'After click in Search button or type Enter key on above text box notes to fill IT Form Notes Field are copied in the clipboard, just paste them in IT Form Notes, after that copy from IT Summary text boxes what you need in IT Form Summary if no errors occurs'
$form.Controls.Add($labelInstruc)
$labelSummaryEU = New-Object System.Windows.Forms.Label ##Create a label
$labelSummaryEU.Location = New-Object System.Drawing.Point(10,260)
$labelSummaryEU.Size = New-Object System.Drawing.Size(380,20)
$labelSummaryEU.Text = 'IT Form Summary Existing User'
$form.Controls.Add($labelSummaryEU)
$copyEUButton = New-Object System.Windows.Forms.Button ##Create a button
$copyEUButton.Location = New-Object System.Drawing.Point(280,260) 
$copyEUButton.Size = New-Object System.Drawing.Size(380,20)
$form.Controls.Add($copyEUButton)
$tBSummaryEU = New-Object System.Windows.Forms.TextBox ##Create a text box to input user information
$tBSummaryEU.Location = New-Object System.Drawing.Point(10,280)
$tBSummaryEU.Size = New-Object System.Drawing.Size(380,20)
$tBSummaryEU.Multiline = $true
$form.Controls.Add($tBSummaryEU)
$copyEUButton = New-Object System.Windows.Forms.Button ##Create a button
$copyEUButton.Location = New-Object System.Drawing.Point(10,310) 
$copyEUButton.Size = New-Object System.Drawing.Size(380,20)
$copyEUButton.Text = "Copy IT Form Summary Existing User"
$exiUser = "123123123"
$copyEUButton.Add_Click({$Global:exiUser = $tBSummaryEU.Text;Set-Clipboard -Value $exiUser}) ##get textbox.text value from global scope with heritage
$form.Controls.Add($copyEUButton)
$labelBSummaryNE = New-Object System.Windows.Forms.Label ##Create a label
$labelBSummaryNE.Location = New-Object System.Drawing.Point(10,370)
$labelBSummaryNE.Size = New-Object System.Drawing.Size(280,20)
$labelBSummaryNE.Text = 'IT Form Summary New User'
$form.Controls.Add($labelBSummaryNE)
$tBSummaryNE = New-Object System.Windows.Forms.TextBox ##Create a text box to output user information
$tBSummaryNE.Location = New-Object System.Drawing.Point(10,390)
$tBSummaryNE.Size = New-Object System.Drawing.Size(380,20)
$form.Controls.Add($tBSummaryNE)
$copyNUButton = New-Object System.Windows.Forms.Button ##Create a button
$copyNUButton.Location = New-Object System.Drawing.Point(10,420) 
$copyNUButton.Size = New-Object System.Drawing.Size(380,20)
$copyNUButton.Text = "Copy IT Form Summary New User"
$newUser = "jojojojoj"
$copyNUButton.Add_Click({$Global:newUser = $tBSummaryNE.Text;Set-Clipboard -Value $newUser})  ##get textbox.text value from global scope with heritage
$form.Controls.Add($copyNUButton)
$Searchbutton = New-Object System.Windows.Forms.Button ##Create a button
$Searchbutton.Location = New-Object System.Drawing.Size(140,100) 
$Searchbutton.Size = New-Object System.Drawing.Size(100,30) 
$Searchbutton.Text = "Search" 
$Searchbutton.Add_Click({GetUserData($tBInputData.Text)}) ##use a function to get data from active directory
$Form.Controls.Add($Searchbutton)
$Form.ShowDialog()
