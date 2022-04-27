
Add-Type -AssemblyName PresentationFramework


Function Get-Information {
    [CmdletBinding()]
    # This param() block indicates the start of parameters declaration
    param (
        <# 
            This parameter accepts the name of the target computer.
            It is also set to mandatory so that the function does not execute without specifying the value.
        #>
        [Parameter(Mandatory)]
        [string]$identity,
        [string]$fromEmail,
        [string]$toEmail,
        [string]$sentDate,
        [string]$emailSubject,
        [string]$hasAttachment,
        [string]$attachmentName
    )

    $KQLString = "from:$fromEmail AND to:$toEmail AND sent:$sentDate AND subject:""$emailSubject"""


    If( $hasAttachment -eq "true") {
        $KQLString = $KQLString + " AND hasAttachment:$hasAttachment AND attachmentnames:$attachmentName"
    } elseif ( $hasAttachment -eq "false") {
        $KQLString = $KQLString + "  AND hasAttachment:$hasAttachment"
    }

    $KQLString
}

# See if the user is signed in
$sessionState = Get-PSSession | Select-Object State -ExpandProperty State



    If( $sessionState -contains "Opened" ) {
        Write-Host "Already Signed In"
    } else {
        Write-Host "Please see window to sign in"
        Connect-IPPSSession -UserPrincipalName $userPrincipalName
    }



# What is the XAML file?
$xamlFile = @"
<Window x:Class="complianceSearch.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:complianceSearch"
        mc:Ignorable="d"
        Title="Compliance Search - Credit Card Info" Height="450" Width="800">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="119*"/>
            <ColumnDefinition Width="681*"/>
        </Grid.ColumnDefinitions>
        <Label x:Name="Ticket_Number" Content="Ticket Number&#xD;&#xA;" HorizontalAlignment="Center" Margin="0,20,0,0" VerticalAlignment="Top" Width="110" Height="27"/>
        <Label Content="Sender Email&#xD;&#xA;" HorizontalAlignment="Center" Margin="0,47,0,0" VerticalAlignment="Top" Width="112" Height="29"/>
        <Label Content="Received Email" Margin="0,76,0,0" VerticalAlignment="Top" Height="31" HorizontalAlignment="Center" Width="110"/>
        <Label Content="Sent Date" Margin="0,107,0,0" VerticalAlignment="Top" Height="26" HorizontalAlignment="Center" Width="110"/>
        <Label Content="Email Subject" Margin="0,138,0,0" VerticalAlignment="Top" Height="58" HorizontalAlignment="Center" Width="110"/>
        <Label Content="Attachment?" Margin="0,201,0,0" VerticalAlignment="Top" Height="35" HorizontalAlignment="Center" Width="110"/>
        <Label Content="Attachment Name or Extension (file.dox, .pdf)" Margin="0,236,0,0" VerticalAlignment="Top" Height="35" HorizontalAlignment="Center" Width="108"/>
        <RadioButton x:Name="hasAttachmentY" Content="Yes" Margin="21,202,0,0" VerticalAlignment="Top" Grid.Column="1" GroupName="attachment" Height="15" HorizontalAlignment="Left" Width="62" TabIndex="5"/>
        <RadioButton x:Name="hasAttachmentN" Content="No" HorizontalAlignment="Left" Margin="22,218,0,0" VerticalAlignment="Top" Grid.Column="1" GroupName="attachment" Height="15" Width="61" TabIndex="6"/>
        <TextBox x:Name="toEmail" Grid.Column="1" HorizontalAlignment="Left" Margin="16,76,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="164" Height="31" Opacity="0.96" TabIndex="2"/>
        <TextBox x:Name="fromEmail" Grid.Column="1" HorizontalAlignment="Left" Margin="16,47,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="164" Height="29" TabIndex="1"/>
        <TextBox x:Name="identity" Grid.Column="1" HorizontalAlignment="Left" Margin="16,20,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="164" Height="27" TabIndex="0"/>
        <TextBox x:Name="emailSubject" Grid.Column="1" HorizontalAlignment="Left" Margin="16,140,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="165" Height="57" TabIndex="4"/>
        <TextBox x:Name="attachmentName" Grid.Column="1" HorizontalAlignment="Left" Margin="16,236,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="170" Height="35" TabIndex="7"/>
        <TextBox x:Name="txtOutput" Grid.Column="1" HorizontalAlignment="Left" Margin="242,0,0,0" TextWrapping="Wrap" VerticalAlignment="Center" Width="418" Height="302"/>
        <Button x:Name="btnSubmit" Content="Submit" Grid.Column="1" HorizontalAlignment="Left" Margin="10,296,0,0" VerticalAlignment="Top" Width="72" Height="20" TabIndex="8"/>
        <DatePicker x:Name="sentDate" Grid.Column="1" HorizontalAlignment="Left" Margin="16,107,0,0" VerticalAlignment="Top" Height="33" Width="164" TabIndex="3" DisplayDate="2022-04-22" SelectedDateFormat="Short"/>
        <Button x:Name="btnExit" Content="Exit" HorizontalAlignment="Center" Margin="0,372,0,0" VerticalAlignment="Top" Width="76"/>
        <Button x:Name="btnSearch" Content="Start Search" Grid.Column="1" HorizontalAlignment="Left" Margin="242,382,0,0" VerticalAlignment="Top" Height="27" Width="82"/>

    </Grid>
</Window>
"@

# Create the window
$inputXML = $xamlFile
$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
[XML]$xaml = $inputXML

# Read XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
    $window = [Windows.Markup.XamlReader]::Load( $reader )
} catch {
    Write-Warning $_.Exception
    throw
}

$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    #"trying item $($_.Name)"
    try {
        Set-Variable -Name "var_$($_.Name)" -Value $window.FindName($_.Name) -ErrorAction Stop
    } catch {
        throw
    }
}

#Get-Variable var_*


#Action when submit button is clicked
$var_btnSubmit.Add_Click( {

    $this.IsEnabled = $False


    $global:identity = $var_identity.Text
    $fromEmail = $var_fromEmail.Text
    $toEmail = $var_toEmail.Text
    $sentDate = $var_sentDate.Text
    $sentDate = Get-Date $sentDate -Format "yyyy-MM-dd"
    $emailSubject = $var_emailSubject.Text
    if ($var_hasAttachmentY.IsChecked) {
        $hasAttachment = "true"
    } elseif ($var_hasAttachmentN.IsChecked) {
        $hasAttachment = "false"
    }

    $attachmentName = $var_attachmentName.Text

    $var_txtOutput.Text = ""

    $KQLString = Get-Information -identity $identity -fromEmail $fromEmail -toEmail $toEmail -sentDate $sentDate -emailSubject $emailSubject -hasAttachment $hasAttachment -attachmentName $attachmentName

    #$var_txtOutput.Text = $var_txtOutput.Text + "[Debugging]:KQL String - $KQLString`n"
    $var_txtOutput.Text = $var_txtOutput.Text + "Ensure information is correct:`n"
    $var_txtOutput.Text = $var_txtOutput.Text + "Ticket Number: $identity`n"
    $var_txtOutput.Text = $var_txtOutput.Text + "From Address: $fromEmail`n"
    $var_txtOutput.Text = $var_txtOutput.Text + "To Address: $toEmail`n"
    $var_txtOutput.Text = $var_txtOutput.Text + "Sent Date: $sentDate `n"
    $var_txtOutput.Text = $var_txtOutput.Text + "Email Subject: $emailSubject`n"
    $var_txtOutput.Text = $var_txtOutput.Text + "Has Attachment: $hasAttachment`n"
    $var_txtOutput.Text = $var_txtOutput.Text + "Attachment name or extension: $attachmentName`n"
    $var_txtOutput.Text = $var_txtOutput.Text + " `n"
    
    New-ComplianceSearch -Name "$identity" -ExchangeLocation All -ContentMatchQuery $KQLString

    $var_txtOutput.Text = $var_txtOutput.Text + "Click Start Search to begin compliance search.`n"
    $var_txtOutput.Text = $var_txtOutput.Text + "Window will freeze until search completes.`n"




})

#Activates when Search button is clicked. Starts the compliance search and waits for more info.

$var_btnSearch.Add_Click( {
    Start-ComplianceSearch -Identity "$global:identity"
    $var_txtOutput.Text += "Starting search, please wait...`n"
    Start-Sleep -Seconds 5
    $complianceSearchStatus = Get-ComplianceSearch -Identity "$global:identity" | Select-Object -ExpandProperty Status


        while($complianceSearchStatus -ne "Completed" ) {
            Start-Sleep -Seconds 60

        $complianceSearchStatus = Get-ComplianceSearch -Identity "$global:identity" | Select-Object -ExpandProperty Status

    
        
        
        If($complianceSearchStatus -eq "Error") {
            $msgBoxInput = [System.Windows.MessageBox]::Show('An error occured, please contact IT.','Error','Exit','Error')
            if ($msgBoxInput -eq 'Exit') {
                $window.close()
                exit
            }
        }
        If($complianceSearchStatus -eq "Completed"){
            break
        }
    }


    $numberOfItems = Get-ComplianceSearch -Identity "$global:identity" | Select-Object -Property Items -ExpandProperty Items
    
    If($numberOfItems -gt 0 -and $numberOfItems -lt 25) {
        $var_txtOutput.Text = $var_txtOutput.Text + "Number of items found is $numberOfItems`n"
        $purge = [System.Windows.MessageBox]::Show('Ready to delete items? Items deleted in this way are unrecoverable.','Attention', 'YesNoCancel','Warning')
        If($purge -eq "Yes"){
            New-ComplianceSearchAction -SearchName "$global:identity" -Force -Confirm:$false -Purge -PurgeType HardDelete
            $var_txtOutput.Text += "Items have been marked for permanent deletion`n"
        }
    }
    If($numberOfItems -eq 0) {
        $var_txtOutput.Text = $var_txtOutput.Text + "No items returned.`n"
    }
})

#Exit button
$var_btnExit.Add_Click( {
    $window.close()
    exit
})

$Null = $window.ShowDialog()