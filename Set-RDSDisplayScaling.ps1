####################################################################################
##
## (C) 2018 Bechtle GmbH & Co. KG IT-Systemhaus Neckarsulm - Tim Bitzer
##
##
## Filename:      Set-RDSDisplayScaling.ps1
##
## Version:       1.0
##
## Requirements:  -none-
##
## Changes:       2018-11-28    Initial creation 
##	            
##
####################################################################################

<#
			.SYNOPSIS
			Allows you to change the Scaling factor (DPI scaling)

			.DESCRIPTION
            A little GUI allows you to choose between 100%, 125%, 150%, 200% Scaling factor.
            Sets the corresponding value for the DWORD property in the registry "HKCU:\Control Panel\Desktop" 

			.EXAMPLE
			Set-ScalingFactor.ps1
#>

### Load Assemblies
Add-Type -AssemblyName PresentationFramework # Gui
Add-Type -AssemblyName System.Windows.Forms


#___________________________________________________________________
#region FUNCTION Set-ScalingFactor -> Button Save Click
function Set-ScalingFactor
{
    <#
        .SYNOPSIS
        Button 'Save' Click method
        
        .DESCRIPTION
        sets the reg key and asks for logoff
        
        .EXAMPLE
        Set-ScalingFactor
    #>


    ### Check if Combobox has a selected value
    if ($CB_ChangePercent.Text -eq "")
    {
        ### If no value is selected, mark combobox with red border
        $BORDER_CB.BorderThickness = "5,5,5,5"
    }
    else
    {
        try
        {
            # Set registry value
            New-ItemProperty -Path $Key -Name $PropertyName -PropertyType DWord -Value $ScalingFactorLookupTable[($CB_ChangePercent.Text)] -Force | Out-Null
        }
        catch
        {
            throw 
        }

        ### MessageBox asking for logoff
        $MessageBoxText = ("Damit die {0}nderungen wirksam werden, ist eine Abmeldung erforderlich`n`nM{1}chten Sie sich jetzt abmelden?" -f [Char]0x00C4, [char]0x00F6)
        $MessageBoxTitle = 'Abmeldung erforderlich'
        $MessageBoxResult = [System.Windows.Forms.MessageBox]::Show($MessageBoxText, $MessageBoxTitle, 4)
    
        ### Evaluate users choice
        if ($MessageBoxResult -eq "Yes")
        {
            ### User choice was 'yes'
            Start-Process -FilePath "C:\windows\system32\logoff.exe"
        }
        else
        {
            ### User choice was 'no'
            $LBL_Status.Visibility = 'Visible'
            $LBL_CurrentValue.Content = $CB_ChangePercent.Text
        }
    }
}


#___________________________________________________________________
#region FUNCTION Get-CurrentScalingFactor
function Get-CurrentScalingFactor
{
    <#
        .SYNOPSIS
        Returns the current Scaling factor
        
        .DESCRIPTION
        Returns the current Scaling factor 
        -> Value of 'HKCU:\Control Panel\Desktop' 'LogPixels'
        
        .EXAMPLE
        Get-CurrentScalingFactor
    #>

    try
    {
        # Read and return current reg value if exists
        $Value = (Get-ItemPropertyValue -Path $Key -Name $PropertyName -ErrorAction SilentlyContinue)  
        return ($ScalingFactorLookupTable.GetEnumerator() | Where-Object {$_.Value -eq $Value}).Name   
    }    
    catch
    {
        # Reg value not found
        return "N/A"
    }
}

function Show-GUI
{
    <#
        .SYNOPSIS
        Show GUI function
        
        .DESCRIPTION
        Initializes all WPF controls and starts the GUI
        
        .EXAMPLE
        Show-GUI
    #>

    ### Parameters
    $Key = 'HKCU:\Control Panel\Desktop'
    $PropertyName = "LogPixels"

    $ScalingFactorLookupTable = @{
        "100%" = 96
        "125%" = 120
        "150%" = 144
        "200%" = 192
    }

    # Icon as Base64 encoded data
    $Base64Image = @"
    AAABAAEAICAAAAAAIAANAwAAFgAAAIlQTkcNChoKAAAADUlIRFIAAAAgAAAAIAgGAAAAc3p69AAAAAFzUkdCAK7OHOkAAAAEZ0FNQQAAsY8L/GEFAA
    AACXBIWXMAAA7DAAAOwwHHb6hkAAACoklEQVRYR+2WWU9TURSFb6GWWsBaaUspVMSqgDiP0RhN/CPOI0ZfnWfjbzhPiozOsz76z+5yrfZuPBBq7MCT
    PclKk96bfGvtvc85N2iv9mqvxXXzF1ZbMU8R1Vv8c3AhbEprpoDgxR/FqI6XQCcV5zM9T1Cdt1bBQM/0UrBBBex6BSSptVSKStxusYG+2XAJ2JIatJ
    vmZLB3pqpkKw0MzIeVMktxysApAgUWcB2VngXWUxkqdaeFBroIFFypVWpLLLCgAm6YA/qoLJWb57O7LTKgZAa31D0CUwIbNE9oP1VYAAao9L0WGMjO
    VUvvw5VaZVZiAxt08DUwRJWozP0mDajvAhvcSi64UqvM/R649AbYSA1Tm97ynQdNGlCfNXDquZ9ccEtdAVMG3UyV3wFbqFwzBtIzYWWradplpBZciQ
    0s6Lb3wCg19oHvPGzQQJ59t31updfAqecquw8fiRJvjcDjBG+ndnwEio8aMGB9lwFtPSu9pl0Dp54vhyu1wBOE7qR2fQJ2U0OPGzCgI1QnnaXXyaZ9
    bqUv0oCGTWX34Uos8J7PwD5q/xeafFKnATtqte3U++XprfQaOPXchyvx3gh8kDr8lSbrMVCIjloZsMn3e++nV+k1cGM0oLIrueAHIvCRb8BRqvy0Dg
    MquV00fvlt8nXCrZRePVfZDS7wse/A8R8czGf/aMC/YlUFGbCtp/LnovLrhLPea+ItvXp+KEou+AnCT/7kO7UMhCHBN6Yryk5OITjlFhU77RA/45A4
    65Ckus859J53SF9wyFx0yF5yyFOFyw7FKw6lqw7Dkw5a5WvV39Hr1d/x5zUM/E0dVJxfMvqY0H2e4o2mW00Xi852Ha85HjA6ZLTPtdU07SOUeq6yK7
    ngEysaaK/2+j9XEPwGPD/Hf2DZ7gMAAAAASUVORK5CYII=
"@
    
    # WPF XAML
    [XML]$XAML = @"
    <Window x:Class="Set_ScalingFactor.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Set_ScalingFactor"
        mc:Ignorable="d"
        Title="Set Scaling Factor" 
        Height="240" 
        Width="320">
        <Grid>
            <Border x:Name="BORDER_CB" HorizontalAlignment="Right" Margin="0,78,10,0" VerticalAlignment="Top" Width="73" Height="33" BorderBrush="Red" BorderThickness="5">
                <ComboBox x:Name="CB_ChangePercent" FontSize="16" Grid.RowSpan="2" Background="#FFE7F0F9">
                    <ComboBoxItem Content="100%"/>
                    <ComboBoxItem Content="125%"/>
                    <ComboBoxItem Content="150%"/>
                </ComboBox>
            </Border>
            <Label x:Name="LBL_CurrentTitle" Content="Current Scaling factor: " Margin="10,33,0,0" Height="38" VerticalAlignment="Top" FontSize="16" HorizontalAlignment="Left" Width="181"/>
            <Label x:Name="LBL_CurrentValue" Content="100%" Margin="0,33,10,0" FontSize="16" HorizontalAlignment="Right" Width="73" Height="38" VerticalAlignment="Top"/>
            <Label x:Name="LBL_ChangeTitle" Content="Change Scaling factor to:" Margin="10,83,0,0" Height="35" VerticalAlignment="Top" FontSize="16" HorizontalAlignment="Left" Width="181"/>
            <Label x:Name="LBL_Status" Content="Logoff required" HorizontalAlignment="Left" Margin="10,2,0,0" VerticalAlignment="Top" Foreground="Red" Width="282" HorizontalContentAlignment="Center"/>
            <Button x:Name="BTN_Save" Content="Save" Margin="10,0,10,10" Height="43" VerticalAlignment="Bottom" Grid.Row="1" BorderThickness="1" Background="#FFEEEEEE" FontSize="16" BorderBrush="#FFD1D1D1"/>            
        </Grid>
    </Window>
"@ -replace 'mc:Ignorable="d"', '' -replace 'x:Name', 'Name'  

    # Initialize XML node reader
    $XAML.Window.RemoveAttribute("x:Class") 
    $NodeReader = New-Object System.Xml.XmlNodeReader $XAML

    # Initialize Base64 icon 
    $Bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
    $Bitmap.BeginInit()
    $Bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($Base64Image)
    $Bitmap.EndInit()
    $Bitmap.Freeze()
	
    # Main window
    $Window = [Windows.Markup.XamlReader]::Load($NodeReader)
    $Window.Title = "RDS Display Scaling"
    $Window.Topmost = $true
    $Window.Icon = $Bitmap

    # ComboBox
    $CB_ChangePercent = $Window.FindName('CB_ChangePercent')
    $CB_ChangePercent.add_GotFocus( {$BORDER_CB.BorderThickness = "0"})
    $CB_ChangePercent.Items.Clear()
    foreach ($Item in ($ScalingFactorLookupTable.GetEnumerator() | Sort-Object))
    {
        # Add each element of $ScalingFactorLookupTable to the ComboBox
        $CB_ChangePercent.Items.Add($Item.Key) | Out-Null
    }

    # Border for ComboBox
    $BORDER_CB = $Window.FindName('BORDER_CB')
    $BORDER_CB.BorderThickness = "0"

    # Label Current Title
    $LBL_CurrentTitle = $Window.FindName('LBL_CurrentTitle')
    $LBL_CurrentTitle.Content = "Aktuelle Vergr{0}{1}erung:" -f [char]0x00F6, [char]0x00DF # German special chars ä,ö,ü,ß,...

    # Label Current Value
    $LBL_CurrentValue = $Window.FindName('LBL_CurrentValue')
    $LBL_CurrentValue.Content = (Get-CurrentScalingFactor)

    # Label Change Title
    $LBL_ChangeTitle = $Window.FindName('LBL_ChangeTitle')
    $LBL_ChangeTitle.Content = "Vergr{0}{1}erung {2}ndern:" -f [char]0x00F6, [char]0x00DF, [char]0x00E4 # German special chars ä,ö,ü,ß,...

    # Label Change Title
    $LBL_Status = $Window.FindName('LBL_Status')
    $LBL_Status.Content = "Abmeldung noch ausstehend..." -f [Char]0x00C4
    $LBL_Status.Visibility = 'Hidden'

    # Button Save
    $BTN_Save = $Window.FindName('BTN_Save')
    $BTN_Save.Content = "Speichern"
    $BTN_Save.add_MouseEnter( {$BTN_Save.Background = "#FFEEEEEE"})
    $BTN_Save.Add_Click( {Set-ScalingFactor})

    # Show main winow -> start GUI
    $Window.ShowDialog() | Out-Null    
}

Show-GUI