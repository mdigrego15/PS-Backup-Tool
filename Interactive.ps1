Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName ReachFramework

# XAML generated from Visual Studio
[xml]$Form = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Backup Tool" Height="400" Width="600">
    <Grid>
        <Button Name="openFile" HorizontalAlignment="Left" Height="24" Margin="230,124,0,0" VerticalAlignment="Top" Width="118" Content="Choose Folder..." FontFamily="Microsoft YaHei UI"/>
        <TextBlock Height="33" Margin="203,63,203,0" Text="IT Backup Tool" TextWrapping="Wrap" VerticalAlignment="Top" FontFamily="Microsoft YaHei UI" FontWeight="Normal" FontSize="20" TextAlignment="Center"/>
        <TextBox Name="Path" HorizontalAlignment="Left" Height="20" Margin="67,126,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="145" FontFamily="Microsoft YaHei UI"/>
        <Button Name="start" Content="Begin Backup" Height="40" Margin="366,116,67,0" VerticalAlignment="Top" Background="#FFE6E0E0" FontSize="14" FontFamily="Microsoft YaHei UI"/>
        <Label Content="Robocopy Summary" HorizontalAlignment="Center" Height="34" VerticalAlignment="Center" Width="232" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" FontFamily="Microsoft YaHei UI"/>
        <TextBox Name="logDisplay" Height="152" Margin="10,200,10,0" Text="" TextWrapping="Wrap" VerticalAlignment="Top" FontSize="14" FontFamily="Microsoft YaHei UI" TextAlignment="Center"/>
    </Grid>
</Window>

"@

# Linking the XAML to the code
$NR=(New-Object System.Xml.XmlNodeReader $Form)
$Win=[Windows.Markup.XamlReader]::Load($NR)

function Tech-Copy{
# This CmdletBinding line makes the function an advanced function which gives it access to standard common parameters like -Verbose
[CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] $Source,
        [Parameter(Mandatory=$true)] $Destination
    )
    # Credit to Trevor Sullivan for the Staging Code: https://stackoverflow.com/questions/13883404/custom-robocopy-progress-bar-in-powershell#comment47098288_25334958
    # Regular expression to hold the number of Bytes
    $RegexBytes = '(?<=\s+)\d+(?=\s+)';

    #Setting the parameters for robocopy
    # /MIR to mirror the contents of the folder
    # /NJH for no Job Header, this is so that the regular expression does not pick up numbers displayed in the header and add it to the bytestotal value
    # /NJS for no Job Summary, same reason as the regular expression
    $CommonRobocopyParams = '/MIR /NJH /NJS /NDL /NP /NC /R:0 /W:0 /BYTES /xj';
    $StagingLogPath = '{0}\temp\{1} robocopy staging.log' -f $env:windir, (Get-Date -Format 'yyyy-MM-dd HH-mm-ss');
    $StagingArgumentList = '"{0}" "{1}" /LOG:"{2}" /L {3}' -f $Source, $Destination, $StagingLogPath, $CommonRobocopyParams;
    # Running robocopy in a logging mode to find out the number of Bytes that need to be transported
    Start-Process -FilePath robocopy.exe -ArgumentList $StagingArgumentList -NoNewWindow -Wait


    # Get the total number of files that will be copied
    $StagingContent = Get-Content -Path $StagingLogPath;
    $TotalFileCount = $StagingContent.Count - 1;
    [RegEx]::Matches(($StagingContent -join "`n"), $RegexBytes) | % { $BytesTotal = 0; } { $BytesTotal += $_.Value; };

    #Output the number of Bytes to be copied
    
    
     #Setting the log path to the logs folder on the server
     $RoboParam = '/MIR /NDL /NJH /NJS /NP /NC /R:0 /W:0 /BYTES /xj'
     $RobocopyLogPath = '{0}\temp\{1} robocopy.log' -f $env:windir, (Get-Date -Format 'yyyy-MM-dd HH-mm-ss');
     $ArgumentList = '"{0}" "{1}" /LOG:"{2}" {3}' -f $Source, $Destination, $RobocopyLogPath, $RoboParam;
     #This line is not running: Can't run robocopy from a variable? But it works in the backup script
     $Robocopy = Start-Process -FilePath robocopy.exe -ArgumentList $ArgumentList -NoNewWindow -PassThru 
     Start-Sleep -Milliseconds 100

     #Allow the script some time to generate entries in the log file

     #Progress bar loop
     while (!$Robocopy.HasExited) {
        $BytesCopied = 0;
        $LogContent = Get-Content -Path $RobocopyLogPath;
        $BytesCopied = [Regex]::Matches($LogContent, $RegexBytes) | ForEach-Object -Process { $BytesCopied += $_.Value; } -End { $BytesCopied; };
        $CopiedFileCount = $LogContent.Count - 1;
        $Percentage = 0;
        if ($BytesCopied -gt 0) {
           $Percentage = (($BytesCopied/$BytesTotal)*100)
        }
        Write-Progress -Activity Robocopy -Status ("Copied {0} of {1} files; Copied {2} of {3} bytes" -f $CopiedFileCount, $TotalFileCount, $BytesCopied, $BytesTotal) -PercentComplete $Percentage
        }

    #Stores the content of the log file into a string so that it can be displayed in the scroll box in the app
    #$log = Get-Content $RobocopyLogPath | Out-String
    $convert = 1024
    $KBCopied = $BytesCopied / $convert
    $MBCopied = $KBCopied / $convert
    $GBCopied = $MBCopied / $convert 
    $output.text = "`n The Robocopy job is complete!`n`n You have moved $BytesCopied bytes`n`n Total GB`n------------`n $GBCopied"
    #robocopy $Source $Destination /MIR /R:0 /W:0 /xj /xd "OneDrive - Providence College"
    return 
    
}

function Open-Folder{
    # Creating a folder browser object and displaying the "choose a folder" dialog
    $FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog 
    $FileBrowser.RootFolder = 'MyComputer'
    $FileBrowser.ShowDialog()
    # Setting the Path text box to have the value of the Selected path in the Folder Dialog
    $userPath.Text = $FileBrowser.SelectedPath
}



#Initializing variables to manipulate
$userPath = $Win.FindName("Path")
$output = $Win.FindName("logDisplay")
$open = $Win.FindName("openFile")
$backup = $Win.FindName("start")


# Open File Picker on Click for openFile Button
$open.Add_Click({Open-Folder})

# Adding the click function for the Start button
# Calls the Tech-Copy function and passes the Client and Server variables as the Source and Destination parameters 
$backup.Add_Click({
    #Initialize Client and Server
    $Client = $userpath.text
    # Initialize the User object to be just the name of the user folder by replacing the other parths of the path with an empty string
    # This makes it so that on the server, the folder will display with the same name as the user
    $User = $userPath.Text.Replace("C:\Users\","")
    #Your Server here
    $Server = " "
    $robocopyPath = '{0}\{1}' -f $Server, (Get-Date -Format 'yyyy-MM-dd HH-mm-ss')
    #[System.Windows.MessageBox]::Show($User)
    Tech-Copy -Source $Client -Destination $robocopyPath
    
})

#Shows the GUI (Should go at bottom)
$Win.showdialog()
