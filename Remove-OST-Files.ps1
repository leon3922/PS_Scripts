# Check Outlook process is active
$outlook = Get-Process OUTLOOK -ErrorAction SilentlyContinue
if($outlook)

{
    $outlook.CloseMainWindow() 
    if(!$outlook.HasExited)
    
    {
        $outlook | Stop-Process -Force # Closing Outlook down
    }

}

# Checking if closed & If not open

if($outlook.HasExited)

{
    Set-Location -Path "C:\Users\$([Environment]::UserName)\AppData\Local\Microsoft\Outlook"
    sleep 3
    Remove-Item * -Include *.ost
} 

elseif(!$outlook)

{
    Set-Location -Path "C:\Users\$([Environment]::UserName)\AppData\Local\Microsoft\Outlook"
    Remove-Item * -Include *.ost
}

echo "RELAUNCH OUTLOOK"