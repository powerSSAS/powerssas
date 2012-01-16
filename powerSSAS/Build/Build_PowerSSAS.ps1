param(
    [string]$ReleaseVersion = ""
)

#IntPtr in 64bit is 8 bytes
function is64bit() 
{    if ([IntPtr].Size -eq 4) { return $false }    else { return $true }}


function createZip($SqlVer)
{
write-Host "Starting Zip"
    $ver = $ReleaseVersion -replace "\.", "_"
    &($zip) a -tzip "$base_dir\Setup\powerSSAS_$($SqlVer)_module_$ver.zip" "$base_dir\bin\release\powerSSAS.dll"
    &($zip) a -tzip "$base_dir\Setup\powerSSAS_$($SqlVer)_module_$ver.zip" "$base_dir\powerSSAS.psd1"
    &($zip) a -tzip "$base_dir\Setup\powerSSAS_$($SqlVer)_module_$ver.zip" "$base_dir\help\about_powerSSAS.help.txt"
    &($zip) a -tzip "$base_dir\Setup\powerSSAS_$($SqlVer)_module_$ver.zip" "$base_dir\help\powerSSAS.dll-Help.xml"
    &($zip) a -tzip "$base_dir\Setup\powerSSAS_$($SqlVer)_module_$ver.zip" "$base_dir\FormatData\powerSSAS.Format.Ps1Xml"
    &($zip) a -tzip "$base_dir\Setup\powerSSAS_$($SqlVer)_module_$ver.zip" "$base_dir\TypeData\powerSSAS.Types.Ps1Xml"
}

#$framework_dir_2 = "$env:systemroot\microsoft.net\framework\v2.0.50727"
$framework_dir_3_5 = "$env:systemroot\microsoft.net\framework\v3.5"
$base_dir = [System.IO.Directory]::GetParent("$pwd")
$build_dir = "$base_dir\build"
$sln_file = "$base_dir\powerSSAS.sln"
$setup_2005  ="$base_dir\setup\PowerSSAS_Setup.nsi"
$setup_2008 = "$base_dir\setup\PowerSSAS_Setup2008.nsi"

#utility locations
$msbuild = "$framework_dir_3_5\msbuild.exe"  
$tf = "$env:ProgramFiles\Microsoft Visual Studio 9.0\Common7\IDE\tf.exe"
$nsisPath = "${env:ProgramFiles(x86)}\NSIS\makensis.exe"
$zip = "$env:ProgramFiles\7-zip\7z.exe"

if(is64bit -eq $true) 
{
$tf = "$env:ProgramFiles (x86)\Microsoft Visual Studio 9.0\Common7\IDE\tf.exe"
$nsisPath = "$env:ProgramFiles (x86)\NSIS\makensis.exe"
}


#version files
$lastReleaseFile = "$base_dir\My Project\assemblyinfo.vb"
$versionFiles = @("$lastReleaseFile", "ascii"),
        @("$setup_2008", "ascii"),
        @("$setup_2005", "ascii")

#Clean 

#checkVersion 
    if($ReleaseVersion.Length -eq 0)
    {
        # if we have not been given an explicit build, then we increment the build number
        # of the current release by 1.
        [string]$(get-Content "$($lastReleaseFile)") -cmatch '\d+\.\d+\.\d+\.\d+'
        $ver = $matches[0]
        write-host $ver
        $ver -cmatch "(?'main'\d+\.\d+\.\d+\.)(?'build'\d+)"
        $ReleaseVersion = "$($matches['main'])$([int]$matches['build']+1)"
    }
    
# updateVersion 
    write-Host $ReleaseVersion
    
    $user = read-Host "Codeplex UserName"
    $pass =  read-Host "Codeplex Password" -assecurestring
    
    # I'm only using secure string to hide the password on screen, so we convert it
    # back to a regular string so that it can be passed to codeplex
    $BasicString = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass)
    $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BasicString)
    
    #get the lastest version from codeplex
    &($tf) get "`$/powerSSAS" "/login:$user,$pass" *.* /v:T /recursive
    
    #checkout the files with version information
    $versionFiles | foreach-Object {
        write-Host "checking out: $($_[0])"
        &($tf) checkout "/login:$user,$pass" $_[0]
    }
    
    $versionFiles | foreach-Object {
        $newContent =  $(get-Content "$($_[0])") -replace "\d+\.\d+\.\d+\.\d+", $ReleaseVersion
        set-Content $_[0] $newContent -encoding $_[1]
    }

#Compile
    write-Host "Path: $msbuild"
    &($msbuild) $sln_file /t:Rebuild /p:Configuration=Release /v:q /p:SqlServerTargetEdition=2005
    Write-Host "Executed Compile!"

#BuildSetup 
    write-Host "Starting NSIS"
    &($nsisPath) "$SETUP_2005"
    write-Host "Completed NSIS"

#Create .zip file for SSAS 2005 v2 Module
CreateZip(2005)	
	
#Compile2008 
	write-Host "Path:  $msbuild"
    &($msbuild) $sln_file /t:Rebuild /p:Configuration=Release /v:q /p:SqlServerTargetEdition=2008
    Write-Host "Executed Compile!"

#BuildSetup2008
    write-Host "Starting NSIS"
    &($nsisPath) "$setup_2008"
    write-Host "Completed NSIS"

#Create .zip file for SSAS 2008 v2 Module
CreateZip 2008 




# CheckinVersionFiles
    $cont = read-host "
Release Files Built!    
====================

Upload the files to codeplex now.

Hit 'Y' and enter to continue and checkin and label the files on codeplex. 
Just hit Enter to abort without updating the source repository."
    write-Host $cont
    if ($cont -match "^Y$")
    {
        write-Host "checking files into codeplex and applying release label"
        $versionFiles | foreach-Object {
            write-Host "checking out: $($_[0])"
            &($tf) checkin "/login:$user,$pass" "/comment:`"Updating to version $ReleaseVersion`"" "$($_[0])"
        }   
        &($tf) label "`"Release $ReleaseVersion`"" "/login:$user,$pass" /v:T /recursive *.*
    }
    else
    {
        write-Host "No Files checked-in"
    }