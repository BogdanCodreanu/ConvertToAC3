<#
.PARAMETER Path
    The path to the folder that contains files which you want to convert.
.PARAMETER FileTypes
    Type of files which you want to convert. 
    Default: *.mkv
.PARAMETER DefaultBitrate
    If bitrate cannot be detected automatically, will fallback to this bitrate.
    Use 640k if you have 5.1 channel stream.
    Default: 320k.
.PARAMETER PossibleSubtitleExtensions
    A list of possible subtitle extensions.
    If a subtitle is detected with the same name as the found input file it will also copy it
    to the new location, of the ac3 converted file.
    Default: @("srt", "sub")
.PARAMETER OutputDirectory
    If provided, will output the files in this folder location.
.NOTES
    Author: Bogdan Codreanu
    Date:   August 12, 2020
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]
    $Path,
    [Parameter(Mandatory=$false)]
    [string]
    $FileTypes="*.mkv",
    [Parameter(Mandatory=$false)]
    [string]
    $DefaultBitrate="320k",
    [Parameter(Mandatory=$false)]
    [string[]]
    $PossibleSubtitleExtensions=@("srt", "sub"),
    [Parameter(Mandatory=$false)]
    [string]
    $OutputDirectory
)

if (!(Test-Path $Path)) {
    Write-Output "Provided path to $FileTypes files does not exist"
    return
}
$files = (Get-Item -Path (Join-Path -Path $Path -ChildPath $FileTypes))

if ([string]::IsNullOrEmpty($OutputDirectory)) {
    $OutputDirectory = $Path
    Write-Output "* * *`nWill output files in the same directory: $OutputDirectory`n* * *"
} else {
    if (![System.IO.Directory]::Exists($OutputDirectory)) {
        Write-Output "Provided output directory path does not exist"
        return
    }
    Write-Output "* * *`nWill output files to: $OutputDirectory`n* * *"
}

foreach ($file in $files) {
    $filePath=($file.FullName)

    Write-Output "* * *`nAnalyzing file: $newPath`n* * *"

    $newPath = [System.IO.Path]::GetFileNameWithoutExtension($filePath) + 
        "_AC3" + [System.IO.Path]::GetExtension($filePath)
    $newPath = (Join-Path -Path $OutputDirectory -ChildPath $newPath)
    
    # get current data
    & ffprobe -i $filePath 2>ffProbestderr.txt
    $ffOutput = Get-Content ffProbestderr.txt

    $regexData = [regex]::Match($ffOutput, '\bAudio: \b(?<audioType>[a-zA-Z0-9]{2,}), \b(?<bitrate>\d*)\b Hz')

    if ($regexData.Groups["audioType"].Length -eq 0) {
        Write-Output "Could not determine audio type. Converting it to AC3 anyway."
    } else {
        $audioType = $regexData.Groups["audioType"].Value
        Write-Output "Detected audio stream type of: $audioType."
        if ($audioType -eq "ac3") {
            Write-Output "File is already AC3. Skipping."
            continue
        }
    }

    $bitrate = $DefaultBitrate
    if ($regexData.Groups["bitrate"].Length -eq 0) {
        Write-Output "Could not determine bitrate. Will use default bitrate: $DefaultBitrate."
    } else {
        $bitrate = $regexData.Groups["bitrate"].Value
        $bitrate = $bitrate.Substring(0, $bitrate.Length - 2) + "k"
        Write-Output "Detected bitrate of: $bitrate."
    }

    Remove-Item ffProbestderr.txt

    # convert audio with ffmpeg
    ffmpeg -i $filePath -c:v copy -c:a ac3 -b:a $bitrate $newPath

    Write-Output "* * *`nConverted file: $newPath`n* * *"

    # also move subtitles
    foreach ($subExtension in $PossibleSubtitleExtensions) {
        $subtitlePath = [System.IO.Path]::ChangeExtension($filePath, $subExtension)

        if (Test-Path $subtitlePath) {
            $newSubPath = [System.IO.Path]::ChangeExtension($newPath, $subExtension)
            Copy-Item -Path $subtitlePath -Destination $newSubPath
            break
        }
    }
}
