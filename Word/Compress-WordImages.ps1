[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

function Compress-Image ([string]$Path, [int]$Quality)
{

    $SourceFile = Get-Item -Path $Path

    # Create bitmap from file
    $OriginalBitmap = New-Object System.Drawing.Bitmap($Path)

    # Create a copy of the bitmap
    $Bitmap = New-Object System.Drawing.Bitmap($OriginalBitmap)

    # Dispose of the original bitmap, so the file can be deleted
    $OriginalBitmap.Dispose()

    # Delete the original file
    Remove-Item $SourceFile

    $Encoder = [System.Drawing.Imaging.Encoder]::Quality
    $EncoderParameters = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $EncoderParameters.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($Encoder, $Quality)
    $ImageCodecInfo = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object {$_.MimeType -eq 'image/jpeg'}

    # Save the bitmap to file using original filename
    $Bitmap.Save($Path,$ImageCodecInfo, $($EncoderParameters))

    # Dispose of the bitmap
    $Bitmap.Dispose()
}

Function Compress-Document ([string[]]$SourcePath, [ValidateRange(0,100)] [int]$Quality)
{
    $results = @()

    ForEach($File in Get-ChildItem -Path $SourcePath) {

        $Base = ("{0}\{1}" -f $File.Directory, $File.BaseName)

        $OriginalFileSize = (Get-Item -Path "$Base.docx" | Measure-Object -Sum Length).Sum

        # Rename the Word document to zip, so it can be expanded
        Rename-Item -Path "$Base.docx" -NewName "$Base.zip"

        # Expand the zip
        Expand-Archive "$Base.zip" -DestinationPath "$Base"

        # Remove the original zip file
        Remove-Item -Path "$Base.zip"

        # Compress all images inside the document folder structure
        ForEach($Image in Get-ChildItem -Path "$Base" -Filter *.jpeg -Recurse) {
            Compress-Image -Path $Image.FullName -Quality $Quality
        }

        # Compress the folder to new zip file
        Compress-Archive -Path "$Base\*" -DestinationPath "$Base.zip"

        # Remove the Word document folder
        Remove-Item "$Base" -Force -Recurse

        # Rename the zip file to Word document
        Rename-Item -Path "$Base.zip" -NewName "$Base.docx"

        $CompressedFileSize = (Get-Item -Path "$Base.docx" | Measure-Object -Sum Length).Sum

        # Results
        $item = New-Object PSObject
        $item | Add-Member -MemberType NoteProperty -Name 'FileName' -Value $File.Name
        $item | Add-Member -MemberType NoteProperty -Name 'OriginalFileSize' -Value $OriginalFileSize
        $item | Add-Member -MemberType NoteProperty -Name 'CompressedFileSize' -Value $CompressedFileSize
        $item | Add-Member -MemberType NoteProperty -Name 'CompressionRatio' -Value ("{0:P2}" -f ($CompressedFileSize / $OriginalFileSize))

        $results += $item
    }

    return $results
}

$path = 'E:\SEMSjablonen\Gedeelde documenten\Documents\RES\Letters'
$docs = Get-ChildItem -Path "$path\*.docx" | Where-Object {$_.LastWriteTime -lt (Get-Date).AddMonths(-3) -and $_.Length -gt 2048000} | Sort-Object -Property Length,LastWriteTime -Descending
foreach ($doc in $docs)
{
    # Get original file dates
    $orgLwt = $doc.LastWriteTime
    $orgLat = $doc.LastAccessTime

    # Compress Word file
    Compress-Document -SourcePath $doc.FullName -Quality 50

    # Restore file dates on compressed Word file
    $newDoc = (Get-ChildItem -Path $doc.FullName)
    $newDoc.LastWriteTime = $orgLwt
    $newDoc.LastAccessTime = $orgLat
}
