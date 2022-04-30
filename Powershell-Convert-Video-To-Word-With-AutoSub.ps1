
param (
    [string]$targetFilePath = "",
    [string]$locale = "zh-TW",
    [decimal]$similarityPass = 0.99906
    <#
        af      Afrikaans
        ar      Arabic
        az      Azerbaijani
        be      Belarusian
        bg      Bulgarian
        bn      Bengali
        bs      Bosnian
        ca      Catalan
        ceb     Cebuano
        cs      Czech
        cy      Welsh
        da      Danish
        de      German
        el      Greek
        en      English
        eo      Esperanto
        es      Spanish
        et      Estonian
        eu      Basque
        fa      Persian
        fi      Finnish
        fr      French
        ga      Irish
        gl      Galician
        gu      Gujarati
        ha      Hausa
        hi      Hindi
        hmn     Hmong
        hr      Croatian
        ht      Haitian Creole
        hu      Hungarian
        hy      Armenian
        id      Indonesian
        ig      Igbo
        is      Icelandic
        it      Italian
        iw      Hebrew
        ja      Japanese
        jw      Javanese
        ka      Georgian
        kk      Kazakh
        km      Khmer
        kn      Kannada
        ko      Korean
        la      Latin
        lo      Lao
        lt      Lithuanian
        lv      Latvian
        mg      Malagasy
        mi      Maori
        mk      Macedonian
        ml      Malayalam
        mn      Mongolian
        mr      Marathi
        ms      Malay
        mt      Maltese
        my      Myanmar (Burmese)
        ne      Nepali
        nl      Dutch
        no      Norwegian
        ny      Chichewa
        pa      Punjabi
        pl      Polish
        pt      Portuguese
        ro      Romanian
        ru      Russian
        si      Sinhala
        sk      Slovak
        sl      Slovenian
        so      Somali
        sq      Albanian
        sr      Serbian
        st      Sesotho
        su      Sudanese
        sv      Swedish
        sw      Swahili
        ta      Tamil
        te      Telugu
        tg      Tajik
        th      Thai
        tl      Filipino
        tr      Turkish
        uk      Ukrainian
        ur      Urdu
        uz      Uzbek
        vi      Vietnamese
        yi      Yiddish
        yo      Yoruba
        zh-CN   Chinese (Simplified)
        zh-TW   Chinese (Traditional)
        zu      Zulu
    #>
)

$CurrentPS1File = $(Get-Item -Path "$PSCommandPath")
Set-Location "$($CurrentPS1File.PSParentPath)" | Out-Null

$Result = ''

$ProgramPath = ".\autosub_app\autosub_app.exe"
$ProgramPara = @('--list-languages')
$ProgramItem = Get-Item $ProgramPath
$psi = New-object System.Diagnostics.ProcessStartInfo 
$psi.CreateNoWindow = $true 
$psi.UseShellExecute = $false 
$psi.RedirectStandardOutput = $true 
$psi.RedirectStandardError = $true 
$psi.FileName = $ProgramItem.FullName
$psi.Arguments = $ProgramPara
$process = New-Object System.Diagnostics.Process 
$process.StartInfo = $psi 
[void]$process.Start()
$autosubLanList = $process.StandardOutput.ReadToEnd() 
$process.WaitForExit() 
$autosubLanArray = @()
forEach( $lis in $($autosubLanList | Select-String -pattern '(.+)\t(.+)' -AllMatches).Matches ) {
    $autosubLanArray += $($lis | Select-Object @{Label='LocaleCode';Expression={$lis.groups[1].Value}},@{Label='LanguageFullNameFromLocale';Expression={$lis.groups[2].Value}})
}

$chooseLan = $($autosubLanArray | Out-GridView -Title '選擇影片語言' -PassThru)
if (-not ([string]::IsNullOrEmpty($chooseLan.LocaleCode))) {
    $locale = $chooseLan.LocaleCode
}

$LanguageFullNameFromLocale = $($autosubLanArray | Where-Object{$_.LocaleCode -eq $locale}).LanguageFullNameFromLocale

$Result = $Result + "`r`n" + '選擇語言：' + $LanguageFullNameFromLocale
clear
Write-Output $Result

if (-not ([string]::IsNullOrEmpty($targetFilePath))) {
    if ( $(Test-Path -Path $targetFilePath -PathType Leaf) -eq $false ) {
        [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
        $targetFileItem = New-Object System.Windows.Forms.OpenFileDialog
        $targetFileItem.Filter = "mp4 files (*.mp4)|*.mp4|All files (*.*)|*.*" 
        If($targetFileItem.ShowDialog() -eq "OK") {
            $targetFilePath = $targetFileItem.FileName
            $targetFileItem = Get-Item $targetFilePath
            $Result_Target = $('來源檔案："' + $targetFileItem.FullName + '"')
            $Result = $Result + "`r`n"  + $Result_Target
            clear
            Write-Output $Result
        }



    }
} else {
    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
    $targetFileItem = New-Object System.Windows.Forms.OpenFileDialog
    $targetFileItem.Filter = "mp4 files (*.mp4)|*.mp4|All files (*.*)|*.*"
    If($targetFileItem.ShowDialog() -eq "OK") {
        $targetFilePath = $targetFileItem.FileName
        $targetFileItem = Get-Item $targetFilePath
        $Result_Target = $('來源檔案："' + $targetFileItem.FullName + '"')
        $Result = $Result + "`r`n"  + $Result_Target
        clear
        Write-Output $Result
    }
}

if (-not ([string]::IsNullOrEmpty($targetFilePath))) {
    if ( $(Test-Path -Path $targetFilePath -PathType Leaf) -eq $true ) {

        $st = Get-Date

        $SubPSOList = @()

        $targetSubPath = $($targetFileItem.Directory.FullName + '\' +  $targetFileItem.BaseName + '.srt')

        $Result_AutoSub = '聲音轉文字 (0/1)'

        if ( $(Test-Path -Path $targetSubPath -PathType Leaf) -eq $false ) {
            $ProgramPath = ".\autosub_app\autosub_app.exe"
            $ProgramPara = ' -S ' + $locale + ' -D ' + $locale + ' -F srt "' + $targetFilePath + '"'
            $ProgramItem = Get-Item $ProgramPath
            Invoke-Command -ScriptBlock {
                $process = Start-Process -FilePath "$($ProgramItem.FullName)" -ArgumentList "$ProgramPara" -Wait -NoNewWindow
            }
            $et = $(get-date) - $st
            $tt = "{0:HH:mm:ss}" -f ([datetime]$et.Ticks)
            $Result_AutoSub = '聲音轉文字 (1/1) [此階段已費時 ' + $tt + ' ]'
        } else {
            $Result_AutoSub = '聲音轉文字 (0/1) [檔案已存在]'
        }
        $Result = $Result + "`r`n" + $Result_AutoSub
        clear
        Write-Output $Result

        try{
            $targetSubItem = Get-Item $targetSubPath
        }catch{}

        if ( $(Test-Path -Path $targetSubPath -PathType Leaf) -eq $true ) {

            $st = Get-Date

            forEach( $itemInSub in $($(Get-Content $targetSubPath -Raw -Encoding UTF8) | Select-String -pattern "([0-9]+)`n([0-9]{2,2}:[0-9]{2,2}:[0-9]{2,2},[0-9]{3,3}) \-\-> ([0-9]{2,2}:[0-9]{2,2}:[0-9]{2,2},[0-9]{3,3})`n(.*)`n" -AllMatches).Matches ) {
                $Index = $itemInSub.groups[1].Value
                $startTime = $itemInSub.groups[2].Value
                $endTime = $itemInSub.groups[3].Value
                $Content = $itemInSub.groups[4].Value

                $startTimeItem = $($startTime | Select-String -pattern "([0-9]{2,2}):([0-9]{2,2}):([0-9]{2,2}),([0-9]{3,3})").Matches
                $startTimeStamp = $(($($startTimeItem.groups[1].Value -as [int])*3600 + $($startTimeItem.groups[2].Value -as [int])*60 + $($startTimeItem.groups[3].Value -as [int])*1) -as [string]) + '.' + $($startTimeItem.groups[4].Value)

                $endTimeItem = $($endTime | Select-String -pattern "([0-9]{2,2}):([0-9]{2,2}):([0-9]{2,2}),([0-9]{3,3})").Matches
                $endTimeStamp = $(($($endTimeItem.groups[1].Value -as [int])*3600 + $($endTimeItem.groups[2].Value -as [int])*60 + $($endTimeItem.groups[3].Value -as [int])*1) -as [string]) + '.' + $($endTimeItem.groups[4].Value)

                $randomTimeStamp = Get-Random -Minimum $startTimeStamp -Maximum $endTimeStamp

                New-Item -ItemType Directory -Force -Path "$(($targetSubItem.Directory.FullName + '\' +  $targetSubItem.BaseName) + '-Frames\')" | Out-Null
                $imagePath = "$(($targetSubItem.Directory.FullName + '\' +  $targetSubItem.BaseName + '-Frames\' +  $targetSubItem.BaseName) + '-' + $Index + '.png')"
                $imagePreinsert = $true

                $SubPSO = $itemInSub | select-object @{n='Index';e={$Index}},@{n='startTime';e={$startTime}},@{n='startTimeStamp';e={$startTimeStamp}},@{n='endTime';e={$endTime}},@{n='endTimeStamp';e={$endTimeStamp}},@{n='randomTimeStamp';e={$randomTimeStamp}},@{n='Content';e={$Content}},@{n='imagePath';e={$imagePath}},@{n='imagePreinsert';e={$imagePreinsert}},@{n='similarity';e={''}}
                $SubPSOList += $SubPSO

            }

            $SubPSOListCount = $SubPSOList.Count

            $randomTimeTemp = ''
            $LostImage = $false
            $ImageIndex = 0
            $Result_ffmpegOutFrames = '輸出圖片 (0/' + $SubPSOListCount + ')'
            foreach($so in $SubPSOList){
                $soIndex = $($so.Index -as [int])
                $randomTimeTemp += 'lt(prev_pts*TB\,' + $so.randomTimeStamp + ')*gte(pts*TB\,' + $so.randomTimeStamp + ') +'
                $et = $(get-date) - $st
                $tt = "{0:HH:mm:ss}" -f ([datetime]$et.Ticks)
                $Result_ffmpegOutFrames = '輸出圖片 (' + $ImageIndex + '/' + $SubPSOListCount + ') [此階段已費時 ' + $tt + ' ]'
                clear
                Write-Output $($Result + "`r`n" + $Result_ffmpegOutFrames)
                if ( $(Test-Path -Path $so.imagePath -PathType Leaf) -eq $false ) {
                    $LostImage = $true
                }
                if($($soIndex%100) -eq 1){
                    $rangeStartTimeStamp = $so.startTimeStamp
                }
                if( ($($soIndex%100) -eq 0) -and ($LostImage -eq $true)){
                    $randomTimeTemp = $randomTimeTemp -replace ' ?\+$',''
                    $ProgramPath = ".\autosub_app\ffmpeg.exe"
                    $imagePath = "$(($targetSubItem.Directory.FullName + '\' +  $targetSubItem.BaseName + '-Frames\' +  $targetSubItem.BaseName) + '-%d.png')"
                    #$ProgramPara = ' -skip_frame nokey -i "' + $targetFilePath + '" -ss ' + $startTimeStamp + ' -to ' + $endTimeStamp + ' -vsync 0 "' + $imagePath
                    $ProgramPara = '-i "' + $targetFilePath + '" -ss ' + $rangeStartTimeStamp + ' -to ' + $so.endTimeStamp + ' -filter:v "select=''' + $randomTimeTemp + '''" -vsync drop -start_number ' + $([Math]::Truncate($soIndex/100)*100-99 -as [string]) + ' "' + $imagePath + '"'
                    $ProgramItem = Get-Item $ProgramPath
                    Invoke-Command -ScriptBlock {
                        Start-Process -FilePath "$($ProgramItem.FullName)" -WindowStyle Hidden -ArgumentList "$ProgramPara" -Wait
                    }
                    $randomTimeTemp = ''
                    $LostImage = $false
                    $ImageIndex = $ImageIndex + 100
                }
                if ($so.Index -eq $SubPSOListCount) {
                    if ($LostImage -eq $true) {
                        $randomTimeTemp = $randomTimeTemp -replace ' ?\+$',''
                        $ProgramPath = ".\autosub_app\ffmpeg.exe"
                        $imagePath = "$(($targetSubItem.Directory.FullName + '\' +  $targetSubItem.BaseName + '-Frames\' +  $targetSubItem.BaseName) + '-%d.png')"
                        $ProgramPara = '-i "' + $targetFilePath + '" -ss ' + $rangeStartTimeStamp + ' -to ' + $so.endTimeStamp + ' -filter:v "select=''' + $randomTimeTemp + '''" -vsync drop -start_number ' + $([Math]::Truncate($soIndex/100+1)*100-99 -as [string]) + ' "' + $imagePath + '"'
                        $ProgramItem = Get-Item $ProgramPath
                        Invoke-Command -ScriptBlock {
                            Start-Process -FilePath "$($ProgramItem.FullName)" -WindowStyle Hidden -ArgumentList "$ProgramPara" -Wait
                        }
                        $randomTimeTemp = ''
                        $LostImage = $false
                        $ImageIndex = $ImageIndex + ($SubPSOListCount % 100)
                    }
                    $et = $(get-date) - $st
                    $tt = "{0:HH:mm:ss}" -f ([datetime]$et.Ticks)
                    $Result_ffmpegOutFrames = '輸出圖片 (' + $ImageIndex + '/' + $SubPSOListCount + ') [此階段已費時 ' + $tt + ' ]'
                    $Result = $Result + "`r`n" + $Result_ffmpegOutFrames
                    clear
                    Write-Output $Result
                    break;
                }
                clear
                Write-Output $($Result + "`r`n" + $Result_ffmpegOutFrames)
            }

            $st = Get-Date
            for($index=0;$index -lt $SubPSOListCount;$index++){
                $et = $(get-date) - $st
                $tt = "{0:HH:mm:ss}" -f ([datetime]$et.Ticks)
                $Result_FrameSimilarity = '比較圖片相似度 (' + $index + '/' + ($SubPSOListCount-1) + ') [此階段已費時 ' + $tt + ' ]'
                if ($(Test-Path -Path $SubPSOList[$index].imagePath -PathType Leaf) -eq $true) {
                    if (($similarityPass -gt 0) -and ($(Test-Path -Path $SubPSOList[$index+1].imagePath -PathType Leaf) -eq $true) ) {
                        clear
                        Write-Output $($Result + "`r`n" + $Result_FrameSimilarity)
                        $ProgramPath = ".\autosub_app\ffmpeg.exe"
                        $ProgramPara = @('-i "' + $SubPSOList[$index].imagePath + '"','-i "' + $SubPSOList[$index+1].imagePath + '"','-lavfi "ssim" ','-f null - ')
                        $ProgramItem = Get-Item $ProgramPath
                        $psi = New-object System.Diagnostics.ProcessStartInfo 
                        $psi.CreateNoWindow = $true 
                        $psi.UseShellExecute = $false 
                        $psi.RedirectStandardOutput = $true 
                        $psi.RedirectStandardError = $true 
                        $psi.FileName = $ProgramItem.FullName
                        $psi.Arguments = $ProgramPara
                        $process = New-Object System.Diagnostics.Process 
                        $process.StartInfo = $psi
                        [void]$process.Start()
                        $process.WaitForExit()
                        $similarityInfo = $process.StandardOutput.ReadToEnd() + $process.StandardError.ReadToEnd()

                        $SubPSOList[$index].similarity = $($similarityInfo | Select-String -pattern '.*SSIM.*All:([0-9]\.[0-9]+).*' -AllMatches).Matches.groups[1].Value

                        $($SubPSOList[$index].similarity -as [decimal])
                        $similarityPass
                        ($($SubPSOList[$index].similarity -as [decimal]) -lt $similarityPass)

                        if($($SubPSOList[$index].similarity -as [decimal]) -lt $similarityPass){
                            $SubPSOList[$index].imagePreinsert = $true
                        } else {
                            $SubPSOList[$index].imagePreinsert = $false
                        }

                        $et = $(get-date) - $st
                        $tt = "{0:HH:mm:ss}" -f ([datetime]$et.Ticks)
                        $Result_FrameSimilarity = '比較圖片相似度 (' + $($index+1) + '/' + ($SubPSOListCount-1) + ') [此階段已費時 ' + $tt + ' ]'
                    }
                } else {
                    $SubPSOList[$index].imagePreinsert = $false
                }
            }
            $Result = $Result + "`r`n" + $Result_FrameSimilarity
            clear
            Write-Output $Result

            $imageSimilarityList = $($targetFileItem.Directory.FullName + '\' +  $targetFileItem.BaseName + '-SimilarityList.csv')
            $SubPSOList | Select-Object -Property Content,imagePath,similarity,imagePreinsert | ConvertTo-Csv | Out-File -FilePath $imageSimilarityList -Encoding utf8 -Force

            try{
                $word = New-Object -ComObject word.application
            }catch{}
            $ImageIndex = 0
            $Result_WriteToWord_Sub = '寫入字幕到Word (0/' + $SubPSOListCount + ')'
            $Result_WriteToWord_Image = '插入圖片到Word (' + $ImageIndex + '/' + $SubPSOListCount + ')'
            $st = Get-Date
            if($word){
                $wordDocPath = $($targetFileItem.Directory.FullName + '\' +  $targetFileItem.BaseName + '.docx')
                $word.Visible = $false
                $doc = $word.documents.add()
                $doc.Styles["Normal"].ParagraphFormat.SpaceAfter = 0
                $doc.Styles["Normal"].ParagraphFormat.SpaceBefore = 0
                $margin = 36 # 1.26 cm
                $doc.PageSetup.LeftMargin = $margin
                $doc.PageSetup.RightMargin = $margin
                $doc.PageSetup.TopMargin = $margin
                $doc.PageSetup.BottomMargin = $margin
                $selection = $word.Selection
                foreach($so in $SubPSOList){
                    $soimagePath = $so.imagePath
                    if ($so.imagePreinsert -eq $true) {
                        $selection.InlineShapes.AddPicture($soimagePath)
                        $ImageIndex++
                        $Result_WriteToWord_Image = '插入圖片到Word (' + $ImageIndex + '/' + $SubPSOListCount + ')'
                    }
                    $selection.TypeText($so.Content)
                    $Result_WriteToWord_Sub = '寫入字幕到Word (' + $so.Index + '/' + $SubPSOListCount + ')'
                    $selection.TypeParagraph() | Out-Null
                    $et = $(get-date) - $st
                    $tt = "{0:HH:mm:ss}" -f ([datetime]$et.Ticks)
                    clear
                    Write-Output $($Result + "`r`n" + $Result_WriteToWord_Sub + "    " + $Result_WriteToWord_Image + ' [此階段已費時 ' + $tt + ' ]')
                }
                $selection.TypeParagraph() | Out-Null
                $doc.SaveAs($wordDocPath)
                $doc.Close()
                $word.Quit()
                $et = $(get-date) - $st
                $tt = "{0:HH:mm:ss}" -f ([datetime]$et.Ticks)
                clear
                Write-Output $($Result + "`r`n" + $Result_WriteToWord_Sub + "    " + $Result_WriteToWord_Image + ' [此階段已費時 ' + $tt + ' ]')
            }
            $et = $(get-date) - $st
            $tt = "{0:HH:mm:ss}" -f ([datetime]$et.Ticks)
            $Result = $Result + "`r`n" + $Result_WriteToWord_Sub + "    " + $Result_WriteToWord_Image + ' [此階段已費時 ' + $tt + ' ]'
        }
    clear
    Write-Output $Result
    }
}

Write-Output "`r`n按任意鍵或直接關閉結束`r`n"
[void][System.Console]::ReadKey(1)
