<#
    2016.12.22
    фотобот проброса фото файлов
#>

#############################
# Variables
#############################

# путь до рабочей папки
$Path = Split-Path -Path ($MyInvocation.MyCommand.Path) -Parent
# путь куда складывать скачанные файлы
$OutputFolder = "$Path\foto"

<####################################################
    создаем переменные в глобальном скопе
#>###################################################
# переменная сообщений, сделано чтобы не пробрасывать везде сообщение
if (test-path variable:\Msg) { Remove-Variable Msg }
New-Variable -Name MSG -Option AllScope -Value $false

# переменная сообщений, длительность ожидания сообщения
if (test-path variable:\botUpdateId) { Remove-Variable botUpdateId }
New-Variable -Name botUpdateId -Option AllScope -Value $false

$botUpdateId = 0
$token = "277834943:AAELzqDhk7S1ZhJMsVtH4K384jDlUq5OIew"

# файл лога
$logFile = "$Path\log.txt"
# символ новой строки
$br = "%0A"

<############################################################################################
    логер
#>###########################################################################################
function Resize-Image {

    Param([Parameter(Mandatory=$true)][string]$InputFile, [string]$OutputFile, [int32]$Width, [int32]$Height, [int32]$Scale, [Switch]$Display)

    # Add System.Drawing assembly
    Add-Type -AssemblyName System.Drawing

    # Open image file
    $img = [System.Drawing.Image]::FromFile((Get-Item $InputFile))

    # Define new resolution
    if($Width -gt 0) { [int32]$new_width = $Width }
    elseif($Scale -gt 0) { [int32]$new_width = $img.Width * ($Scale / 100) }
    else { [int32]$new_width = $img.Width / 2 }

    if($Height -gt 0) { [int32]$new_height = $Height }
    elseif($Scale -gt 0) { [int32]$new_height = $img.Height * ($Scale / 100) }
    else { [int32]$new_height = $img.Height / 2 }

    # Create empty canvas for the new image
    $img2 = New-Object System.Drawing.Bitmap($new_width, $new_height)

    # Draw new image on the empty canvas
    $graph = [System.Drawing.Graphics]::FromImage($img2)
    $graph.DrawImage($img, 0, 0, $new_width, $new_height)

    # Create window to display the new image
    if($Display)
    {
        Add-Type -AssemblyName System.Windows.Forms
        $win = New-Object Windows.Forms.Form
        $box = New-Object Windows.Forms.PictureBox
        $box.Width = $new_width
        $box.Height = $new_height
        $box.Image = $img2
        $win.Controls.Add($box)
        $win.AutoSize = $true
        $win.ShowDialog()
    }

    # Save the image
    if($OutputFile -ne "")
    {
        $img2.Save($OutputFile);
    }
}

<############################################################################################
    логер
#>###########################################################################################
function log {
	param ( [parameter(Mandatory = $true)] [string]$Message )
	
    # проверка на длину лога
    if ( ($(Get-ChildItem $logFile).Length / 1mb) -gt 20 ) { Clear-Content $logFile }

	$DT = Get-Date -Format "yyyy.MM.dd HH:mm:ss"
	$MSGOut = $DT + "`t" + $Message
	Out-File -FilePath $logFile -InputObject $MSGOut -Append -encoding unicode
}

<############################################################################################
    получаем сообщения от телеграмм бота, парсим json
#>###########################################################################################
function Bot-Listen {
    param(
            [string]$UpdateId
         )

    $URL = "https://api.telegram.org/bot$token/getUpdates?offset=$UpdateId&timeout=1"

    $Request = Invoke-WebRequest -Uri ( $URL ) -Method Get
    $str = $Request.content

    $ok = ConvertFrom-Json $Request.content

    $str = $ok.result | select -First 1
    $UpdId = ($str).update_id
    $str = ($str).message

    $isJPG = $false
    $docFileName = "!"
    $docFileID = ""
    $docFileSize = 0

    $PhotoFileID = "!"
    $photoFileSize = 0


    ############# проверки на тип сообщения
    if ( $($str.document).mime_type -eq "image/jpeg" ) {  $isJPG = $true  }

    if ( $($str.document).file_name -ne $null ) {
        $docFileName = ($str.document).file_name
        $docFileID = ($str.document).file_id
        $docFileSize = ($str.document).file_size
    }

    # это фотография, подготавливаем структуру для нее
    # выбираем самый большой из файлов на передачу и сохраняем его параметры
    $temp = $str.photo.file_id | select -Last 1

    if ( $($temp) -ne $null ) {
        $BigPhoto = $str.photo | select -Last 1
        $PhotoWidth = $BigPhoto.width
        $PhotoHeight = $BigPhoto.height
        $PhotoFileID = $BigPhoto.file_id
        $PhotoSize = $BigPhoto.file_size
    }

    ############# проверки на тип сообщения
    $props = [ordered]@{  ok=$ok.ok
                        UpdateId = $UpdId
                        Message_ID=$str.message_id
                        first_name=($str.from).first_name
                        last_name =($str.from).last_name
                        chat_id=($str.chat).id
                        text=$str.text
                        isJPG = $isJPG
                        docFileName = $docFileName
                        docFileID = $docFileID
                        docFileSize = $docFileSize
                        PhotoWidth  = $PhotoWidth
                        PhotoHeight = $PhotoHeight
                        PhotoFileID = $PhotoFileID
                        PhotoSize   = $PhotoSize
                       }

    $obj = New-Object -TypeName PSObject -Property $props

    return $obj
}

<############################################################################################
    отправляет сообщение по телеграмму
#>###########################################################################################
function BotSay2 {
    param(  [string]$chat_id = $(Throw "'-chat_id' argument is mandatory"),
            [string]$text = $(Throw "'-text' argument is mandatory"),
            [switch]$markdown,
            [switch]$nopreview
         )

    if($nopreview) { $preview_mode = "True" }
    if($markdown) { $markdown_mode = "Markdown" } else {$markdown_mode = ""}

    $payload = @{   "chat_id" = $chat_id;
                    "text" = $text
                    "parse_mode" = $markdown_mode;
                    "disable_web_page_preview" = $preview_mode;
                }

    $request = Invoke-WebRequest -Uri ("https://api.telegram.org/bot{0}/sendMessage" -f $token) `
                -Method Post -ContentType "application/json;charset=utf-8" `
                -Body (ConvertTo-Json -Compress -InputObject $payload)
}

<############################################################################################
    отправляет сообщение
#>###########################################################################################
function BotSay {
    param(  [string]$text = $(Throw "'-text' argument is mandatory")  )

    $payload = @{ "parse_mode" = "Markdown";
                  "disable_web_page_preview" = "True";
                }

    $URL = "https://api.telegram.org/bot$token/sendMessage?chat_id=$($Msg.Chat_id)&text=$text"

    #$request = Invoke-WebRequest -Uri $URL -Method Post -ContentType "application/json;charset=utf-8" -Body (ConvertTo-Json -Compress -InputObject $payload)
}

<############################################################################################
    Начало работы отсюда
#>###########################################################################################
Write-Host 'fotobot start' -ForegroundColor Green

$ExitFlag = $false

# главный цикл, циклимся пока не будет сброс
while ($ExitFlag -eq $False) {
    Write-Host "новый цикл"

    $Msg = Bot-Listen -UpdateId $botUpdateId

    if ($Msg.UpdateId -gt 1) {
        log "пришло сообщение: $($msg.text) от $($msg.first_name) $($msg.last_name) из чата №$($msg.chat_id)"
        Write-Host "     пришло сообщение: $($msg.text) от $($msg.first_name) $($msg.last_name) из чата №$($msg.chat_id)"  -ForegroundColor Magenta

        $temp = $msg.UpdateID
        $botUpdateId = [System.Convert]::ToInt32( $temp ) + 1

        ###################################
        # смотрим знакомый ли чат айди?
        $UserName = ""
        Switch ( $Msg.chat_id ) 
        {
            # тут номер чата с которого сгружаем фотки
            '123456789' {
                            $UserName = "UserName"
                            $UserAccount = "AccountName"
                            $UserMail = "User Mail"
                            break
                        }
            ### сюда добавляем чаты
            default { BotSay -text "вам необходимо показать это утсройство системному администратору" }
        }

        ################################### скачиваем файл
        if ( $UserName.Length -gt 0 ) {
            if ( $(Test-Path -Path "$Path\$UserAccount") -eq $false ){
                New-Item -Path "$Path\$UserAccount" -ItemType Directory -Force -Verbose
            }

            if ( $($Msg.PhotoFileID) -ne "!" ) {

                Write-Host "пришла фотография $($Msg.PhotoFileID)" -ForegroundColor Yellow

                $FileID = $Msg.PhotoFileID
                $URL = "https://api.telegram.org/bot$token/getFile?file_id=$FileID"
                $Request = Invoke-WebRequest -Uri $URL
        
                foreach ( $JSON in $((ConvertFrom-Json $Request.Content).result) ){
                    $FilePath = $json.file_path
                    $URL = "https://api.telegram.org/file/bot$token/$FilePath"

                    $FilePath = Split-Path -Leaf $FilePath
                    $OutputFile = "$OutputFolder\$UserAccount\$FilePath"
                    Invoke-WebRequest -Uri $URL -OutFile $OutputFile
                    BotSay2 -chat_id $Msg.chat_id -text "скачан файл : file name is ""$($JSON.file_path)""; size $($json.file_size) kb"

                    $DateName = $(Get-Date -Format "yyyy.MM.%d_HH.mm_") + "$FilePath"
                    $NormalizedImageFile = "$OutputFolder\$UserAccount\$DateName"
                    
                    # telegramm сам умеет неплохо сжимать файлы
                    #Resize-Image -InputFile $OutputFile -Scale 50 -OutputFile $NormalizedImageFile
                    #Remove-Item $OutputFile -Force -Verbose
                    
                    BotSay2 -chat_id $Msg.chat_id -text "файл уменьшен и передан"
                }
            }
        }
        ################################### скачиваем файл
    }

    Start-Sleep -Seconds 1
}