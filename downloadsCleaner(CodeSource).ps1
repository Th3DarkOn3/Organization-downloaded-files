#####################################################################
#                                                                   #
# Script pour l'organisation automatique des fichier de télécharger #
# Create by Th3DarkOn3                                              #
# ###################################################################

#Demande d'autorisations d'administrateur
if(!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList "-File `"$($MyInvocation.MyCommand.Path)`"  `"$($MyInvocation.MyCommand.UnboundArguments)`""
    Exit
}
#######################################
$user = $env:UserName #GetUser
#Création du chemin des dossiers
$pdfFolder = "C:\Users\$user\Downloads\PDF"
$exeFolder = "C:\Users\$user\Downloads\Exe"
$wordFolder = "C:\Users\$user\Downloads\Word"
$excelFolder = "C:\Users\$user\Downloads\Excel"
$pPointFolder = "C:\Users\$user\Downloads\Powerpoint"
$zipFolder = "C:\Users\$user\Downloads\Zip"
$textFolder = "C:\Users\$user\Downloads\Text"
$isoFolder = "C:\Users\$user\Downloads\ISO"
$pTracerFolder = "C:\Users\$user\Downloads\PacketTracer"
$torrentFolder = "C:\Users\$user\Downloads\Torrent"
$otherFolder = "C:\Users\$user\Downloads\Other"
############################################
#Création d'une "base de données" avec les extensions les plus importantes
$EXECUTABLE = @('.exe', '.msi')
$WORD = @('.doc','.docm','.docx','.dot','.dotm','.dotx','.odt','.rtf','.wps')
$EXCEL = @('.csv','.dbf','.dif','.ods','.xls','.prn','.xlam','.xlsb','.xlsm','.xlsx','.xla','.xlam','.xlt','.xltm','.xltx','.xlw')
$POWEPOINT = @('.odp','.pot','.potm','.potx','.ppa','.ppam','.pps','.ppsm','.ppsx','.ppt','.pptm','.pptx','.rtf','.wmf')
$COMPRESSED = @('.rar', '.zip', '.cab', '.arj', '.lzh', '.tar', '.gzip', '.uue', '.bzip2', '.z', '.7-zip')
$IMAGES = @('.apng','.avif','.gif','.jpg','.jpeg','.jfif','.pjpeg','.pjp','.png','.svg','.webp','.bmp','.ico','.cur','.tif','.tiff')
$VIDEO = @('.webm','.mkv','.flv','.vob','.ogv','.ogg','.drc','.gifv','.mng','.avi','.mts','.m2ts','.ts','.mov','.qt','.wmv','.yum','.rm','.rmvb','.viv','.asf','.amv','.mp4','.m4p','.m4v','.mpg','.mp2','.mpeg','.mpe','.mpv','.m2v','.svi','.3gp','.3g2','.mxf','.roq','.nsv','.f4v','.f4b','.f4p','.f4a')
$AUDIO = @('.aif','.cda','.mid','.midi','.mp3','.mpa','.ogg','.wav','.wma','.wpl')
##############################################
#Déplacer des fichiers pdf
if (Test-Path -Path $pdfFolder) {
    Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *.pdf |
        Where-Object { $_.Extension -eq '.pdf' } | Move-Item -Destination C:\Users\$user\Downloads\PDF
} else {
    New-Item -Path C:\Users\$user\Downloads -Name "PDF" -ItemType "directory"
    Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *.pdf |
        Where-Object { $_.Extension -eq '.pdf' } | Move-Item -Destination C:\Users\$user\Downloads\PDF
}

# Déplacer des fichiers exe
if (Test-Path -Path $exeFolder) {
    foreach($EXT in $EXECUTABLE){
        Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *$EXT | 
            Where-Object { $_.Extension -eq $EXT } | Move-Item -Destination C:\Users\$user\Downloads\Exe
    }
}else{
    New-Item -Path C:\Users\$user\Downloads -Name "Exe" -ItemType "directory"
    foreach($EXT in $EXECUTABLE){
        Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *$EXT | 
            Where-Object { $_.Extension -eq $EXT } | Move-Item -Destination C:\Users\$user\Downloads\Exe
    }
}
##############################################
# Déplacer des fichiers word
if (Test-Path -Path $wordFolder) {
    foreach($EXT in $WORD){
        Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *$EXT | 
            Where-Object { $_.Extension -eq $EXT } | Move-Item -Destination C:\Users\$user\Downloads\Word
    }
}else {
    New-Item -Path C:\Users\$user\Downloads -Name "Word" -ItemType "directory"
    foreach($EXT in $WORD){
        Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *$EXT | 
            Where-Object { $_.Extension -eq $EXT } | Move-Item -Destination C:\Users\$user\Downloads\Word
    }
}
##############################################
# Déplacer des fichiers EXCEL
if (Test-Path -Path $excelFolder) {
    foreach($EXT in $EXCEL){
        Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *$EXT | 
            Where-Object { $_.Extension -eq $EXT } | Move-Item -Destination C:\Users\$user\Downloads\Excel
    }
}else{
    New-Item -Path C:\Users\$user\Downloads -Name "Excel" -ItemType "directory"
    Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *$EXT | 
        Where-Object { $_.Extension -eq $EXT } | Move-Item -Destination C:\Users\$user\Downloads\Excel
}
##############################################
# Déplacer des fichiers POWERPOINT
if (Test-Path -Path $pPointFolder) {
    foreach($EXT in $POWEPOINT){
        Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *$EXT | 
            Where-Object { $_.Extension -eq $EXT } | Move-Item -Destination C:\Users\$user\Downloads\Powerpoint
    }
}else{
    New-Item -Path C:\Users\$user\Downloads -Name "Powerpoint" -ItemType "directory"
    foreach($EXT in $POWEPOINT){
        Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *$EXT | 
            Where-Object { $_.Extension -eq $EXT } | Move-Item -Destination C:\Users\$user\Downloads\Powerpoint
    }
}
##############################################
# Déplacer des fichiers RAR
if (Test-Path -Path $zipFolder) {
    foreach($EXT in $COMPRESSED){
        Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *$EXT | 
            Where-Object { $_.Extension -eq $EXT } | Move-Item -Destination C:\Users\$user\Downloads\Zip
    }
}else{
    New-Item -Path C:\Users\$user\Downloads -Name "Zip" -ItemType "directory"
    foreach($EXT in $COMPRESSED){
        Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *$EXT | 
            Where-Object { $_.Extension -eq $EXT } | Move-Item -Destination C:\Users\$user\Downloads\Zip
    }
}
##############################################
# Déplacer des fichiers Images
foreach($EXT in $IMAGES){
    Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *$EXT | 
        Where-Object { $_.Extension -eq $EXT } | Move-Item -Destination C:\Users\$user\OneDrive\Immagini
}
##############################################
# Déplacer des fichiers VIDEOS
foreach($EXT in $VIDEO){
    Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *$EXT | 
        Where-Object { $_.Extension -eq $EXT } | Move-Item -Destination C:\Users\$user\OneDrive\Video
}
##############################################
# Déplacer des fichiers AUDIO
foreach($EXT in $AUDIO){
    Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *$EXT | 
        Where-Object { $_.Extension -eq $EXT } | Move-Item -Destination C:\Users\$user\OneDrive\Musica
}
##############################################
#Déplacer des fichiers text
if (Test-Path -Path $textFolder) {
    Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *.txt |
        Where-Object { $_.Extension -eq '.txt' } | Move-Item -Destination C:\Users\$user\Downloads\Text
}else{
    New-Item -Path C:\Users\$user\Downloads -Name "Text" -ItemType "directory"
    Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *.txt |
        Where-Object { $_.Extension -eq '.txt' } | Move-Item -Destination C:\Users\$user\Downloads\Text
}
##############################################
#Déplacer des fichiers iso
if (Test-Path -Path $isoFolder) {
    Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *.iso |
        Where-Object { $_.Extension -eq '.iso' } | Move-Item -Destination C:\Users\$user\Downloads\ISO
}else{
    New-Item -Path C:\Users\$user\Downloads -Name "ISO" -ItemType "directory"
    Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *.iso |
        Where-Object { $_.Extension -eq '.iso' } | Move-Item -Destination C:\Users\$user\Downloads\ISO
}
##############################################
#Déplacer des fichiers packet tracer
if (Test-Path -Path $pTracerFolder) {
    Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *.pka |
        Where-Object { $_.Extension -eq '.pka' } | Move-Item -Destination C:\Users\$user\Downloads\PacketTracer
}else{
    New-Item -Path C:\Users\$user\Downloads -Name "PacketTracer" -ItemType "directory"
    Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *.pka |
        Where-Object { $_.Extension -eq '.pka' } | Move-Item -Destination C:\Users\$user\Downloads\PacketTracer
}
##############################################
#Déplacer des fichiers torrent
if (Test-Path -Path $torrentFolder) {
    Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *.torrent |
        Where-Object { $_.Extension -eq '.torrent' } | Move-Item -Destination C:\Users\$user\Downloads\Torrent
}else{
    New-Item -Path C:\Users\$user\Downloads -Name "Torrent" -ItemType "directory"
    Get-ChildItem -Path C:\Users\$user\Downloads -ErrorAction SilentlyContinue -Filter *.torrent |
        Where-Object { $_.Extension -eq '.torrent' } | Move-Item -Destination C:\Users\$user\Downloads\Torrent
}
##############################################
#Déplacer des fichiers non reconnu
if (Test-Path -Path $otherFolder) {
    Get-ChildItem -Path C:\Users\$user\Downloads -Attribute !directory | Move-Item -Destination C:\Users\$user\Downloads\Other
}else{
    New-Item -Path C:\Users\$user\Downloads -Name "Other" -ItemType "directory"
    Get-ChildItem -Path C:\Users\$user\Downloads -Attribute !directory | Move-Item -Destination C:\Users\$user\Downloads\Other
}