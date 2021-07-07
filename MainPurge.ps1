$CSVFile = Import-Csv -Path "D:\Sources\ScriptPurge\listeserveurs.csv" -Delimiter ";"


foreach ($server in $CSVFile)
{
    
    $FolderPathLenght = $server.FolderPath | Measure-Object -Character

    #FolderPath
    if ($server.FolderPath -eq 'NO','*','*.*','\\','\\*','\*' )
    {
        $FolderPathClean = ''
        $FolderPathArg = ''
    }
    else
    {
        $FolderPathClean = $server.FolderPath
        $FolderPathArg = '-FolderPath'
    }




    #FileAge
    if ($server.FileAge -eq 'NO')
    {
        $FileAgeClean = ''
        $FileAgeArg = ''
    }
    else
    {
        $FileAgeClean = $server.FileAge
        $FileAgeArg = '-FileAge'
    }




    #LogFile
    if ($server.LogFile -eq 'NO')
    {
        $LogFileClean = ''
        $LogFileArg = ''
    }
    else 
    {
        $LogFileClean = $server.LogFile 
        $LogFileArg = '-LogFile'
    }




    #AutoLog
    if ($server.AutoLog -eq 'NO')
    {
        $AutoLogArg = ''
    }
    else {
        $AutoLogArg = '-AutoLog'
    }




    #ExclurePath
    if ($server.ExclurePath -eq 'NO')
    {
        $ExclurePathClean = ''
        $ExclurePathArg = ''
    }
    else {
        $ExclurePathClean = $server.ExclurePath
        $ExclurePathArg = '-ExclurePath'
    }




    #InclurePath
    if ($server.InclurePath -eq 'NO','*','*.*','\\','\\*','\*')
    {
        $InclurePathClean = ''
        $InclurePathArg = ''
    }
    else {
        $InclurePathClean = $server.InclurePath
        $InclurePathArg = '-InclurePath'
    }




    #RegExPath
    if ($server.RegExPath -eq 'NO')
    {
        $RegExPathArg = ''
    }
    else {
        $RegExPathArg = '-RegExPath'
    }





    #ExcludeFileExtension
    if ($server.ExcludeFileExtension -eq 'NO')
    {
        $ExcludeFileExtensionClean = ''
        $ExcludeFileExtensionArg = ''
    }
    else 
    {
        $ExcludeFileExtensionClean = $server.ExcludeFileExtension
        $ExcludeFileExtensionArg = '-ExcludeFileExtension'
    }





    #IncludeFileExtension
    if ($server.IncludeFileExtension -eq 'NO')
    {
        $IncludeFileExtensionClean = ''
        $IncludeFileExtensionArg = ''
    }
    else {
        $IncludeFileExtensionClean = $server.IncludeFileExtension
        $IncludeFileExtensionArg = '-IncludeFileExtension'
    }





    #IncludeFileExtension paramètre de crash du système
    if ($server.IncludeFileExtension -eq '.dll','.exe','.ecf','.bin','.xml','.cab','.sdb','.mui','.ini','.sdi','.com','.ttf','.wim','.ps1','.lex','.inf','.orp','.admx','.xsd')
    {
        $IncludeFileExtensionClean = 'crash:crash'
        $IncludeFileExtensionArg = 'crash:crash'
    }
    else {
        $IncludeFileExtensionClean = $server.IncludeFileExtension
        $IncludeFileExtensionArg = '-IncludeFileExtension'
    }

    



    #KeepFile
    if ($server.KeepFile -eq 'NO')
    {
        $KeepFileClean = ''
        $KeepFileArg = ''
    }
    else {
        $KeepFileClean = $server.KeepFile
        $KeepFileArg = '-KeepFile'
    }





    #EmailTo
    if ($server.EmailTo -eq 'NO')
    {
        $EmailToClean = ''
        $EmailToArg = ''
    }
    else {
        $EmailToClean = $server.EmailTo
        $EmailToArg = '-EmailTo'
    }





    #EmailFrom
    if ($server.EmailFrom -eq 'NO')
    {
        $EmailFromClean = ''
        $EmailFromArg = ''
    }
    else {
        $EmailFromClean = $server.EmailFrom
        $EmailFromArg = '-EmailFrom'
    }





    #EmailSmtpServer
    if ($server.EmailSmtpServer -eq 'NO')
    {
        $EmailSmtpServerClean = ''
        $EmailSmtpServerArg = ''
    }
    else {
        $EmailSmtpServerClean = $server.EmailSmtpServer
        $EmailSmtpServerArg = '-EmailSmtpServer'
    }




    #EmailSmtpPort
    if ($server.EmailSmtpPort -eq 'NO')
    {
        $EmailSmtpPortClean = ''
        $EmailSmtpPortArg = ''
    }
    else {
        $EmailSmtpPortClean = $server.EmailSmtpPort
        $EmailSmtpPortArg = '-EmailSmtpPort'
    }





    #EmailSubject avec les paramètres EmailTo, EmailFrom, EmailSmtpServer
    if ($server.EmailTo -eq 'NO' -or $server.EmailFrom -eq 'NO' -or $server.EmailSmtpServer -eq 'NO')
    {
        $EmailSubjectClean = ''
        $EmailSubjectArg = ''
    }
    else {
        $EmailSubjectClean = $server.EmailSubject
        $EmailSubjectArg = '-EmailSubject'
    }






    #EmailBody
    if ($server.EmailTo -eq 'NO'-or $server.EmailFrom -eq 'NO' -or $server.EmailSmtpServer -eq 'NO')
    {
        $EmailBodyClean = ''
        $EmailBodyArg = ''
    }
    else {
        $EmailBodyClean = $server.EmailBody
        $EmailBodyArg = '-EmailBody'
    }





    #ExcludeDate
    if ($server.ExcludeDate -eq 'NO')
    {
        $ExcludeDateClean = ''
        $ExcludeDateArg = ''
    }
    else {
        $ExcludeDateClean = $server.ExcludeDate
        $ExcludeDateArg = '-ExcludeDate'
    }





    #ListOnly
    if ($server.ListOnly -eq 'NO')
    {
        $ListOnlyArg = ''
    }
    else {
        $ListOnlyArg = '-ListOnly'
    }





    #VerboseLog
    if ($server.VerboseLog -eq 'NO')
    {
        $VerboseLogArg = ''
    }
    else {
        $VerboseLogArg = '-VerboseLog'
    }





    #AppendLog
    if ($server.AppendLog -eq 'NO')
    {
        $AppendLogArg = ''
    }
    else {
        AppendLogArg = '-AppendLog'
    }





    #LastAccessTime
    if ($server.LastAccessTime -eq 'NO')
    {
        $LastAccessTimeClean = ''
        $LastAccessTimeArg = ''
    }
    else {
        $LastAccessTimeClean = $server.LastAccessTime
        $LastAccessTimeArg = '-LastAccessTime'
    }





    #CleanFolders
    if ($server.CleanFolders -eq 'NO')
    {
        $CleanFoldersClean = ''
        $CleanFoldersArg = ''
    }
    else {
        $CleanFoldersClean = $server.CleanFolders
        $CleanFoldersArg = '-CleanFolders'
    }





    #NOFolder
    if ($server.NOFolder -eq 'NO')
    {
        $NOFolderClean = ''
        $NOFolderArg = ''
    }
    else {
        $NOFolderClean = $server.NOFolder
        $NOFolderArg = '-NOFolder'
    }





    #ArchivedOnly
    if ($server.ArchivedOnly -eq 'NO')
    {
        $ArchivedOnlyArg = ''
    }
    else {
        $ArchivedOnlyArg = '-ArchivedOnly'
    }





    
    powershell.exe -ExecutionPolicy ByPass -File D:\Sources\ScriptPurge\deleteold.ps1 $FolderPathArg $FolderPathClean $FileAgeArg $FileAgeClean $LogFileArg $LogFileClean $AutoLogArg $ExclurePathArg $ExclurePathClean $InclurePathArg $InclurePathClean $RegExPathArg $ExcludeFileExtensionArg $ExcludeFileExtensionClean $IncludeFileExtensionArg $IncludeFileExtensionClean $KeepFileArg $KeepFileClean $EmailToArg $EmailToClean $EmailFromArg $EmailFromClean $EmailSmtpServerArg $EmailSmtpServerClean $EmailSmtpPortArg $EmailSmtpPortClean $EmailSubjectArg $EmailSubjectClean $EmailBodyArg $EmailBodyClean $ExcludeDateArg $ExcludeDateClean $ListOnlyArg $VerboseLogArg $AppendLogArg $LastAccessTimeArg $LastAccessTimeClean $CleanFoldersArg $CleanFoldersClean $NOFolderArg $NOFolderClean $ArchivedOnlyArg
}