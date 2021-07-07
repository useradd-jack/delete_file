<#
Script pour supprimer les fichiers de plus de x jours. Le script est conçu pour être utilisé comme une tâche planifiée, il génère automatiquement un nom de fichier journal basé sur l'emplacement de la copie et la date/heure actuelle.
Il y a plusieurs niveaux de journalisation disponibles et le script peut aussi fonctionner en mode liste seulement dans 
lequel il ne liste que les fichiers qu'il supprimerait autrement. Il y a deux routines principales, l'une pour supprimer les fichiers et l'autre pour vérifier s'il reste des dossiers vides qui pourraient être supprimés.
   
.SYNOPSIS   
Script pour supprimer ou lister les anciens fichiers d'un dossier
    
.DESCRIPTION 
Script pour supprimer les fichiers de plus de x jours. Le script est conçu pour être utilisé comme une tâche planifiée, il génère automatiquement un nom de fichier journal basé sur l'emplacement de la copie et la date/heure actuelle. Il y a plusieurs niveaux de journalisation disponibles et le script peut aussi fonctionner en mode liste seulement dans lequel il ne liste que les fichiers qu'il supprimerait autrement. Il y a deux routines principales, l'une pour supprimer les fichiers et l'autre pour vérifier s'il reste des dossiers vides qui pourraient être supprimés.
	
.PARAMETER FolderPath 
Le chemin d'accès qui sera parcouru de façon récusive pour les anciens fichiers.

.PARAMETER Fileage
Filtre pour l'âge du fichier, entré en jours. Utilisez -1 pour tous les fichiers à supprimer.
	
.PARAMETER LogFile
Spécifie le chemin complet et le nom de fichier du fichier journal. Lorsque le paramètre LogFile est utilisé en combinaison avec -autolog, seul le chemin est requis.

.PARAMETER AutoLog
Génère automatiquement le nom de fichier au chemin spécifié dans -logfile. Si un nom de fichier est spécifié dans le paramètre LogFile et que le paramètre AutoLog est utilisé, seul le chemin spécifié dans LogFile est utilisé. Le nom de fichier est créé avec la convention d'appellation suivante :
"Autolog_<CheminDossier><dd-MM-MM-yyyyyy_HHHmm.ss>.log"

.PARAMETER ExcludePath
Spécifie un ou plusieurs chemins entre guillemets séparés par des virgules. Le paramètre Exclure n'accepte que les chemins complets, les chemins relatifs ne doivent pas être utilisés.

.PARAMETER IncludePath
Spécifie un ou plusieurs chemins entre guillemets séparés par des virgules. Le paramètre Exclure n'accepte que les chemins complets, les chemins relatifs ne doivent pas être utilisés. IncludePath est traité avant ExcludePath.

.PARAMETER RegExPath
Ce paramètre affecte à la fois les paramètres IncludePath et ExcludePath. Au lieu de comparer avec un nom de chemin d'accès, une expression régulière est utilisée. Pour plus d'informations sur les expressions régulières, consultez le fichier d'aide : Obtenir de l'aide sur les expressions courantes. L'expression régulière est seulement comparée au chemin d'un fichier, donc aucun nom de fichier ne peut être exclu en utilisant ExcludePath.

.PARAMETER ExcludeFileExtension
Spécifie une ou plusieurs extensions entre guillemets, séparées par des virgules. Les extensions seront exclues de la suppression. L'astérisque peut être utilisé comme joker.

.PARAMETER IncludeFileExtension
Spécifie une ou plusieurs extensions entre guillemets, séparées par des virgules. Les extensions seront incluses dans la suppression, toutes les autres extensions seront implicitement exclues. L'astérisque peut être utilisé comme joker.

.PARAMETER KeepFile
Spécifie le nombre de fichiers qui doivent être conservés dans chaque dossier. Ceci peut être utile si un dossier doit être nettoyé mais que x-nombre de fichiers doit toujours être conservé, le script regardera le paramètre spécifié, LastWriteTime,CreationTime

.PARAMETER EmailTo
Doit être utilisé conjointement avec les paramètres EmailFrom et EmailSmtpServer, ce paramètre peut prendre une adresse email ou un tableau d'adresses email à qui le fichier journal sera envoyé.

.PARAMETER EmailFrom
Doit être utilisé conjointement avec les paramètres EmailTo et EmailSmtpServer, ce paramètre peut prendre une adresse email qui est définie comme l'adresse email dans le champ from.

.PARAMETER EmailSmtpServer
A utiliser avec les paramètres EmailTo et EmailFrom, ce paramètre prend le nom de domaine complet de votre serveur smtp.

.PARAMETER EmailSmtpPort
Option paramètre email, permet de définir un port personnalisé, en omettant cette option, le port 25 sera utilisé par défaut.

.PARAMETER EmailSubject
Option paramètre email, permet de définir un sujet différent pour l'email contenant le fichier journal. Le formatage par défaut du sujet est'deleteold.ps1 a commencé à : DossierStartTime FolderPath : $FolderPath'.

.PARAMETER EmailBody
Option paramètre email, permet de définir un corps d'email personnalisé, par défaut un corps d'email vide.

.PARAMETER ExcludeDate
Si le paramètre ExcludeDate est spécifié, la requête est convertie par la fonction ConvertFrom-Query. La sortie de cette table est une table de hachage qui est splattée dans la fonction ConvertTo-DateObject qui retourne un tableau de dates. Tous les fichiers qui correspondent à une date dans le tableau retourné seront exclus de la suppression.


Quelques exemples:

Week:
'Week,sat,-1'
Liste tous les samedis jusqu'à ce que le maximum LimitYear soit atteint.
'Week,wed,5'
Liste des 5 derniers mercredis

Month:
'Month,first,4'
Énumérera le premier jour des quatre derniers mois.
'Month,last,-1'
Inscrira le dernier jour de jusqu'à ce que le maximum LimitYear soit atteint. Si la date du jour est le dernier jour du mois, le jour courant est également indiqué.
'Month,30,3'
Énumérera le 30 des trois derniers mois, si le mois de février est dans les résultats, il sera ignoré car il n'a pas 30 jours.
'Month,31,-1'
N'indiquera que le 31 du mois, tous les mois qui ont moins de 31 jours sont exclus. Énumérera les cas où le maximum de l'année-limite n'a pas été atteint.
'Month,15,4','Month,last,-1'
Indiquera le premier jour des quatre derniers mois et le dernier jour jusqu'à ce que le maximum LimitYear soit atteint. Si la date du jour est le dernier jour du mois, le jour courant est également indiqué.

Quarter:
'Quarter,first,-1'
Inscrira le premier jour d'un trimestre jusqu'à ce que le maximum LimitYear soit atteint.
'Quarter,last,6'
Je vais énumérer le dernier jour des six derniers trimestres. Si la date du jour est le dernier jour du trimestre, le jour courant est également indiqué.
'Quarter,91,5'
N'indiquera que le 91e jour de chaque trimestre, les trois derniers trimestres pour les années non bissextiles. Dans les années bissextiles, le premier trimestre a également 91 jours et sera donc inclus dans les résultats.
'Quarter,92,-1'
N'affichera que le 92e jour de chaque trimestre, donc seulement le 30 septembre et le 31 décembre. Les deux premiers trimestres de l'année ont moins de jours et ne seront pas inscrits. Fonctionne jusqu'à ce que le maximum de l'année limite soit atteint

Year:
'Year,last,4'
Liste au 31 décembre des 4 dernières années
'Year,first,-1'
Inscrira le 1er janvier jusqu'à ce que le maximum de l'année limite soit atteint.
'Year,15,-1'
Inscrira le 15 janvier jusqu'à ce que le maximum LimitYear ait été atteint.
'Year,366,5'
N'indiquera que le 366e jour, seulement le dernier jour des 5 dernières années bissextiles.

Specific Date:
'Date,2010-05-15'
Liste à paraître le 15 mai 2010
'Date,2012/12/12'
A paraître le 12 décembre 2012

Date Ranges:
'DateRange,2010-05-05,10'
Énumérera 10 dates, commençant le 5 mai 2010 et se poursuivant jusqu'au 14 mai 2010
'LimitYear,2008'
Place la limite de LimitYear à 2008, la valeur par défaut de ce paramètre est 2010

Toute combinaison ou requête est permise en séparant par exemple les requêtes par des virgules. Les éléments de requête Week/Month/Quarter/Year ne peuvent pas être utilisés deux fois pour combiner des requêtes. La valeur Date peut être utilisée plusieurs fois :
'Week,Fri,10','Year,last,-1','LimitYear,1950'
Énumérera les 10 derniers vendredis et le 31 décembre pour toutes les années jusqu'à ce que l'année limite soit atteinte.
'Week,Thu,4','Month,last,-1','Quarter,first,6','Year,last,10','LimitYear,2012','Date,1990-12-31','Date,1995-5-31'
Inscrira les quatre derniers jeudis, le dernier jour du mois jusqu'à l'atteinte du maximum LimitYear, le premier jour des 6 premiers trimestres et le 31 décembre pour les 10 dernières années et les deux dates spécifiques 1990-12-31 & 1995-5-31.

.PARAMETER ListOnly
Uniquement des listes, ne supprime pas ou ne modifie pas les fichiers. Ce paramètre peut être utilisé pour établir quels fichiers seront supprimés si le script est exécuté.

.PARAMETER VerboseLog
Enregistre toutes les opérations de suppression à enregistrer, le comportement par défaut du script est d'enregistrer uniquement les échecs.

.PARAMETER AppendLog
Ajoute au fichier journal existant, le comportement par défaut du script est de remplacer les fichiers journaux existants si le fichier journal existe déjà. Si le fichier journal n'existe pas, le fichier journal sera créé comme d'habitude.

.PARAMETER CreateTime
Supprime les fichiers basés sur CreationTime, le comportement par défaut du script est de supprimer basé sur LastWriteTime.

.PARAMETER CompareCreateTimeLastModified
Supprime les fichiers basés sur CreationTime ou LastWriteTime, selon celui qui a été modifié en dernier. L'entrée la plus récente sera utilisée dans la comparaison, ce qui est particulièrement utile pour les fichiers dont la date CreateTime est postérieure à la date LastModified. Cela se produit lorsqu'un fichier est copié dans un nouvel emplacement.

.PARAMETER LastAccessTime
Supprime les fichiers basés sur LastAccessTime, le comportement par défaut du script est de supprimer basé sur LastWriteTime.

.PARAMETER CleanFolders
Si ce commutateur est spécifié, tout dossier vide sera supprimé. Le comportement par défaut de ce script est de ne supprimer que les dossiers contenant d'anciens fichiers.

.PARAMETER NoFolder
Si ce commutateur est spécifié, seuls les fichiers seront supprimés et le dossier existant sera conservé.

.PARAMETER ArchivedOnly
Si ce commutateur est spécifié, seuls les fichiers dont le bit d'archive est effacé (c'est-à-dire sauvegardé) seront purgés.

.NOTES   
Name: deleteold.ps1
Author: Tony Stark
Version: JARVIS.deleteoldlog
DateCreated: 2019/10/23

.EXAMPLE   
.\deleteold.ps1 -FolderPath H:\scripts -FileAge 100 -ListOnly -LogFile H:\log.log

Description:
Recherche dans le dossier H:\scripts et écrit un fichier journal contenant les derniers fichiers modifiés il y a 100 jours et plus.

.EXAMPLE
.\deleteold.ps1 -FolderPath H:\scripts -FileAge 30 -LogFile H:\log.log -VerboseLog

Description:
Recherche dans le dossier H:\scripts et supprime les fichiers qui ont été modifiés il y a 30 jours ou avant, écrit toutes les opérations, réussies ou non, dans un fichier journal sur le lecteur :H.

.EXAMPLE
.\deleteold.ps1 -FolderPath C:\docs -FileAge 30 -LogFile H:\log.log -ExcludePath "C:\docs\finance\","C:\docs\hr\"

Description:
Recherche dans le dossier C:\docs et supprime les fichiers, en excluant les dossiers finance et hr dans C:\docs.

.EXAMPLE
.\deleteold.ps1 -FolderPath C:\Folder -FileAge 30 -LogFile H:\log.log -IncludePath "C:\Folder\Docs\","C:\Folder\Users\" -ExcludePath "C:\docs\finance\","C:\docs\hr\"

Description:
Ne vérifiez que les fichiers dans les dossiers C:\Dossier\Docs\ et C:\Dossier\Utilisateurs\ Dossiers et non les autres dossiers dans C:\Dossiers et excluez explicitement les dossiers Finance et HR dans C:\Dossier\Docs.

.EXAMPLE
.\deleteold.ps1 -FolderPath C:\Folder -FileAge 30 -LogFile H:\log.log -IncludePath "C:\Folder\Docs\","C:\Folder\Users\" -ExcludePath "C:\docs\finance\","C:\docs\hr\" -ExcludeDate 'Week,Fri,10','Year,last,-1','LimitYear,1950'

Description:
Ne vérifiez que les fichiers dans les dossiers C:\Folder\Docs\ et C:\Folder\Users\Folder et non les autres dossiers dans C:\Folderet excluez explicitement les dossiers Finance et HR dans C:\Folder\Docs. Exclut également les dossiers basés sur la date, sauf les 10 derniers vendredis et le 31 décembre pour toutes les années antérieures à 1950.

.EXAMPLE
.\deleteold.ps1 -FolderPath C:\Folder -FileAge 60 -LogFile H:\log.log -IncludePath .*images.* -RegExPath

Description:
Supprimez les fichiers de plus de 60 jours et ne supprimez que les fichiers qui contiennent des " images " dans leur chemin d'accès.

.EXAMPLE
.\deleteold.ps1 -FolderPath C:\Folder -FileAge 45 -LastAccessTime -LogFile H:\log.log

Description:
Supprimer les fichiers qui n'ont pas été consultés depuis plus de 45 jours.

.EXAMPLE
PowerShell.exe deleteold.ps1 -FolderPath 'H:\admin_jaap' -FileAge 10 -LogFile C:\log -AutoLog

Description:
Lance le script à partir d'un fichier batch ou d'une invite de commande, un nom de fichier est automatiquement généré puisque le paramètre -AutoLog est utilisé. Notez les guillemets '' qui sont utilisés pour le paramètre FolderPath.

.EXAMPLE
.\deleteold.ps1 -FolderPath H:\SQL\BackUp -FileAge 10 -LogFile C:\log -AutoLog -KeepFile 5

Description:
Le nom du fichier journal est généré automatiquement depuis que le paramètre -AutoLog est utilisé. Tous les fichiers de plus de dix fichiers seront supprimés sauf s'il reste moins de 5 fichiers, les 5 derniers fichiers de chaque dossier seront supprimés.

.EXAMPLE
.\deleteold.ps1 -FolderPath C:\docs -FileAge 30 -logfile h:\log.log -CreateTime -NoFolder

Description:
Supprime tous les fichiers qui ont été créés il y a 30 jours ou avant dans le dossier C:\docs. Aucun dossier n'est supprimé.

.EXAMPLE
.\deleteold.ps1 -FolderPath C:\docs -FileAge 30 -logfile h:\log.log -CreateTime -CleanFolders

Description:
Supprime tous les fichiers qui ont été créés il y a 30 jours ou avant dans le dossier C:\docs. Seuls les dossiers qui contenaient d'anciens fichiers et qui sont vides après la suppression de ces fichiers seront supprimés.

.EXAMPLE
.\deleteold.ps1 -FolderPath C:\docs -FileAge 30 -logfile h:\log.log -CompareCreateTimeLastModified

Description:
Supprime tous les fichiers qui n'ont pas été créés ou modifiés il y a 30 jours ou avant dans le dossier C:\docs.

.EXAMPLE
.\deleteold.ps1 -folderpath c:\users\jaapbrasser\desktop -fileage 10 -log c:\log.txt -autolog -verboselog -IncludeFileExtension '.xls*','.doc*'

Description:
Supprime les fichiers de plus de 10 jours, ne supprime que les fichiers correspondant aux modèles .xls* et .doc*, par exemple les fichiers .doc et .docx. Le fichier journal est stocké à la racine du disque C avec un nom généré automatiquement.

.EXAMPLE
.\deleteold.ps1 -folderpath c:\users\jaapbrasser\desktop -fileage 10 -log c:\log.txt -autolog -verboselog -ExcludeFileExtension .xls

Description:
Supprime les fichiers de plus de 10 jours, à l'exception des fichiers xls. Le fichier journal est stocké à la racine du disque C avec un nom généré automatiquement.

.EXAMPLE
.\deleteold.ps1 -FolderPath C:\docs -FileAge 30 -LogFile h:\log.log -ExcludeDate 'Week,Thu,4','Month,last,-1','Quarter,first,6','Year,last,10','LimitYear,2012','Date,1990-12-31','Date,1995-5-31'

Description:
Supprime tous les fichiers qui ont été créés il y a 30 jours ou avant dans le dossier C:\docs. A l'exclusion des derniers fichiers modifiés/créés spécifiés dans la requête -ExcludeDate.

.EXAMPLE   
.\deleteold.ps1 -FolderPath H:\scripts -FileAge 100 -ListOnly -LogFile H:\log.log -ExcludeDate 'DateRange,2005-05-16,8'

Description:
Recherche dans le dossier H:\scripts et écrit un fichier journal contenant les derniers fichiers modifiés il y a 100 jours et plus. Excluant les fichiers modifiés le 16 mai 2005 et les sept jours suivants.

.EXAMPLE
.\deleteold.ps1 -FolderPath C:\docs -ListOnly -FileAge 30 -LogFile h:\log.log -ExcludeDate 'Month,15,5','Month,16,5' -EmailTo jaapbrasser@corp.co -EmailFrom tonystark@grp.com -EmailSmtpServer smtp.grp.com

Description:
Supprime tous les fichiers qui ont été créés il y a 30 jours ou avant dans le dossier C:\docs. À l'exclusion des fichiers modifiés/créés pour la dernière fois le 15 ou le 16 des cinq derniers mois. Une fois le script terminé, le fichier journal sera envoyé par courriel à tonystark@grp.com via le serveur smtp.grp.com smtp.
#>

#region Parameters
param(
    [string]   $FolderPath,
	[decimal]  $FileAge,
	[string]   $LogFile,
    [string[]] $ExcludePath,
    [string[]] $IncludePath,
	[string[]] $ExcludeFileExtension,
    [string[]] $IncludeFileExtension,
    [string[]] $ExcludeDate,
    [int]      $KeepFile,
    [string[]] $EmailTo,
    [string]   $EmailFrom,
    [string]   $EmailSmtpServer,
    [int]      $EmailSmtpPort,
    [string]   $EmailSubject,
    [string]   $EmailBody,
    [switch]   $ListOnly,
	[switch]   $VerboseLog,
	[switch]   $AutoLog,
    [switch]   $AppendLog,
	[switch]   $CreateTime,
    [switch]   $CompareCreateTimeLastModified,
    [switch]   $LastAccessTime,
    [switch]   $CleanFolders,
    [switch]   $NoFolder,
    [switch]   $ArchivedOnly,
    [switch]   $RegExPath
)
#endregion

#region Functions
# Fonction pour convertir la requête fournie dans -ExcludeDate dans un format qui peut être analysé par la fonction ConvertTo-DateObject
function ConvertFrom-DateQuery {
param (
    $Query
)
    try {
        $CsvQuery = Convertfrom-Csv -InputObject $Query -Delimiter ',' -Header 'Type','Day','Repeat'
        $ConvertCsvSuccess = $true
    } catch {
        Write-Warning 'Query is in incorrect format, please supply query in proper format'
        $ConvertCsvSuccess = $false
    }
    if ($ConvertCsvSuccess) {
        $Check=$HashOutput = @{}
        foreach ($Entry in $CsvQuery) {
            switch ($Entry.Type) {
                'week' {
                    # Convertir les dates nommées en format correct
                    switch ($Entry.Day)
                    {
                        # DayOfWeek commence à compter à 0, en se référant à la propriété [datetime] DayOfWeek
                        'sun' {
                            $HashOutput.DayOfWeek  = 0
                            $HashOutput.WeekRepeat = $Entry.Repeat -as [int]
                        }
                        'mon' {
                            $HashOutput.DayOfWeek  = 1
                            $HashOutput.WeekRepeat = $Entry.Repeat -as [int]
                        }
                        'tue' {
                            $HashOutput.DayOfWeek  = 2
                            $HashOutput.WeekRepeat = $Entry.Repeat -as [int]
                        }
                        'wed' {
                            $HashOutput.DayOfWeek  = 3
                            $HashOutput.WeekRepeat = $Entry.Repeat -as [int]
                        }
                        'thu' {
                            $HashOutput.DayOfWeek  = 4
                            $HashOutput.WeekRepeat = $Entry.Repeat -as [int]
                        }
                        'fri' {
                            $HashOutput.DayOfWeek  = 5
                            $HashOutput.WeekRepeat = $Entry.Repeat -as [int]
                        }
                        'sat' {
                            $HashOutput.DayOfWeek  = 6
                            $HashOutput.WeekRepeat = $Entry.Repeat -as [int]
                        }
                        Default {$Check.WeekSuccess = $false}
                    }
                }
                'month' {
                    # Convertir les dates nommées en format correct
                    switch ($Entry.Day)
                    {
                        # DayOfMonth commence le décompte à 0, se référant au dernier jour du mois avec zéro.
                        'first' {
                            [array]$HashOutput.DayOfMonth  += 1
                            [array]$HashOutput.MonthRepeat += $Entry.Repeat -as [int]
                        }
                        'last' {
                            [array]$HashOutput.DayOfMonth  += 0
                            [array]$HashOutput.MonthRepeat += $Entry.Repeat -as [int]
                        }
                        {(1..31) -contains $_} {
                            [array]$HashOutput.DayOfMonth  += $Entry.Day
                            [array]$HashOutput.MonthRepeat += $Entry.Repeat -as [int]
                        }
                        Default {$Check.MonthSuccess = $false}
                    }
                }
                'quarter' {
                    # Count the number of times the quarter argument is used, used in final check of values
                    $QuarterCount++

                    # Convert named dates to correct format
                    switch ($Entry.Day)
                    {
                        # DayOfMonth starts count at 0, referring to the last day of the month with zero
                        'first' {
                            $HashOutput.DayOfQuarter   = 1
                            $HashOutput.QuarterRepeat  = $Entry.Repeat
                        }
                        'last' {
                            $HashOutput.DayOfQuarter   = 0
                            $HashOutput.QuarterRepeat  = $Entry.Repeat
                        }
                        {(1..92) -contains $_} {
                            $HashOutput.DayOfQuarter   = $Entry.Day
                            $HashOutput.QuarterRepeat  = $Entry.Repeat
                        }
                        Default {$Check.QuarterSuccess = $false}
                    }
                }
                'year' {
                    # Convert named dates to correct format
                    switch ($Entry.Day)
                    {
                        # DayOfMonth starts count at 0, referring to the last day of the month with zero
                        'first' {
                            $HashOutput.DayOfYear = 1
                            $HashOutput.DayOfYearRepeat = $Entry.Repeat
                        }
                        'last' {
                            $HashOutput.DayOfYear = 0
                            $HashOutput.DayOfYearRepeat = $Entry.Repeat
                        }
                        {(1..366) -contains $_} {
                            $HashOutput.DayOfYear       = $Entry.Day
                            $HashOutput.DayOfYearRepeat = $Entry.Repeat
                        }
                        Default {$Check.YearSuccess = $false}
                    }
                }
                'date' {
                    # Verify if the date is in the correct format
                    switch ($Entry.Day)
                    {
                        {try {[DateTime]"$($Entry.Day)"} catch{}} {
                            [array]$HashOutput.DateDay += $Entry.Day
                        }
                        Default {$Check.DateSuccess = $false}
                    }
                }

                'daterange' {
                    # Verify if the date is in the correct format
                    switch ($Entry.Day)
                    {
                        {try {[DateTime]"$($Entry.Day)"} catch{}} {
                            $HashOutput.DateRange       += $Entry.Day
                            $HashOutput.DateRangeRepeat += $Entry.Repeat
                        }
                        Default {$Check.DateRangeSuccess = $false}
                    }
                }

                'limityear' {
                    switch ($Entry.Day)
                    {
                        {(1000..2100) -contains $_} {
                            $HashOutput.LimitYear        = $Entry.Day
                        }
                        Default {$Check.LimitYearSuccess = $false}
                    }
                }
                Default {
                    $QueryContentCorrect = $false
                }
            }
        }
        $HashOutput
    }
}

# Function that outputs an array of date objects that can be used to exclude certain files from deletion
function ConvertTo-DateObject {
param(
    [validaterange(0,6)]
    $DayOfWeek,
    [int]$WeekRepeat=1,
    [validaterange(0,31)]
    $DayOfMonth,
    $MonthRepeat=1,
    [validaterange(0,92)]
    $DayOfQuarter,
    [int]$QuarterRepeat=1,
    [validaterange(0,366)]
    $DayOfYear,
    [int]$DayOfYearRepeat=1,
    $DateDay,
    $DateRange,
    [int]$DateRangeRepeat=1,
    [validaterange(1000,2100)]
    [int]$LimitYear = 2010
)
    # Define variable
    $CurrentDate = Get-Date

    if ($DayOfWeek -ne $null) {
        $CurrentWeekDayInt = $CurrentDate.DayOfWeek.value__

            # Loop runs for number of times specified in the WeekRepeat parameter
            for ($j = 0; $j -lt $WeekRepeat; $j++)
                { 
                    $CheckDate = $CurrentDate.Date.AddDays(-((7*$j)+$CurrentWeekDayInt-$DayOfWeek))

                    # Only display date if date is larger than current date, this is to exclude dates in the current week
                    if ($CheckDate -le $CurrentDate) {
                        $CheckDate
                    } else {
                        # Increase weekrepeat, to ensure the correct amount of repeats are executed when date returned is
                        # higher than current date
                        $WeekRepeat++
                    }
                }
            
            # Loop runs until $LimitYear parameter is exceeded
			if ($WeekRepeat -eq -1) {
                $j=0
                do {
                    $CheckDate = $CurrentDate.AddDays(-((7*$j)+$CurrentWeekDayInt-$DayOfWeek))
                    $j++

                    # Only display date if date is larger than current date, this is to exclude dates in the current week
                    if ($CheckDate -le $CurrentDate) {
                        $CheckDate
                    }
                } while ($LimitYear -le $CheckDate.Adddays(-7).Year)
            }
        }

    if (-not [string]::IsNullOrEmpty($DayOfMonth)) {
        for ($MonthCnt = 0; $MonthCnt -lt $DayOfMonth.Count; $MonthCnt++) {
            # Loop runs for number of times specified in the MonthRepeat parameter
            for ($j = 0; $j -lt $MonthRepeat[$MonthCnt]; $j++)
                { 
                    $CheckDate = $CurrentDate.Date.AddMonths(-$j).AddDays($DayOfMonth[$MonthCnt]-$CurrentDate.Day)

                    # Only display date if date is larger than current date, this is to exclude dates ahead of the current date and
                    # to list only output the possible dates. If a value of 29 or higher is specified as a DayOfMonth value
                    # only possible dates are listed.
                    if ($CheckDate -le $CurrentDate -and $(if ($DayOfMonth[$MonthCnt] -ne 0) {$CheckDate.Day -eq $DayOfMonth[$MonthCnt]} else {$true})) {
                        $CheckDate
                    } else {
                        # Increase MonthRepeat integer, to ensure the correct amount of repeats are executed when date returned is
                        # higher than current date
                        $MonthRepeat[$MonthCnt]++
                    }
                }
            
            # Loop runs until $LimitYear parameter is exceeded
		    if ($MonthRepeat[$MonthCnt] -eq -1) {
                $j=0
                do {
                    $CheckDate = $CurrentDate.Date.AddMonths(-$j).AddDays($DayOfMonth[$MonthCnt]-$CurrentDate.Day)
                    $j++

                    # Only display date if date is larger than current date, this is to exclude dates ahead of the current date and
                    # to list only output the possible dates. For example if a value of 29 or higher is specified as a DayOfMonth value
                    # only possible dates are listed.
                    if ($CheckDate -le $CurrentDate -and $(if ($DayOfMonth[$MonthCnt] -ne 0) {$CheckDate.Day -eq $DayOfMonth[$MonthCnt]} else {$true})) {
                        $CheckDate
                    }
                } while ($LimitYear -le $CheckDate.Adddays(-31).Year)
            }
        }
    }

    if ($DayOfQuarter -ne $null) {
        # Set quarter int to current quarter value $QuarterInt
        $QuarterInt = [int](($CurrentDate.Month+1)/3)
        $QuarterYearInt = $CurrentDate.Year
        $QuarterLoopCount = $QuarterRepeat
        $j = 0
        
        do {
            switch ($QuarterInt) {
                1 {
                    $CheckDate = ([DateTime]::ParseExact("$($QuarterYearInt)0101",'yyyyMMdd',$null)).AddDays($DayOfQuarter-1)
                    
                    # Check for number of days in the 1st quarter, this depends on leap years
                    $DaysInFeb = ([DateTime]::ParseExact("$($QuarterYearInt)0301",'yyyyMMdd',$null)).AddDays(-1).Day
                    $DaysInCurrentQuarter = 31+$DaysInFeb+31
                        
                    # If the number of days is larger that the total number of days in this quarter the quarter will be excluded
                    if ($DayOfQuarter -gt $DaysInCurrentQuarter) {
                        $CheckDate = $null
                    }

                    # This check is built-in to return the date last date of the current quarter, to ensure consistent results
                    # in case the command is executed on the last day of a quarter
                    if ($DayOfQuarter -eq 0) {
                        $CheckDate = [DateTime]::ParseExact("$($QuarterYearInt)0331",'yyyyMMdd',$null)
                    }

                    $QuarterInt = 4
                    $QuarterYearInt--
                }
                2 {
                    $CheckDate = ([DateTime]::ParseExact("$($QuarterYearInt)0401",'yyyyMMdd',$null)).AddDays($DayOfQuarter-1)
                        
                    # Check for number of days in the 2nd quarter
                    $DaysInCurrentQuarter = 30+31+30
                        
                    # If the number of days is larger that the total number of days in this quarter the quarter will be excluded
                    if ($DayOfQuarter -gt $DaysInCurrentQuarter) {
                        $CheckDate = $null
                    }

                    # This check is built-in to return the date last date of the current quarter, to ensure consistent results
                    # in case the command is executed on the last day of a quarter                       
                    if ($DayOfQuarter -eq 0) {
                        $CheckDate = [DateTime]::ParseExact("$($QuarterYearInt)0630",'yyyyMMdd',$null)
                    }
                        
                    $QuarterInt = 1
                }
                3 {
                    $CheckDate = ([DateTime]::ParseExact("$($QuarterYearInt)0701",'yyyyMMdd',$null)).AddDays($DayOfQuarter-1)
                        
                    # Check for number of days in the 3rd quarter
                    $DaysInCurrentQuarter = 31+31+30
                        
                    # If the number of days is larger that the total number of days in this quarter the quarter will be excluded
                    if ($DayOfQuarter -gt $DaysInCurrentQuarter) {
                        $CheckDate = $null
                    }
                        
                    # This check is built-in to return the date last date of the current quarter, to ensure consistent results
                    # in case the command is executed on the last day of a quarter                       
                    if ($DayOfQuarter -eq 0) {
                        $CheckDate = [DateTime]::ParseExact("$($QuarterYearInt)0930",'yyyyMMdd',$null)
                    }

                    $QuarterInt = 2
                }
                4 {
                    $CheckDate = ([DateTime]::ParseExact("$($QuarterYearInt)1001",'yyyyMMdd',$null)).AddDays($DayOfQuarter-1)
                        
                    # Check for number of days in the 4th quarter
                    $DaysInCurrentQuarter = 31+30+31
                        
                    # If the number of days is larger that the total number of days in this quarter the quarter will be excluded
                    if ($DayOfQuarter -gt $DaysInCurrentQuarter) {
                        $CheckDate = $null
                    }

                    # This check is built-in to return the date last date of the current quarter, to ensure consistent results
                    # in case the command is executed on the last day of a quarter                       
                    if ($DayOfQuarter -eq 0) {
                        $CheckDate = [DateTime]::ParseExact("$($QuarterYearInt)1231",'yyyyMMdd',$null)
                    }                        
                    $QuarterInt = 3
                }
            }

            # Only display date if date is larger than current date, and only execute check if $CheckDate is not equal to $null
            if ($CheckDate -le $CurrentDate -and $CheckDate -ne $null) {
                    
                # Only display the date if it is not further in the past than the limit year
                if ($CheckDate.Year -ge $LimitYear -and $QuarterRepeat -eq -1) {
                    $CheckDate
                }

                # If the repeat parameter is not set to -1 display results regardless of limit year                    
                if ($QuarterRepeat -ne -1) {
                    $CheckDate
                    $j++
                } else {
                    $QuarterLoopCount++
                }
            }
            # Added if statement to catch errors regarding 
        } while ($(if ($QuarterRepeat -eq -1) {$LimitYear -le $(if ($CheckDate) {$CheckDate.Year} else {9999})} 
                else {$j -lt $QuarterLoopCount}))
    }

    if ($DayOfYear -ne $null) {
        $YearLoopCount = $DayOfYearRepeat
        $YearInt = $CurrentDate.Year
        $j = 0

        # Mainloop containing the loop for selecting a day of a year
        do {
            $CheckDate = ([DateTime]::ParseExact("$($YearInt)0101",'yyyyMMdd',$null)).AddDays($DayOfYear-1)
            
            # If the last day of the year is specified, a year is added to get consistent results when the query is executed on last day of the year 
            if ($DayOfYear -eq 0) {
                $CheckDate = $CheckDate.AddYears(1)
            }
            
            # Set checkdate to null to allow for selection of last day of leap year
            if (($DayOfYear -eq 366) -and !([DateTime]::IsLeapYear($YearInt))) {
                $CheckDate = $null
            }

            # Only display date if date is larger than current date, and only execute check if $CheckDate is not equal to $null
            if ($CheckDate -le $CurrentDate -and $CheckDate -ne $null) {
                # Only display the date if it is not further in the past than the limit year
                if ($CheckDate.Year -ge $LimitYear -and $DayOfYearRepeat -eq -1) {
                    $CheckDate
                }

                # If the repeat parameter is not set to -1 display results regardless of limit year
                if ($DayOfYearRepeat -ne -1) {
                    $CheckDate
                    $j++
                } else {
                    $YearLoopCount++
                }
            }
            $YearInt--
        } while ($(if ($DayOfYearRepeat -eq -1) {$LimitYear -le $(if ($CheckDate) {$CheckDate.Year} else {9999})} 
                else {$j -lt $YearLoopCount}))
    }

    if ($DateDay -ne $null) {
        foreach ($Date in $DateDay) {
            try {
                $CheckDate     = [DateTime]::ParseExact($Date,'yyyy-MM-dd',$null)
            } catch {
                try {
                    $CheckDate = [DateTime]::ParseExact($Date,'yyyy\/MM\/dd',$null)
                } catch {}
            }
            
            if ($CheckDate -le $CurrentDate) {
                $CheckDate
            }
            $CheckDate=$null
        }
    }

    if ($DateRange -ne $null) {
        $CheckDate=$null
        try {
            $CheckDate     = [DateTime]::ParseExact($DateRange,'yyyy-MM-dd',$null)
        } catch {
            try {
                $CheckDate = [DateTime]::ParseExact($DateRange,'yyyy\/MM\/dd',$null)
            } catch {}
        }
        if ($CheckDate) {
            for ($k = 0; $k -lt $DateRangeRepeat; $k++) { 
                if ($CheckDate -le $CurrentDate) {
                    $CheckDate
                }
                $CheckDate = $CheckDate.AddDays(1)
            }
        }
    }
}
#endregion

# Check if correct parameters are used
if (-not $FolderPath) {Write-Warning 'Please specify the -FolderPath variable, this parameter is required. Use Get-Help .\deleteold.ps1 to display help.';exit}
if (-not $FileAge) {Write-Warning 'Please specify the -FileAge variable, this parameter is required. Use Get-Help .\deleteold.ps1 to display help.';exit}
if (-not $LogFile) {Write-Warning 'Please specify the -LogFile variable, this parameter is required. Use Get-Help .\deleteold.ps1 to display help.';exit}
if ($Autolog) {
    # Section that is triggered when the -autolog switch is active
	# Gets date and reformats to be used in log filename
	$TempDate = (get-date).ToString('dd-MM-yyyy_HHmm.ss')
	# Reformats $FolderPath so it can be used in the log filename
	$TempFolderPath = $FolderPath -replace '\\','_'
	$TempFolderPath = $TempFolderPath -replace ':',''
	$TempFolderPath = $TempFolderPath -replace ' ',''
	# Checks if the logfile is either pointing at a folder or a logfile and removes
	# Any trailing backslashes
	$TestLogPath = Test-Path $LogFile -PathType Container
	if (-not $TestLogPath) {
        $LogFile = Split-Path $LogFile -Erroraction SilentlyContinue
    }
	if ($LogFile.SubString($LogFile.Length-1,1) -eq '\') {
        $LogFile = $LogFile.SubString(0,$LogFile.Length-1)
    }
	# Combines the date and the path scanned into the log filename
	$LogFile = "$LogFile\Autolog_$TempFolderPath$TempDate.log"
}

#region Variables
# Sets up the variables
$Startdate = Get-Date
$LastWrite = $Startdate.AddDays(-$FileAge)
$StartTime = $Startdate.ToShortDateString()+', '+$Startdate.ToLongTimeString()
$Switches = "`r`n`t`t-FolderPath`r`n`t`t`t$FolderPath`r`n`t`t-FileAge $FileAge`r`n`t`t-LogFile`r`n`t`t`t$LogFile"
    # Populate the switches string with the switches and parameters that are set
    if ($IncludePath) {
	    $Switches += "`r`n`t`t-IncludePath"
	    for ($j=0;$j -lt $IncludePath.Count;$j++) {$Switches+= "`r`n`t`t`t";$Switches+= $IncludePath[$j]}
    }
    if ($ExcludePath) {
	    $Switches += "`r`n`t`t-ExcludePath"
	    for ($j=0;$j -lt $ExcludePath.Count;$j++) {$Switches+= "`r`n`t`t`t";$Switches+= $ExcludePath[$j]}
    }
    if ($IncludeFileExtension) {
	    $Switches += "`r`n`t`t-IncludeFileExtension"
	    for ($j=0;$j -lt $IncludeFileExtension.Count;$j++) {$Switches+= "`r`n`t`t`t";$Switches+= $IncludeFileExtension[$j]}
    }
    if ($ExcludeFileExtension) {
	    $Switches += "`r`n`t`t-ExcludeFileExtension"
	    for ($j=0;$j -lt $ExcludeFileExtension.Count;$j++) {$Switches+= "`r`n`t`t`t";$Switches+= $ExcludeFileExtension[$j]}
    }
    if ($KeepFile) {
	    $Switches += "`r`n`t`t-KeepFile $KeepFile"
    }
    if ($ExcludeDate) {
	    $Switches+= "`r`n`t`t-ExcludeDate"
        $ExcludeDate | ConvertFrom-Csv -Header:'Item1','Item2','Item3' -ErrorAction SilentlyContinue | ForEach-Object {
            $Switches += "`r`n`t`t`t"
            $Switches += ($_.Item1,$_.Item2,$_.Item3 -join ',').Trim(',')
        }	    
    }
    if ($EmailTo) {
	    $Switches += "`r`n`t`t-EmailTo"
	    for ($j=0;$j -lt $EmailTo.Count;$j++) {$Switches+= "`r`n`t`t`t";$Switches+= $EmailTo[$j]}
    }
    if ($EmailFrom) {
        $Switches += "`r`n`t`t-EmailFrom`r`n`t`t`t$EmailFrom"
    }
    if ($EmailSubject) {
        $Switches += "`r`n`t`t-EmailSubject`r`n`t`t`t$EmailSubject"
    }
    if ($EmailSmtpServer) {
        $Switches += "`r`n`t`t-EmailSmtpServer`r`n`t`t`t$EmailSmtpServer"
    }
    if ($EmailSmtpPort) {
        $Switches += "`r`n`t`t-EmailSmtpPort`r`n`t`t`t$EmailSmtpPort"
    }
    if ($ListOnly)       {$Switches+="`r`n`t`t-ListOnly"}
    if ($VerboseLog)     {$Switches+="`r`n`t`t-VerboseLog"}
    if ($Autolog)        {$Switches+="`r`n`t`t-AutoLog"}
    if ($Appendlog)      {$Switches+="`r`n`t`t-AppendLog"}
    if ($CreateTime)     {$Switches+="`r`n`t`t-CreateTime"}
    if ($LastAccessTime) {$Switches+="`r`n`t`t-LastAccessTime"}
    if ($CleanFolders)   {$Switches+="`r`n`t`t-CleanFolders"}
    if ($EmailBody)      {$Switches+="`r`n`t`t-EmailBody"}
    if ($NoFolder)       {$Switches+="`r`n`t`t-NoFolder"}
    if ($ArchivedOnly)   {$Switches+="`r`n`t`t-ArchivedOnly"}
    if ($RegExPath)      {$Switches+="`r`n`t`t-RegExPath"}
    if ($CompareCreateTimeLastModified) {$Switches+="`r`n`t`t-CompareCreateTimeLastModified"}
    
[long]$FilesSize    = 0
[long]$FailedSize   = 0
[int]$FilesNumber   = 0
[int]$FilesFailed   = 0
[int]$FoldersNumber = 0
[int]$FoldersFailed = 0

# Sets up the email splat, displays a warning if not all variables have been correctly entered
if ($EmailTo -or $EmailFrom -or $EmailSmtpServer) {
    if (($EmailTo,$EmailFrom,$EmailSmtpServer) -contains '') {
        Write-Warning 'EmailTo EmailFrom and EmailSmtpServer parameters only work if all three parameters are used, no email sent...'
    } else {
        $EmailSplat = @{
            To          = $EmailTo
            From        = $EmailFrom
            SmtpServer  = $EmailSmtpServer
            Attachments = $LogFile
        }
        if ($EmailSubject) {
            $EmailSplat.Subject = $EmailSubject
        } else {
            $EmailSplat.Subject = "deleteold.ps1 started at: $StartTime FolderPath: $FolderPath"
        }
        if ($EmailBody) {
            $EmailSplat.Body    = $EmailBody
        }
        if ($EmailSmtpPort) {
            $EmailSplat.Port    = $EmailSmtpPort
        }
    }
}
#endregion

# Output text to console and write log header
Write-Output ('-'*79)
Write-Output "  Deleteold`t::`tScript to delete old files from folders"
Write-Output ('-'*79)
Write-Output "`n   Started  :   $StartTime`n   Folder   :`t$FolderPath`n   Switches :`t$Switches`n"
if ($ListOnly) {
    Write-Output "`t*** Running in Listonly mode, no files will be modified ***`n"
}
Write-Output ('-'*79)

# If AppendLog switch is present log will be appended, not replaced
if ($AppendLog) {
    ('-'*79) | Add-Content -LiteralPath $LogFile
} else {
    ('-'*79) | Set-Content -LiteralPath $LogFile
}

"  Deleteold`t::`tScript to delete old files from folders" | Add-Content -LiteralPath $LogFile
('-'*79) | Add-Content -LiteralPath $LogFile
' ' | Add-Content -LiteralPath $LogFile
"   Started  :   $StartTime" | Add-Content -LiteralPath $LogFile
' ' | Add-Content -LiteralPath $LogFile
"   Folder   :   $FolderPath" | Add-Content -LiteralPath $LogFile
' ' | Add-Content -LiteralPath $LogFile
"   Switches :   $Switches" | Add-Content -LiteralPath $LogFile
' ' | Add-Content -LiteralPath $LogFile
('-'*79) | Add-Content -LiteralPath $LogFile
' ' | Add-Content -LiteralPath $LogFile

# Define the properties to be selected for the array, if createtime switch is specified 
# CreationTime is added to the list of properties, this is to conserve memory space
$SelectProperty = @{'Property'='Fullname','Length','PSIsContainer'}
if ($CreateTime) {
	$SelectProperty.Property += 'CreationTime'
} elseif ($LastAccessTime) {
    $SelectProperty.Property += 'LastAccessTime'
} elseif ($CompareCreateTimeLastModified) {
    $SelectProperty.Property += @{
        name = 'CustomTime'
        expression = {if ($_.lastwritetime -ge $_.CreationTime){$_.LastWriteTime} else {$_.CreationTime}}
    }
} else {
	$SelectProperty.Property += 'LastWriteTime'
}
if ($ExcludeFileExtension -or $IncludeFileExtension) {
    $SelectProperty.Property += 'Extension'
}
if ($ArchivedOnly) {
    $SelectProperty.Property += 'Attributes'
}

# Get the complete list of files and save to array
Write-Output "`n   Retrieving list of files and folders from: $FolderPath"
$CheckError = $Error.Count
if ($FolderPath -match '\[|\]') {
    $null = New-PSDrive -Name TempDrive -PSProvider FileSystem -Root $FolderPath
    $FullArray = @(Get-ChildItem -LiteralPath TempDrive:\ -Recurse -ErrorAction SilentlyContinue -Force | Select-Object @SelectProperty)
} else {
    $FullArray = @(Get-ChildItem -LiteralPath $FolderPath -Recurse -ErrorAction SilentlyContinue -Force | Select-Object @SelectProperty)
}

# Split the complete list of items into a separate list containing only the files
$FileList   = @($FullArray | Where-Object {$_.PSIsContainer -eq $false})
$FolderList = @($FullArray | Where-Object {$_.PSIsContainer -eq $true})

# If the IncludePath parameter is included then this loop will run. This will clear out any path not specified in the
# include parameter. If the ExcludePath parameter is also specified
if ($IncludePath) {
    # If RegExpath has not been specified the script will escape all regular expressions from values specified
    if (!$RegExPath) {
        for ($j=0;$j -lt $IncludePath.Count;$j++) {
		    [array]$NewFileList   += @($FileList   | Where-Object {$_.FullName -match [RegEx]::Escape($IncludePath[$j])})
            [array]$NewFolderList += @($FolderList | Where-Object {$_.FullName -match [RegEx]::Escape($IncludePath[$j])})
        }
    } else {
    # Process the list of files when RegExPath has been specified
        for ($j=0;$j -lt $IncludePath.Count;$j++) {
		    [array]$NewFileList   += @($FileList   | Where-Object {$_.FullName -match $IncludePath[$j]})
            [array]$NewFolderList += @($FolderList | Where-Object {$_.FullName -match $IncludePath[$j]})
        }        
    }
    $FileList = $NewFileList
    $FolderList = $NewFolderList
    $NewFileList=$NewFolderList = $null
}

# If the ExcludePath parameter is included then this loop will run. This will clear out the 
# excluded paths for both the filelist.
if ($ExcludePath) {
    # If RegExpath has not been specified the script will escape all regular expressions from values specified
    if (!$RegExPath) {
        for ($j=0;$j -lt $ExcludePath.Count;$j++) {
            $FileList   = @($FileList   | Where-Object {$_.FullName -notmatch [RegEx]::Escape($ExcludePath[$j])})
            $FolderList = @($FolderList | Where-Object {$_.FullName -notmatch [RegEx]::Escape($ExcludePath[$j])})
	    }
    } else {
    # Process the list of files when RegExPath has been specified
        for ($j=0;$j -lt $ExcludePath.Count;$j++) {
		    $FileList =   @($FileList   | Where-Object {$_.FullName -notmatch $ExcludePath[$j]})
            $FolderList = @($FolderList | Where-Object {$_.FullName -notmatch $ExcludePath[$j]})
	    }
    }
}

# If the -IncludeFileExtension is specified all filenames matching the criteria specified
if ($IncludeFileExtension) {
    for ($j=0;$j -lt $IncludeFileExtension.Count;$j++) {
        # If no dot is present the dot will be added to the front of the string
        if ($IncludeFileExtension[$j].Substring(0,1) -ne '.') {$IncludeFileExtension[$j] = ".$($IncludeFileExtension[$j])"}
        [array]$NewFileList += @($FileList | Where-Object {$_.Extension -like $IncludeFileExtension[$j]})
    }
    $FileList = $NewFileList
    $NewFileList=$null
}

# If the -ExcludeFileExtension is specified all filenames matching the criteria specified
if ($ExcludeFileExtension) {
    for ($j=0;$j -lt $ExcludeFileExtension.Count;$j++) {
        # If no dot is present the dot will be added to the front of the string
        if ($ExcludeFileExtension[$j].Substring(0,1) -ne '.') {$ExcludeFileExtension[$j] = ".$($ExcludeFileExtension[$j])"}
        $FileList = @($FileList | Where-Object {$_.Extension -notlike $ExcludeFileExtension[$j]})
    }
}

# Catches errors during read stage and writes to log, mostly catches permissions errors. Placed after Exclude/Include portion
# of the script to ensure excluded paths are not generating errors.
$CheckError = $Error.Count - $CheckError
if ($CheckError -gt 0) {
	for ($j=0;$j -lt $CheckError;$j++) {
        # Verifies is the error does not match an excluded path, only errors not matching excluded paths will be written to the Log	
        if ($ExcludePath) {
            if (!$RegExPath) {
                if ($(for ($k=0;$k -lt $ExcludePath.Count;$k++) {$Error[$j].TargetObject -match [RegEx]::Escape($ExcludePath[$k].SubString(0,$ExcludePath[$k].Length-2))}) -notcontains $true) {
                    $TempErrorVar = "$($Error[$j].ToString()) ::: $($Error[$j].TargetObject)"
		            "`tFAILED ACCESS`t$TempErrorVar" | Add-Content -LiteralPath $LogFile
                }
            } else {
                if ($(for ($k=0;$k -lt $ExcludePath.Count;$k++) {$Error[$j].TargetObject -match $ExcludePath[$k]}) -notcontains $true) {
                    $TempErrorVar = "$($Error[$j].ToString()) ::: $($Error[$j].TargetObject)"
		            "`tFAILED ACCESS`t$TempErrorVar" | Add-Content -LiteralPath $LogFile
                }            
            }
	    } else {
            $TempErrorVar = "$($Error[$j].ToString()) ::: $($Error[$j].TargetObject)"
		    "`tFAILED ACCESS`t$TempErrorVar" | Add-Content -LiteralPath $LogFile
        }
    }
}

# Counter for prompt output
$AllFileCount = $FileList.Count

# If the -CreateTime switch has been used the script looks for file creation time rather than
# file modified/lastwrite time
if ($CreateTime) {
	$FileList = @($FileList | Where-Object {$_.CreationTime -le $LastWrite})
} elseif ($LastAccessTime) {
    $FileList = @($FileList | Where-Object {$_.LastAccessTime -le $LastWrite})
} elseif ($CompareCreateTimeLastModified) {
    $FileList = @($FileList | Where-Object {$_.CustomTime -le $LastWrite})
} else {
    $FileList = @($FileList | Where-Object {$_.LastWriteTime -le $LastWrite})
}

# If the ExcludeDate parameter is specified the query is converted by the ConvertFrom-Query function. The
# output of that table is a hashtable that is splatted to the ConvertTo-DateObject function which returns
# an array of dates. All files that match a date in the returned array will be excluded from deletion which
# allows for more specific exclusions.
if ($ExcludeDate) {
    $SplatDate = ConvertFrom-DateQuery $ExcludeDate
    $ExcludedDates = @(ConvertTo-DateObject @SplatDate | Select-Object -Unique | Sort-Object -Descending)
    if ($CreateTime) {
        $FileList = @($FileList | Where-Object {$ExcludedDates -notcontains $_.CreationTime.Date})
    } elseif ($LastAccessTime) {
        $FileList = @($FileList | Where-Object {$ExcludedDates -notcontains $_.LastAccessTime.Date})
    } elseif ($CompareCreateTimeLastModified) {
        $FileList = @($FileList | Where-Object {$ExcludedDates -notcontains $_.CustomTime.Date})
    } else {
        $FileList = @($FileList | Where-Object {$ExcludedDates -notcontains $_.LastWriteTime.Date})
    }
    [string]$DisplayExcludedDates = for ($j=0;$j -lt $ExcludedDates.Count;$j++) {
        if ($j -eq 0) {
            "`n   ExcludedDates: $($ExcludedDates[$j].ToString('yyyy-MM-dd'))"
        } else {
            $ExcludedDates[$j].ToString('yyyy-MM-dd')
        }
        # After every fifth date start on the next line
        if ((($j+1) % 6) -eq 0) {"`n`t`t "}
    }
    $DisplayExcludedDates
}

# If -KeepFile is specified this block will ensure that x-number of files will remain in the folder
if ($KeepFile) {
    $FileList | Select-Object -Property *,@{
        name       = 'ParentFolder'
        expression = {
            Split-Path -Path $_.FullName
        }
    } | Group-Object -Property ParentFolder | Where-Object {$_.Count -ge $KeepFile} | ForEach-Object {
        if ($CreateTime) {
            $FileList = @($_.Group | Sort-Object -Property CreationTime   | Select-Object -Last ($_.Count-$KeepFile))
        } elseif ($LastAccessTime) {
            $FileList = @($_.Group | Sort-Object -Property LastAccessTime | Select-Object -Last ($_.Count-$KeepFile))
        } elseif ($CompareCreateTimeLastModified) {
            $FileList = @($_.Group | Sort-Object -Property CustomTime     | Select-Object -Last ($_.Count-$KeepFile))
        } else {
            $FileList = @($_.Group | Sort-Object -Property LastWriteTime  | Select-Object -Last ($_.Count-$KeepFile))
        }
    }
}
# Defines the list of folders, either a complete list of all folders if -CleanFolders
# was specified or just the folders containing old files. The -NoFolder switch will ensure
# the folder structure is not modified and only files are deleted.
if ($CleanFolders) {
    # Uses the FolderList variable defined at the start of the script, including any exclusions/inclusions
} elseif ($NoFolder) {
    $FolderList = @()
} else {
    $FolderList = @($FileList | ForEach-Object {
        Split-Path -Path $_.FullName} |
        Select-Object -Unique | ForEach-Object {
        Get-Item -LiteralPath $_ -ErrorAction SilentlyContinue | Select-Object @SelectProperty
    })
}

# If -ArchivedOnly switch is set then eliminate any files that still have their archive bit set.
if ($ArchivedOnly)
{
    $FileList = @($FileList | Where-Object {$_.Attributes -notmatch 'Archive'})
}

# Clear original array containing files and folders and create array with list of older files
$FullArray = $null

# Write totals to console
Write-Output 	 "`n   Files`t: $AllFileCount`n   Folders`t: $($FolderList.Count) `n   Old files`t: $($FileList.Count)"

# Execute main functions of script
if (-not $ListOnly) {
    Write-Output "`n   Starting with removal of old files..."
} else {
    Write-Output "`n   Listing files..."
}

#region Delete Files
# This section determines in a loop which files are deleted. If a file fails to be deleted
# an error is logged and the error message is written to the log.
# $count is used to speed up the delete fileloop and will also be used for other large loops in the script
$Count = $FileList.Count
for ($j=0;$j -lt $Count;$j++) {
	$TempFile = $FileList[$j].FullName
	$TempSize = $FileList[$j].Length
	if (-not $ListOnly) {Remove-Item -LiteralPath $Tempfile -Force -ErrorAction SilentlyContinue}
	if (-not $?) {
		$TempErrorVar = "$($Error[0].ToString()) ::: $($Error[0].targetobject)"
		"`tFAILED FILE`t`t$TempErrorVar" | Add-Content -LiteralPath $LogFile
		$FilesFailed++
		$FailedSize+=$TempSize
	} else {
		if (-not $ListOnly) {
            $FilesNumber++
            $FilesSize+=$TempSize
            if ($VerboseLog) {
                switch ($true) {
                    {$CreateTime} {"`tDELETED FILE`t$($FileList[$j].CreationTime.ToString('yyyy-MM-dd hh:mm:ss'))`t$($FileList[$j].Length.ToString().PadLeft(15))`t$tempfile" | Add-Content -LiteralPath $LogFile}
                    {$LastAccessTime} {"`tDELETED FILE`t$($FileList[$j].LastAccessTime.ToString('yyyy-MM-dd hh:mm:ss'))`t$($FileList[$j].Length.ToString().PadLeft(15))`t$tempfile" | Add-Content -LiteralPath $LogFile}
                    {$CompareCreateTimeLastModified} {"`tDELETED FILE`t$($FileList[$j].CustomTime.ToString('yyyy-MM-dd hh:mm:ss'))`t$($FileList[$j].Length.ToString().PadLeft(15))`t$tempfile" | Add-Content -LiteralPath $LogFile}
                    Default {"`tDELETED FILE`t$($FileList[$j].LastWriteTime.ToString('yyyy-MM-dd hh:mm:ss'))`t$($FileList[$j].Length.ToString().PadLeft(15))`t$tempfile" | Add-Content -LiteralPath $LogFile}
                }
            }
        }
	}
	if($ListOnly) {
        if ($VerboseLog) {
            switch ($true) {
                {$CreateTime} {"`tLISTONLY`t$($FileList[$j].CreationTime.ToString('yyyy-MM-dd hh:mm:ss'))`t$($FileList[$j].Length.ToString().PadLeft(15))`t$tempfile" | Add-Content -LiteralPath $LogFile}
                {$LastAccessTime} {"`tLISTONLY`t$($FileList[$j].LastAccessTime.ToString('yyyy-MM-dd hh:mm:ss'))`t$($FileList[$j].Length.ToString().PadLeft(15))`t$tempfile" | Add-Content -LiteralPath $LogFile}
                {$CompareCreateTimeLastModified} {"`tLISTONLY`t$($FileList[$j].CustomTime.ToString('yyyy-MM-dd hh:mm:ss'))`t$($FileList[$j].Length.ToString().PadLeft(15))`t$tempfile" | Add-Content -LiteralPath $LogFile}
                Default {"`tLISTONLY`t$($FileList[$j].LastWriteTime.ToString('yyyy-MM-dd hh:mm:ss'))`t$($FileList[$j].Length.ToString().PadLeft(15))`t$tempfile" | Add-Content -LiteralPath $LogFile}
            }
        } else {
            "`tLISTONLY`t$TempFile" | Add-Content -LiteralPath $LogFile
        }
		$FilesNumber++
		$FilesSize+=$TempSize
	}
}
#endregion

if (-not $ListOnly) {
    Write-Output "   Finished deleting files`n"
} else {
    Write-Output "   Finished listing files`n"
}
if (-not $ListOnly) {
	Write-Output '   Check/remove empty folders started...'

#region Delete Folders
    # Checks whether folder is empty and uses temporary variables
    # Main loop goes through list of folders, only deleting the empty folders
    # The if(-not $tempfolder) is the verification whether the folder is empty
	$FolderList = @($FolderList | sort-object @{Expression={$_.FullName.Length}; Ascending=$false})
	$Count = $FolderList.Count
	for ($j=0;$j -lt $Count;$j++) {
		$TempFolder = Get-ChildItem -LiteralPath $FolderList[$j].FullName -ErrorAction SilentlyContinue -Force
		if (-not $TempFolder) {
		    $TempName = $FolderList[$j].FullName
		    Remove-Item -LiteralPath $TempName -Force -Recurse -ErrorAction SilentlyContinue
			if(-not $?) {
				$TempErrorVar = "$($Error[0].ToString()) ::: $($Error[0].targetobject)"
				"`tFAILED FOLDER`t$TempErrorVar" | Add-Content -LiteralPath $LogFile
				$FoldersFailed++
			} else {
				if ($VerboseLog) {
                    switch ($true) {
                        {$CreateTime} {"`tDELETED FOLDER`t$($FolderList[$j].CreationTime.ToString('yyyy-MM-dd hh:mm:ss'))`t`t`t$TempName" | Add-Content -LiteralPath $LogFile}
                        {$LastAccessTime} {"`tDELETED FOLDER`t$($FolderList[$j].LastAccessTime.ToString('yyyy-MM-dd hh:mm:ss'))`t`t`t$TempName" | Add-Content -LiteralPath $LogFile}
                        {$CompareCreateTimeLastModified} {"`tDELETED FOLDER`t$($FolderList[$j].CustomTime.ToString('yyyy-MM-dd hh:mm:ss'))`t`t`t$TempName" | Add-Content -LiteralPath $LogFile}
                        Default {"`tDELETED FOLDER`t$($FolderList[$j].LastWriteTime.ToString('yyyy-MM-dd hh:mm:ss'))`t`t`t$TempName" | Add-Content -LiteralPath $LogFile}
                    }
                }
				$FoldersNumber++
			}
		}
	}
#endregion

	Write-Output "   Empty folders deleted`n"
}

# Pre-format values for footer
$TimeTaken          = ((Get-Date) - $StartDate).ToString().SubString(0,8)
$FilesSize          = $FilesSize/1MB
[string]$FilesSize  = $FilesSize.ToString()
$FailedSize         = $FailedSize/1MB
[string]$FailedSize = $FailedSize.ToString()
$EndDate            = "$((Get-Date).ToShortDateString()), $((Get-Date).ToLongTimeString())"

# Write footer to log and output to console
Write-Output ($Footer = @"

$('-'*79)

   Files               : $FilesNumber
   Filesize(MB)        : $FilesSize
   Files Failed        : $FilesFailed
   Failedfile Size(MB) : $FailedSize
   Folders             : $FoldersNumber
   Folders Failed      : $FoldersFailed

   Finished Time       : $EndDate
   Total Time          : $TimeTaken

$('-'*79)
"@)

$Footer | Add-Content -LiteralPath $LogFile

# Section of script that emails the logfile if required parameters are specified.
if ($EmailSplat) {
    Send-MailMessage @EmailSplat
}

# Clean up variables at end of script
$FileList=$FolderList = $null