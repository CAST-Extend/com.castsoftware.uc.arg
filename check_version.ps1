if ($args.count -lt 1) {
    write-host 'Not enough arguments'
    write-host 'check_version <minimum version>'
    exit -1
}

$rslt = $$ -match '(\d+)\.(\d+)\.(\d+)'
if ($rslt -eq $false) {
    write-host 'Minimum version must be in the format of ##.##.## ($$)'
    exit -1
}
$min_version = [int]$matches[1]*10000 + [int]$matches[2]*100 + [int]$matches[3]

$text = py --version
$rslt = $text -match '([pP]ython)\s(\d+\.\d+)\.(\d+)'
if ($rslt) {
    $system_version = [double]$matches[2]*10000+[int]$matches[3]

    write-host ' The system python version is:' $system_version
    write-host 'The minimum python version is:' $min_version

    if ($system_version -lt $min_version) {
        echo upgrade 
        exit 1
    } else {
        echo 'no upgrade'
    }
} else {
    # python is not installed
    exit 1
}
exit 0

