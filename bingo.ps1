# Define Variables
$ContentPath = "C:\Users\User\Desktop\Lines.txt" # Path to get lines from
$ExportPath = "C:\Users\User\Desktop" # Path (Without Filename) to export finished bingo to
$ExportName = "Bingo.html" # Filename to export finished bingo
$SheetsToGenerate = 25 # How many sheets should be generated?
$FreeFieldText = "Frei" # Change this for the "free" middle field, e.g. "Frei" for german, "Free" for english.

# Get content for Bingo from file
$array = Get-Content -Encoding UTF8 $ContentPath

# How many different Bingo-sheets should be generated
$repeats = $SheetsToGenerate

# Define HTML-snippets
$top = '<style>
th, td {
padding: 10px;
border: 2px #666 solid;
text-align: center;
width: 20%;
}
</style>
<body style="color: black;font-size: 20px;font-family: Verdana,Arial, Helvetica, sans-serif;">'
$tableTop = '<table style="border-collapse: collapse;page-break-before: always;">'
$tableBottom = '</table>'
$bottom = '</body>'

# Begin generation
$html = $top
$html += "`n"

# Begin iteration for sheets
for ($i = 0; $i -lt $repeats; $i++){
    # Set array for lines of current sheet
    $arrayTMP = $array

    $html += $tableTop

    # Begin iteration for rows of current sheet
    for ($j = 0; $j -lt 5; $j++){
        $html += "<tr>"
        $html += "`n"

        # Begin iteration for single fields of current row
        for ($k = 0; $k -lt 5; $k++){
            
            # If middle of sheet -> Free
            if ($k -eq 2 -and $j -eq 2){
                $html += "<td>"
                $html += "Frei / Free"
                $html += "</td>"
                $html += "`n"
            }
            else{
                # Get random line from array
                $random = Get-Random -Maximum ($arrayTMP.Count) -Minimum 0
                
                # Add line to HTML
                $html += "<td>"
                $html += $arrayTMP[$random]
                $html += "</td>"
                $html += "`n"

                # Remove line from array
                $arrayTMP = @($arrayTMP | where{$_ -notlike $arrayTMP[$random]})
            }
        }
        $html += "</tr>"
        $html += "`n"
    }
    $html += $tableBottom
}
$html += $bottom

# Export finished sheets
New-Item -Path $ExportPath -Name $ExportName -ItemType File -Value $html -Force