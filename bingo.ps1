$array = @(
    'Kunde trägt die Maske nicht richtig'
    'Kunde nimmt zum Sprechen die Maske ab'
    'Kunde hat die Brille beschlagen'
    'Kunde beschwert sich über Maskentragen'
    'Kunde fragt was die Zählkarten sollen'
    'Kunde hinterfragt Ausagen / "Das habe ich nicht hier." "Haben sie das nicht hier?"'
    '"Wie lautet Ihr Nachname?" "Ich bin schon im ihrem System"'
    '"Haben Sie das auch nochmal neu eingepackt da?"'
    '"Ich hab da mal eine kurze Frage." -> Kunde erzählt seine Lebensgeschichte'
    'Kunde hat einen süßen Hund dabei.'
    '"Das kostet im Internet aber nur..."'
    '"Für welches alter soll es sein?" -> Kunde zeigt nur auf ein Kind'
    'Kunde spricht eine fremde Sprache'
    '"Früher hatten Sie das aber mal."'
    '"Gibt es das auch als Taschenbuch?"'
    '"Packen Sie das auch ein?"'
    '"Haben Sie hier eine Toilette?"'
    '"Haben Sie davon noch mehr auf Lager?"'
    'Kunde will am Infoschalter (bar) bezahlen'
    'frisch geraucht/ Deo vergessen'
    '"Nein ich weiß nicht was der Beschenkte mag"'
    '"Ich hab online gesehen, dass Sie das haben."'
    'Kunde hat wärend dem Gespräch Kopfhörer drin / telefoniert'
    '"Ist das bis Weihnachten da?"'
    '"Ist das hier Etage XY?"'
    '"Haben Sie auch fremdsprachige Bücher?"'
    '"Ich brauch das aber heute!" -> Erwartungsvoller Blick'
    'Kunde unterbricht ein anderes Gespräch'
    'Anklagende Fragestellung Bsp. "Das haben Sie bestimmt nicht hier!"'
)
$repeats = 25

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
$tableTop2 = '<table style="border-collapse: collapse;">'
$tableBottom = '</table>'
$bottom = '</body>'

$html = ""
$html += $top
$html += "`n"
for ($i = 0; $i -lt $repeats; $i++){
    $arrayTMP = $array
    $html += $tableTop
    for ($j = 0; $j -lt 5; $j++){
        $html += "<tr>"
        $html += "`n"
        for ($k = 0; $k -lt 5; $k++){
            if ($k -eq 2 -and $j -eq 2){
                $html += "<td>"
                $html += "Frei"
                $html += "</td>"
                $html += "`n"
            }
            else{
                $random = Get-Random -Maximum ($arrayTMP.Count) -Minimum 0
            
                $html += "<td>"
                $html += $arrayTMP[$random]
                $html += "</td>"
                $html += "`n"

                $arrayTMP = @($arrayTMP | where{$_ -notlike $arrayTMP[$random]})
            }
        }
        $html += "</tr>"
        $html += "`n"
    }
    $html += $tableBottom
}
$html += $bottom

#$html
New-Item -Path "C:\Users\User\Desktop" -Name Bingo.html -ItemType File -Value $html -Force
