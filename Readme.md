
# Info
Mein Vorgesetzter erhielt die Aufgabe, viele Stundenzettel in Form von Excel-Dateien mit zahlreichen Tabellen zu lesen und in einer Zieldatei zu speichern. Ursprünglich überließ er es den Mitarbeitern, dies selbst zu erledigen. Infolgedessen bot ich an, die Aufgabe mithilfe eines Skripts schneller zu bewältigen. Ich entschied mich für Python, da ich bereits Erfahrung mit dieser Sprache hatte und sie mir gefällt.
<br>
<br>Nach vielen schlaflosen Nächten und trotz meiner regulären Arbeit war das Programm endlich fertig.


# Zum Programm selbst:

Das fertige Programm ist eine .exe-Datei. Hierfür habe ich PyAutoGui verwendet, da es mir ermöglicht, besser zu steuern, welche Module in die .exe-Datei konvertiert werden und welche nicht. Die vorherige Version war als .py-Datei gedacht, ich entschied mich jedoch für die .exe-Version, da sie im Allgemeinen flexibler ist. Zur Bearbeitung der xlsx-Dateien wählte ich openpyxl, weil es sich als die schnellste Bibliothek für diese Aufgabe herausstellte.

Das Lesen der Stundenzettel kann in Python aufgrund der Vielzahl der Einträge in der Woche (bei 52-53 Wochen im Jahr und 15 Mitarbeitern) mehrere Minuten dauern. Deshalb habe ich Print-Befehle eingefügt(old-version), um die Geschwindigkeit der Verarbeitung zu überprüfen und so die Leistung zu optimieren. Letztendlich verwendete ich ein Jupyter-Notebook, da es mir ermöglicht, jede einzelne Zelle zu isolieren.

Ich finde die Dot-Notation sehr angenehm und habe daher die Klasse 'AttrDict' verwendet, obwohl sie sonst keinen besonderen Sinn ergibt.

# Bilder
Hauptfenster, mit blockierten Knöpfen, da die jeweiligen Ordner & Dateien fehlen.

![programm_main](https://github.com/Joe19922/Extraction/assets/132180983/7a0c8bc7-ef6b-46a3-9983-27ccd7804c00)is

Infofenster

![programm_info](https://github.com/Joe19922/Extraction/assets/132180983/5fa95bae-8d12-4bc6-aeec-6b64a86f1c52)

Nach dem Auswählen der Ordner und der Zieldatei, werden werden diese grün und der Startknopf wird freigegeben.

![programm_main_allok](https://github.com/Joe19922/Extraction/assets/132180983/a4f54974-9392-4810-9423-8c7c48a91e2e)

Die dateien werden aus dem source ordner gelesen, dabei wird der Startknopf blockiert und am Ende sieht man im Status, was gemacht wurden ist.
Werden lückenhafte Einträge gefunden oder wenn Kostenstellen in der Ausgabedatei nicht gefunden werden, wird eine .txt Datei im vorher festgelegten "ERROR-Ordner" erstellt.

![programm_main_process](https://github.com/Joe19922/Extraction/assets/132180983/75ab60ba-2872-4cf0-918f-bc2c78526285)
![programm_main_finished](https://github.com/Joe19922/Extraction/assets/132180983/01a8b790-308f-46a4-a66a-b086bf6a5e2e)

fehlermeldung.txt

![error](https://github.com/Joe19922/Extraction/assets/132180983/c103e149-6629-4cb1-9c13-08b09454f953)
