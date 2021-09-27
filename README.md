# USC_Stimmberechtigte
Erzeugt eine Liste aller Stimmberechtigten für Mitgliederversammlungen des Ulmer Spatzen Chors

Die Herausforderung liegt darin, dass laut Satzung die aktiven Mitglieder (Sänger:innen) stimmberechtigt sind (pro Person eine Stimme), das Stimmrecht bei unter 18-Jährigen aber von deren Eltern wahrgenommen wird (die somit mehrere Stimmen haben können). Weiterhin sind Fördermitglieder selbst stimmberechtigt, aber nur, wenn sie nicht schon ein Chormitglied vertreten. Diese Liste berücksichtigt diese Regeln und gibt für jede Stimme an, wer sie ausübt und ggf. wessen Stimmrecht übertragen wurde.

### Wie funktioniert's? ###

* Starten der Abfrage "Stimmberechtigte Mitgliederversammlung"
* Ziel der Abfrage: Zwischenablage
* Starten des Skripts Stimmberechtigten-Liste.py
* Es wird eine CSV-Datei im aktuellen Verzeichnis abgelegt.

### Was brauche ich? ###

* Installation von ComMusic
* Import der Abfrage in den ComMusic-Reporter
* Installation von Python 3.7 oder höher
* Installation des Moduls win32clipboard
