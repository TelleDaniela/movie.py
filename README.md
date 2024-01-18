
# Projekts FILM SORTER
# Projekta apraksts
Šis Python projekts ļauj lietotājam veidot Excel tabulas un turēt tajā savu filmu sarakstu. Programma piedāvā iespēju pievienot jaunas filmas tabulā, dzēst esošas filmas, rediģēt esošo filmu datus, izvēlēties nejaušu filmu no tabulas un izdrukāt visu filmu sarakstu uz konsoles. Projekts izmanto openpyxl bibliotēku, jo tā nodrošina ērtu un efektīvu veidu, kā manipulēt ar Excel failiem, saglabājot datu kvalitāti un organizāciju.

## Izmantotās Python bibliotēkas
Projekta izstrādē tika izmantota openpyxl bibliotēka, jo tā piedāvā iespēju ērti manipulēt ar Excel failiem, veicot tajos lasīšanu un rakstīšanu. Šī bibliotēka nodrošina spēju pievienot jaunas rindiņas, dzēst esošas, rediģēt šūnas un veikt citas darbības, kas ir būtiskas, lai veidotu un pārvaldītu filmu sarakstu Excel formātā. Tas padara projektu piemērotu filmu mīļotājiem, kuri vēlas sistematizēti uzturēt savu filmu kolekciju. Ari izmantota biblioteka random, lai izveliet nejausi filmu.

## Programmatūras izmantošanas metodes
### Pievienot jaunu filmu:

Izsauciet funkciju, kas ļauj pievienot jaunu filmu, norādot nepieciešamos parametrus, piemēram, nosaukumu.
### Dzēst filmu:

Norādiet filmas nosaukumu vai id, lai izdzēstu konkrētu filmu no saraksta.
### Rediģēt filmas datus:

Izsauciet funkciju, kas ļauj rediģēt konkrētas filmas laiks, norādot filmas jauno nosaukumu un laiks.
### Izvēlēties nejaušu filmu:

Izmantojiet funkciju, kas izvēlas un atgriež nejauši izvēlētu filmu no saraksta.
### Izdrukāt visu sarakstu uz konsoles:

Izmantojiet funkciju, kas iegūst visu filmu sarakstu un to izdrukā uz konsoles.