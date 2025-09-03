# data.ods.adios.google

Gestisce un file excel online, senza dipendere da google o microsoft, DB indipendente; dopo il tag "liscio" è stato pesantemente personalizzato per un caso d'uso specifico.

## info

+ nel ```index.html``` alcune funzioni sono commentate, è stato riadattato per un uso personale specifico.

## Config

Prima riga: ```data	importo	causale	med mese	stima	totale	note```


```Config.php```

```php
<?php

namespace App\Config;

class Config
{
	static string $db_host = 'localhost';
	static string $db_name = 'excel_webapp';
	static string $db_username = 'root';
	static string $db_password = '';
}
```

## TODOS

+ ~~mobile non lascia modificare le celle~~
+ ~~meno colonne, più righe, specie mobile~~
+ ~~modale dopo che inserisci importo~~
+ ~~template data~~
+ ~~calcoli~~
+ ~~formattazione importi, fissi 2 decimali e simbolo dell'euro?~~
+ ~~autoscroll~~
+ ~~cambiare label~~
+ ~~nascondere i fogli e mostrarli solo in modalità libera, come anche tutti i pulsanti che di default sono display: none~~
+ ~~nascondere i numeri riga e le lettere delle colonne in modalità non libera~~
+ puoi ottimizzare e ripulire anche il js?
