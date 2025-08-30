# data.ods.adios.google

Gestisce un file excel online, senza dipendere da google o microsoft, DB indipendente

## info

+ nel ```index.html``` alcune funzioni sono commentate, è stato riadattato per un uso personale specifico.

## Config

Prima riga: ```date	import	note```


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
+ autoscroll
+ calcoli
