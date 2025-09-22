# Planeaci-n_SEO
El objetivo de este proyecto es crear una automatizacion para la asignacion de tareas con respecto al calendario mensual para cada uno de los agentes de SEO.

Funcionamiento de la automatizacion:

En esta automatizacion tomamos los archivos mensuales de planeacion que normalmente Andreina llena el mes vencido, al cual le agregamos 4 columnas para su funcionamiento las cuales son:

-Responsable de produccion
-Tiempo estimado produccion
-Responsable de publicacion
-Tiempo estimado publicacion

En base a estas columnas el algoritmo detecta a que agente en su respectiva hoja asignar cada tarea, en relacion con la fecha, gracias a las columnas de tiempo estimado produccion y tiempo estimado publicacion, el algoritmo 
tiene un tiempo estimado el cual en el encabezado del calendario suma cada uno de los tiempos y da la informacion de cuantas horas de carga laboral tiene cada uno de los agentes al dia. Cuando la carga laboral sobrepasa las 8 horas al dia, se marca en color rojo el fondo del encabezado de la fecha.

NOTA: En las columnas de los tiempos estimados es crucial que se ponga el tiempo en minutos, es decir, para una hora se escribe "60" y asi sucesivamente.

El algoritmo funciona leyendo en la carpeta Archivos/Archivos_origen todos los archivos de cada partner pais, diligenciados por andreina, posteriormente, recorre cada uno de los archivos en un bucle y genera un dataframe general, creando una tabla general con todos los datos y todas las columnas, agregando una columna nueva donde nos indica a que archivo pertenece, y tambien a partir d eesta creando otra columna con el nombre del partner, para agregar en la tarea. 

Luego se crea un dataframe a partir de esta tabla unicamente con las columnas con la informacion necesaria para llenar las tareas de cada uno de los agentes.

Posteriormente este dataframe se guarda en SQLite, donde luego dependiendo del mes que se asigne y el año, se va a hacer una consulta de la stareas por fecha, el codigo va a tomar una plantilla que va a llenar los encabezados dinamicamente dependiendo del dia de la semana donde empieza el dia 1 del mes, luego se va a llenar la plantilla en base al mes seleccionado para cada partner y cada agente.

NOTA: la plantilla se debe renombrar siempre con el mes que se va a automatizar, ejemplo: (octubre.xlsx) y en minuscula todo, y ya en el archivo de python cambiar unicamente el numero de mes, de un digito, y el año, si es necesario, al cambiar de año, es necesario cambiar en el codigo el nombre del archivo, manteniendo la misma estructura, cambiando solo el año ejemplo: 2025 a 2026.
