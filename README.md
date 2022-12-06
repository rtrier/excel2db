# excel2db

Imports Excel-Files into a Database.

It imports all worksheets with the same structure into one table. If validation is enabled all worksheets have to have the same structure.

The type of an column will be determined by evalutaion the Excel tables.<br>Possible types: 
<br>&nbsp;&nbsp;&nbsp;&nbsp;String => varchar
<br>&nbsp;&nbsp;&nbsp;&nbsp;Double => float
<br>&nbsp;&nbsp;&nbsp;&nbsp;Boolean => boolean
<br>&nbsp;&nbsp;&nbsp;&nbsp;java.util.Date => timestamp
<br>&nbsp;&nbsp;&nbsp;&nbsp;java.sql.Time => time

Main-Class is Excel2DB. If the Excel file doesnt fit into memory Excel2DBSeq can be used. Then all column types are strings.

To use it, a file archiv.properties has to be in the classpath. As an example use [ archiv.properties-default](https://raw.githubusercontent.com/rtrier/excel2db/master/archiv.properties-default). 
If you want to access not a postgresql database you have to put the jdbc driver in the classpath.

__Parmeters:__

&nbsp;&nbsp;&nbsp;&nbsp;schema=targetSchema<br>
&nbsp;&nbsp;&nbsp;&nbsp;file=path2file or dir=path (if dir is specified file will be ignored)<br>
&nbsp;&nbsp;&nbsp;&nbsp;createSchemaIfNotExits=true|false Standard=false<br>

to import with validation:<br>
&nbsp;&nbsp;&nbsp;&nbsp;importTableTypes=DatabaseTableWithDescription<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;columns=(table_type, source_column_name, target_column_name, data_type, is_nullable)
<br>&nbsp;&nbsp;&nbsp;&nbsp;expectedTableType=value of table_type from above table
<br>&nbsp;&nbsp;&nbsp;&nbsp;stopOnValidationErrors=bool Standard=true

# access2db

Imports all table from access db into postgres

__Parmeters:__
&nbsp;&nbsp;&nbsp;&nbsp;host=myhost<br>
&nbsp;&nbsp;&nbsp;&nbsp;port=123<br>
&nbsp;&nbsp;&nbsp;&nbsp;database=mydb<br>
&nbsp;&nbsp;&nbsp;&nbsp;user=username<br>
&nbsp;&nbsp;&nbsp;&nbsp;password=password<br>
&nbsp;&nbsp;&nbsp;&nbsp;zielschema=targetschema<br>
&nbsp;&nbsp;&nbsp;&nbsp;datei=accesFile<br>
<br>

__Use (Docker):__<br>
download Dockerfile<br>
download archiv.properties-default to archiv.properties<br>
adjust archiv.properties<br>
docker build -t excel2db:latest .<br>
docker run --rm -v dirWithExcelFile:/exceldata/ excel2db:latest -schema=import_schema -file=excelFile [-params1=...]
