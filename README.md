# excel2db

Imports Excel-Files into a Database.

to use it, a file archiv.properties has to be in the classpath. As an example use [ archiv.properties-default](https://raw.githubusercontent.com/rtrier/excel2db/master/archiv.properties-default). 
If you want to access not a postgresql database you have to put the jdbc driver in the classpath.

It imports all worksheets with the same structure into one table. If validation is enabled all worksheets have to have the same structure.

__Parmeters:__

&nbsp;&nbsp;&nbsp;&nbsp;schema=targetSchema<br>
&nbsp;&nbsp;&nbsp;&nbsp;file=path2file or dir=path (if dir is specified file will be ignored)<br>
&nbsp;&nbsp;&nbsp;&nbsp;createSchemaIfNotExits=true|false Standard=false<br>

to import with validation:<br>
&nbsp;&nbsp;&nbsp;&nbsp;importTableTypes=DatabaseTableWithDescription<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;columns=(table_type, source_column_name, target_column_name, data_type, is_nullable)
<br>&nbsp;&nbsp;&nbsp;&nbsp;expectedTableType=value of table_type from above table
<br>&nbsp;&nbsp;&nbsp;&nbsp;stopOnValidationErrors=bool Standard=true
