GExcelBuilder
=============

This work is basically fork of public code snippet from blogpost: http://www.technipelago.se/content/technipelago/blog/44
It provide great way read `XLS` and `XLSX` in easy `Groovy`-style syntax.

Groovy builder makes reading Microsoft Excel documents a breeze. F.e. With it you can write the following code to insert customers into your Grails database:
```groovy
new ExcelBuilder("customers.xls").eachLine([labels:true]) {
    new Person(name:"$firstname $lastname", address:address, telephone:phone).save()
}
````
If the spreadsheet has no labels on the first row, you can use numeric index to access cells:
```
new ExcelBuilderX("customers.xlsx").eachLine {
  println "First column on row ${it.rowNum} = ${cell(0)}"
}
```

In that repository version supports additional xls and xlsx files for the sane way via automatically detection file format in factory like:
```
ExcelBuilder xls = ExcelBuilder.factory(xlsfile);
```

which return apropriate builder: `ExcelBuilder` or `ExcelBuilderX` and you mostly will not care of differencies.

Added also simple gradle build file.

All feedback welcome.
