option explicit

sub create(xlsName as variant) ' {

    with sheet1 ' {

        .name = "Cockpit"

        .cells(2,2) = "Provider"              : .cells(2,3).name = "oleDbProv"           : range("oleDbProv"    ) = "Microsoft.ACE.OLEDB.12.0"
        .cells(3,2) = "Data Source"           : .cells(3,3).name = "oleDbDataSrc"        : range("oleDbDataSrc" ) =  xlsName
        .cells(4,2) = "Extended Properties"   : .cells(4,3).name = "oleDbExtProps"       : range("oleDbExtProps") = "Excel 12.0 Xml;HDR=YES;IMEX=1"

        .columns(2).autoFit

         with addTextBox(.range(.cells(6, 2), .cells(18, 26)), "sqlText") ' {
             .font              = "Courier New"
             .enterKeyBehavior  =  true
             .multiLine         =  true
             .borderStyle       =  fmBorderStyleSingle
             .backColor         =  rgb(255, 250, 200)
         end with ' }

         with addButton(.range(.cells(20,2), .cells(21,3)), "btn", "Run SQL") ' {

         end with ' }

    end with ' }

end sub ' }

sub processQuery() ' {

    dim sht as workSheet
    set sht = worksheets.add

    dim dataSource as string
    dataSource = sheet1.range("oleDbDataSrc")

    if dir(dataSource) = "" then ' {
       msgBox "Data source not found:" & vbCrLf & dataSource
       exit sub
    end if ' }

    dim source as string
    source = "OLEDB;provider=" & sheet1.range("oleDbProv") & ";data source=" & dataSource

    if sheet1.range("oleDbExtProps") <> "" then
       source = source & ";Extended Properties='" & sheet1.range("oleDbExtProps") & "'"
    end if

    insertListObject _
       source       := source           , _
       sqlStatement := sheet1.sqlText   , _
       destCell     := sht.cells(3,1)

end sub ' }

sub insertListObject( source as string, sqlStatement as string, destCell as range) ' {

 on error goto err_

    dim listObj as listObject

    set listObj = activeSheet.listObjects.add( _
        sourceType  := xlSrcExternal         , _
        source      := array(source)         , _
        destination := destCell)

    with listObj ' {

        .displayName = "Data_from_other_worksheet" ' Must not contain white spaces

         with .queryTable ' {

'            .adjustColumnWidth      = true                  ' True is default anyway

             .commandType            = xlCmdSql
             .commandText            = array(sqlStatement)
'            .rowNumbers             = false

             .refreshOnFileOpen      = false                 ' Get newest data when worksheet is opened (Default is false)
             .backgroundQuery        = true                  ' Update data asynchronously
             .refreshStyle           = xlInsertDeleteCells   ' Partial rows are inserted or deleted to match the exact number of rows required for the new recordset.
             .saveData               = true
             .refreshPeriod          = 0                     ' Refresh period in minuts. 0 disables refreshing.
             .preserveColumnInfo     = true                  ' Preserve sorting, filtering, and layout information when data is refreshed.


             .refresh backgroundQuery := false               ' Refresh the data NOW.

         end with ' }

    end with ' }

    exit sub

  err_:

    msgBox err.number & chr(10) & err.description

end sub ' }

sub createSourceWorksheet(fileName as string) ' {

  '
  '  Delete source workbook file if it alread exists.
  '
    if dir(fileName) <> "" then ' {
       kill fileName
    end if ' }

    dim otherWorkbook as workbook
    set otherWorkbook = workbooks.add

    with otherWorkbook ' {

      dim firstCell as range

      with .sheets(1) ' {

        dim r as long : r = 3
        set firstCell = .cells(r,2)

       .range( .cells(r, 2), .cells(r, 4) ).value = array("Col one", "Col two", "Col three"  ) : r = r + 1
       .range( .cells(r, 2), .cells(r, 4) ).value = array("Baz"    ,       42 , #2020-03-03# ) : r = r + 1
       .range( .cells(r, 2), .cells(r, 4) ).value = array("Bar"    ,       99 , #2018-05-17# ) : r = r + 1
       .range( .cells(r, 2), .cells(r, 4) ).value = array("Baz"    ,   123456 , #2019-11-13# ) : r = r + 1
       .range( .cells(r, 2), .cells(r, 4) ).value = array("Foo"    ,      518 , #2018-07-19# ) : r = r + 1
       .range( .cells(r, 2), .cells(r, 4) ).value = array("Baz"    ,      219 , #2014-10-02# ) : r = r + 1
       .range( .cells(r, 2), .cells(r, 4) ).value = array("Foo"    ,       21 , #2015-09-09# )

    '
    '   Name a source data range
    '
       .range( firstCell, .cells(r,4) ).name = "srcTable"

       .usedRange.columns.autoFit

      end with ' }

     .saveAs                            _
        fileName   := fileName,         _
        fileFormat := xlOpenXMLWorkbook

     .close

    end with ' }

end sub ' }
