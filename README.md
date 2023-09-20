
# Spring Excel Utilities Lib

This is a basic Excel library which helps of creating and reading Excels.
There are 4 classes to support the Excel generation and parsing.

In SpringBootMainApplication class add the following scan base to activate the library
```Java
@SpringBootApplication(scanBasePackages = {"com.utku"});
```
## Usage/Examples

Creating excel from data object
```Java
@Autowire
private ExcelGenerationService excelGenerationService;

..

List<CellSourceData> cellList = new ArrayList<>();
List<RowSourceData> rowSourceDataList = new ArrayList<>();

cellList.add(CellSourceData.builder().columnNumber(0).cellType(CellType.STRING).isHeader(true).value("Row 1 Cell 1").build());

rowSourceDataList.add(RowSourceData.builder().rowNumber(0).cellList(cellList).build());
cellList = new ArrayList<>();

cellList.add(CellSourceData.builder().columnNumber(0).cellType(CellType.STRING).isHeader(false).value("Row 2 Cell 1").build());
rowSourceDataList.add(RowSourceData.builder().rowNumber(1).cellList(cellList).build());

SheetSourceData sheetSourceData = new SheetSourceData();
sheetSourceData.setRowSourceDataList(rowSourceDataList);
sheetSourceData.setSheetName("Sheet name");
sheetSourceData.setColumnSize(10);

ExcelSourceData excelSourceData = new ExcelSourceData();
excelSourceData.setSheetList(List.of(sheetSourceData));
excelSourceData.setValueBorderStyle(BorderStyle.MEDIUM);
excelSourceData.setHeaderBorderStyle(BorderStyle.MEDIUM);

//creates and returns excel stream for given data
excelGenerationService.streamExcel(excelSourceData)
```

Reading object from stream
```Java
//file is a MultiPartFile or alternatives.
try(InputStream inputStream = file.getInputStream()){
    if(excelGenerationService.isFileExcel(file)){
        ExcelSourceData sourceData = excelGenerationService.readExcelFromStream(inputStream);
    }catch (Exception e){
        throw new RequestException("Document could not be read");
    }
}

```

