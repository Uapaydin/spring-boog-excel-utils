package com.utku.excelutilities.service.impl;

import com.utku.excelutilities.exception.ExcelGenerationException;
import com.utku.excelutilities.model.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

@TestInstance(TestInstance.Lifecycle.PER_CLASS)
@Slf4j
class ExcelGenerationServiceImplTest {

    private ExcelGenerationServiceImpl excelGenerationService;

    public InputStream inputStream;
    @BeforeEach
    public void init (){
        excelGenerationService = new ExcelGenerationServiceImpl();
        inputStream = getClass().getClassLoader().getResourceAsStream("test.xlsx");
    }
    @Test
    void hasExcelFormat() throws IOException {
        MultipartFile multipartFile = new CustomMultipartFile(inputStream.readAllBytes(),"test.xlsx",ExcelGenerationServiceImpl.DOCUMENT_TYPE);
        boolean actual = ExcelGenerationServiceImpl.hasExcelFormat(multipartFile);
        assertTrue(actual);
    }

    @Test
    void generateWorkBook() {

        ExcelSourceData excelSourceData = createSourceObject();
        try(Workbook workBook = excelGenerationService.generateWorkBook(excelSourceData)){
            assertEquals(excelSourceData.getSheetList().size(),workBook.getNumberOfSheets());
            assertEquals("Row 1 Cell 1", workBook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
        }catch (Exception e){
            log.error(e.getMessage());
        }

    }
    @Test
    void streamExcel() {
        ExcelSourceData excelSourceData = createSourceObject();
        byte[] stream = new byte[0];
        try {
            stream = excelGenerationService.streamExcel(excelSourceData);
            assertTrue(stream.length > 0);
        } catch (ExcelGenerationException ex) {
            throw new RuntimeException(ex);
        }
    }

    @Test
    void readExcelFromStream() {
        try {
            ExcelSourceData excelSourceData = excelGenerationService.readExcelFromStream(inputStream);
            assertEquals(1, excelSourceData.getSheetList().size());
            assertEquals(2, excelSourceData.getSheetList().get(0).getRowSourceDataList().size());
        } catch (ExcelGenerationException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void checkExtension() throws IOException {
        MultipartFile multipartFile = new CustomMultipartFile(inputStream.readAllBytes(),"test.xlsx",ExcelGenerationServiceImpl.DOCUMENT_TYPE);
        boolean actual = excelGenerationService.checkExtension(multipartFile);
        assertTrue(actual);
    }

    @Test
    void isFileExcel() throws IOException {
        MultipartFile multipartFile = new CustomMultipartFile(inputStream.readAllBytes(),"test.xlsx",ExcelGenerationServiceImpl.DOCUMENT_TYPE);
        boolean actual = excelGenerationService.isFileExcel(multipartFile);
        assertTrue(actual);

    }

    private static ExcelSourceData createSourceObject() {
        List<CellSourceData> cellList = new ArrayList<>();
        List<RowSourceData> rowSourceDataList = new ArrayList<>();
        cellList.add(CellSourceData.builder().columnNumber(0).cellType(CellType.STRING).isHeader(true).value("Row 1 Cell 1").build());
        rowSourceDataList.add(RowSourceData.builder().rowNumber(0).cellList(cellList).build());
        cellList = new ArrayList<>();
        cellList.add(CellSourceData.builder().columnNumber(0).cellType(CellType.STRING).isHeader(false).value("Row 2 Cell 1").build());
        cellList.add(CellSourceData.builder().columnNumber(1).cellType(CellType.BOOLEAN).isHeader(false).value(true).build());
        cellList.add(CellSourceData.builder().columnNumber(2).cellType(CellType.NUMERIC).isHeader(false).value(1).build());
        cellList.add(CellSourceData.builder().columnNumber(3).cellType(CellType.BLANK).isHeader(false).value(1f).build());

        rowSourceDataList.add(RowSourceData.builder().rowNumber(1).cellList(cellList).build());

        SheetSourceData sheetSourceData = new SheetSourceData();
        sheetSourceData.setRowSourceDataList(rowSourceDataList);
        sheetSourceData.setSheetName("Sheet name");
        sheetSourceData.setColumnSize(10);
        CellRangeAddress cellAddresses = new CellRangeAddress(1,1,1,2);
        sheetSourceData.getMergeRegions().add(cellAddresses);
        ExcelSourceData excelSourceData = new ExcelSourceData();
        excelSourceData.setSheetList(List.of(sheetSourceData));
        excelSourceData.setValueBorderStyle(BorderStyle.MEDIUM);
        excelSourceData.setHeaderBorderStyle(BorderStyle.MEDIUM);
        return excelSourceData;
    }
}
