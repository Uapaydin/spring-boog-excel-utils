package com.utku.excelutilities.service;

import com.utku.excelutilities.exception.ExcelGenerationException;
import com.utku.excelutilities.model.ExcelSourceData;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;

public interface ExcelGenerationService {
    Workbook generateWorkBook(ExcelSourceData excelSourceData);
    byte[] streamExcel(ExcelSourceData excelSourceData) throws ExcelGenerationException;
    ExcelSourceData readExcelFromStream(InputStream stream) throws ExcelGenerationException;

    boolean isFileExcel(MultipartFile file);

}
