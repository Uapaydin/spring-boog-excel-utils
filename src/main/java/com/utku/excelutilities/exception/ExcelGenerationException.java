package com.utku.excelutilities.exception;

import org.springframework.http.HttpStatus;

public class ExcelGenerationException extends Exception {
    private final HttpStatus httpStatus;

    public ExcelGenerationException(HttpStatus httpStatus, String message) {
        super(message);
        this.httpStatus = httpStatus;
    }

    public ExcelGenerationException(HttpStatus httpStatus) {
        this.httpStatus = httpStatus;
    }

    public HttpStatus getHttpStatus() {
        return httpStatus;
    }
}
