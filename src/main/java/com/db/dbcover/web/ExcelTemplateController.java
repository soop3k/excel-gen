package com.db.dbcover.web;

import com.db.dbcover.service.ExcelGeneratorService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.server.ResponseStatusException;

import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

@RestController
@RequestMapping("/excel")
public class ExcelTemplateController {

    private final ExcelGeneratorService excelGeneratorService;

    public ExcelTemplateController(ExcelGeneratorService excelGeneratorService) {
        this.excelGeneratorService = excelGeneratorService;
    }

    @GetMapping(value = "/template", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
    public ResponseEntity<byte[]> downloadTemplate(@RequestParam("instrumentType") String instrumentType) {
        try {
            byte[] file = excelGeneratorService.generateTemplate(instrumentType);
            String filename = buildFilename(instrumentType);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .contentLength(file.length)
                    .body(file);
        } catch (IllegalArgumentException ex) {
            throw new ResponseStatusException(HttpStatus.BAD_REQUEST, ex.getMessage(), ex);
        } catch (IOException ex) {
            throw new ResponseStatusException(HttpStatus.INTERNAL_SERVER_ERROR, "Failed to generate Excel template", ex);
        }
    }

    private String buildFilename(String instrumentType) {
        String date = LocalDate.now().format(DateTimeFormatter.ofPattern("yyyyMMdd"));
        return String.format("%s_bulk_upload_%s.xlsx", instrumentType.toLowerCase(), date);
    }
}
