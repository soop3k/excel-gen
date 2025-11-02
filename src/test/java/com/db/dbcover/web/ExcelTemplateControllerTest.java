package com.db.dbcover.web;

import com.db.dbcover.service.ExcelGeneratorService;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.WebMvcTest;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.boot.test.mock.mockito.MockBean;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.test.context.TestPropertySource;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.test.web.servlet.MvcResult;
import org.springframework.web.server.ResponseStatusException;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.Mockito.when;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.get;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.*;

@SpringBootTest
@TestPropertySource(locations = "classpath:application.yml")
class ExcelTemplateControllerTest {

    @Autowired
    private ExcelTemplateController excelTemplateController;

    @Test
    void downloadTemplate_MORTGAGE_VerifiesExcelColumnsStructure() throws Exception {
        // Given
        String instrumentType = "MORTGAGE";

        // When
        ResponseEntity<byte[]> response = excelTemplateController.downloadTemplate(instrumentType);
        byte[] excelData = response.getBody();

        // Then
        assertThat(excelData).isNotNull();
        
        // Verify filename in Content-Disposition header
        String contentDisposition = response.getHeaders().getFirst(HttpHeaders.CONTENT_DISPOSITION);
        assertThat(contentDisposition).isNotNull();
        
        String currentDate = LocalDate.now().format(DateTimeFormatter.ofPattern("yyyyMMdd"));
        String expectedFilename = String.format("mortgage_bulk_upload_%s.xlsx", currentDate);
        assertThat(contentDisposition).contains(expectedFilename);
        
        try (ByteArrayInputStream inputStream = new ByteArrayInputStream(excelData);
             Workbook workbook = new XSSFWorkbook(inputStream)) {
            
            // Verify we have the expected number of sheets
            assertThat(workbook.getNumberOfSheets()).isEqualTo(6);
            
            // Verify sheet names
            assertThat(workbook.getSheetName(0)).isEqualTo("INSTRUMENT_DETAILS");
            assertThat(workbook.getSheetName(1)).isEqualTo("LINKED_DEALS");
            assertThat(workbook.getSheetName(2)).isEqualTo("LINKED_ASSETS");
            assertThat(workbook.getSheetName(3)).isEqualTo("PERSISTED_IDS");
            assertThat(workbook.getSheetName(4)).isEqualTo("LINKED_INSTRUMENTS");
            assertThat(workbook.getSheetName(5)).isEqualTo("LINKED_PARTIES");
            
            // Verify INSTRUMENT_DETAILS sheet columns
            Sheet instrumentDetailsSheet = workbook.getSheet("INSTRUMENT_DETAILS");
            Row headerRow = instrumentDetailsSheet.getRow(0);
            assertThat(headerRow.getCell(0).getStringCellValue()).isEqualTo("INSTRUMENT_ID");
            assertThat(headerRow.getCell(1).getStringCellValue()).isEqualTo("INSTRUMENT_NAME");
            assertThat(headerRow.getCell(2).getStringCellValue()).isEqualTo("CURRENCY");
            assertThat(headerRow.getCell(3).getStringCellValue()).isEqualTo("ISSUE_DATE");
            
            // Verify LINKED_DEALS sheet columns
            Sheet linkedDealsSheet = workbook.getSheet("LINKED_DEALS");
            Row dealsHeaderRow = linkedDealsSheet.getRow(0);
            assertThat(dealsHeaderRow.getCell(0).getStringCellValue()).isEqualTo("DEAL_ID");
            assertThat(dealsHeaderRow.getCell(1).getStringCellValue()).isEqualTo("DEAL_TYPE");
            assertThat(dealsHeaderRow.getCell(2).getStringCellValue()).isEqualTo("DEAL_DATE");
            assertThat(dealsHeaderRow.getCell(3).getStringCellValue()).isEqualTo("NOTIONAL");
            
            // Verify LINKED_ASSETS sheet columns
            Sheet linkedAssetsSheet = workbook.getSheet("LINKED_ASSETS");
            Row assetsHeaderRow = linkedAssetsSheet.getRow(0);
            assertThat(assetsHeaderRow.getCell(0).getStringCellValue()).isEqualTo("ASSET_ID");
            assertThat(assetsHeaderRow.getCell(1).getStringCellValue()).isEqualTo("ASSET_CLASS");
            assertThat(assetsHeaderRow.getCell(2).getStringCellValue()).isEqualTo("ASSET_VALUE");
        }
    }
}