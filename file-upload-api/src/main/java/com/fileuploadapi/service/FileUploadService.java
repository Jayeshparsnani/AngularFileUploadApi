package com.fileuploadapi.service;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.util.Iterator;

@Service
public class FileUploadService {

    public JSONObject parseFileToJson(MultipartFile file) throws Exception {
        String fileName = file.getOriginalFilename();
        JSONObject jsonResult = new JSONObject();

        try (InputStream is = file.getInputStream()) {
            Workbook workbook = null;

            // Determine if the file is XLS or XLSX based on its extension
            if (fileName.endsWith(".xls")) {
                workbook = new HSSFWorkbook(is); // Use HSSFWorkbook for .xls files
            } else if (fileName.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(is); // Use XSSFWorkbook for .xlsx files
            } else {
                throw new OLE2NotOfficeXmlFileException("Unsupported file format.");
            }

            Sheet sheet = workbook.getSheetAt(0); // Assuming the first sheet is used
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                JSONObject rowJson = new JSONObject();

                // Assuming the first row contains headers
                for (int cellIndex = 0; cellIndex < row.getPhysicalNumberOfCells(); cellIndex++) {
                    Cell cell = row.getCell(cellIndex);
                    if (cell != null) {
                        String cellValue = null;
                        switch (cell.getCellType()) {
                            case STRING:
                                cellValue = cell.getStringCellValue();
                                break;
                            case NUMERIC:
                                cellValue = String.valueOf(cell.getNumericCellValue());
                                break;
                            case BOOLEAN:
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            default:
                                cellValue = "Unsupported cell type";
                        }

                        rowJson.put("column" + (cellIndex + 1), cellValue);
                    } else {
                        rowJson.put("column" + (cellIndex + 1), "null"); // Handle null cells
                    }
                }
                jsonResult.put("row" + (row.getRowNum() + 1), rowJson);
            }
        }

        return jsonResult;
    }
}