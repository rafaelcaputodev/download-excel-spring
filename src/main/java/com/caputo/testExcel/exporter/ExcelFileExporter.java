package com.caputo.testExcel.exporter;

import com.caputo.testExcel.entities.Client;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

public class ExcelFileExporter {
    public static ByteArrayInputStream exportCustomerListToExcelFile(List<Client> clients) {

        try(Workbook workbook = new XSSFWorkbook()){

            Sheet sheet = workbook.createSheet("Customers");

            Row row = sheet.createRow(0);
            CellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
            headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            Cell cell = row.createCell(0);
            cell.setCellValue("id");
            cell.setCellStyle(headerCellStyle);

            cell = row.createCell(1);
            cell.setCellValue("age");
            cell.setCellStyle(headerCellStyle);

            cell = row.createCell(2);
            cell.setCellValue("name");
            cell.setCellStyle(headerCellStyle);

            for(int i = 0; i<clients.size(); i++){
                Row dataRow = sheet.createRow(i + 1);
                dataRow.createCell(0).setCellValue(clients.get(i).getId());
                dataRow.createCell(1).setCellValue(clients.get(i).getAge());
                dataRow.createCell(2).setCellValue(clients.get(i).getName());
            }

            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);
            sheet.autoSizeColumn(2);

            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);

            return new ByteArrayInputStream(outputStream.toByteArray());

        } catch (IOException e){
            e.printStackTrace();
            return null;
        }
    }
}
