package me.renzo.planterritorio.service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Set;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;
import lombok.extern.slf4j.Slf4j;
import me.renzo.planterritorio.model.DBRecord;
import me.renzo.planterritorio.model.ExtraFieldValue;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
@Slf4j
public class FileWriterService {

  public byte[] createDatabaseFile(List<DBRecord> dbRecords) throws IOException {
    XSSFWorkbook workbook = new XSSFWorkbook();

    Sheet sheet = workbook.createSheet("Base de datos");
    sheet.setColumnWidth(0, 6000);
    sheet.setColumnWidth(1, 4000);

    AtomicInteger currentRowIndex = new AtomicInteger(0);
    Row headerRow = sheet.createRow(currentRowIndex.getAndIncrement());

    CellStyle headerStyle = workbook.createCellStyle();
    headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.index);
    headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

    XSSFFont font = workbook.createFont();
    font.setFontName("Calibri (Body)");
    font.setFontHeightInPoints((short) 11);
    font.setBold(true);
    font.setColor(IndexedColors.WHITE.index);
    headerStyle.setFont(font);

    AtomicInteger currentHeaderColumnIndex = new AtomicInteger(0);

    Stream.of("GERENTE", "KAM", "Cuenta", "Producto")
        .forEach(
            headerName -> {
              Cell headerCell = headerRow.createCell(currentHeaderColumnIndex.getAndIncrement());
              headerCell.setCellStyle(headerStyle);
              headerCell.setCellValue(headerName);
            });

    Set<String> extraFields = extractExtraFields(dbRecords);
    extraFields.forEach(
        extraFieldName -> {
          Cell headerCell = headerRow.createCell(currentHeaderColumnIndex.getAndIncrement());
          headerCell.setCellStyle(headerStyle);
          headerCell.setCellValue(extraFieldName);
        });

    dbRecords.forEach(
        dbRecord -> {
          Row row = sheet.createRow(currentRowIndex.getAndIncrement());

          Cell managerCell = row.createCell(0);
          managerCell.setCellValue(dbRecord.getManager());

          Cell kamCell = row.createCell(1);
          kamCell.setCellValue(dbRecord.getKam());

          Cell accountCell = row.createCell(2);
          accountCell.setCellValue(dbRecord.getAccount());

          Cell productCell = row.createCell(3);
          productCell.setCellValue(dbRecord.getProduct());

          AtomicInteger currentColumnIndex = new AtomicInteger(4);
          dbRecord
              .getExtraFields()
              .values()
              .forEach(
                  extraFieldValue -> {
                    Cell extraFieldCell = row.createCell(currentColumnIndex.getAndIncrement());
                    setCellValueSafely(workbook, extraFieldCell, extraFieldValue);
                  });
        });

    ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
    workbook.write(byteArrayOutputStream);
    return byteArrayOutputStream.toByteArray();
  }

  private Set<String> extractExtraFields(List<DBRecord> dbRecords) {
    return dbRecords.get(0).getExtraFields().keySet();
  }

  private void setCellValueSafely(
      XSSFWorkbook workbook, Cell cell, ExtraFieldValue extraFieldValue) {
    switch (extraFieldValue.getCellType()) {
      case STRING, BLANK -> {
        cell.setCellValue((String) extraFieldValue.getValue());
      }
      case NUMERIC -> {
        if (extraFieldValue.isCellDateFormatted()) {
          CellStyle cellStyle = workbook.createCellStyle();
          cellStyle.setDataFormat(
              workbook.getCreationHelper().createDataFormat().getFormat("dd/MM/yy"));

          cell.setCellValue((java.util.Date) extraFieldValue.getValue());
          cell.setCellStyle(cellStyle);
        } else {
          cell.setCellValue((Double) extraFieldValue.getValue());
        }
      }
      case BOOLEAN -> {
        cell.setCellValue((Boolean) extraFieldValue.getValue());
      }
      case FORMULA -> {
        if (CellType.NUMERIC.equals(extraFieldValue.getCachedFormulaResultType())) {
          cell.setCellValue((Double) extraFieldValue.getValue());
        } else {
          cell.setCellValue((String) extraFieldValue.getValue());
        }
      }
      default -> throw new UnsupportedOperationException(
          String.format("Cell type %s is not supported", cell.getCellType()));
    }
  }
}
