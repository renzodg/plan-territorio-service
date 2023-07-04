package me.renzo.planterritorio.service;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import lombok.extern.slf4j.Slf4j;
import me.renzo.planterritorio.model.DBRecord;
import me.renzo.planterritorio.model.ExtraFieldValue;
import org.apache.commons.math3.util.Pair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
@Slf4j
public class FileReaderService {
  private static final String DB_SHEET_NAME = "Base de datos";
  private static final Pair<Integer, Integer> managerLocation = Pair.create(2, 2);
  private static final Pair<Integer, Integer> kamLocation = Pair.create(1, 2);
  private static final int HORIZONTAL_FIELDS_SIZE = 7;
  private static final int VERTICAL_FIELDS_SIZE = 13;
  private static final int ACCOUNTS_ROW_INDEX_START = 5;
  private static final int ACCOUNTS_COLUMN_INDEX = 1;
  private static final int PRIORITIZED_COLUMN_INDEX = 15;
  private static final int VERTICAL_FIELDS_ROW_INDEX = 4;
  private static final int VERTICAL_FIELDS_COLUMN_INDEX_START = 3;

  private static final String EMPTY_VALUE = "-";

  public List<DBRecord> readAccountsData(InputStream inputStream) throws IOException {
    Workbook workbook = new XSSFWorkbook(inputStream);
    List<Sheet> kamSheets = findKAMSheets(workbook);
    List<DBRecord> dbRecords = new ArrayList<>();
    kamSheets.forEach(
        sheet -> {
        log.info("Processing sheet: {}", sheet.getSheetName());
          String manager =
              sheet
                  .getRow(managerLocation.getFirst())
                  .getCell(managerLocation.getSecond())
                  .getStringCellValue();
          String kam =
              sheet
                  .getRow(kamLocation.getFirst())
                  .getCell(kamLocation.getSecond())
                  .getStringCellValue();

          List<Cell> verticalFieldsCells = findVerticalFields(sheet);
          // accounts
          for (int i = ACCOUNTS_ROW_INDEX_START;
              i <= sheet.getLastRowNum();
              i += HORIZONTAL_FIELDS_SIZE) {
            Row accountRow = sheet.getRow(i);
            if (accountRow == null) {
              continue;
            }
            Cell accountCell = accountRow.getCell(ACCOUNTS_COLUMN_INDEX);
            if (accountCell == null) {
              continue;
            }
            String account = accountCell.getStringCellValue();

            Cell prioritizedCell = accountRow.getCell(PRIORITIZED_COLUMN_INDEX);
            if (prioritizedCell == null) {
              continue;
            }


            // vertical vs horizontal fields
            for (Cell verticalFieldCell : verticalFieldsCells) {
              String verticalField = verticalFieldCell.getStringCellValue();
              DBRecord dbRecord =
                  DBRecord.builder()
                      .manager(manager)
                      .kam(kam)
                      .account(account)
                      .product(verticalField)
                      .extraFields(new LinkedHashMap<>())
                      .build();

              IntStream.range(i, i + HORIZONTAL_FIELDS_SIZE)
                  .forEach(
                      fieldRowIndex -> {
                        String horizontalField =
                            sheet.getRow(fieldRowIndex).getCell(2).getStringCellValue();
                        Cell fieldCell =
                            sheet.getRow(fieldRowIndex).getCell(verticalFieldCell.getColumnIndex());

                        if (fieldCell == null) {
                          return;
                        }

                        ExtraFieldValue fieldValue = mapToExtraFieldValue(fieldCell);
                        dbRecord.getExtraFields().put(horizontalField, fieldValue);
                      });
              dbRecord.getExtraFields().put("Priorizado", mapToExtraFieldValue(prioritizedCell));
              dbRecords.add(dbRecord);
            }
          }
        });
        log.info("Database records size: {}", dbRecords.size());
    return dbRecords;
  }

  private List<Cell> findVerticalFields(Sheet sheet) {
    Row row = sheet.getRow(VERTICAL_FIELDS_ROW_INDEX);
    // VERTICAL_FIELDS_SIZE - 1, because we ignore the last vertical field ("Priorizado") as a product
    return IntStream.range(VERTICAL_FIELDS_COLUMN_INDEX_START, VERTICAL_FIELDS_COLUMN_INDEX_START + (VERTICAL_FIELDS_SIZE - 1))
        .mapToObj(row::getCell)
        .filter(Objects::nonNull)
        .collect(Collectors.toList());
  }

  private List<Sheet> findKAMSheets(Workbook workbook) {
    int dbSheetIndex = workbook.getSheetIndex(DB_SHEET_NAME);

    return IntStream.range(dbSheetIndex + 1, workbook.getNumberOfSheets())
        .mapToObj(workbook::getSheetAt)
        .collect(Collectors.toList());
  }

  private ExtraFieldValue mapToExtraFieldValue(Cell cell) {
    ExtraFieldValue extraFieldValue = ExtraFieldValue.builder().build();
    extraFieldValue.setCellType(cell.getCellType());
    switch (cell.getCellType()) {
      case STRING, BLANK -> extraFieldValue.setValue(cell.getStringCellValue());
      case NUMERIC -> {
        boolean cellDateFormatted = DateUtil.isCellDateFormatted(cell);
        extraFieldValue.setCellDateFormatted(cellDateFormatted);
        if (cellDateFormatted) {
          extraFieldValue.setValue(cell.getDateCellValue());
        } else {
          extraFieldValue.setValue(cell.getNumericCellValue());
        }
      }
      case BOOLEAN -> extraFieldValue.setValue(cell.getBooleanCellValue());
      case FORMULA -> {
        extraFieldValue.setCachedFormulaResultType(cell.getCachedFormulaResultType());
        if (CellType.NUMERIC.equals(cell.getCachedFormulaResultType())) {
          extraFieldValue.setValue(cell.getNumericCellValue());
        } else if (CellType.ERROR.equals(cell.getCachedFormulaResultType())) {
          extraFieldValue.setValue(EMPTY_VALUE);
        } else {
          extraFieldValue.setValue(cell.getStringCellValue());
        }
      }
      default -> throw new UnsupportedOperationException(
          String.format("Cell type %s is not supported", cell.getCellType()));
    }

    return extraFieldValue;
  }
}
