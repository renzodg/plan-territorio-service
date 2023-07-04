package me.renzo.planterritorio.model;

import lombok.Builder;
import lombok.Data;
import org.apache.poi.ss.usermodel.CellType;

@Data
@Builder
public class ExtraFieldValue {
  private Object value;
  private CellType cellType;
  private boolean cellDateFormatted;
  private CellType cachedFormulaResultType;
}
