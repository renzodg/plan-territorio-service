package me.renzo.planterritorio.model;

import java.util.Map;
import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class DBRecord {
  private String manager;
  private String kam;
  private String account;
  private String product;
  private Map<String, ExtraFieldValue> extraFields;
}
