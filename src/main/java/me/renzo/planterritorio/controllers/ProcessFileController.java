package me.renzo.planterritorio.controllers;

import java.io.IOException;
import java.util.List;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import me.renzo.planterritorio.model.DBRecord;
import me.renzo.planterritorio.service.FileReaderService;
import me.renzo.planterritorio.service.FileWriterService;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("/process")
@Slf4j
@RequiredArgsConstructor
public class ProcessFileController {

  private final FileReaderService fileReaderService;
  private final FileWriterService fileWriterService;

  @CrossOrigin(
      origins = {"http://localhost:3000", "https://plan-territorio-web.vercel.app"},
      exposedHeaders = HttpHeaders.CONTENT_DISPOSITION)
  @PostMapping(
      path = "",
      consumes = MediaType.MULTIPART_FORM_DATA_VALUE,
      produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
  public ResponseEntity<Resource> uploadFile(@RequestPart("dataFile") MultipartFile multipartFile)
      throws IOException {
    List<DBRecord> dbRecords = fileReaderService.readAccountsData(multipartFile.getInputStream());

    byte[] databaseFile = fileWriterService.createDatabaseFile(dbRecords);

    ByteArrayResource resource = new ByteArrayResource(databaseFile);
    String originalFileNameWithoutExtension = multipartFile.getOriginalFilename().split("\\.")[0];
    String fileName = String.format("%s_BD.xlsx", originalFileNameWithoutExtension);
    return ResponseEntity.ok()
        .header(
            HttpHeaders.CONTENT_DISPOSITION, String.format("attachment; filename=\"%s\"", fileName))
        .body(resource);
  }
}
