package com.gligor.excel.excel.controllers;

import com.fasterxml.jackson.databind.exc.InvalidFormatException;
import com.gligor.excel.excel.logic.DeleteFile;
import com.gligor.excel.excel.logic.DownloadUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletContext;
import java.io.*;
import java.util.*;

@Controller
public class MainController {
    @Autowired
    private ServletContext servletContext;


    @RequestMapping({"/", ""})
    public String index(){
        return "index";
    }


    @RequestMapping(value = "/upload", method = RequestMethod.POST)
    public String uploadFile(@RequestParam("spreadsheet")MultipartFile spreadsheet,
                             @RequestParam("startCell") Integer startCell) throws Exception {
        String fileName = spreadsheet.getOriginalFilename();
        String filePath = "./uploads/" + fileName;

        File uploadedSpreadsheet = new File(filePath);
        if(uploadedSpreadsheet.exists() && !uploadedSpreadsheet.isDirectory()){
            uploadedSpreadsheet.delete();
        }

        try (OutputStream os = new FileOutputStream(uploadedSpreadsheet)) {
            os.write(spreadsheet.getBytes());
        }

        //cell locations
        final int SERVICE_NAME = 2;
        final int SERVICE_DIFFICULTY = 5;
        final int SPF_CELL = 6;
        final int DKIM_CELL = 7;

        XSSFWorkbook workbook = null;
        try {
            workbook = new XSSFWorkbook(uploadedSpreadsheet);
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (org.apache.poi.openxml4j.exceptions.InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }


        XSSFWorkbook outputFile = new XSSFWorkbook();
        Sheet outputSheet = outputFile.createSheet();
        Row header = outputSheet.createRow(0);

        //font of the header
        XSSFFont font = outputFile.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);

        CellStyle headerStyle = outputFile.createCellStyle();
        headerStyle.setFont(font);

        Cell cellToModifyWithTheServiceName = (header.createCell(0));
        cellToModifyWithTheServiceName.setCellValue("Service Name");
        cellToModifyWithTheServiceName.setCellStyle(headerStyle);
        Cell cellToModifyWithTheServiceDifficulty = (header.createCell(1));
        cellToModifyWithTheServiceDifficulty.setCellValue("Service Difficulty");
        cellToModifyWithTheServiceDifficulty.setCellStyle(headerStyle);

        Map<Integer, List<String>> data = new HashMap<>();
//        int i = 0;


        CellStyle cellStyle = outputFile.createCellStyle();
        font.setFontHeightInPoints((short) 12);
        cellStyle.setFont(font);
        font.setBold(false);

        //get the sheet of the uploaded file
        Sheet sheet = workbook.getSheetAt(0);


        for (int i = startCell - 1; i <= 1000; i++) {
            if (sheet.getRow(i).getCell(SERVICE_NAME) == null) {
                break;
            }
            //create a new row and set the name of the service
            Row outputRow = outputSheet.createRow(i - startCell + 2);
            Cell serviceNameCell = outputRow.createCell(0);
            serviceNameCell.setCellValue(sheet.getRow(i).getCell(2).getStringCellValue());
            serviceNameCell.setCellStyle(cellStyle);

            String serviceDifficulty = "";
            String spfCell = "";
            String dkimCell = "";

            //get the values of the cells needed for the calculations
            //various values return errors (numbers, formulas)
            try {   serviceDifficulty = sheet.getRow(i).getCell(SERVICE_DIFFICULTY).getStringCellValue();
            }catch (Exception e){
                e.printStackTrace();
            }
            try {spfCell = sheet.getRow(i).getCell(SPF_CELL).getStringCellValue();
            }catch (Exception e){
                e.printStackTrace();
            }
            try {dkimCell = sheet.getRow(i).getCell(DKIM_CELL).getStringCellValue();
            }catch (Exception e){
                e.printStackTrace();
            }

                //calculate the difficulty of the service
                Cell serviceNameDifficulty = outputRow.createCell(1);
                serviceNameDifficulty.setCellValue(calculateDifficulty(serviceDifficulty, spfCell, dkimCell));
                serviceNameDifficulty.setCellStyle(cellStyle);
            }

            //write output file to disk / send it to download
            String outputPath = "./uploads/" + "results_for_" + fileName;
            File result = new File(outputPath);
            FileOutputStream os = new FileOutputStream(result);
            outputFile.write(os);
            os.close();


        //delete the files
        DeleteFile deleteFile = new DeleteFile(filePath);
        DeleteFile deleteOutput = new DeleteFile(outputPath);
        deleteFile.start();
        deleteOutput.start();
            return "redirect:/download/" + "results_for_" + fileName;
        }


    @GetMapping(value = "/download/{filename}", produces = MediaType.ALL_VALUE)
    public ResponseEntity<InputStreamResource> download(@PathVariable String filename) throws IOException {

        MediaType mediaType = DownloadUtils.getMediaTypeForFileName(this.servletContext, filename);

        File file = new File("./uploads/" + filename);
        InputStreamResource resource = new InputStreamResource(new FileInputStream(file));

        return ResponseEntity.ok()
                // Content-Disposition
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment;filename=" + file.getName())
                // Content-Type
                .contentType(mediaType)
                // Contet-Length
                .contentLength(file.length()) //
                .body(resource);
    }



    //todo re-think the added values
    private double calculateDifficulty(String supports, String spfAligned, String dkimDocumented) throws Exception {

        int rank = 0;

        if(supports.equals("Check internal sources to make sure they are alignable and are passing DMARC")){
            rank+=5;
        }

        if(supports.equals("Supports Aligned SPF")){
            rank +=1;
            if(spfAligned.equals("yes")){
                rank+=1;
            }
            else if(spfAligned.equals("no")){
                rank+=3;
            }
        }

        if(supports.equals("Supports DKIM")){
            rank +=2;
            if(dkimDocumented.equals("yes")){
                rank+=2;
            }
            if(dkimDocumented.equals("no")){
                rank+=4;
            }
        }

        if(supports.equals("Supports Aligned SPF and DKIM")){
            rank+=3;
            //spf
            if(spfAligned.equals("yes")){
                rank+=1;
            }
            else if(spfAligned.equals("no")) {
                rank += 3;
            }
            //dkim
            if(dkimDocumented.equals("yes")){
                rank+=2;
            }
            else if(dkimDocumented.equals("no")){
                rank +=4;
            }
        }


        if(supports.equals("Not a configurable sender (Does not support aligned SPF or DKIM)")){
            rank=11;
        }

        if(supports.equals("No data found, please check manually for service data")){
            rank=12;
        }
        if(rank == 0){
            throw new Exception("check the spreadsheet");
        }
        return rank;
    }
}
