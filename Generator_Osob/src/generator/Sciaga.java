package generator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;




public class Sciaga {

	public static final String SAMPLE_XLSX_FILE_PATH = "./sample.xlsx";
	
	

	
	public static void main(String[] args) throws IOException, InvalidFormatException {
		
		File file=new File(SAMPLE_XLSX_FILE_PATH);
		Workbook workbook = WorkbookFactory.create(file);
		Sheet sheet=workbook.getSheetAt(0);
		DataFormatter dataFormatter = new DataFormatter();
		for (Row row: sheet) {
            for(Cell cell: row) {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }
	
		workbook.close();
		
		String[] columns = {"imie","wiek"};
		List<Osoba> osoby =  new ArrayList<>();
		osoby.add(new Osoba("Olaf",21));
		osoby.add(new Osoba("Oskar",18));
		Workbook workbook2 = new XSSFWorkbook();
		Sheet sheet2=workbook2.createSheet("Osoby");
		
		Font headerFont = workbook2.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());
        
        CellStyle headerCellStyle = workbook2.createCellStyle();
        headerCellStyle.setFont(headerFont);
		
        Row headerRow = sheet2.createRow(0);

        // Create cells
        for(int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }
        
        int rowNum = 1;
        for(Osoba osoba:osoby){
        	Row row=sheet2.createRow(rowNum++);
        	row.createCell(0).setCellValue(osoba.imie);
        	row.createCell(1).setCellValue(osoba.wiek);
        }
        
        for(int i = 0; i < columns.length; i++) {
            sheet2.autoSizeColumn(i);
        }
        
        FileOutputStream fileOut = new FileOutputStream("./samplewrite.xlsx",true);
        workbook2.write(fileOut);
        fileOut.close();

        
        workbook2.close();
	}

}
