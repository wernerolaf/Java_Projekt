package generator;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.helpers.ColumnHelper;



import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Collections;

public class Generator {
	
	public List<String> headers=new ArrayList<>();
	public List<List<String>> cells=new ArrayList<>();;
	
	
	
	@Override
	public String toString() {
		return "Generator [headers=" + headers + ", cells=" + cells + "]";
	}

	public void read(String Path) throws IOException, InvalidFormatException {
		
		File file=new File(Path);
		List<String> temp=new LinkedList<String>();
		Workbook workbook = WorkbookFactory.create(file);
		Sheet sheet=workbook.getSheetAt(0);
		DataFormatter dataFormatter = new DataFormatter();
		Iterator<Row> rowIterator = sheet.rowIterator();
        
		
		if(rowIterator.hasNext()){
			Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String cellValue = dataFormatter.formatCellValue(cell);
                headers.add(cellValue);
            }
		}
		
		while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String cellValue = dataFormatter.formatCellValue(cell);
                temp.add(cellValue);
            }
            cells.add(new LinkedList<String>(temp));
            temp.clear();
        }
		workbook.close();
	}
		
	public void write(String Path) throws IOException, InvalidFormatException {
		
		
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet=workbook.createSheet();
		
		Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());
        
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
		
        Row headerRow = sheet.createRow(0);

        // Create cells
        for(int i = 0; i < headers.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
            cell.setCellStyle(headerCellStyle);
        }
        
        int rowNum = 1;
        for(List<String> row:cells){
        	Row xrow=sheet.createRow(rowNum++);
        	 for(int i = 0; i < row.size(); i++) {
                 Cell cell = xrow.createCell(i);
                 cell.setCellValue(row.get(i));
                 
             }
        }
        
        for(int i = 0; i < headers.size(); i++) {
            sheet.autoSizeColumn(i);
        }
        
        FileOutputStream fileOut = new FileOutputStream(Path,true);
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
	}

	public void randomize(){
		
	        List<List<String>> ret = new ArrayList<List<String>>();
	        int N = cells.get(0).size();
	        for (int i = 0; i < N; i++) {
	            List<String> col = new ArrayList<String>();
	            for (List<String> row : cells) {
	                col.add(row.get(i));
	            }
	            ret.add(col);
	        }
	        for(List<String> list:ret){
	        	Collections.shuffle(list);
	        }
	        
	        cells.clear();
	        N = ret.get(0).size();
	        for (int i = 0; i < N; i++) {
	            List<String> col = new ArrayList<String>();
	            for (List<String> row : ret) {
	                col.add(row.get(i));
	            }
	            cells.add(col);
	        }
	        Collections.shuffle(cells);
	}
		
		
	}

	

