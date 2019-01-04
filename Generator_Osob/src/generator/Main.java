package generator;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;




public class Main {

	public static final String SAMPLE_XLSX_FILE_PATH = "./sample.xlsx";
	
	

	
	public static void main(String[] args) throws IOException, InvalidFormatException {
		
	
	Generator generator=new Generator();
	generator.read(SAMPLE_XLSX_FILE_PATH);
	System.out.println(generator);
	generator.randomize();
	System.out.println(generator);
	generator.write("./samplewrite.xlsx");
		
}}
