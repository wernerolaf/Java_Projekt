package org.jxls.reader;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.beanutils.ConvertUtilsBean;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Basic implementation of {@link XLSReader} interface
 *
 * @author Leonid Vysochyn
 */
public class XLSReaderImpl implements XLSReader{
    protected final Log log = LogFactory.getLog(getClass());

    Map<String, XLSSheetReader> sheetReaders = new HashMap<String, XLSSheetReader>();
    Map<Integer, XLSSheetReader> sheetReadersByIdx = new HashMap<Integer, XLSSheetReader>();


    XLSReadStatus readStatus = new XLSReadStatus();

    ConvertUtilsBean convertUtilsBean = ReaderConfig.createConvertUtilsBean( ReaderConfig.getInstance().isUseDefaultValuesForPrimitiveTypes() );

    public ConvertUtilsBeanProvider getConvertUtilsBeanProvider(){
    	return new ConvertUtilsBeanProviderDelegate( convertUtilsBean );
    }
    
    public XLSReadStatus read(InputStream inputXLS, Map beans) throws IOException, InvalidFormatException{
        readStatus.clear();
        Workbook workbook = WorkbookFactory.create(inputXLS);
        for(int sheetNo = 0; sheetNo < workbook.getNumberOfSheets(); sheetNo++){
            readStatus.mergeReadStatus(readSheet(workbook, sheetNo, beans));
        }
        workbook.close();
        return readStatus;
    }

    private XLSReadStatus readSheet(Workbook workbook, int sheetNo, Map beans){
        Sheet sheet = workbook.getSheetAt(sheetNo);
        String sheetName = workbook.getSheetName(sheetNo);
        if(log.isInfoEnabled()){
            log.info("Processing sheet " + sheetName);
        }
        XLSSheetReader sheetReader;
        if(sheetReaders.containsKey(sheetName)){
            sheetReader = sheetReaders.get(sheetName);
            return readSheet(sheetReader, sheet, sheetName, beans);
        } else if(sheetReadersByIdx.containsKey(sheetNo)){
            sheetReader = sheetReadersByIdx.get(sheetNo);
            return readSheet(sheetReader, sheet, sheetName, beans);
        } else{
            return null;
        }
    }

    private XLSReadStatus readSheet(XLSSheetReader sheetReader, Sheet sheet, String sheetName, Map beans){
        sheetReader.setSheetName(sheetName);
        return sheetReader.read(sheet, beans);
    }

    public Map getSheetReaders(){
        return sheetReaders;
    }

    public Map getSheetReadersByIdx(){
        return sheetReadersByIdx;
    }

    public void setSheetReadersByIdx(Map sheetReaders){
        sheetReadersByIdx = sheetReaders;
    }

    public void addSheetReader(String sheetName, XLSSheetReader reader){
        sheetReaders.put(sheetName, reader);
        reader.setConvertUtilsBeanProvider( getConvertUtilsBeanProvider() );
    }

    public void addSheetReader(Integer idx, XLSSheetReader reader){
        sheetReadersByIdx.put(idx, reader);
        reader.setConvertUtilsBeanProvider( getConvertUtilsBeanProvider() );
    }

    public void addSheetReader(XLSSheetReader reader){
        addSheetReader(reader.getSheetName(), reader);
        if( reader.getSheetIdx() >= 0 ){
            addSheetReader(reader.getSheetIdx(), reader);
        }
    }

    public void setSheetReaders(Map sheetReaders){
        this.sheetReaders = sheetReaders;
    }
}
