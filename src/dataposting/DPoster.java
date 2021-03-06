/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package dataposting;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import java.text.NumberFormat;

import java.util.Iterator;
import java.util.concurrent.TimeUnit;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.io.FileUtils;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.CellType;
import static org.apache.poi.ss.usermodel.CellType.STRING;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCell;

import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;

/**
 *
 * @author Aman Mambetov "asanbay"@mail.ru
 */
public class DPoster {

    private static final Logger LOG = Logger.getLogger(DPoster.class.getName());

    private final String sXlsxDocPathName;
    private final String sXlsxDocOutputPathName;
    
    private FileInputStream inputStream;
    private FileOutputStream outFile;
    private XSSFWorkbook book;
    private XSSFSheet sheet;
    
    private final String sSiteUrl;
    private Connection.Response resp;
    private final String sUserAgent;
    private int iNumberOfReplacements;
    
    public DPoster() {
        sXlsxDocPathName = "data\\NSW_Master.xlsx";                     // Source for IDs
        sXlsxDocOutputPathName = "data\\output.xlsx";                   // Result here
        sSiteUrl = "http://superfundlookup.gov.au/ABN/View?id=";        // Web link to source site
        sUserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.99 Safari/537.36 OPR/41.0.2353.69";
    }

    public void doTask() throws IOException {
    
        String sID;
        XSSFCell cell;
        NumberFormat nf = NumberFormat.getNumberInstance();
        
        String sStreetAdress;
        String sTypeOfCell = new String();

        
        try {
        // Read XSL file
        inputStream = new FileInputStream(
                new File(sXlsxDocPathName));
 
        // Get the workbook instance for XLS file
        book = new XSSFWorkbook(inputStream);
        
        } catch (FileNotFoundException ex) {
            LOG.log(Level.SEVERE, null, ex);
        } catch ( EncryptedDocumentException | IOException e) {
            LOG.log(Level.SEVERE, null, e);
        }
        
        sheet = book.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();
 
        iNumberOfReplacements = 0;
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if (row == null) continue;

            sTypeOfCell ="";
            cell = (XSSFCell) row.getCell(0);
            if(row.getCell(1) == null) {
                row.createCell(1, STRING);
            }
           
            CellType cellType = cell.getCellTypeEnum();
 
            switch (cellType) {
                case _NONE:
                    sTypeOfCell = "_NONE";
                    break;
                case BOOLEAN:
                    sTypeOfCell = "BOOLEAN";
                    break;
                case BLANK:
                    sTypeOfCell = "BLANK";
                    break;
                case FORMULA:
                    sTypeOfCell = "FORMULA";
                    break;
                case NUMERIC:
                    sID = nf.format(cell.getNumericCellValue());
                    sID = sID.replaceAll("\\D", "");  // remove spaces if any
                    System.out.print(sID);
                    
                    sStreetAdress = getFromWeb(sID);        // Get address from the web site using ID
                    
                    System.out.print("\t");
                    System.out.print(sStreetAdress);
                    System.out.print("\n");
                    row.getCell(1, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).setCellValue( sStreetAdress);
                    iNumberOfReplacements += 1;
                    break;
                case STRING:
                    sTypeOfCell = "STRING";
                    break;
            }            
        }


    inputStream.close();    
    outFile = new FileOutputStream(sXlsxDocOutputPathName);
    book.write(outFile);
    
    }

    
    
    private String getFromWeb(String sIDofPerson) throws IOException {
        Document dDoc;
        String sResult;
        String sGetUrl = sSiteUrl.concat(sIDofPerson);

        myDelay(2);                                 // This delay to be polite
        dDoc = Jsoup.connect(sGetUrl).get();
/*                resp = Jsoup.connect(sSiteUrlPost)
            .postDataCharset("Windows-1251")
            .userAgent(sUserAgent)
            .data("SearchParameters.SearchText", sIDofPerson)
            .data("DefaultSearchType", "ActiveFunds")
            .data("submit", "Search")
            .method(Connection.Method.POST)
            .execute();        
        dDoc = resp.parse();
*/

//        FileWrite(dDoc.toString(), "Debug out159.htm");   // Debug
        sResult = dDoc.select("span[itemprop]").text();
        sResult = sResult.replaceAll("\\<br\\>", "");
    return sResult;
    }

    
    
    
    private void myDelay (long interval) {
        try {
            TimeUnit.SECONDS.sleep(interval);
        }
        catch(InterruptedException e) {
            LOG.log(Level.SEVERE, null, e);
        }
    }
    
    private void FileWrite(String sBuff, String sFileName) {
        // This is for debug only
        final File f = new File(sFileName);
        try {
            FileUtils.writeStringToFile(f, sBuff, "UTF-8");
        }
        catch (IOException e) {
            LOG.log(Level.SEVERE, null, e);
        }
    }
    
    
}
