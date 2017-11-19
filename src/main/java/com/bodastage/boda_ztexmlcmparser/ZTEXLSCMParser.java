/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.bodastage.boda_ztexmlcmparser;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Stack;
import javax.xml.stream.XMLStreamException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
/**
 *
 * @author Emmanuel
 */
public class ZTEXLSCMParser {
    
    /**
     * The base file name of the file being parsed.
     * 
     * @since 1.0.0
     */
    private String baseFileName = "";
    
    /**
     * Output directory
     * 
     * @since 1.0.0
     */
    private String outputDirectory;
    
    /**
     * Export date
     * 
     * @since 1.0.0
     */
    private String dateTime = "";
    
    /**
     * Name of file being processed
     * 
     * @since 1.0.0
     */
    private String fileName;
    
    /**
     * The file to be parsed.
     * 
     * @since 1.0.0
     */
    private String dataFile;
    
    /**
     * Set data source
     * 
     * @since 1.0.0
     */
    private String dataSource;
    
    
    /**
     * The holds the parameters and corresponding values for the moi tag  
     * currently being processed.
     * 
     * @since 1.0.0
     */
    private Map<String,String> moiParameterValueMap 
            = new LinkedHashMap<String, String>();
    
    
    /**
     * Parser states. Currently there are only 2: extraction and parsing
     * 
     * @since 1.1.0
     */
    private int parserState = ParserStates.EXTRACTING_PARAMETERS;
    
    
        /**
     * Tracks Managed Object attributes to write to file. This is dictated by 
     * the first instance of the MO found. 
     *
     * @since 1.0.0
     */
    private Map<String, Stack> moColumns = new LinkedHashMap<String, Stack>();
    
    
    /**
     * Tracks Managed Object key attributes. _id is appended to the parimary key 
     * parameters in the headers of the csv
     *
     * @since 1.0.0
     */
    private Map<String, Stack> moKeyColumns = new LinkedHashMap<String, Stack>();
            
    /**
     * This holds a map of the Managed Object Instances (MOIs) to the respective
     * csv print writers.
     * 
     * @since 1.0.0
     */
    private Map<String, PrintWriter> moiPrintWriters 
            = new LinkedHashMap<String, PrintWriter>();
    
    public static void main( String[] args ) {
        
        try{
            String filename = args[0];
            String outputDirectory = args[1];

            ZTEXLSCMParser parser = new ZTEXLSCMParser();
            parser.setDataSource(args[0]);
            parser.setFileName(args[0]);
            parser.setOuputDirectory(args[1]);
            parser.parse();
        }catch(Exception e){
            System.out.println(e.getMessage());
        }
    }
    
    
    
    /**
     * Get file base name.
     * 
     * @since 1.0.0
     */
     public String getFileBasename(String filename){
        try{
            return new File(filename).getName();
        }catch(Exception e ){
            return filename;
        }
    }
     
    
    /**
     * Parser entry point 
     * 
     * @since 1.0.0
     * @version 1.1.0
     * 
     * @throws XMLStreamException
     * @throws FileNotFoundException
     * @throws UnsupportedEncodingException 
     */
    public void parse() throws XMLStreamException, FileNotFoundException, UnsupportedEncodingException, IOException, InvalidFormatException {
        //Extract parameters
        if (parserState == ParserStates.EXTRACTING_PARAMETERS) {
            processFileOrDirectory();

            parserState = ParserStates.EXTRACTING_VALUES;
        }

        //Extracting values
        if (parserState == ParserStates.EXTRACTING_VALUES) {
            processFileOrDirectory();
            parserState = ParserStates.EXTRACTING_DONE;
        }
        
        closeMOPWMap();
    }
    
    /**
     * Close file print writers.
     *
     * @since 1.0.0
     * @version 1.0.0
     */
    public void closeMOPWMap() {
        Iterator<Map.Entry<String, PrintWriter>> iter
                = moiPrintWriters.entrySet().iterator();
        while (iter.hasNext()) {
            iter.next().getValue().close();
        }
        moiPrintWriters.clear();
    }
    
    /**
     * Set name of file to parser.
     * 
     * @since 1.0.0
     * @version 1.0.0
     * @param directoryName 
     */
    public void setFileName(String filename ){
        this.dataFile = filename;
    }
    
    /**
     * Determines if the source data file is a regular file or a directory and 
     * parses it accordingly
     * 
     * @since 1.1.0
     * @version 1.0.0
     * @throws XMLStreamException
     * @throws FileNotFoundException
     * @throws UnsupportedEncodingException
     */
    public void processFileOrDirectory()
            throws XMLStreamException, FileNotFoundException, UnsupportedEncodingException, IOException, InvalidFormatException {
        
        //this.dataFILe;
        Path file = Paths.get(this.dataSource);
        boolean isRegularExecutableFile = Files.isRegularFile(file)
                & Files.isReadable(file);

        boolean isReadableDirectory = Files.isDirectory(file)
                & Files.isReadable(file);

        if (isRegularExecutableFile) {
            this.setFileName(this.dataSource);
            baseFileName =  getFileBasename(this.dataFile);
            if( parserState == ParserStates.EXTRACTING_PARAMETERS){
                System.out.print("Extracting parameters from " + this.baseFileName + "...");
            }else{
                System.out.print("Parsing " + this.baseFileName + "...");
            }
            this.parseFile(this.dataSource);
            
            if( parserState == ParserStates.EXTRACTING_PARAMETERS){
                 System.out.println("Done.");
            }else{
                System.out.println("Done.");
                //System.out.println(this.baseFileName + " successfully parsed.\n");
            }
        }

        if (isReadableDirectory) {

            File directory = new File(this.dataSource);

            //get all the files from a directory
            File[] fList = directory.listFiles();

            for (File f : fList) {
                this.setFileName(f.getAbsolutePath());
                try {
                    
                    //@TODO: Duplicate call in parseFile. Remove!
                    baseFileName =  getFileBasename(this.dataFile);
                    if( parserState == ParserStates.EXTRACTING_PARAMETERS){
                        System.out.print("Extracting parameters from " + this.baseFileName + "...");
                    }else{
                        System.out.print("Parsing " + this.baseFileName + "...");
                    }
                    
                    //Parse
                    this.parseFile(f.getAbsolutePath());
                    if( parserState == ParserStates.EXTRACTING_PARAMETERS){
                         System.out.println("Done.");
                    }else{
                        System.out.println("Done.");
                        //System.out.println(this.baseFileName + " successfully parsed.\n");
                    }
                   
                } catch (Exception e) {
                    System.out.println(e.getMessage());
                    System.out.println("Skipping file: " + this.baseFileName + "\n");
                }
            }
        }

    }
    
    public void parseFile(String fileName ) throws FileNotFoundException, IOException, InvalidFormatException{
        Workbook wb = WorkbookFactory.create(new File(fileName));
        Sheet sheet = wb.getSheetAt(1);
         
        int rowCount = 0;
        for (Row row : sheet) {
            rowCount++ ;
            
            //Skip first row of headers 
            if(rowCount == 1 ) continue;
            
            Cell cell = row.getCell(1);
            String moName = cell.getStringCellValue();

            Sheet moSheet = wb.getSheet(moName);

            Stack<String> parameters = new Stack();
            Stack<String> keyParameters = new Stack();
            
            String pNameStr = "FileName";
            String pValueStr   = baseFileName;

            if( moColumns.containsKey(moName)){
                parameters = moColumns.get(moName);
                keyParameters = moKeyColumns.get(moName);
            }
            
            if( parserState == ParserStates.EXTRACTING_VALUES && !moiPrintWriters.containsKey(moName)){
                String moiFile = outputDirectory + File.separatorChar + moName + ".csv";

                moiPrintWriters.put(moName, new PrintWriter(moiFile));
                
                Stack moParameterList = moColumns.get(moName);
                Stack moKeyParameterList = moKeyColumns.get(moName);

                for(int i =0; i < moParameterList.size(); i++ ){
                    String p = moParameterList.get(i).toString();

                    //Append _id to Primary key parameters
                    if( moKeyParameterList.contains(p)) { 
                        if(!p.equals("MEID")) p = p + "_id";
                    }
                    pNameStr += "," + p;
                }
                moiPrintWriters.get(moName).println(pNameStr);

            }
            
            //Parameters in the sheet
            Stack<String> sheetParams = new Stack();  
            Stack<String> sheetKeyParams = new Stack();           
            
            int sheetRowCount = 0;
            for (Row sheetRow : moSheet) {
                ++sheetRowCount;
                
                //Do nothing if we are on rows
                if(sheetRowCount >= 2 && sheetRowCount <= 4 ){
                    continue;
                }
                
                if(sheetRowCount == 5 && parserState != ParserStates.EXTRACTING_PARAMETERS){
                    continue;
                }
                
                //Get values from each row
                Stack<String> sheetParamValues = new Stack(); 
                
                int rCount = 0; //cell horizontal count per row
                for(Cell sheetRowCell: sheetRow){
                    ++rCount;
                    
                    String cellValue = sheetRowCell.getStringCellValue();
                    
                    //Exrtract parameters
                    if( sheetRowCount == 1 && parserState == ParserStates.EXTRACTING_PARAMETERS){
                        if(!parameters.contains(cellValue)){ 
                            parameters.add(cellValue);
                        }
                        continue;
                    }
                    
                    //Get key parameters
                    if( sheetRowCount == 5 && parserState == ParserStates.EXTRACTING_PARAMETERS ){
                        if(cellValue.equals("Primary Key")){
                            //sheetKeyParams.add( sheetParams.get(rCount-1));
                            String kParam = parameters.get(rCount-1);
                            if(!keyParameters.contains(kParam)){
                                keyParameters.add(kParam);
                            }
                        }
                        
                        continue;
                    }
            
                    if( sheetRowCount == 1 && parserState == ParserStates.EXTRACTING_VALUES){
                               
                        String parameterName = cellValue;
                        //Add parameter name
                        sheetParams.add(parameterName);
                        //pNameStr += "," + parameterName
                        
                        continue;
                    }
                    
                    //Else for rows > 5
                    if(sheetRowCount>5 && parserState == ParserStates.EXTRACTING_VALUES ){
                        //pValueStr += "," + toCSVFormat(cellValue);
                        sheetParamValues.add(cellValue);

                    }
                    

                    
                }

                
                //Write values
                if(sheetRowCount>5 && parserState == ParserStates.EXTRACTING_VALUES){
                
                    pNameStr = "FileName";
                    pValueStr   = baseFileName;
                    
                    Stack pList = moColumns.get(moName);
                    for(int i =0; i < pList.size(); i++){
                        String p = pList.get(i).toString();

                        int pIndex = sheetParams.indexOf(p);
                        
                        //System.out.println("pIndex:" + pIndex);
                        
                        String value = sheetParamValues.get(pIndex);
                        //System.out.println("value:" + value);
                        
                        pValueStr += "," + toCSVFormat(value);
                    }

                    moiPrintWriters.get(moName).println(pValueStr);
                    sheetParamValues.clear();
                    continue;

                }
            }

            if( parserState == ParserStates.EXTRACTING_PARAMETERS){
                moColumns.put(moName, parameters);
                moKeyColumns.put(moName, keyParameters);
                //parameters.clear();
                //keyParameters.clear();
                continue;
            }

        }
    }
    
    public void setOuputDirectory(String directoryName ){
        outputDirectory = directoryName;
    }
    
    /**
     * Set the data source 
     * 
     * @since 1.0.0
     * 
     * @param dataSource 
     */
    public void setDataSource(String dataSource){
        this.dataSource = dataSource;
    }
    
    /**
     * Process given string into a format acceptable for CSV format.
     *
     * @since 1.0.0
     * @param s String
     * @return String Formated version of input string
     */
    public String toCSVFormat(String s) {
        String csvValue = s;

        //Check if value contains comma
        if (s.contains(",")) {
            csvValue = "\"" + s + "\"";
        }

        if (s.contains("\"")) {
            csvValue = "\"" + s.replace("\"", "\"\"") + "\"";
        }

        return csvValue;
    }
    
}
