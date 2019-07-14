/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.bodastage.boda_ztexmlcmparser;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
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
import org.apache.commons.cli.Options;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.ParseException;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.poi.ss.usermodel.CellType;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
/**
 *
 * @author Emmanuel
 */
public class ZTEXLSCMParser {
    /**
     * Current release version 
     * 
     * Since 1.3.0
     */
    final static String VERSION = "1.2.1";
    
    private static final Logger LOGGER = LoggerFactory.getLogger(ZTEXLSCMParser.class);
    
    public ZTEXLSCMParser() {
        DateFormat dateFormat  = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        Date date = new Date();
        dateTime = dateFormat.format(date);
    }
    
    
    
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
    
    public void setOutputDirectory(String directoryName) {
        this.outputDirectory = directoryName;
    }
    
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
     * Parser start time.
     *
     * @since 1.1.0
     * @version 1.1.0
     */
    final long startTime = System.currentTimeMillis();
    
    /**
     * Tracks Managed Object key attributes. _id is appended to the primary key 
     * parameters in the headers of the csv
     *
     * @since 1.0.0
     */
    private Map<String, Stack> moKeyColumns = new LinkedHashMap<String, Stack>();
            
    private String neType = "";
    private String templateType = "";
    private String templateVersion = "";
    private String dataType = "";
    
    /**
     * This holds a map of the Managed Object Instances (MOIs) to the respective
     * csv print writers.
     * 
     * @since 1.0.0
     */
    private Map<String, PrintWriter> moiPrintWriters 
            = new LinkedHashMap<String, PrintWriter>();
    
    public static void main( String[] args ) {
        
        

       //Define
       Options options = new Options();
       CommandLine cmd = null;
       String outputDirectory = null;   
       String inputFile = null;
       String parameterConfigFile = null;
       Boolean onlyExtractParameters = false;
       Boolean showHelpMessage = false;
       Boolean showVersion = false;
       Boolean attachMetaFields = false; //Attach mattachMetaFields FILENAME,DATETIME,NE_TECHNOLOGY,NE_VENDOR,NE_VERSION,NE_TYPE
       
       try{ 
            //options.addOption( "p", "extract-parameters", false, "extract only the managed objects and parameters" );
            options.addOption( "v", "version", false, "display version" );
            //options.addOption( "m", "meta-fields", false, "add meta fields to extracted parameters. FILENAME,DATETIME" );
            options.addOption( Option.builder("i")
                    .longOpt( "input-file" )
                    .desc( "input file or directory name")
                    .hasArg()
                    .argName( "INPUT_FILE" ).build());
            options.addOption(Option.builder("o")
                    .longOpt( "output-directory" )
                    .desc( "output directory name")
                    .hasArg()
                    .argName( "OUTPUT_DIRECTORY" ).build());
            options.addOption(Option.builder("c")
                    .longOpt( "parameter-config" )
                    .desc( "parameter configuration file")
                    .hasArg()
                    .argName( "PARAMETER_CONFIG" ).build() );
            options.addOption( "h", "help", false, "show help" );
            
            //Parse command line arguments
            CommandLineParser parser = new DefaultParser();
            cmd = parser.parse( options, args);

            if( cmd.hasOption("h")){
                showHelpMessage = true;
            }

            if( cmd.hasOption("v")){
                showVersion = true;
            }
            
            if(cmd.hasOption('o')){
                outputDirectory = cmd.getOptionValue("o"); 
            }
            
            if(cmd.hasOption('i')){
                inputFile = cmd.getOptionValue("i"); 
            }
            
            if(cmd.hasOption('c')){
                parameterConfigFile = cmd.getOptionValue("c"); 
            }
            
            if(cmd.hasOption('p')){
                onlyExtractParameters  = true;
            }
            
            if(cmd.hasOption('m')){
                attachMetaFields  = true;
            }
            
       }catch(IllegalArgumentException e){
           
       } catch (ParseException ex) {
//            java.util.logging.Logger.getLogger(HuaweiCMObjectParser.class.getName()).log(Level.SEVERE, null, ex);
        }

        
       

        try{
            
            if(showVersion == true ){
                System.out.println(VERSION);
                System.out.println("Copyright (c) 2019 Bodastage Solutions(http://www.bodastage.com)");
                System.exit(0);
            }
            
            //show help
            if( showHelpMessage == true || 
                inputFile == null || 
                ( outputDirectory == null && onlyExtractParameters == false) ){
                     HelpFormatter formatter = new HelpFormatter();
                     String header = "Parses ZTE template excel document to csv\n\n";
                     String footer = "\n";
                     footer += "Examples: \n";
                     footer += "java -jar boda-ztexlscmparser.jar -i workbook.xls -o out_folder\n";
                     footer += "java -jar boda-ztexlscmparser.jar -i input_folder -o out_folder\n";
                     footer += "\nCopyright (c) 2019 Bodastage Solutions(http://www.bodastage.com)";
                     formatter.printHelp( "java -jar boda-ztexlscmparser.jar", header, options, footer );
                     System.exit(0);
            }
        
            //Confirm that the output directory is a directory and has write 
            //privileges
            if(outputDirectory != null ){
                File fOutputDir = new File(outputDirectory);
                if (!fOutputDir.isDirectory()) {
                    System.err.println("ERROR: The specified output directory is not a directory!.");
                    System.exit(1);
                }

                if (!fOutputDir.canWrite()) {
                    System.err.println("ERROR: Cannot write to output directory!");
                    System.exit(1);
                }
            }
            
            

            //Get parser instance
            ZTEXLSCMParser cmParser = new ZTEXLSCMParser();
            
            if(  parameterConfigFile != null ){
                File f = new File(parameterConfigFile);
                if(f.isFile()){
                    cmParser.setParameterFile(parameterConfigFile);
                    cmParser.getParametersToExtract(parameterConfigFile);
                    cmParser.parserState = ParserStates.EXTRACTING_VALUES;
                }
            }
            
            cmParser.setDataSource(inputFile);
            if(outputDirectory != null ) cmParser.setOutputDirectory(outputDirectory);
            
            cmParser.parse();
        }catch(Exception e){
            System.out.println(e.getMessage());
            System.exit(1);
        }
    }
    
    /**
     * Show parser help.
     * 
     * @since 1.0.0
     * @version 1.0.0
     */
    static public void showHelp(){
        System.out.println("boda-ztexlscmparser " + VERSION + " Copyright (c) 2019 Bodastage(http://www.bodastage.com)");
        System.out.println("Parses ZTE CM Dumps from Netnumen in excel to csv.");
        System.out.println("Usage: java -jar boda-ztexlscmparser.jar -i <fileToParse.xls> -o <outputDirectory> [parameterFile]");
    }
    
    /**
     * File containing a list of parameters to export
     * 
     * @since 1.2.0
     */
    private String parameterFile = null;

    /**
     * Set the parameter file name 
     * 
     * @param filename 
     */
    public void setParameterFile(String filename){
        parameterFile = filename;
    }
    
  /**
     * Extract parameter list from  parameter file
     * 
     * @param filename 
     */
    public  void getParametersToExtract(String filename) throws FileNotFoundException, IOException{
        BufferedReader br = new BufferedReader(new FileReader(filename));
        for(String line; (line = br.readLine()) != null; ) {
           String [] moAndParameters =  line.split(":");
           String mo = moAndParameters[0];
           String [] parameters = moAndParameters[1].split(",");
           
           Stack parameterStack = new Stack();
           Stack keyParameterStack = new Stack();
           for(int i =0; i < parameters.length; i++){
               if( parameters[i].endsWith("_id")){
                   String p = parameters[i].replace("_id", "");
                   parameterStack.push(p);
                   keyParameterStack.push(p);
               }else{
                   parameterStack.push(parameters[i]);
               }
               
           }
           
           moColumns.put(mo, parameterStack);
           moKeyColumns.put(mo, keyParameterStack);

        }
        
        //Move to the parameter value extraction stage
        parserState = ParserStates.EXTRACTING_VALUES;
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
        
        printExecutionTime();
    }
    
    /**
     * Print program's execution time.
     * 
     * @since 1.0.0
     */
    public void printExecutionTime(){
        float runningTime = System.currentTimeMillis() - startTime;
        
        String s = "Parsing completed. ";
        s = s + "Total time:";
        
        //Get hours
        if( runningTime > 1000*60*60 ){
            int hrs = (int) Math.floor(runningTime/(1000*60*60));
            s = s + hrs + " hours ";
            runningTime = runningTime - (hrs*1000*60*60);
        }
        
        //Get minutes
        if(runningTime > 1000*60){
            int mins = (int) Math.floor(runningTime/(1000*60));
            s = s + mins + " minutes ";
            runningTime = runningTime - (mins*1000*60);
        }
        
        //Get seconds
        if(runningTime > 1000){
            int secs = (int) Math.floor(runningTime/(1000));
            s = s + secs + " seconds ";
            runningTime = runningTime - (secs/1000);
        }
        
        //Get milliseconds
        if(runningTime > 0 ){
            int msecs = (int) Math.floor(runningTime/(1000));
            s = s + msecs + " milliseconds ";
            runningTime = runningTime - (msecs/1000);
        }

        System.out.println(s);
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
        
        //Handle sheet 0 i.e. first sheet
        Sheet templateInfoSheet = wb.getSheetAt(0);
        for (Row row : templateInfoSheet) {
            String key = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue();
            String value = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue();
            
                     
            if(key.equals("NE Type:")){
                neType = value;
            }
            
            if(key.equals("Template Type:")){
                templateType = value;
            }
            
            if(key.equals("Template Version:")){
                templateVersion = value;
            }
            
            if(key.equals("Data Type:")){
                dataType = value;
            }
        }
        
        //Get the MOs from the 2nd sheet which has the MO list
        Sheet sheet = wb.getSheetAt(1);
                
        int rowCount = 0;
        for (Row row : sheet) {
            rowCount++ ;
            
            //Skip first row of headers on sheet 2
            if(rowCount == 1 ) continue;
                        
            //Get MO name from 2nd column of sheet i.e. index 1
            Cell cell = row.getCell(1,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            String moName = cell.getStringCellValue(); 
            if(moName.replace(" ", "").length()==0)continue;
                       
            //Skip MOs not in parameter file
            if(parameterFile != null && !moColumns.containsKey(moName)) continue;
            
            //Get MO Sheet
            Sheet moSheet = wb.getSheet(moName);
                        
            Stack<String> parameters = new Stack();
            Stack<String> keyParameters = new Stack();

            if( moColumns.containsKey(moName)){
                parameters = moColumns.get(moName);
                keyParameters = moKeyColumns.get(moName);
            }
            
            if( parserState == ParserStates.EXTRACTING_VALUES && !moiPrintWriters.containsKey(moName)){
                String moiFile = outputDirectory + File.separatorChar + moName + ".csv";

                moiPrintWriters.put(moName, new PrintWriter(moiFile));
                
                Stack moParameterList = moColumns.get(moName);
                Stack moKeyParameterList = moKeyColumns.get(moName);

                String pNameStr = "FileName,varDateTime,NeType,TemplateType,TemplateVersion,DataType";
                String pValueStr   = baseFileName + ","+ dateTime + "," + neType + 
                        "," + templateType + "," + templateVersion + 
                        "," + dataType ;
                
                for(int i =0; i < moParameterList.size(); i++ ){
                    String p = moParameterList.get(i).toString();

                    //Skip filename and vardatetime
                    if(parameterFile != null && 
                        ( p.toLowerCase().equals("filename") || 
                            p.toLowerCase().equals("vardatetime") || 
                            p.toLowerCase().equals("netype") || 
                            p.toLowerCase().equals("templatetype") || 
                            p.toLowerCase().equals("templateversion") || 
                            p.toLowerCase().equals("datatype") ) ){
                        continue;
                    }

                    //Append _id to Primary key parameters
                    if( moKeyParameterList.contains(p)) { 
                        if(!p.equals("MEID")) p = p + "_id";
                    }
                    pNameStr += "," + p;
                }
                moiPrintWriters.get(moName).println(pNameStr);

            }
            
            
            //Extract the parameters in the MO sheet
            Stack<String> sheetParams = new Stack();  
            Stack<String> sheetKeyParams = new Stack();           
            
            int sheetRowCount = 0;
            for (Row sheetRow : moSheet) {
                ++sheetRowCount;
                
                //Do nothing if we are on rows 2 to 4
                if(sheetRowCount >= 2 && sheetRowCount <= 4 ){
                    continue;
                }
                
                // Skip row 5 if we are not extracting parameters. This row as the 
                //parameter names 
                if(sheetRowCount == 5 && parserState != ParserStates.EXTRACTING_PARAMETERS){
                    continue;
                }
                
                //Get values from each row
                Stack<String> sheetParamValues = new Stack(); 
                
                //Iterate over the row cells
                for(int rCount=0; rCount < sheetRow.getLastCellNum(); rCount++) {
                    Cell sheetRowCell = sheetRow.getCell(rCount, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    
                    sheetRowCell.setCellType(CellType.STRING);
                    String cellValue = sheetRowCell.getStringCellValue();

                    //Exrtract parameters
                    if( sheetRowCount == 1 && parserState == ParserStates.EXTRACTING_PARAMETERS){
                        
                        //Ignore empty cells
                        if(!parameters.contains(cellValue) && cellValue.replace(" ", "").length() != 0){ 
                            parameters.add(cellValue);
                        }
                        continue;
                    }

                    
                    //Get key parameters
                    if( sheetRowCount == 5 && parserState == ParserStates.EXTRACTING_PARAMETERS ){
//                        if(cellValue.equals("Primary Key")){
//                            //sheetKeyParams.add( sheetParams.get(rCount-1));
//                            String kParam = parameters.get(rCount-1);
//                            if(!keyParameters.contains(kParam)){
//                                keyParameters.add(kParam);
//                            }
//                        }
//                        
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

                    String pNameStr = "FileName,varDateTime,NeType,TemplateType,TemplateVersion,DataType";
                    String pValueStr   = baseFileName + ","+ dateTime + "," + neType + 
                        "," + templateType + "," + templateVersion + 
                        "," + dataType ;
                    Stack pList = moColumns.get(moName);
                    for(int i =0; i < pList.size(); i++){
                            
                        String p = pList.get(i).toString();
                        
                        //Skip filename and vardatetime
                        if(parameterFile != null && 
                            ( p.toLowerCase().equals("filename") || 
                                p.toLowerCase().equals("vardatetime") || 
                                p.toLowerCase().equals("netype") || 
                                p.toLowerCase().equals("templatetype") || 
                                p.toLowerCase().equals("templateversion") || 
                                p.toLowerCase().equals("datatype") ) ){
                            continue;
                        }
                        
                        int pIndex = sheetParams.indexOf(p);

                        String value = "";
                        
                        if(pIndex < sheetParamValues.size() ) value = sheetParamValues.get(pIndex);
                        
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
