/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.bodastage.boda_ztexmlcmparser;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.Arrays;
import java.util.logging.Level;
import java.util.logging.Logger;
import junit.framework.Test;
import junit.framework.TestCase;
import junit.framework.TestSuite;
import org.slf4j.LoggerFactory;

/**
 *
 * @author ADMIN
 */
public class ZTEXLSCMParserTest extends TestCase {
    
    private static final org.slf4j.Logger LOGGER = LoggerFactory.getLogger(ZTEXLSCMParser.class);
    
    public void testGeneralParsing(){
    
        ClassLoader classLoader = getClass().getClassLoader();
        File inFile = new File(classLoader.getResource("templatedata.xlsx").getFile());
        String inputFile = inFile.getAbsolutePath();
        
        String outputFolder = System.getProperty("java.io.tmpdir");
        
        ZTEXLSCMParser parser = new ZTEXLSCMParser();
        
        String[] args = { "-i", inputFile, "-o", outputFolder};
        
        parser.main(args);
        
        String expectedResult [] = {
            "FileName,varDateTime,NeType,TemplateType,TemplateVersion,DataType,SomeMO1Param1,SomeMO1Param2,SomeMO1Param3",
            "templatedata.xlsx,YYYY-MM-DD HH:MI:SS,Multi-mode Controller,Plan,V0123,tech_radio,1,2,3",
            "templatedata.xlsx,YYYY-MM-DD HH:MI:SS,Multi-mode Controller,Plan,V0123,tech_radio,4,5,6"
        };
        
        try{
            String csvFile = outputFolder + File.separator + "SomeMO1.csv";
            
            BufferedReader br = new BufferedReader(new FileReader(csvFile)); 
            String csvResult [] = new String[3];
            
            int i = 0;
            String st; 
            while ((st = br.readLine()) != null) {
                //Repalce the date with YYYY-MM-DD HH:MI:SS as the parser generates 
                //as unique  datetime whenever it runs
                String c [] = st.split(",");
                c[1] = "YYYY-MM-DD HH:MI:SS";
                 
                csvResult[i] = "";
                for(int idx =0; idx < c.length; idx++){ 
                    if( idx > 0) csvResult[idx] += ",";
                    csvResult[i] += c[idx];
                }
                i++;
            }
            
            assertTrue(Arrays.equals(expectedResult, csvResult));
            
        }catch(FileNotFoundException ex){
            Logger.getLogger(ZTEXLSCMParser.class.getName()).log(Level.SEVERE, null, ex);
        }catch(IOException ex){
            Logger.getLogger(ZTEXLSCMParser.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    
    }
    
}
