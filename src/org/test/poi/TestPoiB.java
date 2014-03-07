package org.test.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;


public class TestPoiB {
	private static WordExtractor we;

	/**
     * @param args
     */
    public static void main(String[] args) {
    // TODO Auto-generated method stub
    // Declare test var.
    POIFSFileSystem fs = null;
    File file;
        
         try {  
        	 System.out.println("Starting POI testing!");
             file = new File("C:/msword_result.doc");  
             fs = new POIFSFileSystem(new FileInputStream("C:/msword_template.doc"));
             HWPFDocument doc = new HWPFDocument(fs); 
             we = new WordExtractor(doc);

             Range range = doc.getRange();
             range.replaceText("PR_NUMBER", "201403071242");
             range.replaceText("CLIENT_NAME", "Sherlock Holmes");
             range.replaceText("CLIENT_ADDRESS", "221B Baker Street London");
             String[] paragraphs = we.getParagraphText();
             for (int i = 0; i < paragraphs.length; i++) {  

                 //org.apache.poi.hwpf.usermodel.Paragraph pr = range.getParagraph(i);
                 //System.out.println(pr.toString());
                 //CharacterRun run = pr.getCharacterRun(i);
                 //run.setBold(true);
                 //run.setCapitalized(true);
                 //run.setItalic(true);
                 
                 paragraphs[i] = paragraphs[i].replaceAll("\\cM?\r?\n", ""); 
                 paragraphs[i] = paragraphs[i].replaceAll("PR_NUMBER", "201403071242");
                 
             //System.out.println("Length:" + paragraphs[i].length());  
             //System.out.println("Paragraph" + i + ": " + paragraphs[i].toString()); 
             } 
             
             if (!file.exists()){
            	   file.createNewFile();
              }
             doc.write(new FileOutputStream(file));
             
             
             System.out.println("POI testing completed!");  
         } catch (Exception e) {  
             System.out.println("Exception during test!");  
             e.printStackTrace();  
         } 
    }  
}