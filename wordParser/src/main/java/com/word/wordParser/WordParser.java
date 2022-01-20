package com.word.wordParser;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import org.apache.poi.ooxml.POIXMLProperties.CoreProperties;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.json.JSONArray;
import org.json.JSONObject;


public class WordParser {

	public static void main(String[] args) throws IOException {

        String fileName = "c:\\Users\\GOWTHAM\\Desktop\\ADA.docx";
//        C:\Users\GOWTHAM\Desktop
        JSONObject finalObj=new JSONObject();

        try (XWPFDocument doc = new XWPFDocument(
                Files.newInputStream(Paths.get(fileName)))) {

            XWPFWordExtractor xwpfWordExtractor = new XWPFWordExtractor(doc);
            CoreProperties docText11 = xwpfWordExtractor.getCoreProperties();
//            System.out.println(docText11.getTitle());
            String u=docText11.getTitle();
            String url=u.substring(u.indexOf("http"),u.length());

            String docText=xwpfWordExtractor.getText();
            String Date=(docText.substring(docText.indexOf("produced on"),docText.indexOf("Overall Quality")-1)).replace("produced on ", "").replaceAll("\\.", "");
            Date d1=new SimpleDateFormat("EEE, MMM dd, yyyy").parse(Date);  
            String Date_Stamp=new SimpleDateFormat("MMM dd yyyy").format(d1);
            String s=docText.substring(docText.indexOf("Totals"),docText.length()-1);
            
            JSONObject errorsObject=getErrors(s);
			JSONObject accessibilityAObject=getAccessibilityA(s);
			JSONObject accessibilityAAObject=getAccessibilityAA(s);
			JSONObject compatibility1Object=getCompatibility1(s);
			JSONObject compatibility2Object=getCompatibility2(s);
			JSONObject refObj=new JSONObject();
			JSONArray refArray=new JSONArray();
			refObj.put("Url",url );
			refObj.put("Date", Date_Stamp);
			refObj.put("ErrorObject", errorsObject);
			refObj.put("Acces-A", accessibilityAObject);
			refObj.put("Acces-AA", accessibilityAAObject);
			refObj.put("Comp-1", compatibility1Object);
			refObj.put("Comp-2", compatibility2Object);
			refArray.put(refObj);
			finalObj.put("chartData",refArray);
			System.out.println("==============="+finalObj);
            
        }catch(Exception e) {
        	e.printStackTrace();
        }

    }
////////////////////////////////////////////////////////////////////////////////////////////////////////////

	static String extractInt(String str)
    {
        // Replacing every non-digit number
        // with a space(" ")
        str = str.replaceAll("[^\\d]", " ");
  
        // Remove extra spaces from the beginning
        // and the ending of the string
        str = str.trim();
  
        // Replace all the consecutive white
        // spaces with a single space
        str = str.replaceAll(" +", " ");
  
        if (str.equals(""))
            return "-1";
  
        return str;
    }
////////////////////////////////////////////////////////////////////////////////////////////////////////////

	static JSONObject getErrors(String s) {
		JSONObject finalObj=new JSONObject();
		JSONArray issuesArray = new JSONArray();
		JSONArray guideLinesArray = new JSONArray();
		JSONArray countArray = new JSONArray();
		 String E=s.substring(s.indexOf("Errors"),s.indexOf("pages"));
         String refined=E.substring(E.indexOf("Priority 1"),E.length()-1).replace("Priority 1", "").trim();
         String result=extractInt(refined);
         int Count=Integer.parseInt(result.split("\\s")[0]);
//         System.out.println("================"+result);
         for (int i = 0; i < Count; i++) {
        	 JSONObject q = new JSONObject();
        	 JSONObject g = new JSONObject();
        	 JSONObject c = new JSONObject();
        	 q.put("G&L", "NA");
        	 g.put("D&U", "NA");
        	 c.put("Count", "NA");
        	 issuesArray.put(q);
        	 guideLinesArray.put(q);
        	 countArray.put(q);  	 
		}
            finalObj.put("Priority", "(Error)Priority1");
			finalObj.put("GuideLines", guideLinesArray);
			finalObj.put("Description", issuesArray);
			finalObj.put("Count", countArray);
       
		return finalObj;
		
	}
////////////////////////////////////////////////////////////////////////////////////////////////////////////

	static JSONObject getAccessibilityA(String s) {
		JSONObject finalObj=new JSONObject();
		JSONArray issuesArray = new JSONArray();
		JSONArray guideLinesArray = new JSONArray();
		JSONArray countArray = new JSONArray();
		String A=s.substring(s.indexOf("Level A"),s.length()-1);
        String E=A.substring(A.indexOf("Level A"),A.indexOf("pages"));
        String refined=E.substring(E.indexOf("Level A"),E.length()-1).replace("Level A", "").trim();
        String result=extractInt(refined);
        int Count=Integer.parseInt(result.split("\\s")[0]);
//      System.out.println("================"+result);
      for (int i = 0; i < Count; i++) {
     	 JSONObject q = new JSONObject();
     	 JSONObject g = new JSONObject();
     	 JSONObject c = new JSONObject();
     	 q.put("G&L", "NA");
     	 g.put("D&U", "NA");
     	 c.put("Count", "NA");
     	 issuesArray.put(q);
     	 guideLinesArray.put(q);
     	 countArray.put(q);  	 
		}
         finalObj.put("Priority", "Level A");
			finalObj.put("GuideLines", guideLinesArray);
			finalObj.put("Description", issuesArray);
			finalObj.put("Count", countArray);
    
		return finalObj;
	}
////////////////////////////////////////////////////////////////////////////////////////////////////////////

    static JSONObject getAccessibilityAA(String s) {
    	JSONObject finalObj=new JSONObject();
		JSONArray issuesArray = new JSONArray();
		JSONArray guideLinesArray = new JSONArray();
		JSONArray countArray = new JSONArray();
		String A=s.substring(s.indexOf("Level AA"),s.length()-1);
        String E=A.substring(A.indexOf("Level AA"),A.indexOf("pages"));
        String refined=E.substring(E.indexOf("Level AA"),E.length()-1).replace("Level AA", "").trim();
        String result=extractInt(refined);
        int Count=Integer.parseInt(result.split("\\s")[0]);
//      System.out.println("================"+result);
      for (int i = 0; i < Count; i++) {
     	 JSONObject q = new JSONObject();
     	 JSONObject g = new JSONObject();
     	 JSONObject c = new JSONObject();
     	 q.put("G&L", "NA");
     	 g.put("D&U", "NA");
     	 c.put("Count", "NA");
     	 issuesArray.put(q);
     	 guideLinesArray.put(q);
     	 countArray.put(q);  	 
		}
         finalObj.put("Priority", "Level AA");
			finalObj.put("GuideLines", guideLinesArray);
			finalObj.put("Description", issuesArray);
			finalObj.put("Count", countArray);
    
		return finalObj;
		
	}
////////////////////////////////////////////////////////////////////////////////////////////////////////////

    static JSONObject getCompatibility1(String s) {
    	JSONObject finalObj=new JSONObject();
		JSONArray issuesArray = new JSONArray();
		JSONArray guideLinesArray = new JSONArray();
		JSONArray countArray = new JSONArray();
    	 String P=s.substring(s.indexOf("Compatibility"),s.length()-1);
         String A=P.substring(P.indexOf("Priority 1"),P.length()-1);
         String E=A.substring(A.indexOf("Priority 1"),A.indexOf("pages"));
         String refined=E.substring(E.indexOf("Priority 1"),E.length()-1).replace("Priority 1", "").trim();
         String result=extractInt(refined);
         int Count=Integer.parseInt(result.split("\\s")[0]);
//       System.out.println("================"+result);
       for (int i = 0; i < Count; i++) {
      	 JSONObject q = new JSONObject();
      	 JSONObject g = new JSONObject();
      	 JSONObject c = new JSONObject();
      	 q.put("G&L", "NA");
      	 g.put("D&U", "NA");
      	 c.put("Count", "NA");
      	 issuesArray.put(q);
      	 guideLinesArray.put(q);
      	 countArray.put(q);  	 
 		}
          finalObj.put("Priority", "(Comp)Priority1");
 			finalObj.put("GuideLines", guideLinesArray);
 			finalObj.put("Description", issuesArray);
 			finalObj.put("Count", countArray);
     
 		return finalObj;
   	}
////////////////////////////////////////////////////////////////////////////////////////////////////////////

    static JSONObject getCompatibility2(String s) {
    	JSONObject finalObj=new JSONObject();
		JSONArray issuesArray = new JSONArray();
		JSONArray guideLinesArray = new JSONArray();
		JSONArray countArray = new JSONArray();
    	 String P=s.substring(s.indexOf("Compatibility"),s.length()-1);
         String A=P.substring(P.indexOf("Priority 2"),P.length()-1);
         String E=A.substring(A.indexOf("Priority 2"),A.indexOf("pages"));
         String refined=E.substring(E.indexOf("Priority 2"),E.length()-1).replace("Priority 2", "").trim();
         String result=extractInt(refined);
         int Count=Integer.parseInt(result.split("\\s")[0]);
//       System.out.println("================"+result);
       for (int i = 0; i < Count; i++) {
      	 JSONObject q = new JSONObject();
      	 JSONObject g = new JSONObject();
      	 JSONObject c = new JSONObject();
      	 q.put("G&L", "NA");
      	 g.put("D&U", "NA");
      	 c.put("Count", "NA");
      	 issuesArray.put(q);
      	 guideLinesArray.put(q);
      	 countArray.put(q);  	 
 		}
          finalObj.put("Priority", "(Comp)Priority2");
 			finalObj.put("GuideLines", guideLinesArray);
 			finalObj.put("Description", issuesArray);
 			finalObj.put("Count", countArray);
     
 		return finalObj;
		
   	}
////////////////////////////////////////////////////////////////////////////////////////////////////////////
}
