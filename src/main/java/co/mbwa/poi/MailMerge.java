package co.mbwa.poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.xwpf.usermodel.PositionInParagraph;
import org.apache.poi.xwpf.usermodel.TextSegement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

public class MailMerge {

	//fields to be used in merge
	public static final String DATE = "date";
	public static final String SALUTATION = "salutation";
	public static final String FIRST_NAME = "firstName";
	public static final String LAST_NAME = "lastName";
	public static final String TITLE = "title";
	public static final String ACCOUNT_NAME = "accountName";
	public static final String ADDRESS = "address";
	public static final String CITY = "city";
	public static final String STATE = "state";
	public static final String POSTAL_CODE = "postalCode";
	public static final String USER_FIRST_NAME = "userFirstName";
	public static final String USER_LAST_NAME = "userLastName";
	

	//array of fields
	private static final String mailMergeFields[] = {DATE, SALUTATION, FIRST_NAME, LAST_NAME, TITLE, ACCOUNT_NAME, ADDRESS, CITY, STATE, POSTAL_CODE, USER_FIRST_NAME, USER_LAST_NAME};

	public static void merge(Map<String, String> fieldMap, String template, String outputFile) throws IOException, XmlException {
		
		//valid input map will hold valid input fields and values
		Map<String, String> validFieldMap = new HashMap<String, String>();
		
		//loop through the valid fields to fill the input map
		for (String field : mailMergeFields) {
			if (fieldMap.containsKey(field)) {
				validFieldMap.put("${" + field + "}", fieldMap.get(field));
			}
		}

		
		try (InputStream is = MailMerge.class.getClassLoader().getResourceAsStream(template);) {
			try (XWPFDocument doc = new XWPFDocument(is)) {

				replaceInParagraphs(validFieldMap, doc.getParagraphs());
				
				try (OutputStream out = new FileOutputStream(outputFile)) {
					doc.write(out);
				}

			} 
		}
	}


	  private static void replaceInParagraphs(Map<String, String> replacements, List<XWPFParagraph> xwpfParagraphs) {
		  
		    for (XWPFParagraph paragraph : xwpfParagraphs) {
		      List<XWPFRun> runs = paragraph.getRuns();

		      for (Map.Entry<String, String> replPair : replacements.entrySet()) {    
		        String find = replPair.getKey();
		        String repl = replPair.getValue();
		        TextSegement found = paragraph.searchText(find, new PositionInParagraph());

		        if ( found != null ) {

		          if ( found.getBeginRun() == found.getEndRun() ) {
		            // whole search string is in one Run
		           XWPFRun run = runs.get(found.getBeginRun());
		           String runText = run.getText(run.getTextPosition());
		           String replaced = runText.replace(find, repl);
		           run.setText(replaced, 0);
		          } else {
		            // The search string spans over more than one Run
		            // Put the Strings together
		            StringBuilder b = new StringBuilder();
		            for (int runPos = found.getBeginRun(); runPos <= found.getEndRun(); runPos++) {
		              XWPFRun run = runs.get(runPos);
		              b.append(run.getText(run.getTextPosition()));
		            }                       
		            String connectedRuns = b.toString();
		            String replaced = connectedRuns.replace(find, repl);

		            // The first Run receives the replaced String of all connected Runs
		            XWPFRun partOne = runs.get(found.getBeginRun());
		            partOne.setText(replaced, 0);

		            // Removing the text in the other Runs.
		            for (int runPos = found.getBeginRun()+1; runPos <= found.getEndRun(); runPos++) {
		              XWPFRun partNext = runs.get(runPos);
		              partNext.setText("", 0);
		            }
		            
		          }
		        }
		      }      
		    }
		  }

}
