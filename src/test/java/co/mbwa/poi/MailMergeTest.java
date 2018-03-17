package co.mbwa.poi;

import static org.junit.Assert.fail;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlException;
import org.junit.Test;

public class MailMergeTest {

	@Test
	public void test() throws IOException, XmlException {

		//When
		String template = "letter-with-table.docx";
		Map<String, String> inputMap = new HashMap<String, String>();
		inputMap.put("date", "March 17, 2018");
		inputMap.put("salutation", "Mr.");
		inputMap.put("firstName", "Jim");
		inputMap.put("lastName", "Beam");
		inputMap.put("title", "Manager");
		inputMap.put("accountName", "MBWA");
		inputMap.put("address", "10 Manage By Walking Around Drive");
		inputMap.put("city", "Danbury");
		inputMap.put("state", "CT");
		inputMap.put("postalCode", "06810");
		inputMap.put("userFirstName", "John");
		inputMap.put("userLastName", "Walker");

		String outputFile = "opty.docx";

		
		MailMerge.merge(inputMap, template, outputFile);

		// Then
		// validate the output file exists
		File wordDoc = new File(outputFile);

		if (!wordDoc.exists() || !wordDoc.isFile()) {

			throw new IllegalArgumentException("Could not read Microsoft Word Doc " + wordDoc);
		}

		// and validate the fields have been replaced
		try (InputStream is = new FileInputStream(wordDoc)) {
			try (XWPFDocument doc = new XWPFDocument(is)) {

				String xmlDoc = doc.getDocument().xmlText();
								
				Iterator<Entry<String, String>> entries = inputMap.entrySet().iterator();
				while (entries.hasNext()) {
					Map.Entry<String, String> entry = entries.next();

					if (! xmlDoc.contains(entry.getValue())) {

						fail(entry.getValue() + " field value not found in merged document");

					}

				}

			}
		}
	}
}
