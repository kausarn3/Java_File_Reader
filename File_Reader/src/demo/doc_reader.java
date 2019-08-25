package demo;

import java.io.FileInputStream;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

class doc_reader{
	
	public static class TextReader {
			
			public static void main(String[] args) {
			 try {
				   FileInputStream file = new FileInputStream("D://lecture//eclipse//test.docx");
				   XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(file));
				   XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
				   System.out.println(extractor.getText());
				} catch(Exception ex) {
				    ex.printStackTrace();
				}
		 }

		}
}
	
