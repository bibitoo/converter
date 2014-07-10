package cc.landking.word2html_poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.w3c.dom.Document;

public class Docx2Html {

	public static void main(String[] args) throws Exception {
		 final String path = "result/";
		 String fileInName = "3";
         final String file = fileInName+".docx";
         InputStream input = new FileInputStream( file);
         XWPFDocument wordDocument = new XWPFDocument(input);
//         XHTMLOptions options = XHTMLOptions.create().indent( 4 );
         File htmlFile = new File(path+fileInName+".html");
         
         OutputStream outStream = new FileOutputStream(htmlFile);
         XHTMLOptions options = XHTMLOptions.create();// .indent( 4 );
         // Extract image
         File imageFolder = new File( path + "/images/" + fileInName );
         options.setExtractor( new FileImageExtractor( imageFolder ) );
         // URI resolver
         options.URIResolver( new FileURIResolver( imageFolder ) );
         XHTMLConverter.getInstance().convert( wordDocument, outStream, options );
        
         outStream.close();

	}

}
