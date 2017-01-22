package cc.landking.converter.word2image;

import java.io.File;
import java.io.IOException;
import java.util.List;

import org.icepdf.core.exceptions.PDFException;
import org.icepdf.core.exceptions.PDFSecurityException;
import org.springframework.context.ConfigurableApplicationContext;
import org.springframework.context.support.ClassPathXmlApplicationContext;

public class AppLocalTest {
	public static void main(String[] args) throws IOException, PDFException, PDFSecurityException{
		ConfigurableApplicationContext context =
				new ClassPathXmlApplicationContext(new String[] {"applicationContext.xml"});
		 File inputFile = new File("test.doc");
	        System.out.printf("-- source %s ", inputFile.getAbsolutePath());
	       
	        SofficeService service = (SofficeService) context.getBean("sofficeService");
	        List<File> files = service.convertToJpg("test.docx", null);
	        for(File file : files){
	        	System.out.println(file.getAbsolutePath());
	        }
	}
}
