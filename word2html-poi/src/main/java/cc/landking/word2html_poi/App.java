package cc.landking.word2html_poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.w3c.dom.Document;



/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws Exception
    {
    	  final String path = "result/";
          final String file = "2.doc";
          InputStream input = new FileInputStream( file);
          HWPFDocument wordDocument = new HWPFDocument(input);
          LandkingHtmlDocumentFacade facade = new LandkingHtmlDocumentFacade(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
          WordToHtmlConverter wordToHtmlConverter = new LandkingWordToHtmlConverter(facade);
          wordToHtmlConverter.setPicturesManager(new PicturesManager() {
              public String savePicture(byte[] content, PictureType pictureType, String suggestedName, float widthInches, float heightInches) {
                  File file = new File(path + suggestedName);
              	try {
            OutputStream os = new FileOutputStream(file);
            os.write(content);
            os.close();
          } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
          } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
          }
              	return path + suggestedName;
              }
          });
          wordToHtmlConverter.processDocument(wordDocument);
          Document htmlDocument = wordToHtmlConverter.getDocument();
          File htmlFile = new File("2.html");
          OutputStream outStream = new FileOutputStream(htmlFile);
          DOMSource domSource = new DOMSource(htmlDocument);
          StreamResult streamResult = new StreamResult(outStream);
   
          TransformerFactory tf = TransformerFactory.newInstance();
          Transformer serializer = tf.newTransformer();
          serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
          serializer.setOutputProperty(OutputKeys.INDENT, "yes");
          serializer.setOutputProperty(OutputKeys.METHOD, "html");
          serializer.transform(domSource, streamResult);
          outStream.close();

    }
}
