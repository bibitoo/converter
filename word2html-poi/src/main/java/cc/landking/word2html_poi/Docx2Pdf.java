package cc.landking.word2html_poi;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.lowagie.text.Font;
import com.lowagie.text.pdf.BaseFont;

import fr.opensagres.xdocreport.itext.extension.font.IFontProvider;
import fr.opensagres.xdocreport.itext.extension.font.ITextFontRegistry;

public class Docx2Pdf {

	public static void main(String[] args) throws Exception {
		// 1) Load DOCX into XWPFDocument
		final String path = "result/";
		String fileInName = "3";
		final String file = fileInName + ".docx";
		InputStream input = new FileInputStream(file);
		XWPFDocument wordDocument = new XWPFDocument(input);
		// 2) Prepare Pdf options
//		PdfOptions options = PdfOptions.create();
		PdfOptions options =  PdfOptions.create().fontEncoding( "windows-936" );
		 options.fontProvider( new IFontProvider()
	        {

	            public Font getFont( String familyName, String encoding, float size, int style, Color color )
	            {
	                try
	                {
	                    BaseFont bfChinese =
	                        BaseFont.createFont( "/home/sunz/git/converter/word2html-poi/arialuni.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED );
	                    Font fontChinese = new Font( bfChinese, size, style, color );
	                    if ( familyName != null )
	                        fontChinese.setFamily( familyName );
	                    return fontChinese;
	                }
	                catch ( Throwable e )
	                {
	                    e.printStackTrace();
	                    // An error occurs, use the default font provider.
	                    return ITextFontRegistry.getRegistry().getFont( familyName, encoding, size, style, color );
	                }
	            }
	        } );

		// 3) Convert XWPFDocument to Pdf
		OutputStream out = new FileOutputStream(new File(fileInName + ".pdf"));
		PdfConverter.getInstance().convert(wordDocument, out, options);

	}

}
