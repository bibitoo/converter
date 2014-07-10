package cc.landking.word2html_poi;

import org.apache.poi.hwpf.converter.HtmlDocumentFacade;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class LandkingHtmlDocumentFacade extends HtmlDocumentFacade {
	int pageCount = 0;
	
	public LandkingHtmlDocumentFacade(Document document) {
		super(document);
	}
    public Element createPageBreak()
    {
    	pageCount ++;
    	Element content = document.createElement("span");
    	content.setTextContent(String.valueOf(pageCount));
        Element el =  (Element) document.createElement( "div" );
        el.appendChild(content);
        return el;
    }


}
