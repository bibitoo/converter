package cc.landking.word2html_poi;

import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.converter.HtmlDocumentFacade;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class LandkingWordToHtmlConverter extends WordToHtmlConverter {
	 private  LandkingHtmlDocumentFacade htmlDocumentFacade;
	public LandkingWordToHtmlConverter(Document document) {
		super(document);
	}
	public LandkingWordToHtmlConverter(LandkingHtmlDocumentFacade htmlDocumentFacade) {
		super(htmlDocumentFacade);
		this.htmlDocumentFacade = (LandkingHtmlDocumentFacade) htmlDocumentFacade;
	}
	int pageCount = 0;
    @Override
    protected void processPageBreak( HWPFDocumentCore wordDocument, Element flow )
    {
    	if(htmlDocumentFacade == null){
    		super.processPageBreak(wordDocument, flow);
    	}else{
    		
    		flow.appendChild( htmlDocumentFacade.createPageBreak() );
    	}
    }

}
