package cc.landking.converter.office.excel;
import org.apache.poi.ddf.EscherClientAnchorRecord;
import org.apache.poi.ddf.EscherRecord;
import org.apache.poi.hssf.record.EscherAggregate;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.format.CellFormatResult;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedReader;
import java.io.Closeable;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Formatter;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import static org.apache.poi.ss.usermodel.CellStyle.*;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 * use bootstrap2.3.2 for css
 * may add bootstrap css，javascript  and jquery like
 *   <link href="http://cdn.bootcss.com/twitter-bootstrap/2.3.2/css/bootstrap.min.css" rel="stylesheet">
 *        <script src="http://cdn.bootcss.com/jquery/1.10.2/jquery.min.js"></script>
 *         <script src="http://cdn.bootcss.com/twitter-bootstrap/2.3.2/js/bootstrap.min.js"></script>
 *         
 * @author sunz
 *
 */

public class ExcelToHtml {

	private void resetSheet(){
		gotBounds = false;
		firstColumn = 0;
		endColumn = 0;
	}

	public void setSheetBorderOut(boolean sheetBorderOut) {
		this.sheetBorderOut = sheetBorderOut;
	}

	public void setStyleOut(boolean styleOut) {
		this.styleOut = styleOut;
	}

	private final Workbook wb;
    private final Appendable output;
    private boolean completeHTML;
    private boolean sheetBorderOut = true;
    private boolean styleOut = true;
    private Formatter out;
    private boolean gotBounds;
    private int firstColumn;
    private int endColumn;
    private HtmlHelper helper;
	  String resultImagePath = "result/images/data/";

    public String getResultImagePath() {
		return resultImagePath;
	}

	public void setResultImagePath(String resultImagePath) {
		this.resultImagePath = resultImagePath;
	}

	Map<String,MyPictureData> myPictureDatas = new HashMap<String,MyPictureData>();

    private static final String TABLE_CLASS = "\"table table-bordered excelDefaults\"";
    private static final String DEFAULTS_CLASS = "excelDefaults";
    private static final String COL_HEAD_CLASS = "colHeader";
    private static final String ROW_HEAD_CLASS = "rowHeader";

    private static final Map<Short, String> ALIGN = mapFor(ALIGN_LEFT, "left",
            ALIGN_CENTER, "center", ALIGN_RIGHT, "right", ALIGN_FILL, "left",
            ALIGN_JUSTIFY, "left", ALIGN_CENTER_SELECTION, "center");

    private static final Map<Short, String> VERTICAL_ALIGN = mapFor(
            VERTICAL_BOTTOM, "bottom", VERTICAL_CENTER, "middle", VERTICAL_TOP,
            "top");

    private static final Map<Short, String> BORDER = mapFor(BORDER_DASH_DOT,
            "dashed 1pt", BORDER_DASH_DOT_DOT, "dashed 1pt", BORDER_DASHED,
            "dashed 1pt", BORDER_DOTTED, "dotted 1pt", BORDER_DOUBLE,
            "double 3pt", BORDER_HAIR, "solid 1px", BORDER_MEDIUM, "solid 2pt",
            BORDER_MEDIUM_DASH_DOT, "dashed 2pt", BORDER_MEDIUM_DASH_DOT_DOT,
            "dashed 2pt", BORDER_MEDIUM_DASHED, "dashed 2pt", BORDER_NONE,
            "none", BORDER_SLANTED_DASH_DOT, "dashed 2pt", BORDER_THICK,
            "solid 3pt", BORDER_THIN, "dashed 1pt");

    @SuppressWarnings({"unchecked"})
    private static <K, V> Map<K, V> mapFor(Object... mapping) {
        Map<K, V> map = new HashMap<K, V>();
        for (int i = 0; i < mapping.length; i += 2) {
            map.put((K) mapping[i], (V) mapping[i + 1]);
        }
        return map;
    }

    /**
     * Creates a new converter to HTML for the given workbook.
     *
     * @param wb     The workbook.
     * @param output Where the HTML output will be written.
     *
     * @return An object for converting the workbook to HTML.
     */
    public static ExcelToHtml create(Workbook wb, Appendable output) {
        return new ExcelToHtml(wb, output);
    }

    /**
     * Creates a new converter to HTML for the given workbook.  If the path ends
     * with "<tt>.xlsx</tt>" an {@link XSSFWorkbook} will be used; otherwise
     * this will use an {@link HSSFWorkbook}.
     *
     * @param path   The file that has the workbook.
     * @param output Where the HTML output will be written.
     *
     * @return An object for converting the workbook to HTML.
     */
    public static ExcelToHtml create(String path, Appendable output)
            throws IOException {
        return create(new FileInputStream(path), output);
    }

    /**
     * Creates a new converter to HTML for the given workbook.  This attempts to
     * detect whether the input is XML (so it should create an {@link
     * XSSFWorkbook} or not (so it should create an {@link HSSFWorkbook}).
     *
     * @param in     The input stream that has the workbook.
     * @param output Where the HTML output will be written.
     *
     * @return An object for converting the workbook to HTML.
     */
    public static ExcelToHtml create(InputStream in, Appendable output)
            throws IOException {
        try {
            Workbook wb = WorkbookFactory.create(in);
            return create(wb, output);
        } catch (InvalidFormatException e){
            throw new IllegalArgumentException("Cannot create workbook from stream", e);
        }
    }

    private ExcelToHtml(Workbook wb, Appendable output) {
        if (wb == null)
            throw new NullPointerException("wb");
        if (output == null)
            throw new NullPointerException("output");
        this.wb = wb;
        this.output = output;
        setupColorMap();
        if(wb instanceof HSSFWorkbook){
        	myPictureDatas = this.getAllPictures((HSSFWorkbook)wb);
        }
        
    }

    private void setupColorMap() {
        if (wb instanceof HSSFWorkbook)
            helper = new HSSFHtmlHelper((HSSFWorkbook) wb);
        else if (wb instanceof XSSFWorkbook)
            helper = new XSSFHtmlHelper((XSSFWorkbook) wb);
        else
            throw new IllegalArgumentException(
                    "unknown workbook type: " + wb.getClass().getSimpleName());
    }

    /**
     * Run this class as a program
     *
     * @param args The command line arguments.
     *
     * @throws Exception Exception we don't recover from.
     */
    public static void main(String[] args) throws Exception {
//        if(args.length < 2){
//            System.err.println("usage: ToHtml inputWorkbook outputHtmlFile");
//            return;
//        }

    	if(args == null || args.length<2){
    		args = new String[2];
    		args[0] = "2.xlsx";
    		args[1] = "2.html";
    	}
    	//if use inputstream and string result,use this:
    	//StringBuffer stringOut = new StringBuffer();
    	//
    	// ExcelToHtml toHtml = create(new FileInputStream(args[0]), stringOut);
        ExcelToHtml toHtml = create(args[0], new PrintWriter(new FileWriter(args[1])));
        toHtml.setCompleteHTML(false);//if true,output the html\title and body tag
        toHtml.setSheetBorderOut(false);//if false,do not output sheet boder ,top A,B,C...and left row number
        toHtml.setStyleOut(false);//if false, do not output style 
        String resultImagePath = "result/images/data/";
        toHtml.setResultImagePath(resultImagePath);//image in sheet store path
        toHtml.printPage();
    }

    public void setCompleteHTML(boolean completeHTML) {
        this.completeHTML = completeHTML;
    }

    public void printPage() throws IOException {
        try {
            ensureOut();
            if (completeHTML) {
                out.format(
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" ?>%n");
                out.format("<html>%n");
                out.format("<head>%n");
                out.format("</head>%n");
                out.format("<body>%n");
            }

            print();

            if (completeHTML) {
                out.format("</body>%n");
                out.format("</html>%n");
            }
        } finally {
            if (out != null)
                out.close();
            if (output instanceof Closeable) {
                Closeable closeable = (Closeable) output;
                closeable.close();
            }
        }
    }

    public void print() {
    	if(styleOut){
        printInlineStyle();
    	}
        printSheets();
    }

    private void printInlineStyle() {
        //out.format("<link href=\"excelStyle.css\" rel=\"stylesheet\" type=\"text/css\">%n");
        out.format("<style type=\"text/css\">%n");
        printStyles();
        out.format("</style>%n");
    }

    private void ensureOut() {
        if (out == null)
            out = new Formatter(output);
    }

    public void printStyles() {
        ensureOut();

        // First, copy the base css
        BufferedReader in = null;
        try {
            in = new BufferedReader(new InputStreamReader(
                    getClass().getResourceAsStream("excelStyle.css")));
            String line;
            while ((line = in.readLine()) != null) {
                out.format("%s%n", line);
            }
        } catch (IOException e) {
            throw new IllegalStateException("Reading standard css", e);
        } finally {
            if (in != null) {
                try {
                    in.close();
                } catch (IOException e) {
                    //noinspection ThrowFromFinallyBlock
                    throw new IllegalStateException("Reading standard css", e);
                }
            }
        }

        // now add css for each used style
        Set<CellStyle> seen = new HashSet<CellStyle>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            Iterator<Row> rows = sheet.rowIterator();
            while (rows.hasNext()) {
                Row row = rows.next();
                for (Cell cell : row) {
                    CellStyle style = cell.getCellStyle();
                    if (!seen.contains(style)) {
                        printStyle(style);
                        seen.add(style);
                    }
                }
            }
        }
    }

    private void printStyle(CellStyle style) {
        out.format(".%s .%s {%n", DEFAULTS_CLASS, styleName(style));
        styleContents(style);
        out.format("}%n");
    }

    private void styleContents(CellStyle style) {
        styleOut("text-align", style.getAlignment(), ALIGN);
        styleOut("vertical-align", style.getAlignment(), VERTICAL_ALIGN);
        fontStyle(style);
        borderStyles(style);
        helper.colorStyles(style, out);
    }

    private void borderStyles(CellStyle style) {
        styleOut("border-left", style.getBorderLeft(), BORDER);
        styleOut("border-right", style.getBorderRight(), BORDER);
        styleOut("border-top", style.getBorderTop(), BORDER);
        styleOut("border-bottom", style.getBorderBottom(), BORDER);
    }

    private void fontStyle(CellStyle style) {
        Font font = wb.getFontAt(style.getFontIndex());
System.out.println("bold weight:"+font.getBoldweight() );
        if (font.getBoldweight() > HSSFFont.BOLDWEIGHT_NORMAL)
            out.format("  font-weight: bold;%n");
        if (font.getItalic())
            out.format("  font-style: italic;%n");

        int fontheight = font.getFontHeightInPoints();
        if (fontheight == 9) {
            //fix for stupid ol Windows
            fontheight = 10;
        }
        out.format("  font-size: %dpt;%n", fontheight);

        // Font color is handled with the other colors
    }

    private String styleName(CellStyle style) {
        if (style == null)
            style = wb.getCellStyleAt((short) 0);
        StringBuilder sb = new StringBuilder();
        Formatter fmt = new Formatter(sb);
        fmt.format("style_%02x", style.getIndex());
        return fmt.toString();
    }

    private <K> void styleOut(String attr, K key, Map<K, String> mapping) {
        String value = mapping.get(key);
        if (value != null) {
            out.format("  %s: %s;%n", attr, value);
        }
    }

    private static int ultimateCellType(Cell c) {
        int type = c.getCellType();
        if (type == Cell.CELL_TYPE_FORMULA)
            type = c.getCachedFormulaResultType();
        return type;
    }

    private void printSheets() {
        ensureOut();
        String sheetAllHead = "<div class=\"tabbable tabs-below\"> <div class=\"tab-content\">%n";
        
        String sheetAllContentFood = " </div>%n";
        String sheetTitleHead = "<ul class=\"nav nav-tabs\">%n";
        String sheetTitleEnd = "</ul>%n";
        
        out.format(sheetAllHead);
        
       int activeSheetIndex =  wb.getActiveSheetIndex();
       int count = wb.getNumberOfSheets();
       for(int i=0;i<count;i++){
    	   resetSheet();
    	   boolean active = (i == activeSheetIndex);
     	   Sheet sheet = wb.getSheetAt(i);
       	   String sheetHead = "<div class=\"tab-pane "+(active?"active":"")+"\" id=\"sheet"+i+"\">%n";
    	   out.format(sheetHead);
          printSheet(sheet);
           out.format(sheetAllContentFood);
       }
       out.format(sheetAllContentFood);
       out.format(sheetTitleHead);
       for(int i=0;i<count;i++){
    	   boolean active = (i == activeSheetIndex);
    	   Sheet sheet = wb.getSheetAt(i);
    	   String sheetHead = "<li class=\""+(active?"active":"")+"\"><a href=\"#sheet"+i+"\" data-toggle=\"tab\">"+sheet.getSheetName()+"</a></li>%n";
    	   out.format(sheetHead);
       }
       out.format(sheetTitleEnd);
       out.format(sheetAllContentFood);
    }

    public void printSheet(Sheet sheet) {
        ensureOut();
        out.format("<table class=%s>%n", TABLE_CLASS);
        if(sheetBorderOut){
        	printCols(sheet);
        }else{
            ensureColumnBounds(sheet);
        }
        try {
			printSheetContent(sheet);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        out.format("</table>%n");
//        try {
//			pringSheetImage(sheet);
//		} catch (Exception e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
    }
    private void pringSheetImage(Sheet sheet) throws Exception{
    	List<? extends PictureData> pictures = sheet.getWorkbook().getAllPictures();  
        List<ClientAnchorInfo> clientAnchorRecords = getClientAnchorRecords( (HSSFWorkbook) sheet.getWorkbook());  

    	if(sheet instanceof HSSFSheet){
    		HSSFSheet hsheet = (HSSFSheet)sheet;
    		 for (HSSFShape shape :hsheet.getDrawingPatriarch().getChildren()) {  
                 HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();  
       
                 if (shape instanceof HSSFPicture) {  
                     HSSFPicture pic = (HSSFPicture) shape;  
                     int row = anchor.getRow1();  
                     System.out.println("--->" + anchor.getRow1() + ":"  
                             + anchor.getCol1());  
                     int pictureIndex = pic.getPictureIndex()-1;  
                     HSSFPictureData picData = (HSSFPictureData) pictures.get(pictureIndex);  
       
                     System.out.println( "--->" + pictureIndex);  
                     ClientAnchorInfo anchorInfo = clientAnchorRecords.get(pic.getPictureIndex());  
                     EscherClientAnchorRecord clientAnchorRecord = anchorInfo.clientAnchorRecord;  
                     MyPictureData data = new MyPictureData((HSSFWorkbook) sheet.getWorkbook(), (HSSFSheet)sheet, picData, clientAnchorRecord);  
                     savePic(row, data);  
                    
                 }  else if(shape instanceof HSSFSimpleShape){
                	 HSSFSimpleShape hssfsimpleShap = (HSSFSimpleShape)shape;
                	 saveShape(hssfsimpleShap);
                 }
             }  
//    		   for (int i = 0; i < pictures.size(); i++) {  
//    	            HSSFPictureData pictureData = (HSSFPictureData) pictures.get(i);  
//    	            ClientAnchorInfo anchor = clientAnchorRecords.get(i);  
//    	            EscherClientAnchorRecord clientAnchorRecord = anchor.clientAnchorRecord;  
//    	            MyPictureData data = new MyPictureData((HSSFWorkbook) sheet.getWorkbook(), hsheet, pictureData, clientAnchorRecord);  
//    	            savePic(i, data);  
//    	        }  
    	}else if(sheet instanceof XSSFSheet){
    		XSSFSheet xsheet = (XSSFSheet)sheet;
     		
    	}
    	
    }
    private void saveShape(HSSFSimpleShape shape){
    	HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
    	
    	System.out.println("x1:"+ shape.getPatriarch().getX1()+" x2:"+shape.getPatriarch().getX2()+" y1:"+shape.getPatriarch().getY1()+" y2:"+shape.getPatriarch().getY2());
    	System.out.println("dx1:"+ shape.getAnchor().getDx1()+" dx2:"+shape.getAnchor().getDx2()+" dy1:"+shape.getAnchor().getDy1()+" dy2:"+shape.getAnchor().getDy2());
    	out.format("<svg width="+(shape.getAnchor().getDx1() - shape.getAnchor().getDx2())+" height="+(shape.getAnchor().getDy2()-shape.getAnchor().getDy1())+" version=\"1.1\" xmlns=\"http://www.w3.org/2000/svg\"></svg>");
    }
    public  Map<String,MyPictureData> getAllPictures(HSSFWorkbook workbook) {
    	 Map<String,MyPictureData> list = new HashMap<String,MyPictureData>();
//        List<MyPictureData> list = new ArrayList<MyPictureData>();  
  
        List<HSSFPictureData> pictureList = workbook.getAllPictures();  
        List<ClientAnchorInfo> clientAnchorRecords = getClientAnchorRecords(workbook);  
          
        if (pictureList.size() != clientAnchorRecords.size()) {  
            throw new RuntimeException("解析文件中的图片信息出错，找到的图片数量和图片位置信息数量不匹配");  
        }  
          
        for (int i = 0; i < pictureList.size(); i++) {  
            HSSFPictureData pictureData = pictureList.get(i);  
            ClientAnchorInfo anchor = clientAnchorRecords.get(i);  
            HSSFSheet sheet = anchor.sheet;  
            EscherClientAnchorRecord clientAnchorRecord = anchor.clientAnchorRecord;  
            MyPictureData data = new MyPictureData(workbook, sheet, pictureData, clientAnchorRecord);
//            try {
//				savePic(i, data);
//			} catch (Exception e) {
//				// TODO Auto-generated catch block
//				e.printStackTrace();
//			}  
            list.put(sheet.getSheetName()+"-"+data.getRow0()+"-"+data.getCol0(),data);  
        }  
          
        return list ;  
    } 
    private static class ClientAnchorInfo {  
        public HSSFSheet sheet;  
        public EscherClientAnchorRecord clientAnchorRecord;  
          
        public ClientAnchorInfo(HSSFSheet sheet, EscherClientAnchorRecord clientAnchorRecord) {  
            super();  
            this.sheet = sheet;  
            this.clientAnchorRecord = clientAnchorRecord;  
        }  
    }  
    private static List<ClientAnchorInfo> getClientAnchorRecords(HSSFWorkbook workbook) {  
        List<ClientAnchorInfo> list = new ArrayList<ClientAnchorInfo>();  
          
        EscherAggregate drawingAggregate = null;  
        HSSFSheet sheet = null;  
        List<EscherRecord> recordList = null;  
        Iterator<EscherRecord> recordIter = null;  
        int numSheets = workbook.getNumberOfSheets();  
        for(int i = 0; i < numSheets; i++) {  
            sheet = workbook.getSheetAt(i);  
            drawingAggregate = sheet.getDrawingEscherAggregate();  
            if(drawingAggregate != null) {  
                recordList = drawingAggregate.getEscherRecords();  
                recordIter = recordList.iterator();  
                while(recordIter.hasNext()) {  
                    getClientAnchorRecords(sheet, recordIter.next(), 1, list);  
                }  
            }  
        }  
          
        return list;  
    }  
  
    private static void getClientAnchorRecords(HSSFSheet sheet, EscherRecord escherRecord, int level, List<ClientAnchorInfo> list) {  
        List<EscherRecord> recordList = null;  
        Iterator<EscherRecord> recordIter = null;  
        EscherRecord childRecord = null;  
        recordList = escherRecord.getChildRecords();  
        recordIter = recordList.iterator();  
        while(recordIter.hasNext()) {  
            childRecord = recordIter.next();  
            if(childRecord instanceof EscherClientAnchorRecord) {  
                ClientAnchorInfo e = new ClientAnchorInfo(sheet, (EscherClientAnchorRecord) childRecord);  
                list.add(e);  
            }  
            if(childRecord.getChildRecords().size() > 0) {  
                getClientAnchorRecords(sheet, childRecord, level+1, list);  
            }  
        }  
    }  
    private   void savePic(int i, MyPictureData picData) throws Exception {  
    	  File file = new File(resultImagePath);
    	  if(!file.exists()){
    		  file.mkdirs();
    	  }
    	  System.out.println("图片位置："+picData.getDx1()+","+picData.getDy1()+ "-"+picData.getDx2()+","+picData.getDy2());
    	  PictureData pic = picData.getPictureData();
        String ext = pic.suggestFileExtension();  
  
        byte[] data = pic.getData();  
       
        if (ext.equals("jpeg")) {  
            FileOutputStream fout = new FileOutputStream(  
                   new File(file, "pict" + i + ".jpg"));  
            fout.write(data);  
            fout.close();  
            out.format("<img src=\""+resultImagePath+"pict" + i + ".jpg\"/>%n");
        }  
        if (ext.equals("png")) {  
            FileOutputStream fout = new FileOutputStream(  
            		  new File(file, "pict"+ i + ".png"));  
            fout.write(data);  
            fout.close();  
            out.format("<img src=\""+resultImagePath+"pict" + i + ".png\"/>%n");
        }  
    }  
    
    private void printCols(Sheet sheet) {
        out.format("<col/>%n");
        ensureColumnBounds(sheet);
        for (int i = firstColumn; i < endColumn; i++) {
            out.format("<col/>%n");
        }
    }

    private void ensureColumnBounds(Sheet sheet) {
        if (gotBounds)
            return;

        Iterator<Row> iter = sheet.rowIterator();
        firstColumn = (iter.hasNext() ? Integer.MAX_VALUE : 0);
        endColumn = 0;
        while (iter.hasNext()) {
            Row row = iter.next();
            short firstCell = row.getFirstCellNum();
            if (firstCell >= 0) {
                firstColumn = Math.min(firstColumn, firstCell);
                endColumn = Math.max(endColumn, row.getLastCellNum());
            }
        }
        gotBounds = true;
    }

    private void printColumnHeads() {
        out.format("<thead>%n");
        out.format("  <tr class=%s>%n", COL_HEAD_CLASS);
        out.format("    <th class=%s>&#x25CA;</th>%n", COL_HEAD_CLASS);
        //noinspection UnusedDeclaration
        StringBuilder colName = new StringBuilder();
        for (int i = firstColumn; i < endColumn; i++) {
            colName.setLength(0);
            int cnum = i;
            do {
                colName.insert(0, (char) ('A' + cnum % 26));
                cnum /= 26;
            } while (cnum > 0);
            out.format("    <th class=%s>%s</th>%n", COL_HEAD_CLASS, colName);
        }
        out.format("  </tr>%n");
        out.format("</thead>%n");
    }

    private void printSheetContent(Sheet sheet) throws Exception {
    	Map<String,String> map[] = getRowSpanColSpanMap(sheet);
    	if(sheetBorderOut){
        printColumnHeads();
    	}
  	  File file = new File(resultImagePath);
  	  if(!file.exists()){
  		  file.mkdirs();
  	  }

        out.format("<tbody>%n");
        Iterator<Row> rows = sheet.rowIterator();

        while (rows.hasNext()) {
            Row row = rows.next();

            out.format("  <tr>%n");
            if(sheetBorderOut){
            out.format("    <td class=%s>%d</td>%n", ROW_HEAD_CLASS,
                    row.getRowNum() + 1);
            }
            for (int i = firstColumn; i < endColumn; i++) {
                String content = "&nbsp;";
                String attrs = "";
                CellStyle style = null;
                if (i >= row.getFirstCellNum() && i < row.getLastCellNum()) {
                    Cell cell = row.getCell(i);
                    if (cell != null) {
                        style = cell.getCellStyle();
                        attrs = tagStyle(cell, style);
                        //Set the value that is rendered for the cell
                        //also applies the format
                        CellFormat cf = CellFormat.getInstance(
                                style.getDataFormatString());
                        CellFormatResult result = cf.apply(cell);
                        content = result.text;
                        if (content.equals(""))
                            content = "&nbsp;";
                        MyPictureData picData =  myPictureDatas.get(sheet.getSheetName()+"-"+row.getRowNum()+"-"+i);
                        if(picData != null){
                        	  System.out.println("图片位置："+picData.getDx1()+","+picData.getDy1()+ "-"+picData.getDx2()+","+picData.getDy2());
                        	
                        	  PictureData pic = picData.getPictureData();
                              String ext = pic.suggestFileExtension();  
                        
                              byte[] data = pic.getData();  
                             
                              if (ext.equals("jpeg")) {  
                                  FileOutputStream fout = new FileOutputStream(  
                                         new File(file, "pict" + i + ".jpg"));  
                                  fout.write(data);  
                                  fout.close();  
                                  out.format("<img src=\""+resultImagePath+"pict" + picData.getRow1() + ".jpg\"/>%n");
                              }  
                              if (ext.equals("png")) {  
                                  FileOutputStream fout = new FileOutputStream(  
                                  		  new File(file, "pict"+ i + ".png"));  
                                  fout.write(data);  
                                  fout.close();  
                                  out.format("<img src=\""+resultImagePath+"pict" + picData.getRow1() + ".png\"/>%n");
                              }  
                        }
                    }
                }
              //add colspan and rowspan
                
                String colrowspan = "";
                int rowNum = row.getRowNum() ;
                int colNum = i;
                
                if(map[0].containsKey(rowNum + "," + colNum)) {
                 
                 String pointString = map[0].get(rowNum + "," + colNum);
                 
                 map[0].remove(rowNum + "," + colNum);
                 
                 int bottomeRow = Integer.valueOf(pointString.split(",")[0]);
                 
                 int bottomeCol = Integer.valueOf(pointString.split(",")[1]);
                 
                 int rowSpan = bottomeRow - rowNum + 1;
                 
                 int colSpan = bottomeCol - colNum + 1;
                 
                 colrowspan = ("rowspan= '" + rowSpan + "' colspan= '" + colSpan + "' " );
                 
                } else if(map[1].containsKey(rowNum + "," + colNum)){
                 
                 map[1].remove(rowNum + "," + colNum);
                 
                 continue;
                 
//                } else {
//                 
//                 sb.append("<td ");
                }

                //add colspan and rowspan end
                out.format("    <td class=%s %s %s>%s</td>%n", styleName(style),
                        attrs,colrowspan, content);
            }
            out.format("  </tr>%n");
        }
        out.format("</tbody>%n");
    }

    private String tagStyle(Cell cell, CellStyle style) {
        if (style.getAlignment() == ALIGN_GENERAL) {
            switch (ultimateCellType(cell)) {
            case HSSFCell.CELL_TYPE_STRING:
                return "style=\"text-align: left;\"";
            case HSSFCell.CELL_TYPE_BOOLEAN:
            case HSSFCell.CELL_TYPE_ERROR:
                return "style=\"text-align: center;\"";
            case HSSFCell.CELL_TYPE_NUMERIC:
            default:
                // "right" is the default
                break;
            }
        }
        return "";
    }
    private  Map<String,String>[] getRowSpanColSpanMap(Sheet sheet){
    	  
    	  Map<String,String> map0 = new HashMap<String,String>();
    	  Map<String,String> map1 = new HashMap<String,String>();

    	  int mergedNum = sheet.getNumMergedRegions();
    	  
    	  CellRangeAddress range = null;
    	  
    	  for(int i = 0; i < mergedNum; i ++){
    	  
    	   range = sheet.getMergedRegion(i);
    	   
    	   int topRow = range.getFirstRow();
    	   
    	   int topCol = range.getFirstColumn();
    	   
    	   int bottomRow = range.getLastRow();
    	   
    	   int bottomCol = range.getLastColumn();
    	   
    	   map0.put(topRow + "," + topCol, bottomRow + "," + bottomCol);
    	   
    	   //System.out.println(topRow + "," + topCol + "," + bottomRow + "," + bottomCol);
    	   
    	   int tempRow = topRow;
    	   
    	   while(tempRow <= bottomRow ){

    	    int tempCol = topCol;
    	    
    	    while(tempCol <= bottomCol ){
    	     
    	     map1.put(tempRow + "," + tempCol,"");
    	     
    	     tempCol ++;
    	    }
    	    
    	    tempRow ++;
    	   }
    	   
    	   map1.remove(topRow + "," + topCol);
    	   
    	  }
    	  
    	  Map[] map = {map0,map1};
    	  
    	  return map;
    	 }
}
