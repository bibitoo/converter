package cc.landking.converter.word2image;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.UUID;

import javax.imageio.IIOImage;
import javax.imageio.ImageIO;
import javax.imageio.ImageWriter;
import javax.imageio.stream.ImageOutputStream;

import org.apache.commons.io.FilenameUtils;
import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.document.DefaultDocumentFormatRegistry;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.ExternalOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;
import org.icepdf.core.exceptions.PDFException;
import org.icepdf.core.exceptions.PDFSecurityException;
import org.icepdf.core.pobjects.Document;
import org.icepdf.core.pobjects.Page;
import org.icepdf.core.util.GraphicsRenderingHints;

public class SofficeService {
	
	private float zoom = 3f;
	


	public void setZoom(float zoom) {
		this.zoom = zoom;
	}

	private String outPathBase;
	
	
	
	public void setOutPathBase(String outPathBase) {
		this.outPathBase = outPathBase;
	}

	private boolean externalProcess = false;
	
	
	
	public void setExternalProcess(boolean externalProcess) {
		this.externalProcess = externalProcess;
	}

	private OfficeManager officeManager;
	
	
	public OfficeManager getOfficeManager() {
		return officeManager;
	}
	public void setOfficeManager(OfficeManager officeManager) {
		this.officeManager = officeManager;
	}

	private int port = 2002;
	
	
	public void setPort(int port) {
		this.port = port;
	}
	
	private String officeHome ;
	
	
	public void setOfficeHome(String officeHome) {
		this.officeHome = officeHome;
	}
	
	private long taskExecutionTimeout = 1000*10;
	
	
	public void setTaskExecutionTimeout(long taskExecutionTimeout) {
		this.taskExecutionTimeout = taskExecutionTimeout;
	}
	
	private long taskQueueTimeout = 1000*60*60*24L;
	
	
	public void setTaskQueueTimeout(long taskQueueTimeout) {
		this.taskQueueTimeout = taskQueueTimeout;
	}
	public void init(){
		 if(isExternalProcess()){
			 initFromExistingOOInstance();
		 }else{
			 initFromNewOOInstance();
		 }
		
	}
	public void destroy(){
		stopOfficeManager();
	}
	
	 private void initFromExistingOOInstance() {
         ExternalOfficeManagerConfiguration extConf = new
ExternalOfficeManagerConfiguration();
         extConf.setConnectOnStart(true);
         extConf.setPortNumber( this.port );
         this.officeManager = extConf.buildOfficeManager();
         this.officeManager.start();
         this.externalProcess = true;
         System.out.println("Attached to existing OpenOffice process ... ");
 }
	 protected String getTimePath(){
		 SimpleDateFormat format = new SimpleDateFormat("yyyy/MM");
		 return format.format(new Date());
	 }
	 public File convertToPdf(String sourcePath) throws IOException, PDFException, PDFSecurityException{
		 return convertToPdf( sourcePath, null);
	 }
	 public File convertToPdf(String sourcePath,String targetPath) throws IOException, PDFException, PDFSecurityException{
		 DocumentFormat format = new DefaultDocumentFormatRegistry().getFormatByExtension("pdf");
		 File inputFile = new File(sourcePath);
		 File outputFile = null;
		 if(targetPath == null && outPathBase!= null){
			 outputFile = new File(outPathBase+"/"+getTimePath()+"/"+UUID.randomUUID());
		 }else if(targetPath != null){
			 outputFile = new File(targetPath);
		 }else{
			 outputFile = new File("/tmp/"+getTimePath()+"/"+UUID.randomUUID());
		 }
		 getDocumentConverter().convert(inputFile, outputFile,format);
		 
		 return outputFile;
	 }
	 public List<File> convertToJpg(String sourcePath,float zoom) throws IOException, PDFException, PDFSecurityException{
		 return convertToJpg( sourcePath, null,zoom);
	 }
	 public List<File> convertToJpg(String sourcePath) throws IOException, PDFException, PDFSecurityException{
		 return convertToJpg( sourcePath, null,this.zoom);
	 }
	 public List<File> convertToJpg(String sourcePath,String targetPath) throws IOException, PDFException, PDFSecurityException{
		 return convertToJpg( sourcePath, targetPath,this.zoom);
	 }

	 
		 public List<File> convertToJpg(String sourcePath,String targetPath,float zoom) throws IOException, PDFException, PDFSecurityException{
		 DocumentFormat format = new DefaultDocumentFormatRegistry().getFormatByExtension("pdf");
		 File inputFile = new File(sourcePath);
		 File outputFile = null;
		 if(targetPath == null && outPathBase!= null){
			 outputFile = new File(outPathBase+"/"+getTimePath()+"/"+UUID.randomUUID());
		 }else if(targetPath != null){
			 outputFile = new File(targetPath);
		 }else{
			 outputFile = new File("/tmp/"+getTimePath()+"/"+UUID.randomUUID());
		 }
		 File pdfFile = File.createTempFile("lk_conv", "pdf");
		 getDocumentConverter().convert(inputFile, pdfFile,format);
		 List<File> retval = SofficeService.tranfer(pdfFile, outputFile.getAbsolutePath(), zoom);
		 return retval;
	 }
		public static List<File> tranfer(File sourceFile, String destFile,float zoom)throws PDFException, PDFSecurityException, IOException {
			List<File> retval = new ArrayList<File>(); 
			String FILETYPE_PNG = "jpeg";
			String FileName = FilenameUtils.getName(sourceFile.getName());
		
		Document document = null;
		BufferedImage img = null;
		float rotation = 0f;
		
		//判断目录是否存在，如果不存在的话则创建
		File file = new File(destFile);
		if (!file.exists()) {
			file.mkdirs();
		}
		
		File inputFile = sourceFile;  
		if (!inputFile.exists()) {  
//			System.out.println("找不到源文件");
//			return -1;// 找不到源文件, 则返回-1
			return retval;
		}  
		document = new Document();

		document.setFile(sourceFile.getAbsolutePath());

		// maxPages = document.getPageTree().getNumberOfPages();
		
		//进行pdf文件图片的转化
		for (int i = 0; i < document.getNumberOfPages(); i++) {
			img = (BufferedImage) document.getPageImage(i,GraphicsRenderingHints.SCREEN,
					Page.BOUNDARY_CROPBOX,rotation,zoom);
			//设置图片的后缀名
			Iterator iter = ImageIO.getImageWritersBySuffix(FILETYPE_PNG);
			
			ImageWriter writer = (ImageWriter) iter.next();
			
			File outFile = new File(destFile+"/"+FileName+"_"+(i+1)+".jpeg");
			
			FileOutputStream out = new FileOutputStream(outFile);
			
			ImageOutputStream outImage = ImageIO.createImageOutputStream(out);
			
			writer.setOutput(outImage);
			
			writer.write(new IIOImage(img, null, null));
			retval.add(outFile);
		}
		img.flush();
		document.dispose();
		
		return retval;
	}

 /**
  * Start a new openoffice instance and create the office manager
  */
 private void initFromNewOOInstance() {
         DefaultOfficeManagerConfiguration defaultConf = new
DefaultOfficeManagerConfiguration();
         defaultConf.setPortNumber( this.port );
         if(officeHome != null){
        	 defaultConf.setOfficeHome(officeHome);
		 }
         defaultConf.setPortNumber(port);
         defaultConf.setTaskExecutionTimeout(taskExecutionTimeout);
         defaultConf.setTaskQueueTimeout(taskQueueTimeout);
   
         this.officeManager = defaultConf.buildOfficeManager();
         this.officeManager.start();
         this.externalProcess = false;
         System.out.println("Created a new OpenOffice process ... ");
 }



 /**
  * @return the externalProcess
  */
 public boolean isExternalProcess() {
         return externalProcess;
 }

 /**
  * Get a new document converter.
  * @return
  */
 public OfficeDocumentConverter getDocumentConverter() {
         OfficeDocumentConverter docConverter = null;
         try  {
                 docConverter = new OfficeDocumentConverter ( this.officeManager );
         } catch (Exception e) {
                 e.printStackTrace();
         }
         return docConverter;
 }

 /**
  * If is externalProcess, officeManager simply disconnects from the
process
  * else it stops the OpenOffice instance.
  */
 public void stopOfficeManager() {
         try {
                 if ( this.officeManager != null  && !isExternalProcess() )
                         this.officeManager.stop();
         } catch ( IllegalStateException e) {
                 e.printStackTrace();
         }
 }

}
