
word转pdf和图片
依赖包在lib.zip中

转换成pdf:
SofficeService service = (SofficeService) context.getBean("sofficeService");
	        File file = service.convertToPdf("test.docx");
转换图片:
SofficeService service = (SofficeService) context.getBean("sofficeService");
	        List<File> files = service.convertToJpg("test.docx");
	        或者 
	      List<File> files = service.convertToJpg("test.docx",1.5f);  
	      1.5f为放大倍数