调用方式见Excel2Html类中的main方法。
实例如下：
StringBuffer stringOut = new StringBuffer();
ExcelToHtml toHtml = create(new FileInputStream(args[0]), stringOut);
        toHtml.setCompleteHTML(false);//if true,output the html\title and body tag
        toHtml.setSheetBorderOut(false);//if false,do not output sheet boder ,top A,B,C...and left row number
        toHtml.setStyleOut(false);//if false, do not output style 
        String resultImagePath = "result/images/data/";//设置的这个文件夹必须能有权限读写
        toHtml.setResultImagePath(resultImagePath);//图片保存路径
        toHtml.printPage();

String result = stringOut.toString();

		<dependency>
			<groupId>cc.landking.converter</groupId>
			<artifactId>office</artifactId>
			<version>0.0.1-SNAPSHOT</version>
		</dependency>

