package com.study;

import com.study.demo.ExportCommentExcel;
import com.study.demo.ExportExcel;
import com.study.demo.ExportImageExcel;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ConfigurableApplicationContext;

@SpringBootApplication
public class FreemarkerExcelApplication {

	public static void main(String[] args) {
		ConfigurableApplicationContext context = SpringApplication.run(FreemarkerExcelApplication.class, args);

		// 1.Freemarker导出xml格式复杂的Excel示例
		ExportExcel excel = context.getBean(ExportExcel.class);
		excel.export();

		// 2.Freemarker导出带有图片的Excel示例
		ExportImageExcel imageExcel = context.getBean(ExportImageExcel.class);
		//Excel 2003+ 版本（有弹框提示数据损坏，兼容性不好）
		imageExcel.export2003();
		//Excel 2007+ 版本(推荐使用，兼容性好，性能佳)
		imageExcel.export2007();

		// 3.导出带有注释的Excel
		ExportCommentExcel exportCommentExcel = new ExportCommentExcel();
		// Excel 2007+ 版本(推荐使用)
		exportCommentExcel.export();

	}

}
