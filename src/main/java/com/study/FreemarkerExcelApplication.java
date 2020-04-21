package com.study;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ConfigurableApplicationContext;

import com.study.demo.ExportExcel;
import com.study.demo.ExportImageExcel;

@SpringBootApplication
public class FreemarkerExcelApplication {

    public static void main(String[] args) {
        ConfigurableApplicationContext context = SpringApplication.run(FreemarkerExcelApplication.class, args);
        // 导出普通Excel
        ExportExcel excel = context.getBean(ExportExcel.class);
//        excel.export();
        // 导出带有图片的Excel
        ExportImageExcel imageExcel = context.getBean(ExportImageExcel.class);
        imageExcel.export();
    }

}
