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
        ExportExcel excel = context.getBean(ExportExcel.class);
        // excel.export();
        ExportImageExcel imageExcel = context.getBean(ExportImageExcel.class);
        imageExcel.export();
    }

}
