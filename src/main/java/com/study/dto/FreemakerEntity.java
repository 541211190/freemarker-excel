package com.study.dto;

import java.util.List;
import java.util.Map;

import lombok.Data;

/**
 * 
 * @project freemarker-excel
 * @description: FreeMarker导出带图片的Excel需要的参数对象
 * @author XuChao
 * @create 2020-04-14 14:21
 */
@Data
public class FreemakerEntity {
    /**
     * 加载数据
     */
    private Map dataMap;
    /**
     * 模版名称
     */
    private String templateName;
    /**
     * 模版路径
     */
    private String templateFilePath;
    /**
     * 生成文件名称
     */
    private String fileName;

    /**
     * 历史存放xml文件路径
     */
    private String temporaryXmlfile;

    /**
     * 加载图片信息
     */
    List<ExcelImageLoadDTO> excelImageLoadDTOs;

}
