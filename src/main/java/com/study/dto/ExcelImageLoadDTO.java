package com.study.dto;

import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;

import java.io.Serializable;

@Data
public class ExcelImageLoadDTO implements Serializable {

    /**
     * 图片地址
     */
    private String imgPath;

    /**
     * sheet索引
     */
    private Integer sheetIndex;

    /**
     * 位置
     */
    private HSSFClientAnchor anchor;

    public ExcelImageLoadDTO(String imgPath, Integer sheetIndex, HSSFClientAnchor anchor) {
        this.imgPath = imgPath;
        this.sheetIndex = sheetIndex;
        this.anchor = anchor;
    }
}
