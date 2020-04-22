package com.study.dto.freemarker.input;

import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;

import java.io.Serializable;

@Data
@AllArgsConstructor
public class ExcelImageInput implements Serializable {

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

}
