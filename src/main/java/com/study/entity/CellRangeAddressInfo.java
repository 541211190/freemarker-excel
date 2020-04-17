package com.study.entity;

import lombok.Data;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

/**
 * 合并单元格信息
 * @project freemarker-excel 
 * @description: TODO 
 * @author 大脑补丁
 * @create 2020-04-14 16:54
 */
@Data
public class CellRangeAddressInfo {


    private CellRangeAddress cellRangeAddress;

    private List<Style.Border> borders;

}
