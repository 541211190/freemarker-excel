package com.study.entity.excel;

import lombok.Data;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

/**
 * @project freemarker-excel
 * @description: 合并单元格信息
 * @author 大脑补丁
 * @create 2020-04-14 16:54
 */
@Data
public class CellRangeAddressEntity {

    private CellRangeAddress cellRangeAddress;

    private List<Style.Border> borders;

}
