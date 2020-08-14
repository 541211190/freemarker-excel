
package com.study.dto.output;

import lombok.Data;

import java.io.Serializable;
import java.math.BigDecimal;

/**
 * @project cne-power-operation-facade
 * @description: 导出Excel-合计行
 * @author 大脑补丁
 * @create 2020-03-31 18:23
 */
@Data
public class StationAmountOutput implements Serializable {
    private BigDecimal power;
    private BigDecimal noTaxMoney;
    private BigDecimal taxAmount;
    private BigDecimal taxmoney;
}
