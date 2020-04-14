
package com.study.dto;

import java.io.Serializable;
import java.math.BigDecimal;

import lombok.Data;

/**
 * @project cne-power-operation-facade
 * @description: 电费申请单Excel中尖峰平谷部分
 * @author XuChao
 * @create 2020-03-30 10:30
 */
@Data
public class PeriodPowerOutput implements Serializable {
    // 尖、峰、平、谷
    private String powerName;
    // 电量(尖、峰、平、谷、合计)
    private BigDecimal power;
    // 含税电价(尖、峰、平、谷、合计)
    private BigDecimal price;
    // 不含税金额(尖、峰、平、谷、合计)
    private BigDecimal noTaxMoney;
    // 税率(尖、峰、平、谷、合计)
    private Integer taxRate;
    // 税额(尖、峰、平、谷、合计)
    private BigDecimal taxAmount;
    // 含税金额(尖、峰、平、谷、合计)
    private BigDecimal taxmoney;
}
