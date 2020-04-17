
package com.study.dto;

import java.io.Serializable;
import java.util.List;

import lombok.Data;

/**
 * @project cne-power-operation-facade
 * @description: 结算申请单导出Excel--单电站信息
 * @author 大脑补丁
 * @create 2020-03-30 10:30
 */
@Data
public class StationBillOutput implements Serializable {
    // 发票数量
    // private Integer invoiceCount;
    // 描述
    private String description;
    // 计费周期
    private String period;
    // 尖峰平谷
    private List<PeriodPowerOutput> periodPowerList;
    // 园区地址
    private String stationName;
    // 发票号码
    private String invoiceNumber;
}
