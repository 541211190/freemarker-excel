
package com.study.dto;

import java.io.Serializable;
import java.util.List;

import lombok.Data;

/**
 * @project cne-power-operation-facade
 * @description: 导出Excel：申请单
 * @author 大脑补丁
 * @create 2020-03-26 15:26
 */
@Data
public class SendBillOutput implements Serializable {

    // 客户名称
    private String customerName;
    // 是否一般纳税人
    private String isGeneralTaxpayer;
    // 税号
    private String taxNumber;
    // 客户公司地址及电话
    private String addressAndPhone;
    // 开户银行和账号
    private String bankAndAccount;
    // 信息列表
    private List<StationBillOutput> stationBillList;
    // 合计栏
    private StationAmountOutput stationAmount;

}
