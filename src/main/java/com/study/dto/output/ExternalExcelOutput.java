package com.study.dto.output;

import lombok.Data;

import java.io.Serializable;

/**
 * @description: 外部项目Excel返回值
 * @create: 2020-08-05 17:48
 */
@Data
public class ExternalExcelOutput implements Serializable {

	/**
	 * 发票日期(后一个月最后一天)
	 */
	private String invoiceDate;

	/**
	 * 租赁编号
	 */
	private String leaseNumber;

	/**
	 * 结算单描述
	 */
	private String description;

	/**
	 * 账单编码
	 */
	private String billCode;

	/**
	 * 增值税税额 Tax-VAT。（公式：taxable_amount / tax_rate / 100）
	 */
	private String taxVat;

	/**
	 * 增值税税率 Tax rate-VAT
	 */
	private String taxRateVat;

	/**
	 * 增值税编码
	 */
	private String taxCode;

	/**
	 * 税额
	 */
	private String taxableAmount;

	/**
	 * 新能源费用日期(后一个月最后一天)
	 */
	private String glDate;

	/**
	 * 到期日期(后一个月最后一天)
	 */
	private String dueDate;

	/**
	 * 总金额（页面保留两位小数）。(公式=grossAmount/(1+taxRateVat/100))
	 */
	private String grossAmount;

}