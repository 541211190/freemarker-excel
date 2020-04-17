
package com.study.commons.utils;

import java.math.BigDecimal;

/**
 * @project cne-power-commons-util
 * @description: BigDecimal运算工具类
 * @author 大脑补丁
 * @create 2020-03-30 14:44
 */
public class DecimalUtils {

    /**
     * 加法计算（result = x + y）
     * 
     * @param x
     *            被加数（可为null）
     * @param y
     *            加数 （可为null）
     * @return 和 （可为null）
     * @author 大脑补丁 on 2020-03-30 14:52
     */
    public static BigDecimal add(BigDecimal x, BigDecimal y) {
        if (x == null) {
            return y;
        }
        if (y == null) {
            return x;
        }
        return x.add(y);
    }

    /**
     * 加法计算（result = a + b + c + d）
     * 
     * @param a
     *            被加数（可为null）
     * @param b
     *            加数（可为null）
     * @param c
     *            加数（可为null）
     * @param d
     *            加数（可为null）
     * @return BigDecimal （可为null）
     * @author 大脑补丁 on 2020-03-30 22:17
     */
    public static BigDecimal add(BigDecimal a, BigDecimal b, BigDecimal c, BigDecimal d) {
        BigDecimal ab = add(a, b);
        BigDecimal cd = add(c, d);
        return add(ab, cd);
    }

    /**
     * 累加计算(result=x + result)
     * 
     * @param x
     *            被加数（可为null）
     * @param result
     *            和 （可为null,若被加数不为为null，result默认值为0）
     * @return result 和 （可为null）
     * @author 大脑补丁 on 2020-03-30 14:16
     */
    public static BigDecimal accumulate(BigDecimal x, BigDecimal result) {
        if (x == null) {
            return result;
        }
        if (result == null) {
            result = new BigDecimal("0");
        }
        return result.add(x);
    }

    /**
     * 减法计算(result = x - y)
     * 
     * @param x
     *            被减数（可为null）
     * @param y
     *            减数（可为null）
     * @return BigDecimal 差 （可为null）
     * @author 大脑补丁 on 2020-03-30 14:47
     */
    public static BigDecimal subtract(BigDecimal x, BigDecimal y) {
        if (x == null || y == null) {
            return null;
        }
        return x.subtract(y);
    }

    /**
     * 乘法计算(result = x × y)
     * 
     * @param x
     *            乘数(可为null)
     * @param y
     *            乘数(可为null)
     * @return BigDecimal 积
     * @author 大脑补丁 on 2020-03-30 14:14
     */
    public static BigDecimal multiply(BigDecimal x, BigDecimal y) {
        if (x == null || y == null) {
            return null;
        }
        return x.multiply(y);
    }

    /**
     * 除法计算(result = x ÷ y)
     * 
     * @param x
     *            被除数（可为null）
     * @param y
     *            除数（可为null）
     * @return 商 （可为null,四舍五入，默认保留20位小数）
     * @author 大脑补丁 on 2020-03-30 14:15
     */
    public static BigDecimal divide(BigDecimal x, BigDecimal y) {
        if (x == null || y == null || y.compareTo(BigDecimal.ZERO) == 0) {
            return null;
        }
        return x.divide(y, 20, BigDecimal.ROUND_HALF_UP);
    }

    /**
     * 转为字符串
     * 
     * @param x
     *            要转字符串的小数
     * @return String
     * @author 大脑补丁 on 2020-03-30 15:08
     */
    public static String toPlainString(BigDecimal x) {
        if (x == null) {
            return null;
        }
        return x.toPlainString();
    }

    /**
     * 保留小数位数
     * 
     * @param x
     *            目标小数
     * @param scale
     *            要保留小数位数
     * @return BigDecimal 结果四舍五入
     * @author 大脑补丁 on 2020-03-30 15:17
     */
    public static BigDecimal scale(BigDecimal x, int scale) {
        if (x == null) {
            return null;
        }
        return x.setScale(scale, BigDecimal.ROUND_HALF_UP);
    }

    /**
     * 整型转为BigDecimal
     * 
     * @param x(可为null)
     * @return BigDecimal (可为null)
     * @author 大脑补丁 on 2020-03-30 22:24
     */
    public static BigDecimal toBigDecimal(Integer x) {
        if (x == null) {
            return null;
        }
        return new BigDecimal(x.toString());
    }

    /**
     * 长整型转为BigDecimal
     * 
     * @param x(可为null)
     * @return BigDecimal (可为null)
     * @author 大脑补丁 on 2020-03-30 22:24
     */
    public static BigDecimal toBigDecimal(Long x) {
        if (x == null) {
            return null;
        }
        return new BigDecimal(x.toString());
    }

    /**
     * 双精度型转为BigDecimal
     * 
     * @param x(可为null)
     * @return BigDecimal (可为null)
     * @author 大脑补丁 on 2020-03-30 22:24
     */
    public static BigDecimal toBigDecimal(Double x) {
        if (x == null) {
            return null;
        }
        return new BigDecimal(x.toString());
    }

    /**
     * 单精度型转为BigDecimal
     * 
     * @param x(可为null)
     * @return BigDecimal (可为null)
     * @author 大脑补丁 on 2020-03-30 22:24
     */
    public static BigDecimal toBigDecimal(Float x) {
        if (x == null) {
            return null;
        }
        return new BigDecimal(x.toString());
    }

    /**
     * 倍数计算，用于单位换算
     * 
     * @param x
     *            目标数(可为null)
     * @param multiple
     *            倍数 (可为null)
     * @return BigDecimal (可为null)
     * @author 大脑补丁 on 2020-03-30 22:43
     */
    public static BigDecimal multiple(BigDecimal x, Integer multiple) {
        if (x == null || multiple == null) {
            return null;
        }
        return DecimalUtils.multiple(x, multiple);
    }

}
