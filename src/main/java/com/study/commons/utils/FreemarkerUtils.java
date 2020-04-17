
package com.study.commons.utils;

import java.awt.image.BufferedImage;
import java.io.BufferedWriter;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import javax.imageio.ImageIO;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.dom4j.Document;
import org.dom4j.io.SAXReader;
import org.springframework.util.CollectionUtils;
import org.springframework.util.ObjectUtils;

import com.study.dto.ExcelImageLoadDTO;
import com.study.dto.FreemakerEntity;
import com.study.entity.Cell;
import com.study.entity.CellRangeAddressInfo;
import com.study.entity.Column;
import com.study.entity.Row;
import com.study.entity.Style;
import com.study.entity.Table;
import com.study.entity.Worksheet;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateExceptionHandler;
import lombok.extern.slf4j.Slf4j;

/**
 * @project freemarker-excel
 * @description: freemarker工具类
 * @author 大脑补丁
 * @create 2020-04-14 09:43
 */
@Slf4j
public class FreemarkerUtils {

    
   
    
    /**
     * 导出Excel到指定文件
     * 
     * @param dataMap
     *            数据源
     * @param templateName
     *            模板名称（包含文件后缀名.ftl）
     * @param templateFilePath
     *            模板所在路径（不能为空，当前路径传空字符：""）
     * @param fileFullPath
     *            文件完整路径（如：usr/local/fileName.xls）
     * @author 大脑补丁 on 2020-04-05 11:51
     */
    @SuppressWarnings("rawtypes")
    public static void exportToFile(Map dataMap, String templateName, String templateFilePath, String fileFullPath) {
        try {
            File file = FileUtils.createFile(fileFullPath);
            FileOutputStream outputStream = new FileOutputStream(file);
            exportToStream(dataMap, templateName, templateFilePath, outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    /**
     * 导出Excel到输出流
     * 
     * @param dataMap
     *            数据源
     * @param templateName
     *            模板名称（包含文件后缀名.ftl）
     * @param templateFilePath
     *            模板所在路径（不能为空，当前路径传空字符：""）
     * @param outputStream
     *            输出流
     * @author 大脑补丁 on 2020-04-05 11:52
     */
    @SuppressWarnings("rawtypes")
    public static void exportToStream(Map dataMap, String templateName, String templateFilePath,
        FileOutputStream outputStream) {
        try {
            Template template = getTemplate(templateName, templateFilePath);
            OutputStreamWriter outputWriter = new OutputStreamWriter(outputStream, "UTF-8");
            Writer writer = new BufferedWriter(outputWriter);
            template.process(dataMap, writer);
            writer.flush();
            writer.close();
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 导出到文件中
     * 
     * @param excelFilePath
     * @param freemakerEntity
     * @author 大脑补丁 on 2020-04-14 15:34
     */
    public static void exportImageExcel(String excelFilePath, FreemakerEntity freemakerEntity) {
        File file = FileUtils.createFile(excelFilePath);
        try {
            FileOutputStream outputStream = new FileOutputStream(file);
            createImageExcleToStream(freemakerEntity, outputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

    }

    /**
     * 导出到response输出流中
     * 
     * @param excelFilePath
     * @param freemakerEntity
     * @author 大脑补丁 on 2020-04-14 15:34
     */
    public static void exportImageExcel(HttpServletResponse response, FreemakerEntity freemakerEntity) {
        try {
            OutputStream outputStream = response.getOutputStream();
            // 写入excel文件
            response.reset();
            response.setContentType("application/msexcel;charset=UTF-8");
            response.setHeader("Content-Disposition", "attachment;filename=\""
                + new String((freemakerEntity.getFileName() + ".xls").getBytes("GBK"), "ISO8859-1") + "\"");
            createImageExcleToStream(freemakerEntity, outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
   

    // 获取项目templates文件夹下的模板
    private static Template getTemplate(String templateName, String filePath) throws IOException {
        Configuration configuration = new Configuration(Configuration.VERSION_2_3_28);
        configuration.setDefaultEncoding("UTF-8");
        configuration.setTemplateUpdateDelayMilliseconds(0);
        configuration.setEncoding(Locale.CHINA, "UTF-8");
        configuration.setTemplateExceptionHandler(TemplateExceptionHandler.RETHROW_HANDLER);
        configuration.setClassForTemplateLoading(FreemarkerUtils.class, "/templates" + filePath);
        configuration.setOutputEncoding("UTF-8");
        return configuration.getTemplate(templateName, "UTF-8");
    }

   

    private static void createImageExcleToStream(FreemakerEntity freemakerEntity, OutputStream outputStream) {
        Writer out = null;
        try {
            // 创建xml文件
            Template template = getTemplate(freemakerEntity.getTemplateName(), freemakerEntity.getTemplateFilePath());
            File outFile = new File(freemakerEntity.getTemporaryXmlfile() + freemakerEntity.getFileName() + ".xml");
            out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile), "UTF-8"));
            template.process(freemakerEntity.getDataMap(), out);
            log.debug("=======》开始写入xml完成");
            System.out.println("=======》开始写入xml完成");
            SAXReader reader = new SAXReader();
            Document document = reader.read(outFile);
            Map<String, Style> styleMap = readXmlStyle(document);
            log.debug("=======》读取样式：" + styleMap.toString());
            List<Worksheet> worksheets = readXmlWorksheet(document);
            log.debug("=======》开始写入excel：" + worksheets.toString());
            System.out.println("=======》开始写入excel：" + worksheets.toString());
            HSSFWorkbook wb = new HSSFWorkbook();
            for (Worksheet worksheet : worksheets) {
                HSSFSheet sheet = wb.createSheet(worksheet.getName());
                Table table = worksheet.getTable();
                List<Row> rows = table.getRows();
                List<Column> columns = table.getColumns();
                // 填充列宽
                int columnIndex = 0;
                for (int i = 0; i < columns.size(); i++) {
                    Column column = columns.get(i);
                    columnIndex = getCellWidthIndex(columnIndex, i, column.getIndex());
                    sheet.setColumnWidth(columnIndex, (int)column.getWidth() * 50);
                }

                int createRowIndex = 0;
                List<CellRangeAddressInfo> cellRangeAddresses = new ArrayList<>();
                for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
                    Row rowInfo = rows.get(rowIndex);
                    if (rowInfo == null) {
                        continue;
                    }
                    createRowIndex = getIndex(createRowIndex, rowIndex, rowInfo.getIndex());
                    HSSFRow row = sheet.createRow(createRowIndex);
                    if (rowInfo.getHeight() != null) {
                        Integer height = rowInfo.getHeight() * 20;
                        row.setHeight(height.shortValue());
                    }
                    List<Cell> cells = rowInfo.getCells();
                    if (CollectionUtils.isEmpty(cells)) {
                        continue;
                    }
                    int startIndex = 0;
                    for (int cellIndex = 0; cellIndex < cells.size(); cellIndex++) {
                        Cell cellInfo = cells.get(cellIndex);
                        if (cellInfo == null) {
                            continue;
                        }
                        // 获取起始列
                        startIndex = getIndex(startIndex, cellIndex, cellInfo.getIndex());
                        HSSFCell cell = row.createCell(startIndex);
                        String styleID = cellInfo.getStyleID();
                        Style style = styleMap.get(styleID);
                        /*设置数据单元格格式*/
                        CellStyle dataStyle = wb.createCellStyle();
                        // 设置边框样式
                        setBorder(style, dataStyle);
                        // 设置对齐方式
                        setAlignment(style, dataStyle);
                        // 填充文本
                        setValue(wb, cellInfo, cell, style, dataStyle);
                        // 填充颜色
                        setCellColor(style, dataStyle);
                        cell.setCellStyle(dataStyle);
                        // 合并单元格
                        startIndex = getCellRanges(createRowIndex, cellRangeAddresses, startIndex, cellInfo, style);
                    }
                }
                // 添加合并单元格
                addCellRange(sheet, cellRangeAddresses);
            }
            // 加载图片到excel
            log.debug("=======》加载图片：" + freemakerEntity.getExcelImageLoadDTOs());
            System.out.println("=======》加载图片：" + freemakerEntity.getExcelImageLoadDTOs());
            if (!CollectionUtils.isEmpty(freemakerEntity.getExcelImageLoadDTOs())) {
                loadImage2Excel(freemakerEntity.getExcelImageLoadDTOs(), wb);
            }
            log.debug("=======》加载图片完成：" + freemakerEntity.getExcelImageLoadDTOs());
            // 写入excel文件,response字符流转换成字节流，template需要字节流作为输出
            wb.write(outputStream);
            outputStream.close();
            outFile.delete();
            System.out.println("=======》导出成功");
        } catch (Exception e) {
            e.printStackTrace();
            log.error("=======》写入excel异常：" + e.getMessage());
        } finally {
            try {
                out.close();
            } catch (Exception e) {

            }
        }
    }

    public static Map<String, Style> readXmlStyle(Document document) {
        Map<String, Style> styleMap = XmlReader.getStyle(document);
        return styleMap;
    }

    public static List<Worksheet> readXmlWorksheet(Document document) {
        List<Worksheet> worksheets = XmlReader.getWorksheet(document);
        return worksheets;
    }

    private static int getIndex(int columnIndex, int i, Integer index) {
        if (index != null) {
            columnIndex = index - 1;
        }
        if (index == null && columnIndex != 0) {
            columnIndex = columnIndex + 1;
        }
        if (index == null && columnIndex == 0) {
            columnIndex = i;
        }
        return columnIndex;
    }

    private static int getCellWidthIndex(int columnIndex, int i, Integer index) {
        if (index != null) {
            columnIndex = index;
        }
        if (index == null && columnIndex != 0) {
            columnIndex = columnIndex + 1;
        }
        if (index == null && columnIndex == 0) {
            columnIndex = i;
        }
        return columnIndex;
    }

    /**
     * 设置边框
     * 
     * @param style:
     * @param dataStyle:
     * @return void
     */
    private static void setBorder(Style style, CellStyle dataStyle) {
        if (style != null && style.getBorders() != null) {
            for (int k = 0; k < style.getBorders().size(); k++) {
                Style.Border border = style.getBorders().get(k);
                if (border != null) {
                    if ("Bottom".equals(border.getPosition())) {
                        dataStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
                        dataStyle.setBorderBottom(BorderStyle.THIN);
                    }
                    if ("Left".equals(border.getPosition())) {
                        dataStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
                        dataStyle.setBorderLeft(BorderStyle.THIN);
                    }
                    if ("Right".equals(border.getPosition())) {
                        dataStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
                        dataStyle.setBorderRight(BorderStyle.THIN);
                    }
                    if ("Top".equals(border.getPosition())) {
                        dataStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
                        dataStyle.setBorderTop(BorderStyle.THIN);
                    }
                }

            }
        }
    }

    private static void loadImage2Excel(List<ExcelImageLoadDTO> excelImageLoadDTOs, HSSFWorkbook wb)
        throws IOException {
        BufferedImage bufferImg = null;
        if (!CollectionUtils.isEmpty(excelImageLoadDTOs)) {
            for (ExcelImageLoadDTO excelImageLoadDTO : excelImageLoadDTOs) {
                Sheet sheet = wb.getSheetAt(excelImageLoadDTO.getSheetIndex());
                if (sheet == null) {
                    continue;
                }
                // 画图的顶级管理器，一个sheet只能获取一个（一定要注意这点）
                Drawing patriarch = sheet.createDrawingPatriarch();
                // anchor主要用于设置图片的属性
                HSSFClientAnchor anchor = excelImageLoadDTO.getAnchor();
                anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
                // 插入图片
                String imagePath = excelImageLoadDTO.getImgPath();
                ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
                bufferImg = ImageIO.read(new File(imagePath));
                String imageType = imagePath.substring(imagePath.lastIndexOf(".") + 1, imagePath.length());
                ImageIO.write(bufferImg, imageType, byteArrayOut);
                patriarch.createPicture(anchor,
                    wb.addPicture(byteArrayOut.toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));
            }
        }
    }

    /**
     * 添加合并单元格
     * 
     * @param sheet:
     * @param cellRangeAddresses:
     * @return void
     */
    private static void addCellRange(HSSFSheet sheet, List<CellRangeAddressInfo> cellRangeAddresses) {
        if (!CollectionUtils.isEmpty(cellRangeAddresses)) {
            for (CellRangeAddressInfo cellRangeAddressInfo : cellRangeAddresses) {
                CellRangeAddress cellRangeAddress = cellRangeAddressInfo.getCellRangeAddress();
                sheet.addMergedRegion(cellRangeAddress);
                if (CollectionUtils.isEmpty(cellRangeAddressInfo.getBorders())) {
                    continue;
                }
                for (int k = 0; k < cellRangeAddressInfo.getBorders().size(); k++) {
                    Style.Border border = cellRangeAddressInfo.getBorders().get(k);
                    if (border == null) {
                        continue;
                    }
                    if ("Bottom".equals(border.getPosition())) {
                        RegionUtil.setBorderBottom(BorderStyle.THIN, cellRangeAddress, sheet);
                    }
                    if ("Left".equals(border.getPosition())) {
                        RegionUtil.setBorderLeft(BorderStyle.THIN, cellRangeAddress, sheet);
                    }
                    if ("Right".equals(border.getPosition())) {
                        RegionUtil.setBorderRight(BorderStyle.THIN, cellRangeAddress, sheet);
                    }
                    if ("Top".equals(border.getPosition())) {
                        RegionUtil.setBorderTop(BorderStyle.THIN, cellRangeAddress, sheet);
                    }
                }
            }
        }
    }

    /**
     * 设置对齐方式
     * 
     * @param style:
     * @param dataStyle:
     * @return void
     */
    private static void setAlignment(Style style, CellStyle dataStyle) {
        if (style != null && style.getAlignment() != null) {
            // 设置水平对齐方式
            String horizontal = style.getAlignment().getHorizontal();
            if (!ObjectUtils.isEmpty(horizontal)) {
                if ("Left".equals(horizontal)) {
                    dataStyle.setAlignment(HorizontalAlignment.LEFT);
                } else if ("Center".equals(horizontal)) {
                    dataStyle.setAlignment(HorizontalAlignment.CENTER);
                } else {
                    dataStyle.setAlignment(HorizontalAlignment.RIGHT);
                }
            }

            // 设置垂直对齐方式
            String vertical = style.getAlignment().getVertical();
            if (!ObjectUtils.isEmpty(vertical)) {
                if ("Top".equals(vertical)) {
                    dataStyle.setVerticalAlignment(VerticalAlignment.TOP);
                } else if ("Center".equals(vertical)) {
                    dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                } else if ("Bottom".equals(vertical)) {
                    dataStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
                } else if ("JUSTIFY".equals(vertical)) {
                    dataStyle.setVerticalAlignment(VerticalAlignment.JUSTIFY);
                } else {
                    dataStyle.setVerticalAlignment(VerticalAlignment.DISTRIBUTED);
                }
            }
            // 设置换行
            String wrapText = style.getAlignment().getWrapText();
            if (!ObjectUtils.isEmpty(wrapText)) {
                dataStyle.setWrapText(true);
            }
        }
    }

    /**
     * 设置单元格背景填充色
     * 
     * @param style:
     * @param dataStyle:
     * @return void
     */
    private static void setCellColor(Style style, CellStyle dataStyle) {
        if (style != null && style.getInterior() != null) {
            if ("#FF0000".equals(style.getInterior().getColor())) {
                // 填充单元格
                dataStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                dataStyle.setFillBackgroundColor(IndexedColors.RED.getIndex());
            } else if ("#92D050".equals(style.getInterior().getColor())) {
                // 填充单元格
                dataStyle.setFillForegroundColor(IndexedColors.LIME.getIndex());
            }
            if ("Solid".equals(style.getInterior().getPattern())) {
                dataStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
        }
    }

    /**
     * 构造合并单元格集合
     * 
     * @param createRowIndex:
     * @param cellRangeAddresses:
     * @param startIndex:
     * @param cellInfo:
     * @param style:
     * @return int
     */
    private static int getCellRanges(int createRowIndex, List<CellRangeAddressInfo> cellRangeAddresses, int startIndex,
        Cell cellInfo, Style style) {
        if (cellInfo.getMergeAcross() != null || cellInfo.getMergeDown() != null) {
            CellRangeAddress r1 = null;
            if (cellInfo.getMergeAcross() != null && cellInfo.getMergeDown() != null) {
                int mergeAcross = startIndex;
                if (cellInfo.getMergeAcross() != 0) {
                    // 获取该单元格结束列数
                    mergeAcross += cellInfo.getMergeAcross();
                }
                int mergeDown = createRowIndex;
                if (cellInfo.getMergeDown() != 0) {
                    // 获取该单元格结束列数
                    mergeDown += cellInfo.getMergeDown();
                }
                r1 = new CellRangeAddress(createRowIndex, mergeDown, (short)startIndex, (short)mergeAcross);
            } else if (cellInfo.getMergeAcross() != null && cellInfo.getMergeDown() == null) {
                int mergeAcross = startIndex;
                if (cellInfo.getMergeAcross() != 0) {
                    // 获取该单元格结束列数
                    mergeAcross += cellInfo.getMergeAcross();
                    // 合并单元格
                    r1 = new CellRangeAddress(createRowIndex, createRowIndex, (short)startIndex, (short)mergeAcross);
                }

            } else if (cellInfo.getMergeDown() != null && cellInfo.getMergeAcross() == null) {
                int mergeDown = createRowIndex;
                if (cellInfo.getMergeDown() != 0) {
                    // 获取该单元格结束列数
                    mergeDown += cellInfo.getMergeDown();
                    // 合并单元格
                    r1 = new CellRangeAddress(createRowIndex, mergeDown, (short)startIndex, (short)startIndex);
                }
            }

            if (cellInfo.getMergeAcross() != null) {
                int length = cellInfo.getMergeAcross().intValue();
                for (int i = 0; i < length; i++) {
                    startIndex += cellInfo.getMergeAcross();
                }
            }
            CellRangeAddressInfo cellRangeAddressInfo = new CellRangeAddressInfo();
            cellRangeAddressInfo.setCellRangeAddress(r1);
            if (style != null && style.getBorders() != null) {
                cellRangeAddressInfo.setBorders(style.getBorders());
            }
            cellRangeAddresses.add(cellRangeAddressInfo);
        }
        return startIndex;
    }

    /**
     * 设置文本值内容
     * 
     * @param wb:
     * @param cellInfo:
     * @param cell:
     * @param style:
     * @param dataStyle:
     * @return void
     */
    private static void setValue(HSSFWorkbook wb, Cell cellInfo, HSSFCell cell, Style style, CellStyle dataStyle) {
        if (cellInfo.getData() != null) {
            HSSFFont font = wb.createFont();
            if (style != null && style.getFont() != null) {
                String color = style.getFont().getColor();
                if ("#FF0000".equals(color)) {
                    font.setColor(IndexedColors.RED.getIndex());
                } else if ("#000000".equals(color)) {
                    font.setColor(IndexedColors.BLACK.getIndex());
                }
            }
            if (!ObjectUtils.isEmpty(cellInfo.getData().getType()) && "Number".equals(cellInfo.getData().getType())) {
                cell.setCellType(CellType.NUMERIC);
            }
            if (style != null && style.getFont().getBold() > 0) {
                font.setBold(true);
            }
            if (style != null && !ObjectUtils.isEmpty(style.getFont().getFontName())) {
                font.setFontName(style.getFont().getFontName());
            }
            if (style != null && style.getFont().getSize() > 0) {
                // 设置字体大小道
                font.setFontHeightInPoints((short)style.getFont().getSize());
            }

            if (cellInfo.getData().getFont() != null) {
                if (cellInfo.getData().getFont().getBold() > 0) {
                    font.setBold(true);
                }
                if ("Number".equals(cellInfo.getData().getType())) {
                    cell.setCellValue(Float.parseFloat(cellInfo.getData().getFont().getText()));
                } else {
                    cell.setCellValue(cellInfo.getData().getFont().getText());
                }
                if (!ObjectUtils.isEmpty(cellInfo.getData().getFont().getCharSet())) {
                    font.setCharSet(Integer.valueOf(cellInfo.getData().getFont().getCharSet()));
                }
            } else {
                if ("Number".equals(cellInfo.getData().getType())) {
                    if (!ObjectUtils.isEmpty(cellInfo.getData().getText())) {
                        // cell.setCellValue(Float.parseFloat(cellInfo.getData().getText()));
                        cell.setCellValue(Float.parseFloat(cellInfo.getData().getText().replaceAll(",", "")));
                    }
                } else {
                    cell.setCellValue(cellInfo.getData().getText());

                }
            }

            if (style != null) {
                if (style.getNumberFormat() != null) {
                    String color = style.getFont().getColor();
                    if ("#FF0000".equals(color)) {
                        font.setColor(IndexedColors.RED.getIndex());
                    } else if ("#000000".equals(color)) {
                        font.setColor(IndexedColors.BLACK.getIndex());
                    }
                    if ("0%".equals(style.getNumberFormat().getFormat())) {
                        HSSFDataFormat format = wb.createDataFormat();
                        dataStyle.setDataFormat(format.getFormat(style.getNumberFormat().getFormat()));
                    } else {
                        dataStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0.00"));
                    }
                    // HSSFDataFormat format = wb.createDataFormat();
                    // dataStyle.setDataFormat(format.getFormat(style.getNumberFormat().getFormat()));

                }
            }
            dataStyle.setFont(font);
        }
    }
}
