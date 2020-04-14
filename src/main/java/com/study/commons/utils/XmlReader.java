package com.study.commons.utils;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.dom4j.Document;
import org.dom4j.Element;
import org.springframework.util.CollectionUtils;
import org.springframework.util.ObjectUtils;

import com.study.dto.Data;
import com.study.entity.Cell;
import com.study.entity.Column;
import com.study.entity.Font;
import com.study.entity.Row;
import com.study.entity.Style;
import com.study.entity.Table;
import com.study.entity.Worksheet;

public class XmlReader {

    /**
     * 获取样式
     * 
     * @param document:
     * @return java.util.Map<java.lang.String,com.power.commons.freemarker.entity.Style>
     */
    public static Map<String, Style> getStyle(Document document) {
        // 创建一个LinkedHashMap用于存放style，按照id查找
        Map<String, Style> styleMap = new LinkedHashMap<String, Style>();
        // 新建一个Style类用于存放节点数据
        Style style = null;

        // 获取根节点
        Element root = document.getRootElement();
        // 获取根节点下的Styles节点
        Element styles = root.element("Styles");
        // 获取Styles下的Style节点
        List lstyle = styles.elements("Style");
        Iterator<?> it = lstyle.iterator();
        while (it.hasNext()) {
            // Element e = (Element) it.next();
            // 新建一个Style类用于存放节点数据
            style = new Style();
            Element e = (Element)it.next();
            String id = e.attributeValue("ID").toString();
            // 设置style的id
            style.setId(id);

            if (e.attributeValue("Name") != null) {
                String name = e.attributeValue("Name").toString();
                // 设置style的name
                style.setName(name);
            }

            // if (id == 1) {//当id为1时，设置为另一中构造方法
            // 获取Style下的NumberFormat节点
            Element enumberFormat = e.element("NumberFormat");
            if (enumberFormat != null) {
                Style.NumberFormat numberFormat = new Style.NumberFormat();
                numberFormat.setFormat(enumberFormat.attributeValue("Format"));
                style.setNumberFormat(numberFormat);
            }

            // continue;
            // }
            Style.Alignment alignment = new Style.Alignment();
            // 获取Style下的Alignment节点

            Element ealignment = e.element("Alignment");
            if (ealignment != null) {
                // 设置aligment的相关属性，并且设置style的aliment属性
                alignment.setHorizontal(ealignment.attributeValue("Horizontal"));
                alignment.setVertical(ealignment.attributeValue("Vertical"));
                alignment.setWrapText(ealignment.attributeValue("WrapText"));
                style.setAlignment(alignment);
            }

            // 获取Style下的Borders节点
            Element Borders = e.element("Borders");
            if (Borders != null) {
                // 获取Borders下的Border节点
                List Border = Borders.elements("Border");
                // 用迭代器遍历Border节点
                Iterator<?> bdIt = Border.iterator();
                List<Style.Border> lborders = new ArrayList<Style.Border>();
                while (bdIt.hasNext()) {
                    Element bd = (Element)bdIt.next();
                    Style.Border border = new Style.Border();
                    border.setPosition(bd.attributeValue("Position"));
                    if (bd.attribute("LineStyle") != null) {
                        border.setLinestyle(bd.attributeValue("LineStyle"));
                        int weight = Integer.parseInt(bd.attributeValue("Weight"));
                        border.setWeight(weight);
                        border.setColor(bd.attributeValue("Color"));
                    }
                    lborders.add(border);
                }
                style.setBorders(lborders);
            }

            // 利用List来存取borders下的多个border并放入style

            // for (int i = 0; i < style.getBorders().size(); i++) {
            // System.out.println(style.getBorders().get(i).getPosition());
            // }

            // 设置font的相关属性，并且设置style的font属性
            Style.Font font = new Style.Font();
            Element efont = e.element("Font");
            font.setFontName(efont.attributeValue("FontName"));
            if (efont.attributeValue("Size") != null) {
                double size = Double.parseDouble(efont.attributeValue("Size"));
                font.setSize(size);
            }
            if (efont.attribute("Bold") != null) {
                int bold = Integer.parseInt(efont.attributeValue("Bold"));
                font.setBold(bold);
            }
            font.setColor(efont.attributeValue("Color"));
            style.setFont(font);
            // 设置Interior的相关属性，并且设置style的interior属性
            Style.Interior interior = new Style.Interior();
            if (e.element("Interior") != null) {
                Element einterior = e.element("Interior");
                interior.setColor(einterior.attributeValue("Color"));
                interior.setPattern(einterior.attributeValue("Pattern"));
            }
            style.setInterior(interior);
            // if (e.element("NumberFormat") != null) {
            // 设置NumberFormat的相关属性，并且设置style的NumberFormat属性
            // Element enumberFormat = e.element("NumberFormat");//获取Style下的Alignment节点
            // numberFormat.setFormat(enumberFormat.attributeValue("Format"));
            // style.setNumberFormat(numberFormat);
            // }
            if (e.element("Protection") != null) {
                Element protectione = e.element("Protection");
                Style.Protection protection = new Style.Protection();
                protection.setProtected1(protectione.attributeValue("Protected"));
                style.setProtection(protection);
            }
            styleMap.put(id, style);

        }
        // 获取根节点下的Names节点
        Element Names = root.element("Names");
        if (Names != null) {
            // 获取根节点下的NamedRange节点
            Element NamedRange = Names.element("NamedRange");
            // 获取ss:Name
            String name = NamedRange.attributeValue("Name");
            // 获取ss:RefersTo
            String namedRange = NamedRange.attributeValue("RefersTo");
        }

        return styleMap;
    }

    // public static List<Column> getColumn(Document document){
    // Element root = document.getRootElement();
    // //读取根节点下的Worksheet节点
    // Element worksheet = root.element("Worksheet");
    // String name = worksheet.attributeValue("Name");
    // //读取Worksheet下的ss:Table节点
    // Element table = worksheet.element("Table");
    // //读取ss:Table下的所有ss:Column节点
    // List column = table.elements("Column");
    // List<Column> columns = new ArrayList<Column>();
    // for (Iterator<?> it = column.iterator(); it.hasNext(); ) {
    // //通过创建column来存取column里的值
    // Column column1 = new Column();
    // Element eColumn = (Element) it.next();
    // double width = Double.parseDouble(eColumn.attributeValue("Width"));
    // column1.setWidth(width);
    // columns.add(column1);
    // }
    // return columns;
    // }

    // public static List<Information> getDatas(Document document){
    // Element root = document.getRootElement();
    // List<Element> worksheet1 = root.elements("Worksheet");
    //
    // return null;
    // }

    /**
     * 获取具体数据
     * 
     * @param document:
     * @return com.power.commons.file.Information
     */
    // public static Information getData(Document document) {
    // Element root = document.getRootElement();
    //// List<Element> worksheet1 = root.elements("Worksheet");
    // //读取根节点下的Worksheet节点
    // Element worksheet = root.element("Worksheet");
    // //读取Worksheet下的ss:Table节点
    // Element table = worksheet.element("Table");
    //
    // Information information = new Information();
    // information.setName(worksheet.attributeValue("Name"));
    // //读取Worksheet下的Row节点
    // List row = table.elements("Row");
    // List<Information.Row> lrow = new ArrayList<Information.Row>();
    // for (Iterator<?> it = row.iterator(); it.hasNext(); ) {
    // //设置row对象属性并存放入list，存入information
    // Information.Row row1 = new Information.Row();
    // Element eRow = (Element) it.next();
    // if(eRow.attributeValue("AutoFitHeight")!=null){
    // int autofitheight = Integer.parseInt(eRow.attributeValue("AutoFitHeight"));
    // row1.setAutofitheight(autofitheight);
    // }
    // if (eRow.attributeValue("Height")!= null) {
    // double height = Double.parseDouble(eRow.attributeValue("Height"));
    // row1.setHeight(height);
    // }
    // //将一个个row对象存放入list lrow中
    // lrow.add(row1);
    // //list用于存放cell
    // List<Information.Row.Cell> lcell = new ArrayList<Information.Row.Cell>();
    // //读取Row下的Cell节点
    // List cell = eRow.elements("Cell");
    // for (Iterator<?> it1 = cell.iterator(); it1.hasNext(); ) {
    // Element eCell = (Element) it1.next();
    // //将cell节点相关数据存入row
    // Information.Row.Cell cell1 = new Information.Row.Cell();
    // if(eCell.attributeValue("Index")!=null){
    // int index = Integer.parseInt(eCell.attributeValue("Index"));
    // cell1.setIndex(index);
    // }
    // String styleid = eCell.attributeValue("StyleID").toString();
    // cell1.setStyleID(styleid);
    // if (eCell.attributeValue("MergeAcross") != null) {
    // int mergeacross = Integer.parseInt(eCell.attributeValue("MergeAcross"));
    // cell1.setMergeAcross(mergeacross);
    // }
    // if (eCell.attributeValue("MergeDown") != null) {
    // int mergedown = Integer.parseInt(eCell.attributeValue("MergeDown"));
    // cell1.setMergeDown(mergedown);
    // }
    // //将一个个cell添加入listlcell中
    // lcell.add(cell1);
    // if(eCell.element("Data")!=null) {
    // //读取Cell下的Data节点
    // Element data = eCell.element("Data");
    // //将data节点相关数据存入cell1
    // Information.Row.Cell.Data data1 = new Information.Row.Cell.Data();
    // if (data.attributeValue("Ticked")!= null) {
    // int ticked = Integer.parseInt(data.attributeValue("Ticked"));
    // data1.setTicked(ticked);
    // }
    // data1.setType(data.attributeValue("Type"));
    // data1.setDat(data.getText());
    // cell1.setData(data1);
    // }
    // }
    // row1.setCell(lcell);
    // }
    // //将lrow这个list存放为information的row属性中
    // information.setRow(lrow);
    // return information;
    // }

    public static List<Worksheet> getWorksheet(Document document) {
        List<Worksheet> worksheets = new ArrayList<>();
        Element root = document.getRootElement();
        // 读取根节点下的Worksheet节点
        List<Element> sheets = root.elements("Worksheet");
        if (CollectionUtils.isEmpty(sheets)) {
            return worksheets;
        }

        for (Element sheet : sheets) {
            Worksheet worksheet = new Worksheet();
            String name = sheet.attributeValue("Name");
            worksheet.setName(name);
            Table table = getTable(sheet);
            worksheet.setTable(table);
            worksheets.add(worksheet);
        }
        return worksheets;
    }

    private static Table getTable(Element sheet) {
        Element tableElement = sheet.element("Table");
        if (tableElement == null) {
            return null;
        }
        Table table = new Table();
        String expandedColumnCount = tableElement.attributeValue("ExpandedColumnCount");
        if (expandedColumnCount != null) {
            table.setExpandedColumnCount(Integer.parseInt(expandedColumnCount));
        }
        String expandedRowCount = tableElement.attributeValue("ExpandedRowCount");
        if (expandedRowCount != null) {
            table.setExpandedRowCount(Integer.parseInt(expandedRowCount));
        }
        String fullColumns = tableElement.attributeValue("FullColumns");
        if (fullColumns != null) {
            table.setFullColumns(Integer.parseInt(fullColumns));
        }

        String fullRows = tableElement.attributeValue("FullRows");
        if (fullRows != null) {
            table.setFullRows(Integer.parseInt(fullRows));
        }
        String defaultColumnWidth = tableElement.attributeValue("DefaultColumnWidth");
        if (defaultColumnWidth != null) {
            table.setDefaultColumnWidth(Double.valueOf(defaultColumnWidth).intValue());
        }

        String defaultRowHeight = tableElement.attributeValue("DefaultRowHeight");
        if (defaultRowHeight != null) {
            table.setDefaultRowHeight(Integer.parseInt(defaultRowHeight));
        }
        // 读取列
        List<Column> columns = getColumns(tableElement, expandedColumnCount, defaultColumnWidth);
        table.setColumns(columns);

        // 读取行
        List<Row> rows = getRows(tableElement);
        table.setRows(rows);
        return table;
    }

    private static List<Row> getRows(Element tableElement) {
        List<Element> rowElements = tableElement.elements("Row");
        if (CollectionUtils.isEmpty(rowElements)) {
            return null;
        }
        List<Row> rows = new ArrayList<>();
        for (Element rowElement : rowElements) {
            Row row = new Row();
            String height = rowElement.attributeValue("Height");
            if (height != null) {
                row.setHeight(Double.valueOf(height).intValue());
            }

            String index = rowElement.attributeValue("Index");
            if (index != null) {
                row.setIndex(Integer.valueOf(index));
            }
            List<Cell> cells = getCells(rowElement);
            row.setCells(cells);
            rows.add(row);
        }
        return rows;
    }

    private static List<Cell> getCells(Element rowElement) {
        List<Element> cellElements = rowElement.elements("Cell");
        if (CollectionUtils.isEmpty(cellElements)) {
            return null;
        }
        List<Cell> cells = new ArrayList<>();
        for (Element cellElement : cellElements) {
            Cell cell = new Cell();
            String styleID = cellElement.attributeValue("StyleID");
            if (styleID != null) {
                cell.setStyleID(styleID);
            }
            String mergeAcross = cellElement.attributeValue("MergeAcross");
            if (mergeAcross != null) {
                cell.setMergeAcross(Integer.valueOf(mergeAcross));
            }

            String mergeDown = cellElement.attributeValue("MergeDown");
            if (mergeDown != null) {
                cell.setMergeDown(Integer.valueOf(mergeDown));
            }

            String index = cellElement.attributeValue("Index");
            if (index != null) {
                cell.setIndex(Integer.valueOf(index));
            }
            Element dataElement = cellElement.element("Data");
            if (dataElement != null) {
                Data data = new Data();
                String type = dataElement.attributeValue("Type");
                String xmlns = dataElement.attributeValue("xmlns");
                data.setType(type);
                data.setXmlns(xmlns);
                data.setText(dataElement.getText());
                Element bElement = dataElement.element("B");
                Integer bold = null;
                Element fontElement = null;
                if (bElement != null) {
                    fontElement = bElement.element("Font");
                    bold = 1;
                }
                Element uElement = dataElement.element("U");
                if (uElement != null) {
                    fontElement = uElement.element("Font");
                }
                if (fontElement == null) {
                    fontElement = dataElement.element("Font");
                }
                if (fontElement != null) {
                    Font font = new Font();
                    String face = fontElement.attributeValue("Face");
                    if (face != null) {
                        font.setFace(face);
                    }
                    String charSet = fontElement.attributeValue("CharSet");
                    if (charSet != null) {
                        font.setCharSet(charSet);
                    }
                    String color = fontElement.attributeValue("Color");
                    if (color != null) {
                        font.setColor(color);
                    }
                    if (bold != null) {
                        font.setBold(bold);
                    }
                    font.setText(fontElement.getText());
                    data.setFont(font);
                }

                cell.setData(data);
            }
            cells.add(cell);
        }
        return cells;
    }

    private static List<Column> getColumns(Element tableElement, String expandedRowCount, String defaultColumnWidth) {
        List<Element> columnElements = tableElement.elements("Column");
        if (CollectionUtils.isEmpty(columnElements)) {
            return null;
        }
        if (ObjectUtils.isEmpty(expandedRowCount)) {
            return null;
        }
        int defaultWidth = 60;
        if (!ObjectUtils.isEmpty(defaultColumnWidth)) {
            defaultWidth = Double.valueOf(defaultColumnWidth).intValue();
        }
        List<Column> columns = new ArrayList<>();
        int indexNum = 0;
        for (int i = 0; i < columnElements.size(); i++) {
            Column column = new Column();
            Element columnElement = columnElements.get(i);
            String index = columnElement.attributeValue("Index");
            if (index != null) {
                if (indexNum < Integer.valueOf(index) - 1) {
                    for (int j = indexNum; j < Integer.valueOf(index) - 1; j++) {
                        column = new Column();
                        column.setIndex(indexNum);
                        column.setWidth(defaultWidth);
                        columns.add(column);
                        indexNum += 1;
                    }
                }
                column = new Column();
            }
            column.setIndex(indexNum);
            String autoFitWidth = columnElement.attributeValue("AutoFitWidth");
            if (autoFitWidth != null) {
                column.setAutofitwidth(Integer.valueOf(autoFitWidth));
            }
            String width = columnElement.attributeValue("Width");
            if (width != null) {
                column.setWidth(Double.valueOf(width).intValue());
            }
            columns.add(column);
            indexNum += 1;
        }
        if (columns.size() < Integer.valueOf(expandedRowCount)) {
            for (int i = columns.size() + 1; i <= Integer.valueOf(expandedRowCount); i++) {
                Column column = new Column();
                column.setIndex(i);
                column.setWidth(defaultWidth);
                columns.add(column);
            }
        }
        return columns;
    }
}
