package com.example.file.util;


import com.google.common.collect.Lists;
import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.math.NumberUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletRequest;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.List;
import java.util.Map;


public class PoiExcelUtils {

    private static Logger logger = LoggerFactory.getLogger(PoiExcelUtils.class);
    /**
     * 默认字体
     */
    private static final String EXCEL_FONT = "微软雅黑";
    /**
     * 偶数行
     */
    private static final int OLD_ROW = 2;
    /**
     * blank
     */
    private static final String BLANK_VALUE = "";

    /**
     * Export.
     *
     * @param sheetName 导出的SHEET名字  当前sheet数目只为1
     * @param headers   导出的表格的表头
     * @param dsTitles  导出的数据 map.get(key) 对应的 key
     * @param data      数据集  List<Map>
     * @return the hssf workbook
     * @throws IOException the io exception
     * @since 2017.07.05
     */
    public static HSSFWorkbook export(String sheetName, String[] headers, String[] dsTitles,
                                      List<Map<String, Object>> data) throws IOException {
        int[] widths = intiWidth(dsTitles);
        int[] dsFormat = initFormat(dsTitles);
        //创建一个工作薄
        HSSFWorkbook wb = new HSSFWorkbook();
        //创建一个sheet
        HSSFSheet sheet = wb.createSheet(StringUtils.isNotEmpty(sheetName) ? sheetName : "excel");
        //创建表头，如果没有跳过
        int headerRow = setHeader(headers, wb, sheet, widths);
        formatData(data, dsTitles, wb, sheet, dsFormat, headerRow);
        return wb;
    }

    @SuppressWarnings("deprecation")
    private static int setHeader(String[] headers, HSSFWorkbook wb, HSSFSheet sheet, int[] widths) {
        int headerRow = 0;
        if (headers == null) {
            return headerRow;
        }
        HSSFRow row = sheet.createRow(headerRow);
        //表头样式
        HSSFCellStyle style = wb.createCellStyle();
        HSSFFont font = wb.createFont();
//        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        font.setFontName(EXCEL_FONT);
        font.setFontHeightInPoints((short) 11);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        for (int i = 0; i < headers.length; i++) {
            sheet.setColumnWidth((short) i, (short) widths[i]);
            HSSFCell cell = row.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(style);
        }
        headerRow++;
        return headerRow;
    }

    @SuppressWarnings("deprecation")
    private static void formatData(List<Map<String, Object>> data, String[] dsTitles, HSSFWorkbook wb, HSSFSheet sheet, int[] dsFormat, int headerRow) {
        if (data == null) {
            return;
        }
        //表格主体  解析list
        List<HSSFCellStyle> styleList = formatHSSFCellStyle(dsTitles, wb, dsFormat);
        //行数
        for (Map<String, Object> aData : data) {
            HSSFRow row = sheet.createRow(headerRow);
            //列数
            for (int j = 0; j < dsTitles.length; j++) {
                HSSFCell cell = row.createCell(j);
                Object o = ((Map) aData).get(dsTitles[j]);
                HSSFCellStyle style = styleList.get(j);
                fillCellValue(style, cell, o);
            }
            headerRow++;
        }
    }

    private static int[] intiWidth(String[] dsTitles) {
        int[] widths = new int[dsTitles.length];
        for (int i = 0; i < dsTitles.length; i++) {
            int width = 256 * 14;
            widths[i] = width;
        }
        return widths;
    }

    private static int[] initFormat(String[] dsTitles) {
        int[] dsFormat = new int[dsTitles.length];
        for (int i = 0; i < dsTitles.length; i++) {
            dsFormat[i] = 2;
        }
        return dsFormat;
    }

    /**
     * 对文件流输出下载的中文文件名进行编码 屏蔽各种浏览器版本的差异性
     *
     * @param request   the request
     * @param pFileName the p file name
     * @return the string
     * @throws Exception the exception
     * @since 2017.07.05
     */
    public static String encodeChineseDownloadFileName(
            HttpServletRequest request, String pFileName) throws Exception {
        String filename = null;
        String agent = request.getHeader("USER-AGENT");
        if (null != agent) {
            if (agent.contains("Firefox")) {//Firefox
                filename = "=?UTF-8?B?" + (new String(org.apache.commons.codec.binary.Base64.encodeBase64(pFileName.getBytes("UTF-8")))) + "?=";
            } else if (agent.contains("Chrome")) {//Chrome
                filename = new String(pFileName.getBytes(), "ISO8859-1");
            } else {//IE7+
                filename = java.net.URLEncoder.encode(pFileName, "UTF-8");
                filename = filename.replace("+", "%20");
            }
        } else {
            filename = pFileName;
        }
        return filename;
    }

    /**
     * @param sheetList
     * @param headersList
     * @param dsTitlesList
     * @param dataList
     * @return
     * @throws IOException
     */
    public static HSSFWorkbook export(List<String> sheetList, List<String[]> headersList, List<String[]> dsTitlesList,
                                      List<List<Map<String, Object>>> dataList) throws IOException {
        //创建一个工作薄
        HSSFWorkbook wb = new HSSFWorkbook();
        for (int i = 0; i < sheetList.size(); i++) {
            int[] widths = intiWidth(dsTitlesList.get(i));
            int[] dsFormat = initFormat(dsTitlesList.get(i));
            //创建一个sheet
            HSSFSheet sheet = wb.createSheet(sheetList.get(i));
            //创建表头，如果没有跳过
            int headerRow = setHeader(headersList.get(i), wb, sheet, widths);
            formatData(dataList.get(i), dsTitlesList.get(i), wb, sheet, dsFormat, headerRow);
        }

        return wb;
    }


    private static List<HSSFCellStyle> formatHSSFCellStyle(String[] dsTitles, HSSFWorkbook wb, int[] dsFormat) {
        //表格主体  解析list
        List<HSSFCellStyle> styleList = Lists.newArrayList();
        for (int i = 0; i < dsTitles.length; i++) {
            HSSFCellStyle style = wb.createCellStyle();
            HSSFFont font = wb.createFont();
            font.setFontName(EXCEL_FONT);
            font.setFontHeightInPoints((short) 10);
            style.setFont(font);
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            style.setBorderTop(BorderStyle.THIN);
            if (dsFormat[i] == 1) {
                style.setAlignment(HorizontalAlignment.LEFT);
            } else if (dsFormat[i] == 2) {
                style.setAlignment(HorizontalAlignment.CENTER);
            } else if (dsFormat[i] == 3) {
                style.setAlignment(HorizontalAlignment.RIGHT);
            }
            styleList.add(style);
        }
        return styleList;
    }

    public static HSSFWorkbook exportHasColor(String sheetName, String[] headers, String[] dsTitles,
                                              List<Map<String, Object>> data) throws IOException {
        int[] widths = intiWidth(dsTitles);
        //创建一个工作薄
        HSSFWorkbook wb = new HSSFWorkbook();
        //创建一个sheet
        HSSFSheet sheet = wb.createSheet(StringUtils.isNotEmpty(sheetName) ? sheetName : "excel");
        //创建表头，如果没有跳过
        int headerRow = setHeader(headers, wb, sheet, widths);
        formatDataHasColor(data, dsTitles, wb, sheet, headerRow);
        return wb;
    }

    public static HSSFWorkbook exportMulSheetHasColor(HSSFWorkbook wb, String sheetName, String[] headers, String[] dsTitles,
                                                      List<Map<String, Object>> data) throws IOException {
        int[] widths = intiWidth(dsTitles);
        //创建一个sheet
        //创建一个sheet
        HSSFSheet sheet = wb.createSheet(StringUtils.isNotEmpty(sheetName) ? sheetName : "excel");
        //创建表头，如果没有跳过
        int headerRow = setHeader(headers, wb, sheet, widths);
        formatDataHasColor(data, dsTitles, wb, sheet, headerRow);
        return wb;
    }

    private static HSSFCellStyle getHSSFCellStyle(HSSFWorkbook wb, int rowNum) {
        HSSFCellStyle style = wb.createCellStyle();
        HSSFFont font = wb.createFont();
        font.setFontName(EXCEL_FONT);
        font.setFontHeightInPoints((short) 10);
        style.setFont(font);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        if (rowNum % OLD_ROW == NumberUtils.INTEGER_ZERO) {
            style.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        } else {
            style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        return style;
    }

    private static void formatDataHasColor(List<Map<String, Object>> data, String[] dsTitles, HSSFWorkbook wb, HSSFSheet sheet, int headerRow) {
        if (data == null) {
            return;
        }
        //表格主体  解析list
        for (Map<String, Object> aData : data) {
            HSSFRow row = sheet.createRow(headerRow);
            HSSFCellStyle style = getHSSFCellStyle(wb, row.getRowNum());
            //列数
            for (int j = 0; j < dsTitles.length; j++) {
                HSSFCell cell = row.createCell(j);
                Object o = ((Map) aData).get(dsTitles[j]);
                fillCellValue(style, cell, o);
            }
            headerRow++;
        }
    }

    private static void fillCellValue(HSSFCellStyle style, HSSFCell cell, Object value) {
        if (value == null || BLANK_VALUE.equals(value)) {
            cell.setCellValue(BLANK_VALUE);
        } else {
            cell.setCellType(CellType.NUMERIC);
            if (value instanceof Integer) {
                cell.setCellValue(Integer.parseInt(value.toString()));
            } else if (value instanceof BigDecimal) {
                cell.setCellValue(((BigDecimal) value).doubleValue());
            } else if (value instanceof Double) {
                cell.setCellValue((double) value);
            } else if (value instanceof Long) {
                cell.setCellValue((long) value);
            } else {
                cell.setCellValue(value.toString());
                cell.setCellType(CellType.STRING);
            }
        }
        cell.setCellStyle(style);
    }
}
