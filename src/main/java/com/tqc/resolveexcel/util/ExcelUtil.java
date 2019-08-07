package com.tqc.resolveexcel.util;

import com.sun.org.apache.xpath.internal.objects.XObject;
import com.tqc.resolveexcel.model.excel.ExcelSet;
import com.tqc.resolveexcel.model.excel.ExcelSheet;
import jdk.internal.org.objectweb.asm.tree.InnerClassNode;
import lombok.SneakyThrows;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.codehaus.groovy.runtime.dgmimpl.arrays.IntegerArrayGetAtMetaMethod;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Creator: qingchang.tang
 * Date: 2018/03/06
 * Time: 上午 9:45
 * <p>
 * 解析Excel工具类
 */
public class ExcelUtil {

    private static final String SHUA_KA_RI_QI = "刷卡日期";

    private static final Integer CELL_INDEX_ONE = 1;

    private static final Integer CELL_INDEX_TWO = 2;

    private static final Integer CELL_INDEX_THREE = 3;

    private static final Integer CELL_INDEX_FOUR = 4;

    private static final Integer CELL_INDEX_FIVE = 5;

    private static final String EMPTY_SPACE = " ";

    private static final Integer INDEX_ZERO = 0;

    private static final Integer INDEX_ONE =1;

    private static final String XLS = ".xls";

    private static final String XLSX = ".xlsx";


    /**
     * 解析Excel表格
     *
     * @param path 文件路径
     * @return
     * @throws Exception
     */
    public static ExcelSet resolveExcel(String path) throws Exception {

        ExcelSet excelSet = new ExcelSet();
        //Excel文件
        File file = new File(path);
        if (file.getName().endsWith(XLS)) {
            excelSet = resolveExcelXls(file);
        } else if (file.getName().endsWith(XLSX)) {
            excelSet = resolveExcelXlsx(file);
        }
        return excelSet;
    }

    private static ExcelSet resolveExcelXlsx(File file) {
        return new ExcelSet();
    }

    @SneakyThrows(value = Exception.class)
    private static ExcelSet resolveExcelXls(File file) {
        try (Workbook workbook = WorkbookFactory.create(file)) {
            List<ExcelSheet> sheets = new ArrayList<>();
            Iterator<Sheet> its = workbook.sheetIterator();
            //处理每个sheet
            while (its.hasNext()) {
                Sheet sheet = its.next();
                List<List<Object>> content = new ArrayList<>();
                ExcelSheet excelSheet = new ExcelSheet();
                excelSheet.setName(sheet.getSheetName());
                contentAddHeadTitle(sheet, content);
                formatDateAndTime(sheet, content);
                List<List<Object>> dataClean = dealDataWithSameNameByList(content);
                List<List<Object>> averageTime = getAverageTime(dataClean);
                excelSheet.setContent(averageTime);
                sheets.add(excelSheet);
            }
            return ExcelSet.builder().sheets(sheets).excelFile(file).build();
        }
    }

    private static List<List<Object>> getAverageTime(List<List<Object>> dataClean) {
        List<Object> names = new ArrayList<>();

        List<List<Object>> result = new ArrayList<>();
        List<Object> head = new ArrayList<>();
        head.add(0, "姓名");
        head.add(1, "平均每月每天的上班时间");
        result.add(head);
        for (int i = 1; i < dataClean.size(); i++) {
            String currentName = String.valueOf(dataClean.get(i).get(0));
            if (!names.contains(currentName)) {
                List<List<Object>> middleResult = new ArrayList<>();
                names.add(currentName);
                for (int j = 1; j < dataClean.size(); j++) {
                    String dealName = String.valueOf(dataClean.get(j).get(0));
                    if (currentName.equals(dealName)) {
                        middleResult.add(dataClean.get(j));
                    }
                }
                dealDataForAverageTime(middleResult, result);
            }
        }
        return result;
    }

    private static void dealDataForAverageTime(List<List<Object>> middleResult, List<List<Object>> result) {
        List<Object> dealMiddleResult = new ArrayList<>();
        dealMiddleResult.add(0, middleResult.get(0).get(0));
        int days = middleResult.size();
        double times = 0;
        for (int i = 1; i < days; i++) {
            Object middle = middleResult.get(i).get(2);
            double time = Double.parseDouble(middle.toString());
            times = time + times;
        }
        double averageTime = times / days;
        dealMiddleResult.add(1, averageTime);
        result.add(dealMiddleResult);
    }

    private static void formatDateAndTime(Sheet sheet, List<List<Object>> content) {
        DateFormat parseDate = new SimpleDateFormat("MM月dd日 HH:mm:ss");
        for (int i = 1; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            List<Object> convertEachExcelLine = new ArrayList<>();
            String username = String.valueOf(row.getCell(CELL_INDEX_TWO));
            convertEachExcelLine.add(username);
            Date signInDate = HSSFDateUtil.getJavaDate(row.getCell(CELL_INDEX_FOUR).getNumericCellValue());
            String specificDate = parseDate.format(signInDate).split(" ")[0];
            convertEachExcelLine.add(specificDate);
            Date time = HSSFDateUtil.getJavaDate(row.getCell(CELL_INDEX_FIVE).getNumericCellValue());
            String time1 = parseDate.format(time).split(" ")[1];
            convertEachExcelLine.add(time1);
            content.add(convertEachExcelLine);
        }
    }

    private static List<List<Object>> dealDataWithSameNameByList(List<List<Object>> content) {
        List<Object> names = new ArrayList<>();
        List<List<Object>> dataCleanResult = new ArrayList<>();
        for (int i = 1; i < content.size(); i++) {
            String currentName = String.valueOf(content.get(i).get(0));
            if (!names.contains(currentName)) {
                List<List<Object>> waitDealList = new ArrayList<>();
                names.add(currentName);
                for (int j = 1; j < content.size(); j++) {
                    String dealName = String.valueOf(content.get(j).get(0));
                    if (currentName.equals(dealName)) {
                        waitDealList.add(content.get(j));
                    }
                }
                dealDataWithSameDataByList(waitDealList, dataCleanResult);
            }
        }
        return dataCleanResult;
    }

    private static void contentAddHeadTitle(Sheet sheet, List<List<Object>> content) {
        Row head = sheet.getRow(0);
        List<Object> headLine = new ArrayList<>();
        headLine.add(head.getCell(2));
        headLine.add(head.getCell(4));
        headLine.add(head.getCell(5));
        content.add(headLine);
    }

    private static void dealDataWithSameDataByList(List<List<Object>> waitDealList, List<List<Object>> dataCleanResult) {
        List<String> date = new ArrayList<>();

        for (int i = 0; i < waitDealList.size(); i++) {
            String currentDate = String.valueOf(waitDealList.get(i).get(1));
            if (!date.contains(currentDate)) {
                List<List<Object>> current = new ArrayList<>();
                date.add(currentDate);
                for (int j = 0; j < waitDealList.size(); j++) {
                    String needDealDate = String.valueOf(waitDealList.get(j).get(1));
                    if (currentDate.equals(needDealDate)) {
                        current.add(waitDealList.get(j));
                    }
                }
                calculateTime(current, dataCleanResult);
            }
        }
    }

    private static List<List<Object>> calculateTime(List<List<Object>> current, List<List<Object>> dataCleanResult) {
        if (current.size() == 1) {
            return dealAddDefault(current, dataCleanResult);
        } else if (current.size() == 2) {
            return dealCurrentDateTime(current, dataCleanResult);
        } else {
            return dealDuplicate(current, dataCleanResult);
        }
    }

    private static List<List<Object>> dealAddDefault(List<List<Object>> current, List<List<Object>> dataCleanResult) {
        DateFormat df = new SimpleDateFormat("HH:mm:ss");
        List<Object> defaultValue = new ArrayList<>();
        defaultValue.add(current.get(0).get(0));
        defaultValue.add(current.get(0).get(1));
        defaultValue.add(current.get(0).get(2));
        List<List<Object>> result = new ArrayList<>(2);
        result.add(current.get(0));
        try {
            Date dealLine = df.parse("12:00:00");
            Date currentTime = df.parse(String.valueOf(current.get(0).get(2)));
            if (currentTime.before(dealLine)) {
                defaultValue.set(2, "19:00:00");
            } else {
                defaultValue.set(2, "10:00:00");
            }
            result.add(defaultValue);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return dealCurrentDateTime(result, dataCleanResult);
    }

    private static List<List<Object>> dealCurrentDateTime(List<List<Object>> result, List<List<Object>> dataCleanResult) {
        DateFormat df = new SimpleDateFormat("HH:mm:ss");
        try {
            String date1 = String.valueOf(result.get(0).get(2));
            String date2 = String.valueOf(result.get(1).get(2));
            Date d1 = df.parse(date1);
            Date d2 = df.parse(String.valueOf(date2));
            long diff = d1.getTime() - d2.getTime();
            long days = diff / (1000 * 60 * 60 * 24);
            long hours = Math.abs((diff - days * (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
            long minutes = Math.abs((diff - days * (1000 * 60 * 60 * 24) - hours * (1000 * 60 * 60)) / (1000 * 60));
            List<Object> middleResult = new ArrayList<>();
            middleResult.add(result.get(0).get(0));
            middleResult.add(result.get(0).get(1));
            middleResult.add(hours + "." + minutes);
            dataCleanResult.add(middleResult);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return dataCleanResult;
    }

    private static List<List<Object>> dealDuplicate(List<List<Object>> current, List<List<Object>> dataCleanResult) {
        List<List<Object>> result = new ArrayList<>();
        try {
            DateFormat df = new SimpleDateFormat("HH:mm:ss");
            Date minDate = df.parse(String.valueOf(current.get(0).get(2)));
            Date maxDate = df.parse(String.valueOf(current.get(1).get(2)));
            if (minDate.after(maxDate)) {
                Date middle = minDate;
                minDate = maxDate;
                maxDate = middle;
            }
            for (int i = 2; i < current.size(); i++) {
                Date date = df.parse(String.valueOf(current.get(i).get(2)));
                if (date.before(minDate)) {
                    minDate = date;
                } else if (date.after(maxDate)) {
                    maxDate = date;
                }
            }
            List<Object> min = new ArrayList<>();
            Object name = current.get(0).get(0);
            min.add(name);
            Object date = current.get(0).get(1);
            min.add(date);
            min.add(df.format(minDate));

            List<Object> max = new ArrayList<>();
            max.add(name);
            max.add(date);
            max.add(df.format(maxDate));
            result.add(min);
            result.add(max);

        } catch (Exception e) {
            e.printStackTrace();
        }
        return dealCurrentDateTime(result, dataCleanResult);
    }
}
