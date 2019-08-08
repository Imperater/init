package com.tqc.resolveexcel.util;

import com.google.common.collect.Lists;
import com.tqc.resolveexcel.model.ExcelDTO;
import com.tqc.resolveexcel.model.ResultVO;
import com.tqc.resolveexcel.model.excel.ExcelSet;
import com.tqc.resolveexcel.model.excel.ExcelSheet;
import lombok.SneakyThrows;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;

import javax.swing.*;
import java.io.File;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Creator: qingchang.tang
 * Date: 2018/03/06
 * Time: 上午 9:45
 * <p>
 * 解析Excel工具类
 */
public class ExcelUtil {

    private static final Integer CELL_INDEX_ONE = 1;

    private static final Integer CELL_INDEX_TWO = 2;


    private static final Integer CELL_INDEX_FOUR = 4;

    private static final Integer CELL_INDEX_FIVE = 5;

    private static final String EMPTY_SPACE = " ";

    private static final Integer INDEX_ZERO = 0;

    private static final Integer INDEX_ONE = 1;

    private static final String XLS = ".xls";

    private static final String XLSX = ".xlsx";

    private static final String MONDAY = "星期一";

    private static final String TUESDAY = "星期二";

    private static final String WEDNESDAY = "星期三";

    private static final String THURSDAY = "星期四";

    private static final String FRIDAY = "星期五";

    private static final String SATURDAY = "星期六";

    private static final String SUNDAY = "星期日";

    private static final String DATE_FORMAT_FOR_DATE = "yyyy年MM月dd日";

    private static final String DATE_FORMAT = "yyyy年MM月dd日 HH:mm:ss";

    private static final String DATE_FORMAT_FOR_TIME = "HH:mm:ss";



    /**
     * 解析Excel表格
     *
     * @param path 文件路径
     * @return
     * @throws Exception
     */
    public static ExcelSet resolveExcel(String path) {
        ExcelSet excelSet = new ExcelSet();
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
            ExcelSheet sheets = new ExcelSheet();
            Iterator<Sheet> its = workbook.sheetIterator();
            while (its.hasNext()) {
                Sheet sheet = its.next();
                List<ExcelDTO> specificValues = Lists.newArrayListWithExpectedSize(sheet.getLastRowNum());
                sheets.setName(sheet.getSheetName());
                formatDateAndTime(sheet, specificValues);
                Map<String, List<ExcelDTO>> groupByUserNumber = specificValues.stream().collect(Collectors.groupingBy(ExcelDTO::getUserNumber));
                List<ExcelDTO> finalResult = dealGroupByUserNumber(groupByUserNumber);
                List<ResultVO> content = ResultVO.convertDtoToVOByTimeSort(finalResult);
                sheets.setContent((content));
            }
            return ExcelSet.builder().sheets(sheets).excelFile(file).build();
        }
    }

    private static List<ExcelDTO> dealGroupByUserNumber(Map<String, List<ExcelDTO>> groupResult) {
        List<ExcelDTO> finalResult = Lists.newArrayList();
        for (Map.Entry<String, List<ExcelDTO>> key : groupResult.entrySet()) {
            List<ExcelDTO> currentGroup = key.getValue();
            Map<String, List<ExcelDTO>> groupByDate = currentGroup.stream().collect(Collectors.groupingBy(ExcelDTO::getSignInDate));
            ExcelDTO currentUser = dealGroupByDate(groupByDate);
            finalResult.add(currentUser);
        }
        return finalResult;
    }

    private static ExcelDTO dealGroupByDate(Map<String, List<ExcelDTO>> groupByDate) {
        List<ExcelDTO> result = Lists.newArrayList();
        for (Map.Entry<String, List<ExcelDTO>> key : groupByDate.entrySet()) {
            List<ExcelDTO> groupByDateList = key.getValue();
            ExcelDTO middleExcelDTO = calculatePerDayWorkTime(groupByDateList);
            result.add(middleExcelDTO);
        }
        ExcelDTO excelDTO = dealDataForNormalTime(result);
        DecimalFormat df = new DecimalFormat("#.00");

        currentUser.setWorkDays(excelDTO.getWorkDays());
        double sumWorkTime = excelDTO.getSumWorkTime();
        currentUser.setSumWorkTime(Double.valueOf(df.format(sumWorkTime)));
        Integer days = currentUser.getWorkDays();
        double averageWorkTime = sumWorkTime / days;
        currentUser.setAverageWorkTime(Double.valueOf(df.format(averageWorkTime)));
        return currentUser;
    }

    private static ExcelDTO dealDataForAverageTime(List<ExcelDTO> normalWorkDays, List<ExcelDTO> overTimeWorkDays) {
        ExcelDTO excelDTO = new ExcelDTO();
        double middleResult = 0;
        for (int i = 0; i < normalWorkDays.size(); i++) {
            final double middle = normalWorkDays.get(i).getCurrentWorkTime();
            middleResult += middle;
        }
        for (int i = 0; i < overTimeWorkDays.size(); i++) {
            final double middle = overTimeWorkDays.get(i).getCurrentWorkTime();
            middleResult += middle;
        }
        excelDTO.setSumWorkTime(middleResult);
        excelDTO.setWorkDays(normalWorkDays.size());
        return excelDTO;
    }

    private static ExcelDTO dealDataForNormalTime(List<ExcelDTO> result) {
        SimpleDateFormat sd = new SimpleDateFormat(DATE_FORMAT_FOR_DATE);
        String[] weekDays = {MONDAY, TUESDAY, WEDNESDAY, THURSDAY, FRIDAY, SATURDAY, SUNDAY};
        Calendar cal = Calendar.getInstance();
        Date dateLine = null;
        List<ExcelDTO> normalWorkDays = Lists.newArrayListWithExpectedSize(result.size());
        List<ExcelDTO> overTimeWorkDays = Lists.newArrayListWithExpectedSize(result.size());
        for (ExcelDTO o : result) {
            try {
                dateLine = sd.parse(o.getSignInDate());
                cal.setTime(dateLine);
            } catch (ParseException e) {
                e.printStackTrace();
            }
            int w = cal.get(Calendar.DAY_OF_WEEK) - 1; // 指示一个星期中的某天。
            if (w < 0) {
                w = 0;
            }
            String currentWeekDay = weekDays[w];
            if (currentWeekDay.equals(SATURDAY) || currentWeekDay.equals(SUNDAY)) {
                overTimeWorkDays.add(o);
            } else {
                normalWorkDays.add(o);
            }
        }
        return dealDataForAverageTime(normalWorkDays, overTimeWorkDays);
    }

    private static void formatDateAndTime(Sheet sheet, List<ExcelDTO> content) {

        DateFormat parseDate = new SimpleDateFormat(DATE_FORMAT);
        for (int i = 1; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            ExcelDTO convertEachExcelLine = new ExcelDTO();
            String username = String.valueOf(row.getCell(CELL_INDEX_TWO));
            convertEachExcelLine.setUserName(username);
            String userNumber = String.valueOf(row.getCell(CELL_INDEX_ONE));
            convertEachExcelLine.setUserNumber(userNumber);
            Date signInDate = HSSFDateUtil.getJavaDate(row.getCell(CELL_INDEX_FOUR).getNumericCellValue());
            String specificDate = parseDate.format(signInDate).split(EMPTY_SPACE)[INDEX_ZERO];
            convertEachExcelLine.setSignInDate(specificDate);
            Date signInTime = HSSFDateUtil.getJavaDate(row.getCell(CELL_INDEX_FIVE).getNumericCellValue());
            String specificTime = parseDate.format(signInTime).split(EMPTY_SPACE)[INDEX_ONE];
            convertEachExcelLine.setSignInTime(specificTime);
            content.add(convertEachExcelLine);
        }
    }

    private static void contentAddHeadTitle(Sheet sheet, List<ExcelDTO> content) {
        Row head = sheet.getRow(INDEX_ZERO);
        String userName = String.valueOf(head.getCell(CELL_INDEX_TWO));
        String userNumber = String.valueOf(String.valueOf(head.getCell(CELL_INDEX_ONE)));
        String signInDate = String.valueOf(head.getCell(CELL_INDEX_FOUR));
        String signInTime = String.valueOf(head.getCell(CELL_INDEX_FIVE));
        ExcelDTO headLine = ExcelDTO.builder().userName(userName).userNumber(userNumber)
                .signInDate(signInDate).signInDate(signInTime).build();
        content.add(headLine);
    }

    private static ExcelDTO calculatePerDayWorkTime(List<ExcelDTO> groupByDateList) {
        if (groupByDateList.size() == 1) {
            return dealSingleSignInDate(groupByDateList);
        } else if (groupByDateList.size() == 2) {
            return dealCorrectSignInDate(groupByDateList);
        } else {
            return dealDuplicateSignInDate(groupByDateList);
        }
    }

    private static ExcelDTO dealSingleSignInDate(List<ExcelDTO> groupByDateList) {
        DateFormat df = new SimpleDateFormat(DATE_FORMAT_FOR_TIME);
        ExcelDTO defaultValue = new ExcelDTO();
        String userName = groupByDateList.get(0).getUserName();
        defaultValue.setUserName(userName);
        String userNumber = groupByDateList.get(0).getUserNumber();
        defaultValue.setUserNumber(userNumber);
        String signInDate = groupByDateList.get(0).getSignInDate();
        defaultValue.setSignInDate(signInDate);

        List<ExcelDTO> middleResult = Lists.newArrayListWithExpectedSize(2);
        middleResult.add(groupByDateList.get(0));
        try {
            Date dealLine = df.parse("12:00:00");
            String currentSignInTime = String.valueOf(groupByDateList.get(0).getSignInTime());
            Date currentTime = df.parse(currentSignInTime);
            if (currentTime.before(dealLine)) {
                defaultValue.setSignInTime("20:00:00");
            } else {
                defaultValue.setSignInTime("10:00:00");
            }
            middleResult.add(defaultValue);
        } catch (ParseException e) {
            System.out.println(groupByDateList.get(0).getSignInDate());
            e.printStackTrace();
        }
        return dealCorrectSignInDate(middleResult);
    }

    private static ExcelDTO dealCorrectSignInDate(List<ExcelDTO> duplicateSignInList) {
        ExcelDTO middleResult = new ExcelDTO();
        try {
            DateFormat df = new SimpleDateFormat(DATE_FORMAT_FOR_TIME);
            Date dealLine = df.parse("12:00:00");

            String username = duplicateSignInList.get(0).getUserName();
            String date = duplicateSignInList.get(0).getSignInDate();
            String startWorkTime = duplicateSignInList.get(0).getSignInTime();
            String startWorkTimeSeconds = startWorkTime.split(":")[2];
            String sameTimeForStartWorkTime = startWorkTime.split(":")[0] + startWorkTime.split(":")[1];
            String endWorkTime = duplicateSignInList.get(1).getSignInTime();
            String endWorkTimeSeconds = endWorkTime.split(":")[2];
            String sameTimeForEndWorkTime = endWorkTime.split(":")[0] + endWorkTime.split(":")[1];
            Integer secondsValue = Math.abs(Integer.parseInt(startWorkTimeSeconds) - Integer.parseInt(endWorkTimeSeconds));

            Date convertStartWorkTime = df.parse(startWorkTime);
            Date convertEndWorkTime = df.parse(endWorkTime);
            if (sameTimeForEndWorkTime.equals(sameTimeForStartWorkTime) && secondsValue < 60) {
                if (convertStartWorkTime.after(dealLine)) {
                    startWorkTime = "10:00:00";
                    convertStartWorkTime = df.parse(startWorkTime);
                } else {
                    endWorkTime = "19:00:00";
                    convertEndWorkTime = df.parse(endWorkTime);
                }
            }

            long diff = convertStartWorkTime.getTime() - convertEndWorkTime.getTime();
            long days = diff / (1000 * 60 * 60 * 24);
            long hours = Math.abs((diff - days * (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
            long minutes = Math.abs((diff - days * (1000 * 60 * 60 * 24) - hours * (1000 * 60 * 60)) / (1000 * 60));

            String userName = duplicateSignInList.get(0).getUserName();
            String userNumber = duplicateSignInList.get(0).getUserNumber();
            String signInDate = duplicateSignInList.get(0).getSignInDate();
            middleResult.setSignInDate(signInDate);

            middleResult.setUserName(userName);
            middleResult.setUserNumber(userNumber);

            DecimalFormat format = new DecimalFormat("0.00");
            String currentPerDayWorkTime = hours + "." + minutes;
            String middleValue = format.format(new BigDecimal(currentPerDayWorkTime));
            Double currentWorkTime = Double.valueOf(middleValue);
            middleResult.setCurrentWorkTime(currentWorkTime);

        } catch (ParseException e) {
            System.out.println(duplicateSignInList.get(0).getSignInDate());
            System.out.println(duplicateSignInList.get(0).getUserName());
            e.printStackTrace();
        }
        return middleResult;
    }

    private static ExcelDTO dealDuplicateSignInDate(List<ExcelDTO> current) {
        List<ExcelDTO> result = Lists.newArrayListWithExpectedSize(2);
        try {
            DateFormat df = new SimpleDateFormat(DATE_FORMAT_FOR_TIME);
            Date MinSignInTime = df.parse(String.valueOf(current.get(0).getSignInTime()));
            Date MaxSignInTime = df.parse(String.valueOf(current.get(1).getSignInTime()));
            if (MinSignInTime.after(MaxSignInTime)) {
                Date middle = MinSignInTime;
                MinSignInTime = MaxSignInTime;
                MaxSignInTime = middle;
            }
            for (int i = 2; i < current.size(); i++) {
                Date waitForCompareDate = df.parse(String.valueOf(current.get(i).getSignInTime()));
                if (waitForCompareDate.before(MinSignInTime)) {
                    MinSignInTime = waitForCompareDate;
                } else if (waitForCompareDate.after(MaxSignInTime)) {
                    MaxSignInTime = waitForCompareDate;
                }
            }
            // 逻辑有问题
            ExcelDTO minSignInTime = new ExcelDTO();
            String username = current.get(0).getUserName();
            minSignInTime.setUserName(username);
            String userNumber = current.get(0).getUserNumber();
            String signInDate = current.get(0).getSignInDate();
            minSignInTime.setSignInDate(signInDate);
            minSignInTime.setUserNumber(userNumber);
            minSignInTime.setSignInTime(df.format(MinSignInTime));

            ExcelDTO maxSigInTime = new ExcelDTO();
            maxSigInTime.setUserNumber(username);
            maxSigInTime.setUserNumber(userNumber);
            maxSigInTime.setSignInDate(signInDate);
            maxSigInTime.setSignInTime(df.format(MaxSignInTime));

            result.add(minSignInTime);
            result.add(maxSigInTime);

        } catch (Exception e) {
            e.printStackTrace();
        }
        return dealCorrectSignInDate(result);
    }
}
