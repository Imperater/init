package com.tqc.resolveexcel.util;

import com.fasterxml.jackson.core.JsonToken;
import com.google.common.collect.Lists;
import com.tqc.resolveexcel.model.ExcelDTO;
import com.tqc.resolveexcel.model.ResultVO;
import com.tqc.resolveexcel.model.excel.ExcelSet;
import com.tqc.resolveexcel.model.excel.ExcelSheet;
import lombok.SneakyThrows;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;

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

    private static final Integer CELL_INDEX_THREE = 3;

    private static final Integer CELL_INDEX_FOUR = 4;

    private static final Integer CELL_INDEX_FIVE = 5;

    private static final String EMPTY_SPACE = " ";

    private static final Integer INDEX_ZERO = 0;

    private static final Integer INDEX_ONE = 1;

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
            ExcelSheet sheets = new ExcelSheet();
            Iterator<Sheet> its = workbook.sheetIterator();
            //处理每个sheet
            while (its.hasNext()) {
                Sheet sheet = its.next();
                List<ExcelDTO> specificValues = Lists.newArrayListWithExpectedSize(sheet.getLastRowNum());
                sheets.setName(sheet.getSheetName());
                formatDateAndTime(sheet, specificValues);
                Map<String, List<ExcelDTO>> groupByUserNumber = specificValues.stream().collect(Collectors.groupingBy(ExcelDTO::getUserNumber));
                List<ExcelDTO> finalResult = dealGroupByUserNumber(groupByUserNumber);
                List<ResultVO> content = ResultVO.convertDtoToVO(finalResult);
                List<ResultVO> sourtData = content.stream().sorted(Comparator.comparing(ResultVO::getAverageWorkTime)).collect(Collectors.toList());
                sheets.setContent((sourtData));
            }
            return ExcelSet.builder().sheets(sheets).excelFile(file).build();
        }
    }

    private static List<ExcelDTO> dealGroupByUserNumber(Map<String, List<ExcelDTO>> groupResult) {
        List<ExcelDTO> finalResult = Lists.newArrayList();

        for (Map.Entry<String, List<ExcelDTO>> key : groupResult.entrySet()) {
            List<ExcelDTO> currentGroup = key.getValue();
            ExcelDTO currentUser = new ExcelDTO();
            String userName = currentGroup.get(0).getUserName();
            currentUser.setUserName(userName);
            String userNumber = currentGroup.get(0).getUserNumber();
            currentUser.setUserNumber(userNumber);
            Map<String, List<ExcelDTO>> groupByDate = currentGroup.stream().collect(Collectors.groupingBy(ExcelDTO::getSignInDate));
            currentUser.setWorkDays(groupByDate.size());
            dealGroupByDate(groupByDate, currentUser);
            finalResult.add(currentUser);
        }
        return finalResult;
    }

    private static ExcelDTO dealGroupByDate(Map<String, List<ExcelDTO>> groupByDate, ExcelDTO currentUser) {
        List<ExcelDTO> result = Lists.newArrayList();
        for (Map.Entry<String, List<ExcelDTO>> key : groupByDate.entrySet()) {
            List<ExcelDTO> groupByDateList = key.getValue();
            ExcelDTO middleExcelDTO = calculatePerDayWorkTime(groupByDateList);
            result.add(middleExcelDTO);
        }
        dealDataForAverageTime(result);
        double middleResult = 0;
        for (int i = 0; i < result.size(); i++) {
            double middle = result.get(i).getCurrentWorkTime();
            middleResult += middle;
        }
        DecimalFormat df = new DecimalFormat("#.00");

        currentUser.setSumWorkTime(middleResult);
        Integer days = currentUser.getWorkDays();
        double averageWorkTime = middleResult / days;
        currentUser.setAverageWorkTime(Double.valueOf(df.format(averageWorkTime)));
        double sumWorkTime = currentUser.getSumWorkTime();
        currentUser.setSumWorkTime(Double.valueOf(df.format(sumWorkTime)));
        return currentUser;
    }

  /*  private static List<List<Object>> getAverageTime(List<ExcelDTO> dataClean) {
        List<Object> names = new ArrayList<>();
        List<ExcelDTO> result = Lists.newArrayListWithExpectedSize(dataClean.size());

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
    }*/

    private static ExcelDTO dealDataForAverageTime(List<ExcelDTO> result) {
        ExcelDTO excelDTO = new ExcelDTO();
        String userName = result.get(0).getUserName();
        excelDTO.setUserName(userName);
        String userNumber = result.get(0).getUserNumber();
        excelDTO.setUserNumber(userNumber);
        Double sumWorkTime = result.stream().filter(O -> Objects.nonNull(excelDTO.getCurrentWorkTime())).mapToDouble(ExcelDTO::getCurrentWorkTime).sum();
        excelDTO.setSumWorkTime(sumWorkTime);
        return excelDTO;
    }

    private static void formatDateAndTime(Sheet sheet, List<ExcelDTO> content) {
        DateFormat parseDate = new SimpleDateFormat("MM月dd日 HH:mm:ss");
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

    /* private static List<ExcelDTO> dealDataWithSameNameByList(List<ExcelDTO> content) {
         List<Object> names = new ArrayList<>();
         List<ExcelDTO> dataCleanResult = new ArrayList<>();
         for (int i = 1; i < content.size(); i++) {
             String currentName = String.valueOf(content.get(i).getUserName());
             if (!names.contains(currentName)) {
                 List<ExcelDTO> waitDealList = new ArrayList<>();
                 names.add(currentName);
                 for (int j = 1; j < content.size(); j++) {
                     String dealName = String.valueOf(content.get(j).getUserName());
                     if (currentName.equals(dealName)) {
                         waitDealList.add(content.get(j));
                     }
                 }
                 dealDataWithSameDataByList(waitDealList, dataCleanResult);
             }
         }
         return dataCleanResult;
     }*/
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

    /* private static void dealDataWithSameDataByList(List<ExcelDTO> waitDealList, List<ExcelDTO> dataCleanResult) {
         List<String> date = new ArrayList<>();

         for (int i = 0; i < waitDealList.size(); i++) {
             String currentDate = String.valueOf(waitDealList.get(i).getSignInDate());
             if (!date.contains(currentDate)) {
                 List<ExcelDTO> current = new ArrayList<>();
                 date.add(currentDate);
                 for (int j = 0; j < waitDealList.size(); j++) {
                     String needDealDate = String.valueOf(waitDealList.get(j).getSignInDate());
                     if (currentDate.equals(needDealDate)) {
                         current.add(waitDealList.get(j));
                     }
                 }
                 calculateTime(current, dataCleanResult);
             }
         }
     }*/
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
        DateFormat df = new SimpleDateFormat("HH:mm:ss");
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
                defaultValue.setSignInTime("19:00:00");
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
        DateFormat df = new SimpleDateFormat("HH:mm:ss");
        ExcelDTO middleResult = new ExcelDTO();
        try {
            String username = duplicateSignInList.get(0).getUserName();
            String date = duplicateSignInList.get(0).getSignInDate();
            String startWorkTime = duplicateSignInList.get(0).getSignInTime();
            String endWorkTime = duplicateSignInList.get(1).getSignInTime();
            Date convertStartWorkTime = df.parse(startWorkTime);
            Date convertEndWorkTime = df.parse(endWorkTime);
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
            DateFormat df = new SimpleDateFormat("HH:mm:ss");
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
