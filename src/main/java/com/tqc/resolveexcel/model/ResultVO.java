package com.tqc.resolveexcel.model;

import com.google.common.collect.Lists;
import lombok.*;
import org.graalvm.compiler.graph.Node;

import java.util.Date;
import java.util.List;

@Getter
@Setter
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ResultVO {
    private String userNumber;

    private String userName;

    private Integer workDays;

    private Double sumWorkTime;

    private Double averageWorkTime;

    public static List<ResultVO> convertDtoToVO(List<ExcelDTO> target) {
        List<ResultVO> resultVOS = Lists.newArrayListWithExpectedSize(target.size());
        target.forEach(o -> resultVOS.add(convertData(o)));
        return resultVOS;
    }

    private static ResultVO convertData(ExcelDTO target) {
        ResultVO resultVO = new ResultVO();
        resultVO.setUserName(target.getUserName());
        resultVO.setUserNumber(target.getUserNumber());
        resultVO.setSumWorkTime(target.getSumWorkTime());
        resultVO.setWorkDays(target.getWorkDays());
        resultVO.setAverageWorkTime(target.getAverageWorkTime());
        return  resultVO;
    }
}