package com.tqc.resolveexcel.model;

import lombok.*;

import java.util.ArrayList;
import java.util.List;

/**
 * @author imperater
 * @date 8/1/19
 */
@Setter
@Getter
@Builder
@ToString
@NoArgsConstructor
@AllArgsConstructor
public class ExcelDTO {

    private Integer no;

    private Integer machineId;

    private String suaKaDate;

    private String suaKaTime;

    private String name;

    private String workTime;

    private String averageTime;

    private List<ExcelDTO> convertResult(List<List<Object>> target) {
        List<ExcelDTO> result = new ArrayList<>();
        for (int i = 0; i < target.size(); i++) {
            ExcelDTO excelDTO = new ExcelDTO();
            List<Object> content = target.get(i);
            excelDTO.setName(content.get(3).toString());
            excelDTO.setSuaKaDate(content.get(4).toString());
            excelDTO.setSuaKaTime(content.get(5).toString());
            result.add(excelDTO);
        }

        return result;
    }
}















