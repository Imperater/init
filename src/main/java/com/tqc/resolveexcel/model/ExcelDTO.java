package com.tqc.resolveexcel.model;

import lombok.*;

import java.util.ArrayList;
import java.util.Date;
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


    private Date date1;

    private Date date2;

    private Integer no;

    private Integer machineId;

    private String suaKaDate;

    private String suaKaTime;

    private String name;

    private String workTime;

    private String averageTime;

    private List<ExcelDTO> convertResult(List<List<Object>> target) {
        List<ExcelDTO> result = new ArrayList<>();
        target.forEach(o ->result.add(ExcelDTO.builder()
                .name(o.get(3).toString())
                .suaKaDate(o.get(4).toString())
                .suaKaTime(o.get(5).toString())
                .build()));

        return result;
    }
}















