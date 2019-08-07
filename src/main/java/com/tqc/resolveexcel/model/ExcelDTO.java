package com.tqc.resolveexcel.model;

import com.google.common.collect.Lists;
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

    private String userNumber;

    private Integer workDays;

    private String userName;

    private String signInDate;

    private String signInTime;

    /**
     * work time
     */
    private String startWorkTime;

    private String endWorkTime;

    private Double currentWorkTime;

    private Double sumWorkTime;

    private Date workDate;

    private Double averageWorkTime;
}















