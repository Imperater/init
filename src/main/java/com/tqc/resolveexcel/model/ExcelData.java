package com.tqc.resolveexcel.model;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.io.Serializable;
import java.util.List;

@Setter
@Getter
@NoArgsConstructor
@AllArgsConstructor
public class ExcelData implements Serializable {

    private List<String> titles;

    private List<ResultVO> rows;

    private String name;
}
