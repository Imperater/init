package com.tqc.resolveexcel.model.excel;

import lombok.Getter;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
public class ExcelSheet {

    /**
     * 工作表名
     */
    private String name;

    /**
     * 单元格内容
     * 第一层List代表每行、第二层List代表每个单元格
     */
    private List<List<Object>> content;
}
