package com.tqc.resolveexcel.model.excel;

import com.tqc.resolveexcel.model.ResultVO;
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
     */
    private List<ResultVO> content;
}
