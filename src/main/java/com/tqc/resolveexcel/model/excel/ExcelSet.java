package com.tqc.resolveexcel.model.excel;


import lombok.*;

import java.io.File;
import java.util.List;

/**
 * Creator: qingchang.tang
 * Date: 2018/03/06
 * Time: 上午 9:45
 * 封装解析Excel完成后的内容
 */
@Getter
@Setter
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelSet {

    /**
     * 工作表列表
     */
    private ExcelSheet sheets;

    /**
     * excel文件信息
     */
    private File excelFile;
}
