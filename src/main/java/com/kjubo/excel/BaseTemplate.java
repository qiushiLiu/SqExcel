package com.kjubo.excel;

import lombok.Data;

/**
 * Created by kjubo on 2017/7/19.
 */
@Data
public abstract class BaseTemplate implements IExcelRowIndex {
    private Boolean hasError = false;
}