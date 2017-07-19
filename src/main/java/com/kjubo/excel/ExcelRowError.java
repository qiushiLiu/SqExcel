package com.kjubo.excel;

import lombok.Data;
import org.apache.commons.collections4.CollectionUtils;

import javax.validation.ConstraintViolation;
import java.util.List;

/**
 * Created by kjubo on 2017/4/19.
 */
@Data
public class ExcelRowError {
    /**
     * excel 对应行号
     */
    Integer rowIndex;

    /**
     * 校验不合法内容
     */
    List<ConstraintViolation> errors;

    @Override
    public String toString() {
        StringBuffer buffer = new StringBuffer();
        if (CollectionUtils.isNotEmpty(this.errors)) {
            buffer.append("第").append(this.getRowIndex()).append("行:\n");
            for (ConstraintViolation item : this.errors) {
                buffer.append(item.getPropertyPath().toString())
                        .append(item.getMessage())
                        .append("\n");
            }
        }
        return buffer.toString();
    }
}
