package com.kjubo.excel.validation.validator;


import com.kjubo.excel.validation.annotation.IsDate;
import org.apache.commons.lang3.time.DateUtils;

import javax.validation.ConstraintValidator;
import javax.validation.ConstraintValidatorContext;
import java.text.ParseException;
import java.util.Date;


public class IsDateValidator implements ConstraintValidator<IsDate, Object> {

    private String[] format;

    @Override
    public void initialize(IsDate constraintAnnotation) {
        this.format = constraintAnnotation.format();
    }

    @Override
    public boolean isValid(Object value, ConstraintValidatorContext context) {

        if (value == null) {
            return true;
        }
        if (value instanceof Date) {
            return true;
        }

        String dateStr = String.valueOf(value);
        Date dt = null;
        try {
            dt = DateUtils.parseDateStrictly(dateStr, format);
        } catch (ParseException e) {
        }
        return dt != null;
    }
}
