package com.kjubo.excel;

import com.kjubo.excel.validation.annotation.IsDate;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.springframework.beans.BeansException;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ApplicationContextAware;
import org.springframework.stereotype.Component;

import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import javax.validation.ValidatorFactory;
import javax.validation.constraints.Digits;
import javax.validation.constraints.NotBlank;
import javax.validation.constraints.NotEmpty;
import javax.validation.constraints.NotNull;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.MessageFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

@Slf4j
@Component
public class ExcelTemplate<T extends BaseTemplate> implements ApplicationContextAware {

    private static ApplicationContext ctx;

    @Override
    public void setApplicationContext(ApplicationContext applicationContext) throws BeansException {
        ctx = applicationContext;
    }

    private static final Integer EXCEL_LIMIT_ROW_NUM = 65535;
    private static final String REQUIRED_MARK = "*";

    /**
     * ExcelTemplate对应类
     */
    private Class<T> clazz = null;

    /**
     * 标注必填字段
     * 凡是被@NotNull, @NotEmpty, @NotBlank 修饰的属性
     * 在标题栏会在标题名称后面默认添加 "*" 表示必填
     */
    private boolean markRequiredProperty = true;

    /**
     * 导入行数
     */
    @Getter
    private int count = 0;

    /**
     * 成功行数
     */
    @Getter
    private int success = 0;

    /**
     * 行错误的具体信息
     */
    @Getter
    private List<ExcelRowError> errors = new ArrayList<>();

    public boolean hasError() {
        return success < count;
    }

    private ExcelTemplate() {
    }

    private ExcelTemplate(Class<T> clazz, boolean markRequiredProperty) {
        this.clazz = clazz;
        this.markRequiredProperty = markRequiredProperty;
    }

    /**
     * 获取实例的方法
     *
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T extends BaseTemplate> ExcelTemplate<T> of(Class<T> clazz) {
        return of(clazz, true);
    }

    public static <T extends BaseTemplate> ExcelTemplate<T> of(Class<T> clazz, boolean markRequiredProperty) {
        return new ExcelTemplate<>(clazz, markRequiredProperty);
    }

    public List<T> importExcel(InputStream inputStream) throws IOException {
        return this.importExcel(inputStream, null, 1);
    }

    public List<T> importExcel(InputStream inputStream, Map<String, String> titleMapper) throws IOException {
        return this.importExcel(inputStream, titleMapper, 1);
    }

    /***
     * 将excel转化为对象列表
     * @param inputStream    excel文件流
     * @param titleMapper    标题转化数据，可以为空
     * @param beginRowNum    excel数据开始行，默认值为1
     * @return
     * @throws IOException
     */
    public List<T> importExcel(final InputStream inputStream,
                               final Map<String, String> titleMapper,
                               final Integer beginRowNum) throws IOException {
        if (inputStream == null) {
            return Collections.emptyList();
        }
        List<T> list = new ArrayList<>();
        List<ExcelColumnInfo> colInfo = this.getTemplateColumnInfo(titleMapper);

        // 读取文件
        try (Workbook wb = WorkbookFactory.create(inputStream)) {
            this.count = 0;
            this.success = 0;
            this.errors.clear();

            ValidatorFactory factory = Validation.buildDefaultValidatorFactory();
            Validator validator = factory.getValidator();
            // 取得对口地域
            Sheet sheet = wb.getSheetAt(0);
            // 得到总行数
            for (int i = beginRowNum, rowNum = sheet.getLastRowNum(); i <= rowNum; i++) {
                Row row = sheet.getRow(i);
                if (this.isRowEmpty(row)) {    //跳过全部是空白的行
                    continue;
                }
                T item = this.createRowObject(row, colInfo);
                //设置excel物理行数
                item.setExcelRowIndex(i + 1);
                Set<ConstraintViolation<T>> violations = validator.validate(item);
                this.count++;

                if (violations.isEmpty()) {
                    this.success++;
                } else {
                    ExcelRowError error = this.getRowError(i + 1, violations, colInfo, titleMapper);
                    this.errors.add(error);
                    item.setHasError(true);
                }
                list.add(item);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return list;
    }

    /**
     * 将excel中一行转化为一个对象
     * @param row
     * @param colInfo
     * @return
     * @throws IllegalAccessException
     * @throws ParseException
     * @throws InstantiationException
     */
    private T createRowObject(final Row row, final List<ExcelColumnInfo> colInfo)
            throws IllegalAccessException, ParseException, InstantiationException {
        T item = clazz.newInstance();
        for (int index = 0; index < colInfo.size(); index++) {
            Cell cell = row.getCell(index, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            ExcelColumnInfo columnInfo = colInfo.get(index);
            Field field = columnInfo.getField();
            field.setAccessible(true);

            if (columnInfo.getExcelColumnCodeable() != null) {
                String codeName = Optional.ofNullable(this.getCellValue(cell.getCellTypeEnum(), cell)).orElse("").toString();
                if (StringUtils.isEmpty(codeName)) {
                    field.set(item, "");
                } else {
                    String codeValue = Optional.ofNullable(columnInfo.getExcelColumnCodeable().getCode(codeName)).orElse("");
                    field.set(item, codeValue);
                }
            } else {
                this.setFieldValue(cell, item, columnInfo);
            }
        }
        return item;
    }

    /**
     * 构建结构化 violations
     *
     * @param rowIndex
     * @param violations
     * @param colInfo
     * @param titles
     * @return
     */
    private ExcelRowError getRowError(Integer rowIndex,
                                      Set<ConstraintViolation<T>> violations,
                                      List<ExcelColumnInfo> colInfo,
                                      Map<String, String> titles) {
        ExcelRowError error = new ExcelRowError();
        error.setRowIndex(rowIndex);
        List<ConstraintViolation> list = new ArrayList<>();
        if (CollectionUtils.isNotEmpty(violations)) {
            list.addAll(violations);
        }
        error.setErrors(list);
        return error;
    }

    private String getFieldName(String field, List<ExcelColumnInfo> colInfo, Map<String, String> titles) {
        String name = colInfo
                .stream()
                .filter(p -> StringUtils.equalsIgnoreCase(field, p.getFieldName()))
                .map(ExcelColumnInfo::getName)
                .findFirst()
                .orElse("");
        if (titles != null) {
            return titles.getOrDefault(name, name);
        }
        return name;
    }

    public List<ExcelColumnInfo> getTemplateColumnInfo() {
        return this.getTemplateColumnInfo(null);
    }

    /**
     * 获取模板的列属性
     * @param titleMapper   excel的标题栏映射关系
     * @return
     */
    public List<ExcelColumnInfo> getTemplateColumnInfo(final Map<String, String> titleMapper) {
        if (this.clazz == null) {
            throw new IllegalArgumentException("Class Type is null");
        }

        Field[] declaredFields = clazz.getDeclaredFields();
        if (declaredFields == null) {
            return Collections.emptyList();
        }

        return Arrays.asList(declaredFields)
                .stream()
                .filter(p -> p.isAnnotationPresent(ExcelColumn.class))
                .map(field -> {
                    ExcelColumn column = field.getAnnotation(ExcelColumn.class);
                    ExcelColumnInfo columnInfo = new ExcelColumnInfo(column);
                    if (!column.coding().equals(ICodeable.None.class)) {
                        columnInfo.setExcelColumnCodeable(ctx.getBean(column.coding()));
                    }

                    if (field.isAnnotationPresent(IsDate.class)) {
                        columnInfo.setDateFormat(field.getAnnotation(IsDate.class).format());
                    }

                    columnInfo.setField(field);
                    columnInfo.setFieldName(field.getName());
                    if (titleMapper != null) {
                        columnInfo.setColName(titleMapper.getOrDefault(column.name(), column.name()));
                    }
                    return columnInfo;
                })
                .sorted(Comparator.comparing(ExcelColumnInfo::getCol))
                .collect(Collectors.toList());
    }

    /**
     * 单元格值取得处理
     *
     * @param cell
     * @return
     * @throws IllegalAccessException
     * @throws IllegalArgumentException
     * @throws ParseException
     */
    private void setFieldValue(Cell cell, Object target, ExcelColumnInfo info) throws IllegalArgumentException, IllegalAccessException, ParseException {
        if (cell == null
                || target == null
                || info == null) {
            return;
        }
        Field field = info.getField();
        Class<?> type = field.getType();

        Object value = this.getCellValue(cell.getCellTypeEnum(), cell);
        if (value == null) {
            return;
        }
        if (type.isAssignableFrom(value.getClass())) {
            field.set(target, value);
        } else {
            if (type == Date.class) {
                field.set(target, DateUtils.parseDateStrictly(value.toString(), info.getDateFormat()));
            } else if (Number.class.isAssignableFrom(type)) {
                BigDecimal dValue;
                dValue = new BigDecimal(value.toString());
                if (field.isAnnotationPresent(Digits.class)) {
                    Digits digits = field.getAnnotation(Digits.class);
                    dValue.setScale(digits.fraction());
                }
                if (BigDecimal.class.isAssignableFrom(type)) {
                    field.set(target, dValue);
                } else if (Integer.class.isAssignableFrom(type) || type.getName().equals("int")) {
                    field.set(target, dValue.intValue());
                } else if (Short.class.isAssignableFrom(type) || type.getName().equals("short")) {
                    field.set(target, dValue.shortValue());
                } else if (Float.class.isAssignableFrom(type) || type.getName().equals("float")) {
                    field.set(target, dValue.floatValue());
                } else if (Byte.class.isAssignableFrom(type) || type.getName().equals("byte")) {
                    field.set(target, dValue.byteValue());
                } else if (Double.class.isAssignableFrom(type) || type.getName().equals("double")) {
                    field.set(target, dValue.doubleValue());
                } else {
                    field.set(target, value);
                }
            } else if (type == String.class) {
                if (value.getClass() == Date.class) {
                    field.set(target, dateFormat((Date) value, info.getDefaultDateFormat()));
                } else if (value.getClass() == Double.class) {
                    BigDecimal decimal = new BigDecimal((Double) value);
                    field.set(target, decimal.stripTrailingZeros().toPlainString());
                } else {
                    field.set(target, value.toString());
                }
            } else {
                field.set(target, value);
            }
        }
    }

    private Object getCellValue(CellType cellType, Cell cell) {
        Object value = null;
        switch (cellType) {
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    value = cell.getDateCellValue();
                } else {
                    value = new Double(cell.getNumericCellValue());
                }
                break;
            case STRING:
                value = cell.getRichStringCellValue().getString().trim();
                break;
            // 公式类型
            case FORMULA:
                value = this.getCellValue(cell.getCachedFormulaResultTypeEnum(), cell);
                break;
            // 布尔类型
            case BOOLEAN:
                value = new Boolean(cell.getBooleanCellValue());
                break;
            // 空值
            case BLANK:
                break;
            // 故障
            case ERROR:
                break;
            default:
                value = cell.getStringCellValue().trim();
        }
        return value;
    }

    /**
     * 添加头部行
     *
     * @param sheet
     * @param workbook
     * @param cellStyle
     * @param colInfo
     * @return 这一行数据每一列的宽度推荐
     */
    private int[] addHeaderRow(Sheet sheet, Workbook workbook, CellStyle cellStyle, List<ExcelColumnInfo> colInfo) {
        if (sheet == null
                || CollectionUtils.isEmpty(colInfo)) {
            throw new IllegalArgumentException("excel column info can not be null");
        }

        int rowNum = sheet.getLastRowNum();
        if (sheet.getPhysicalNumberOfRows() > 0) {
            rowNum += 1;
        }

        Row row = sheet.createRow(rowNum);
        int colNum = 0;
        int[] cellWidth = new int[colInfo.size()];
        for (int colIndex = 0; colIndex < colInfo.size(); colIndex++) {
            ExcelColumnInfo col = colInfo.get(colIndex);
            String cellValue = col.getName();
            if (this.markRequiredProperty
                    && this.isFieldRequired(col.getField())) {
                cellValue += REQUIRED_MARK;
            }
            Cell cell = row.createCell(colNum);
            cell.setCellValue(cellValue);
            cell.setCellStyle(cellStyle);
            //根据CodeName去添加下拉框
            if (col.getExcelColumnCodeable() != null) {
                CellRangeAddressList range = new CellRangeAddressList(1, EXCEL_LIMIT_ROW_NUM, colIndex, colIndex);
                this.createDropDownListDataValidation(workbook, sheet, col, range);
            }
            colNum++;
            if (col.getName() != null) {
                cellWidth[colIndex] = col.getName().getBytes().length;
            }
        }
        return cellWidth;
    }

    /**
     * 添加数据行
     *
     * @param sheet
     * @param data
     * @return 这一行数据每一列的宽度推荐
     */
    private int[] addDataRow(Sheet sheet, T data, List<ExcelColumnInfo> colInfo) {
        Row row = sheet.createRow(sheet.getLastRowNum() + 1);
        int[] cellWidth = new int[colInfo.size()];
        for (int i = 0; i < colInfo.size(); i++) {
            ExcelColumnInfo col = colInfo.get(i);
            Cell cell = row.createCell(i);
            col.getField().setAccessible(true);
            String value = null;
            try {
                Class<?> type = col.getField().getType();
                Object object = col.getField().get(data);
                if (object == null) {
                    continue;
                } else if (type.equals(Date.class)) {
                    value = dateFormat((Date) object, col.getDefaultDateFormat());
                } else if (type.equals(BigDecimal.class)) {
                    value = col.getField().get(data).toString();
                } else if (type.equals(Integer.class)) {
                    value = String.valueOf(col.getField().get(data));
                } else if (col.getExcelColumnCodeable() != null) {
                    value = col.getExcelColumnCodeable().getName(object.toString());
                } else {
                    value = object.toString();
                }
            } catch (Exception e) {
                log.error(e.toString());
            }
            if (value == null) {
                value = "";
            }
            cellWidth[i] = value.getBytes().length;
            cell.setCellType(CellType.STRING);
            cell.setCellValue(value);
        }
        return cellWidth;
    }

    /**
     * 生成excel的workBook
     *
     * @param workbook
     * @param list           数据源
     * @param requireFields  需要导出的字段名称
     * @param titles         属性名称的映射集合
     * @param repeatTitleRow 头部是否需要重复
     */
    public void generateExcel(Workbook workbook,
                              List<T> list,
                              List<String> requireFields,
                              LinkedHashMap<String, String> titles,
                              boolean repeatTitleRow) {
        if (workbook == null) {
            workbook = new SXSSFWorkbook(100);
            ((SXSSFWorkbook) workbook).setCompressTempFiles(true);
        }
        Sheet sheet;
        if (workbook.getNumberOfSheets() > 0) {
            sheet = workbook.getSheetAt(0);
        } else {
            sheet = workbook.createSheet();
        }
        CellStyle cellStyle = this.getTitleStyle(workbook);
        List<ExcelColumnInfo> colInfo = this.getTemplateColumnInfo(titles);
        //根据所需要的字段重新组合ColumnInfo
        if (CollectionUtils.isNotEmpty(requireFields)) {
            colInfo = colInfo.stream()
                    .filter(p -> requireFields.contains(p.getFieldName()))
                    .collect(Collectors.toList());
        }

        if (CollectionUtils.isNotEmpty(colInfo)) {
            int[] cellWidth = new int[colInfo.size()];
            int index = 0;
            do {
                if (index == 0 || repeatTitleRow) {    //并且是第一行，或者指定为重复头部的模式
                    this.mergeMaxValue(this.addHeaderRow(sheet, workbook, cellStyle, colInfo), cellWidth);
                }
                if (index < list.size()) {
                    T data = list.get(index);
                    this.mergeMaxValue(this.addDataRow(sheet, data, colInfo), cellWidth);
                }
                index++;
            } while (index < list.size());
            //调整单元格宽度
            for (int i = 0; i < cellWidth.length; i++) {
                sheet.setColumnWidth(i, Math.min(255, cellWidth[i]) * 256);
            }
        }
    }

    /**
     * 为Excel根据Coding注释，创建下拉选择框
     *
     * @param workbook
     * @param targetSheet
     * @param excelColumn
     * @param range
     */
    public void createDropDownListDataValidation(Workbook workbook, Sheet targetSheet,
                                                 ExcelColumnInfo excelColumn, CellRangeAddressList range) {
        if (workbook == null || targetSheet == null
                || excelColumn == null || excelColumn.getExcelColumnCodeable() == null || range == null) {
            return;
        }

        String codeType = excelColumn.getName();
        if (StringUtils.isEmpty(codeType)) {
            throw new IllegalArgumentException("the name value of excelColumn cannot be empty.");
        }

        List<? extends ICodeBean> beans = excelColumn.getExcelColumnCodeable().loadCodeList();
        if (CollectionUtils.isEmpty(beans)) {
            return;
        }
        Sheet sheet = workbook.getSheet(codeType);
        if (sheet == null) {    //根据CodeList创建Sheet
            sheet = workbook.createSheet(codeType);
            //数据源sheet页不显示
            workbook.setSheetHidden(workbook.getSheetIndex(codeType), true);
            List<String> listData = beans
                    .stream()
                    .map(ICodeBean::codeName)
                    .collect(Collectors.toList());

            for (int i = 0, length = listData.size(); i < length; i++) {
                Row hiddenRow = sheet.createRow(i);
                Cell hiddenCell = hiddenRow.createCell(0);
                hiddenCell.setCellValue(listData.get(i));
            }

            Name namedCell = workbook.createName();
            namedCell.setNameName(codeType);
            namedCell.setRefersToFormula(MessageFormat.format("{0}!$A$1:$A${1}", codeType, listData.size()));
        }
        XSSFDataValidationHelper dvHelper = (XSSFDataValidationHelper) targetSheet.getDataValidationHelper();
        DataValidationConstraint constraint = dvHelper.createFormulaListConstraint(codeType);
        XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(constraint, range);
        targetSheet.addValidationData(validation);
    }

    public Workbook generateExcel(List<T> list, LinkedHashMap<String, String> titles, List<String> requireFields, boolean repeatTitleRow) {
        SXSSFWorkbook workbook = new SXSSFWorkbook(100);
        workbook.setCompressTempFiles(true);
        this.generateExcel(workbook, list, requireFields, titles, repeatTitleRow);
        return workbook;
    }

    private void mergeMaxValue(int[] input, int[] target) {
        if (input == null || target == null) {
            return;
        }
        for (int i = 0; i < input.length && i < target.length; i++) {
            target[i] = Math.max(input[i], target[i]);
        }
    }

    /**
     * 为WorkBook添加Cell的样式和Font的样式
     *
     * @param wb
     * @return Cell的样式
     */
    private CellStyle getTitleStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        // 对齐方式设置
        style.setAlignment(HorizontalAlignment.CENTER);
        // 边框颜色和宽度设置
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        // 设置背景颜色
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 粗体字设置
        Font font = wb.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    /**
     * 判断某一行是否是空白的
     *
     * @param row
     * @return
     */
    private boolean isRowEmpty(Row row) {
        if (row == null) {
            return true;
        }
        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            if (cell != null
                    && cell.getCellTypeEnum() != CellType.BLANK) {
                return false;
            }
        }
        return true;
    }

    /**
     * 是否标注为必填字段
     *
     * @param field
     * @return
     */
    private boolean isFieldRequired(Field field) {
        if (field == null) {
            return false;
        }
        return field.isAnnotationPresent(NotNull.class)
                || field.isAnnotationPresent(NotEmpty.class)
                || field.isAnnotationPresent(NotBlank.class);
    }

    // 把日期转为字符串
    public static String dateFormat(Date date, String format) {
        if (date != null) {
            SimpleDateFormat sdf = new SimpleDateFormat();
            sdf.setLenient(false);
            sdf.applyPattern(format);

            try {
                return sdf.format(date);
            } catch (Exception ex) {
                return null;
            }
        } else {
            return null;
        }
    }
}