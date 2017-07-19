package com.kjubo.excel;

import java.util.Collections;
import java.util.List;

/**
 * Created by kjubo on 2017/4/19
 * ExcelColumn 支持CodeBean的 模式
 */
public interface ICodeable {

    /**
     * 加载CodeList，多用于创建Excel下拉选择框
     * @return
     */
    default List<? extends ICodeBean> loadCodeList() {
        return Collections.emptyList();
    }

    /**
     * 通过ID获取Name，多用于导出
     * @param id
     * @return
     */
    String getName(String id);

    /**
     * 通过Name获取Id，多用于导入
     * @param name
     * @return
     */
    String getCode(String name);

    class None implements ICodeable {

        @Override
        public String getName(String id) {
            return null;
        }

        @Override
        public String getCode(String name) {
            return null;
        }
    }
}
