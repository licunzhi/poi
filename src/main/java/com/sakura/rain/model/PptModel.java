package com.sakura.rain.model;

/**
 * @ClassName PptModel
 * @Description ppt base information
 * @Author lcz
 * @Date 2019/06/04 09:15
 * @notice: @dataType 1 text or number data
 *                    3 picture dataContent is picture path
 *                    4 picture dataContent is picture file
 *                    5 picture dataContent is base64 code
 *                    6 table
 *           PptUtils can find enum type
 */
public class PptModel<T> {
    private int dataType;
    private T dataConent;
    private T defaultContent;

    public int getDataType() {
        return dataType;
    }

    public void setDataType(int dataType) {
        this.dataType = dataType;
    }

    public T getDataConent() {
        return dataConent;
    }

    public void setDataConent(T dataConent) {
        this.dataConent = dataConent;
    }

    public T getDefaultContent() {
        return defaultContent;
    }

    public void setDefaultContent(T defaultContent) {
        this.defaultContent = defaultContent;
    }
}
