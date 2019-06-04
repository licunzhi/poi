package com.sakura.rain.utils;

import com.sakura.rain.model.PptModel;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import sun.misc.BASE64Decoder;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @ClassName PptUtils
 * @Description create ppt base on ppt model
 * @Author lcz
 * @Date 2019/06/04 09:07
 * @notice java version support 1.8 or later
 */
public class PptUtils {
    public static final int TEXT_DATA = 1;
    public static final int NUMBER_DATA = 1;
    public static final int PICTURE_PATH_DATA = 3;
    public static final int PICTRUE_FILE_DATA = 4;
    public static final int PICTRUE_BASE64_DATA = 5;
    public static final int TABLE_DATA = 6;

    /**
     * @desc return file stream to user
     * @param file template file @notnull
     * @param dataMap template data information
     * @throws Exception exceptions
     */
    public static XMLSlideShow createPtt(File file, Map<String, PptModel> dataMap) throws Exception {
        /*validate file exist*/
        if (file == null) {
            throw new FileNotFoundException("文件不存在");
        }

        try (FileInputStream fileInputStream = new FileInputStream(file)) {
            XMLSlideShow xmlSlideShow = new XMLSlideShow(fileInputStream);
            XSLFSlide[] xmlSlideShowSlides = xmlSlideShow.getSlides();
            for (XSLFSlide xslfSlide : xmlSlideShowSlides) {
                for (XSLFShape shape : xslfSlide.getShapes()) {
                    renderShape(xmlSlideShow, xslfSlide, shape, dataMap);
                }
            }
            return xmlSlideShow;
        }

    }

    /**
     * @desc inner method
     * @param dataMap template data information
     */
    private static void renderShape(XMLSlideShow xmlSlideShow, XSLFSlide xslfShapes, XSLFShape shape, Map<String, PptModel> dataMap) throws Exception {
        if (shape instanceof XSLFGroupShape) {
            XSLFShape[] shapes = ((XSLFGroupShape) shape).getShapes();
            for (XSLFShape xslfShape : shapes) {
                renderShape(xmlSlideShow, xslfShapes, xslfShape, dataMap);
            }
        } else if (shape instanceof XSLFTextShape) {
            XSLFTextShape txShape = (XSLFTextShape) shape;
            renderShapeContent(xmlSlideShow, xslfShapes, txShape, dataMap);
        } else {
            System.out.println(shape.getClass());
        }
    }

    /**
     * @desc inner methods
     * @param dataMap template data information
     */
    private static void renderShapeContent(XMLSlideShow xmlSlideShow, XSLFSlide xslfShape, XSLFTextShape shape, Map<String, PptModel> dataMap) throws Exception {
        Set<String> dataKeys = getDataKeys(shape);
        for (String key : dataKeys) {
            key = key.replaceAll("#", "");
            PptModel pptModel = dataMap.get(key);
            int dataType = pptModel.getDataType();
            switch (dataType) {
                case 1 :
                    List<XSLFTextParagraph> paragraphList = shape.getTextParagraphs();
                    for (XSLFTextParagraph p : paragraphList) {
                        for (XSLFTextRun textRun : p.getTextRuns()) {
                            String value = (String) (pptModel.getDataConent() == null ? pptModel.getDefaultContent() : pptModel.getDataConent());
                            String text = textRun.getText().replaceAll(key, value).replaceAll("#", "");
                            textRun.setText(text);
                        }
                    }
                    break;
                case 2 :
                    break;
                case 3 :
                    String value = (String) (pptModel.getDataConent() == null ? pptModel.getDefaultContent() : pptModel.getDataConent());
                    if (value == null) {
                        throw new FileNotFoundException("文件不存在");
                    }

                    File image = new File(pptModel.getDataConent().toString());
                    byte[] pictureData = IOUtils.toByteArray(new FileInputStream(image));
                    int idx = xmlSlideShow.addPicture(pictureData, XSLFPictureData.PICTURE_TYPE_JPEG);
                    XSLFPictureShape pic = xslfShape.createPicture(idx);
                    pic.setAnchor(shape.getAnchor());
                    xslfShape.removeShape(shape);
                    break;
                case 4 :
                    File file = (File) (pptModel.getDataConent() == null ? pptModel.getDefaultContent() : pptModel.getDataConent());
                    if (file == null) {
                        throw new FileNotFoundException("文件不存在");
                    }
                    byte[] pictureFileData = IOUtils.toByteArray(new FileInputStream(file));
                    int idxFile = xmlSlideShow.addPicture(pictureFileData, XSLFPictureData.PICTURE_TYPE_JPEG);
                    XSLFPictureShape picFile = xslfShape.createPicture(idxFile);
                    picFile.setAnchor(shape.getAnchor());
                    xslfShape.removeShape(shape);
                    break;
                case 5 :
                    String base64 = (String) (pptModel.getDataConent() == null ? pptModel.getDefaultContent() : pptModel.getDataConent());
                    if (base64 == null) {
                        throw new FileNotFoundException("base64数据值存在问题");
                    }
                    byte[] decodeBuffer = new BASE64Decoder().decodeBuffer(base64);
                    int idxBaseFile = xmlSlideShow.addPicture(decodeBuffer, XSLFPictureData.PICTURE_TYPE_JPEG);
                    XSLFPictureShape picBaseFile = xslfShape.createPicture(idxBaseFile);
                    picBaseFile.setAnchor(shape.getAnchor());
                    xslfShape.removeShape(shape);
                    break;
                case 6 :
                    String tableData = (String) (pptModel.getDataConent() == null ? pptModel.getDefaultContent() : pptModel.getDataConent());
                    if (tableData == null) {
                        break;
                    }
                    XSLFTable table = xslfShape.createTable();
                    String tableVlaue = pptModel.getDataConent().toString();
                    String[] a = tableVlaue.split(";");
                    for (int m = 0; m < a.length; m++) {
                        XSLFTableRow firstRow = table.addRow();
                        String[] aa = a[m].split(",");
                        for (int n = 0; n < aa.length; n++) {
                            XSLFTableCell firstCell = firstRow.addCell();
                            firstCell.setBorderBottomColor(new Color(10, 100, 120));
                            firstCell.setBorderRightColor(new Color(10, 100, 120));
                            firstCell.setBorderLeftColor(new Color(10, 100, 120));
                            firstCell.setBorderTopColor(new Color(10, 100, 120));
                            firstCell.setText(aa[n]);
                            firstCell.setBorderLeft(10);
                            firstCell.setBorderRight(10);
                            firstCell.setBorderTop(10);
                            firstCell.setBorderBottom(10);
                        }
                    }
                    table.setAnchor(shape.getAnchor());
                    System.out.println(table.getRows().get(0).getCells().get(0).getText());
                    xslfShape.removeShape(shape);
                    break;
                default:
                    break;
            }
        }
    }

    private static Set<String> getDataKeys(XSLFTextShape shape) {
        String regex = "#[^#]*#";
        Set<String> dataMapKeys = new LinkedHashSet<>();
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(shape.getText());
        while (matcher.find()) {
            System.out.println(matcher.group());
            dataMapKeys.add(matcher.group());
        }
        return dataMapKeys;
    }
}
