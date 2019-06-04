package com.sakura.rain;

import com.sakura.rain.model.PptModel;
import com.sakura.rain.utils.PptUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import sun.misc.BASE64Encoder;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * @ClassName Main
 * @Description 程序运行主函数
 * @Author lcz
 * @Date 2019/05/05 11:46
 */
public class Main {
    public static void main(String[] args) throws IOException {
        /*
        * 基础数据准备
        * */
        Map<String, PptModel> dataMap = new HashMap<>();

        PptModel<String> textModel = new PptModel<>();
        textModel.setDataType(PptUtils.TEXT_DATA);
        textModel.setDataConent("Dream on plum park!");
        dataMap.put("text_data", textModel);

        PptModel<String> pathModel = new PptModel<>();
        pathModel.setDataType(PptUtils.PICTURE_PATH_DATA);
        pathModel.setDataConent(Main.class.getResource("/image/sakura.jpg").getFile());
        dataMap.put("image_path", pathModel);

        PptModel<File> fileModel = new PptModel<>();
        fileModel.setDataType(PptUtils.PICTRUE_FILE_DATA);
        fileModel.setDataConent(new File(Main.class.getResource("/image/sakura.jpg").getFile()));
        dataMap.put("image_file", fileModel);

        PptModel<String> base64Model = new PptModel<>();
        base64Model.setDataType(PptUtils.PICTRUE_BASE64_DATA);
        File sakuraFile = new File(Main.class.getResource("/image/sakura.jpg").getFile());
        FileInputStream fileInputStream = new FileInputStream(sakuraFile);
        byte[] bytes = new byte[fileInputStream.available()];
        fileInputStream.read(bytes);
        base64Model.setDataConent(new BASE64Encoder().encode(bytes));
        dataMap.put("image_base", base64Model);

        /*暂不可用，等待优化*/
        PptModel<String> tableModel = new PptModel<>();
        tableModel.setDataType(PptUtils.TABLE_DATA);
        tableModel.setDataConent("name,password,address;sakura,123456,nanjing;plumpark,123456,bozhou");
        dataMap.put("table_data", tableModel);

        String modelPath = Main.class.getResource("/model/template_model.pptx").getFile();
        File file = new File(modelPath);
        try (OutputStream outputStream = new FileOutputStream(new File("result.pptx"))) {
            XMLSlideShow ptt = PptUtils.createPtt(file, dataMap);
            ptt.write(outputStream);
        } catch (Exception e) {
            System.out.println("程序运行出现问题...");
        }

    }
}
