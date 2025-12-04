package com.example.poi_demo;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ConfigureBuilder;
import com.deepoove.poi.plugin.table.LoopRowTableRenderPolicy;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;

public class PoiTest {
    public static void main(String[] args) throws IOException {
        //1.数据处理
        List<Person> personList = Arrays.asList(
                new Person("张三",22,"男","NBKJ","无"),
                new Person("李四",23,"女","NBKJ","无")
        );
        // 2. 创建表格渲染策略
        LoopRowTableRenderPolicy policy = new LoopRowTableRenderPolicy();
        // 3. 配置绑定 - 注意：绑定的是模板中的占位符名称
        ConfigureBuilder builder = Configure.builder();
        builder.buildGramer("${", "}");
        Configure configure = builder.build();
        configure.customPolicy("personList",policy);
        // 4. 指定模板文件路径
        String templatePath = "src/main/resources/templateDoc/poi测试.docx";
        XWPFTemplate template = XWPFTemplate.compile(templatePath ,configure).render(
                new HashMap<String, Object>() {{
                    put("personList", personList);
                }}
        );
        // 5. 输出结果
        FileOutputStream out = new FileOutputStream("src/main/resources/outputDoc/output.docx");
        template.write(out);
        template.close();
        out.close();
    }
}
