package com.example.poi_demo;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ConfigureBuilder;
import com.deepoove.poi.plugin.table.LoopRowTableRenderPolicy;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;

@SpringBootTest
class PoiDemoApplicationTests {

    @Test
    void contextLoads() {
    }

    @Test
    void testRenderTemplatewith() throws IOException {
        List<Person> personList = Arrays.asList(
                new Person("张三", 22, "男", "NBKJ", "无"),
                new Person("李四", 23, "女", "NBKJ", "无")
        );
        ConfigureBuilder builder = Configure.builder();
        builder.buildGramer("${", "}");
        String templatePath = "src/main/resources/templateDoc/poi测试.docx";
        LoopRowTableRenderPolicy policy = new LoopRowTableRenderPolicy();
        Configure configure = builder.build();
        configure.customPolicy("personList", policy);
        XWPFTemplate template = XWPFTemplate.compile(templatePath, configure).render(
                new HashMap<String, Object>() {{
                    put("personList", personList);
                }}
        );
        FileOutputStream out = new FileOutputStream("src/main/resources/outputDoc/output.docx");
        template.write(out);
        template.close();
        out.close();
    }

}
