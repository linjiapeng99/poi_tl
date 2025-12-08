package com.example.poi_demo;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.PictureType;
import com.deepoove.poi.data.Pictures;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class TestRecordRenderer {
    
    public static void main(String[] args) throws Exception {
        // 1. 准备数据
        Map<String, Object> data = new HashMap<>();
        
        // 文档基本信息
        data.put("recordNo", "Lab-20251201003-DCR");
        data.put("version", "V1.0");
        data.put("totalPages", "共15页");
        
        // 2. 准备测试用例记录列表
        List<Map<String, Object>> testCaseRecords = new ArrayList<>();
        
        // 测试用例1
        Map<String, Object> record1 = new HashMap<>();
        record1.put("testCaseCode", "20240900001");
        record1.put("testCaseName", "功能测试");
        record1.put("testCaseReason", "GB/T 25000.51-2016《系统与软件工程...》");
        record1.put("testCaseCondition", "系统可以正常访问");
        record1.put("testCaseStep", "步骤1、以管理员级别帐户登录系统；\n步骤2、执行相关操作。");
        record1.put("testCaseExpectedReslt", "系统正常响应");
        record1.put("testCaseReslt", "通过");
        record1.put("testCaseDescription", "无");
        record1.put("testCaseRank", "高");
        record1.put("testCaseDesigner", "设计员");
        record1.put("testCaseInspector", "检测员");
        record1.put("testCaseTime", "2025年12月01日");
        // 图片（如果有）
        FileInputStream fis = new FileInputStream("src/main/resources/image/poi测试图片.png");
        PictureRenderData picture = Pictures.ofStream(fis, PictureType.PNG)
                .size(400, 200)
                .create();
        record1.put("testCaseImages", picture);
        testCaseRecords.add(record1);
        
        // 测试用例2
        Map<String, Object> record2 = new HashMap<>();
        record2.put("testCaseCode", "20240900002");
        record2.put("testCaseName", "性能测试");
        record2.put("testCaseReason", "GB/T 25000.51-2016");
        record2.put("testCaseCondition", "系统可以正常访问");
        record2.put("testCaseStep", "步骤1：压力测试\n步骤2：记录响应时间");
        record2.put("testCaseExpectedReslt", "响应时间<3秒");
        record2.put("testCaseReslt", "通过");
        record2.put("testCaseDescription", "无");
        record2.put("testCaseRank", "中");
        record2.put("testCaseDesigner", "设计员");
        record2.put("testCaseInspector", "检测员");
        record2.put("testCaseTime", "2025年12月01日");
        // 无图片可以不设置或设置为null
        testCaseRecords.add(record2);
        
        // ... 继续添加其他测试用例（record3-record9）
        
        // 测试用例9
        Map<String, Object> record9 = new HashMap<>();
        record9.put("testCaseCode", "20240900009");
        record9.put("testCaseName", "性能效率");
        record9.put("testCaseReason", "GB/T 25000.51-2016《系统与软件工程...》\n\n《需求规格说明书》");
        record9.put("testCaseCondition", "系统可以正常访问");
        record9.put("testCaseStep", "步骤1、以管理员级别帐户登录系统；\n步骤2、设置备份周期为每小时。");
        record9.put("testCaseExpectedReslt", "系统每小时备份一次业务数据");
        record9.put("testCaseReslt", "T");
        record9.put("testCaseDescription", "无");
        record9.put("testCaseRank", "高");
        record9.put("testCaseDesigner", "检测员");
        record9.put("testCaseInspector", "检测员");
        record9.put("testCaseTime", "2025年12月01日");
        record9.put("testCaseImages", Pictures.ofLocal("media/image1.png").size(400, 200).create());
        testCaseRecords.add(record9);
        
        data.put("testCaseRecords", testCaseRecords);
        
        // 3. 配置模板引擎
        Configure config = Configure.builder()
                .buildGramer("${", "}")  // 使用 ${} 作为标签前后缀
                .build();
        
        // 4. 渲染模板
        XWPFTemplate template = XWPFTemplate
                .compile("src/main/resources/templateDoc/检测用例记录表格.docx", config)
                .render(data);
        
        template.writeAndClose(new FileOutputStream("src/main/resources/outputDoc/输出_检测记录.docx"));
        
        System.out.println("文档生成成功！");
    }
}