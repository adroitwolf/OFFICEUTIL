package com.adroitwolf;

import com.adroitwolf.entity.Worker;
import com.adroitwolf.enums.ExcelWorkEnum;
import com.adroitwolf.util.ExcelUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.JUnit4;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.ArrayList;
import java.util.List;

/**
 * <pre>ServiceTest</pre>
 *
 * @author adroitwolf 2021年03月16日 14:51
 */
@SpringBootTest
@RunWith(JUnit4.class)
public class ServiceTest {


    @Test
    public void createExcel(){
        List<Worker> workerList = new ArrayList<>();

        workerList.add(new Worker(1,"张三",8,"青岛"));
        workerList.add(new Worker(2,"李四",8,"青岛"));
        workerList.add(new Worker(3,"张三",8,"青岛"));
        workerList.add(new Worker(4,"李四",8,"青岛"));
        workerList.add(new Worker(5,"李四",8,"德州"));
        workerList.add(new Worker(6,"张三",8,"德州"));
        workerList.add(new Worker(7,"张三",8,"德州"));
        workerList.add(new Worker(8,"李四",8,"德州"));
        workerList.add(new Worker(9,"张三",8,"德州"));
        workerList.add(new Worker(10,"张三",8,"德州"));

        System.out.println(workerList.size());
        Workbook workbook = ExcelUtil.exportExcel(workerList, Worker.class, null, null);


        ExcelUtil.export2File("d://code/诺亚公司"+  ExcelWorkEnum.SXSSF.getSuffix(),workbook);
    }
}
