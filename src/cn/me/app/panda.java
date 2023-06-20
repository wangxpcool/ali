package cn.me.app;

import cn.hutool.core.io.FileUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class panda {

    public static void main(String[] args) {
        System.out.println(11);

        File file = FileUtil.file("C:\\Users\\tykj\\Desktop\\t\\熊猫有品4.xlsx");
        ExcelReader reader = ExcelUtil.getReader(file,0);
        String fileName = file.getName().replaceAll(".xlsx","");
        List<Map<String,Object>> read = reader.readAll();
        System.out.println("----------");
        Map<String,Map<String,Object>> result = new HashMap<>();
        for (int i=1;i<read.size() ;i++) {
            Map<String,Object> map = read.get(i);

            String comp =(String) map.get("账户名称");
            String key = fileName+"-"+comp;

            Double d1= convert( map.get("账户总支出(¥)"));
            Double d2= convert( map.get("账户现金支出(¥)"));
            Double d3= convert(map.get("账户赠款支出(¥)"));
            Double d4= convert(map.get("账户总存入(¥)"));
            Double d5= convert(map.get("账户总转出(¥)"));

            if (result.containsKey(key)){

                Map<String ,Object > m = result.get(key);
                Double old1 =(Double) m.get("账户总支出(¥)");
                Double newd1 = old1+d1;
                m.put("账户总支出(¥)",newd1);
                Double old2 =(Double) m.get("账户现金支出(¥)");
                Double newd2 = old2+d2;
                m.put("账户现金支出(¥)",newd2);

                Double old3 =(Double) m.get("账户赠款支出(¥)");
                Double newd3 = old3+d3;
                m.put("账户赠款支出(¥)",newd3);

                Double oldcz =(Double) m.get("充值");
                oldcz = oldcz + (d4-d5);
                m.put("充值",oldcz);

                if (map.get("日期").equals("2023-05-31")){
                    m.put("余额",map.get("总余额(¥)"));
                    m.put("非赠款余额(¥)",map.get("非赠款余额(¥)"));
                    m.put("赠款余额(¥)",map.get("赠款余额(¥)"));
                }
                m.put("name",key);
            }else{
                Map<String ,Object > m = new HashMap<>();
                m.put("账户总支出(¥)",d1);
                m.put("账户现金支出(¥)",d2);
                m.put("账户赠款支出(¥)",d3);
                m.put("充值",d4-d5);
                if (map.get("日期").equals("2023-05-31")){
                    m.put("非赠款余额(¥)",map.get("非赠款余额(¥)"));
                    m.put("赠款余额(¥)",map.get("赠款余额(¥)"));
                    m.put("余额",map.get("总余额(¥)"));
                }
                m.put("name",key);
                result.put(key,m);
            }

        }


        List<Map<String,Object>> list = new ArrayList<>();
        result.keySet().forEach(r->{
            list.add(result.get(r));
            System.out.println(r+":"+result.get(r));

        });
        System.out.println("---------");



        reader.close();
        ExcelWriter writer = ExcelUtil.getWriter(file);
        writer.setSheet("final").merge(list.size() - 1, "统计").write(list, true);
        writer.close();



    }

    static Double convert (Object o){
        Double r =null;
        if (o instanceof Long){
            r=((Long) o).doubleValue();
        }else if (o instanceof Double){
            r=(Double)o;
        }
        return r;
    }

}
