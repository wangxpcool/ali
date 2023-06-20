package cn.me.app;

import cn.hutool.core.io.FileUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;

import java.io.File;
import java.math.BigDecimal;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class apl {

    public static void main(String[] args) {
        System.out.println(11);

        File file = FileUtil.file("C:\\Users\\tykj\\Desktop\\t\\晋拓快手2 2023-05.xlsx");
        ExcelReader reader = ExcelUtil.getReader(file,0);
        String fileName = file.getName();
        List<Map<String,Object>> read = reader.readAll();

//        for (Map<String,Object> map : read) {
//            System.out.println(map);
//        }

        System.out.println("----------");


        Map<String,Map<String,Object>> result = new HashMap<>();
        for (int i=1;i<read.size() ;i++) {
            Map<String,Object> map = read.get(i);

//            System.out.println(map);
            String pro =(String) map.get("产品名");
            String comp =(String) map.get("企业名称");
            String key = pro+"-"+comp;

            Double d1= convert( map.get("现金花费"));
            Double d2= convert( map.get("前返花费"));
            Double d3= convert(map.get("后返花费"));
            Double d4= convert(map.get("信用花费"));

            Double c1 = convert( map.get("现金转入"));
            Double c2 = convert( map.get("前返转入"));
            Double c3 = convert( map.get("后返转入"));
            Double c4 = convert( map.get("信用转入"));

            Double c5 = convert( map.get("现金转出"));
            Double c6 = convert( map.get("前返转出"));
            Double c7 = convert( map.get("后返转出"));
            Double c8 = convert( map.get("信用转出"));
            if (result.containsKey(key)){
                if ("康佳KONKA泽美专卖店-义乌市泽美贸易有限公司".equals(key)){
                    System.out.println(d1 + "- "+d2 + "- "+d3 + "- "+d4 + "- ");
                }
                Map<String ,Object > m = result.get(key);
                Double old =(Double) m.get("消耗");
                Double x = d1+d2+d3+d4;
                Double newx = old+x;
                m.put("消耗",BigDecimal.valueOf(newx).setScale(3,BigDecimal.ROUND_UP ).doubleValue());

                Double oldcz =(Double) m.get("充值");
                Double cz = c1+c2+c3+c4-c5-c6-c7-c8;
                Double newcz = oldcz+cz;
                m.put("充值",BigDecimal.valueOf(newcz).setScale(3,BigDecimal.ROUND_UP ).doubleValue());

                result.put(key,m);
            }else{
                Map<String ,Object > m = new HashMap<>();
                Double x = d1+d2+d3+d4;

                m.put("消耗",BigDecimal.valueOf(x).setScale(3,BigDecimal.ROUND_UP ).doubleValue());
                Double cz = c1+c2+c3+c4-c5-c6-c7-c8;
                m.put("充值", BigDecimal.valueOf(cz).setScale(3,BigDecimal.ROUND_UP ).doubleValue());

                result.put(key,m);
            }

        }


        result.keySet().forEach(r->{
            Map<String ,Object> m = result.get(r);
            if ((m.get("消耗")!=null && (Double) m.get("消耗")!=0.0 )
            ||
                    ((m.get("充值")!=null && (Double) m.get("充值")!=0.0))
            ){
                System.out.println(r+":"+result.get(r));
            }
        });
        System.out.println("---------");


        reader.close();



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
