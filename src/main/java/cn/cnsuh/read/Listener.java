package cn.cnsuh.read;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

/**
 * @author Magic
 * @create 2019-11-30 14:24
 */
public class Listener extends AnalysisEventListener {

    /**
     * 读取每行数据都会调用invoke
     * @param o 读取到的行数据
     * @param context
     */
    public void invoke(Object o, AnalysisContext context) {
        int row = context.getCurrentRowNum();

        System.out.println("读取的行号："+row+"---读取的行数据"+o);
    }

    /**
     *
     * @param context
     */
    public void doAfterAllAnalysed(AnalysisContext context) {
        System.out.println("excel读取完成");
    }
}
