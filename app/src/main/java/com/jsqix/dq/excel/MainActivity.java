package com.jsqix.dq.excel;

import android.os.Bundle;
import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.view.View;
import android.widget.Button;

import com.jsqix.dq.excel.bean.OrderBean;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;
import java.util.UUID;

import jxl.Workbook;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.CellFormat;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class MainActivity extends AppCompatActivity {
    private Button button;
    private List<OrderBean> list = new ArrayList<>();

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        button = (Button) findViewById(R.id.button);
        button.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                try {
                    createExcel();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        });
        initData();
    }

    private void initData() {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        for (int i = 0; i < 10; i++) {
            OrderBean bean = new OrderBean();
            int random = new Random().nextInt(10) + 1;
            bean.setId(UUID.randomUUID().toString());
            bean.setAmount(0.1 * random + 0.5 + "");
            bean.setDate(sdf.format(new Date().getTime() + random));
            bean.setStatus("成功");
            bean.setPhone("13512345678");
            list.add(bean);
        }
    }

    private void createExcel() throws Exception {
        String dir = Environment.getExternalStorageDirectory().getAbsolutePath();
        String fileName = "Db_" + System.currentTimeMillis();
        File file = new File(dir, fileName + ".xls");
        WritableWorkbook wwb = Workbook.createWorkbook(file);
        WritableSheet sheet = wwb.createSheet("订单", 0);
        String[] title = {"订单号", "手机号", "金额", "状态", "时间"};
        Label label;
        for (int i = 0; i < title.length; i++) {
            // Label(x,y,z) 代表单元格的第x+1列，第y+1行, 内容z
            // 在Label对象的子对象中指明单元格的位置和内容
            label = new Label(i, 0, title[i], getHeader());
            // 将定义好的单元格添加到工作表中
            sheet.addCell(label);
        }
        CellFormat format = getCenterFormat();
        for (int i = 0; i < list.size(); i++) {
            OrderBean order = list.get(i);

            Label orderNum = new Label(0, i + 1, order.getId(), format);
            Label phone = new Label(1, i + 1, order.getPhone(), format);
            Label amount = new Label(2, i + 1, order.getAmount(), format);
            Label status = new Label(3, i + 1, order.getStatus(), format);
            Label time = new Label(4, i + 1, order.getDate(), format);

            sheet.addCell(orderNum);
            sheet.addCell(phone);
            sheet.addCell(amount);
            sheet.addCell(status);
            sheet.addCell(time);
        }
        wwb.write();
        wwb.close();
    }

    private WritableCellFormat getHeader() {
        WritableFont font = new WritableFont(WritableFont.TIMES, 10,
                WritableFont.BOLD);// 定义字体
        try {
            font.setColour(Colour.BLUE);// 蓝色字体
        } catch (WriteException e1) {
            e1.printStackTrace();
        }
        WritableCellFormat format = new WritableCellFormat(font);
        try {
            format.setAlignment(jxl.format.Alignment.CENTRE);// 左右居中
            format.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);// 上下居中
            format.setBorder(Border.ALL, BorderLineStyle.THIN,
                    Colour.BLACK);// 黑色边框
            format.setBackground(Colour.YELLOW);// 黄色背景
        } catch (WriteException e) {
            e.printStackTrace();
        }
        return format;
    }

    //单元格居中
    private CellFormat getCenterFormat() {
        WritableCellFormat format = new WritableCellFormat();
        try {
            format.setAlignment(jxl.format.Alignment.CENTRE);// 左右居中
            format.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);// 上下居中
        } catch (WriteException e1) {
            e1.printStackTrace();
        }
        return format;
    }
}
