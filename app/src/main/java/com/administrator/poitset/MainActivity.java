package com.administrator.poitset;

import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.TextView;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

public class MainActivity extends AppCompatActivity {
    File file;
    Button searh;
    TextView text;
    EditText row;
    EditText col;
    int rowNum;
    int colNum;
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        searh=(Button)findViewById(R.id.sear);
        text=(TextView)findViewById(R.id.txt);
        row=(EditText)findViewById(R.id.hang);
        col=(EditText)findViewById(R.id.lie);
        file=new File(Environment.getExternalStorageDirectory().getPath()+"/Test");
        if(!file.exists()){
            file.mkdir();
        }
        file=new File(Environment.getExternalStorageDirectory().getPath()+"/Test/test.xls");
        if(!file.exists()){
            try {
                file.createNewFile();
            }
            catch (Exception e)
            {

            }
        }
        searh.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                OutputStream os = null;
                InputStream is = null;
                HSSFWorkbook workbook = null;
                try {
                    is = new FileInputStream(file);
                } catch (FileNotFoundException e) {
                    Toast.makeText(MainActivity.this, "找不到文件", Toast.LENGTH_SHORT).show();
                }
               try {
                    workbook = new HSSFWorkbook(is);
                } catch (IOException e) {
                    Toast.makeText(MainActivity.this, "workbook失败", Toast.LENGTH_SHORT).show();
                }
                rowNum = Integer.valueOf(row.getText().toString());
                colNum = Integer.valueOf(col.getText().toString());
                HSSFSheet sheet = workbook.getSheetAt(0);
                try {
                HSSFRow row = sheet.getRow(rowNum);
                HSSFCell cell = row.getCell(colNum);
                cell.setCellType(HSSFCell.CELL_TYPE_STRING);//先全部转换成String,因为POI分 cell.getStringCellValue()和get其他类型，如果格式不对就要报错推出
                text.setText( cell.getStringCellValue());
                     }
                catch (NullPointerException e) {
                    //如果单元格为空，就跳过此次循环
                     Toast.makeText(MainActivity.this, "内容为空", Toast.LENGTH_SHORT).show();
                }
                 try {

                    is.close();
                } catch (IOException e) {
                    Toast.makeText(MainActivity.this, "关闭失败", Toast.LENGTH_SHORT).show();
                }
            }
            });





    }
}
