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
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
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
    Button write;
    EditText text;
    EditText row;
    EditText col;
    int rowNum;
    int colNum;
    String  filePath2003=Environment.getExternalStorageDirectory().getPath()+"/Test/test2003.xls";
    String  filePath2007=Environment.getExternalStorageDirectory().getPath()+"/Test/test2007.xlsx";
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        searh=(Button)findViewById(R.id.sear);
        write=(Button)findViewById(R.id.write);
        text=(EditText) findViewById(R.id.txt);
        row=(EditText)findViewById(R.id.hang);
        col=(EditText)findViewById(R.id.lie);
        file=new File(Environment.getExternalStorageDirectory().getPath()+"/Test");
        if(!file.exists()){
            file.mkdir();
        }
        searh.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                        text.setText(readExcel2003(filePath2003,Integer.valueOf(row.getText().toString()),Integer.valueOf(col.getText().toString())));
                     }
                });
        write.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                writeExcel2003(filePath2003,Integer.valueOf(row.getText().toString()),Integer.valueOf(col.getText().toString()),text.getText().toString());
            }
        });



    }

    public  void writeExcel2003(String filePath,int rowNum ,int colNum,String content){
        File file=new File(filePath);
        if(!file.exists()){
            HSSFWorkbook workbook=new HSSFWorkbook();
            HSSFSheet sheet=workbook.createSheet("Sheet1");
            HSSFRow row;
            for(int i=0;i<3000;i++){
              sheet.createRow(i);
            }
            row=sheet.getRow(rowNum);
            HSSFCell cell=row.createCell(colNum);
            cell.setCellValue(content);
            OutputStream outputStream=null;
            try {
                outputStream=new FileOutputStream(file);
                workbook.write(outputStream);
                outputStream.flush();
                outputStream.close();
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
             catch (IOException e) {
                e.printStackTrace();
            }

        }else{
            try {
                InputStream inputStream=new FileInputStream(filePath);
                HSSFWorkbook workbook=new HSSFWorkbook(inputStream);
                HSSFSheet sheet=workbook.getSheet("Sheet1");
                HSSFRow row=sheet.getRow(rowNum);
                HSSFCell cell=row.createCell(colNum);
                cell.setCellValue(content);
                OutputStream outputStream=null;
                outputStream=new FileOutputStream(file);
                workbook.write(outputStream);
                outputStream.flush();
                outputStream.close();
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            catch (IOException e) {
                e.printStackTrace();
            }

        }


    }
    public  String readExcel2003(String filePath,int rowNum ,int colNum){
        OutputStream os = null;
        InputStream is = null;
        HSSFWorkbook workbook = null;
        try {
            is = new FileInputStream(filePath);
        } catch (FileNotFoundException e) {
            Toast.makeText(MainActivity.this, "找不到文件", Toast.LENGTH_SHORT).show();
        }
        try {
            workbook = new HSSFWorkbook(is);
        } catch (IOException e) {
            Toast.makeText(MainActivity.this, "workbook失败", Toast.LENGTH_SHORT).show();
        }
       // rowNum = Integer.valueOf(row.getText().toString());
       // colNum = Integer.valueOf(col.getText().toString());
        HSSFSheet sheet = workbook.getSheetAt(0);
        try {
            HSSFRow row = sheet.getRow(rowNum);
            HSSFCell cell = row.getCell(colNum);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);//先全部转换成String,因为POI分 cell.getStringCellValue()和get其他类型，如果格式不对就要报错推出
            return  cell.getStringCellValue();
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
        return  null;
    }





}



