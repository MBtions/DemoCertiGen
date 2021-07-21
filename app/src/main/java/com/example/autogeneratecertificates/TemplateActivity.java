package com.example.autogeneratecertificates;

import android.content.res.AssetManager;
import android.net.Uri;
import android.os.Bundle;
import android.util.Log;
import android.widget.TextView;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.List;

public class TemplateActivity extends AppCompatActivity {

    TextView test;
    String[] a=new String[30];
    String[] b=new String[30];
    String[] c=new String[30];
    String[] d=new String[30];
    String[] e=new String[30];
    String[] name=new String[30];
    String[] course=new String[30];
    String[] college=new String[30];
    String[] position=new String[30];
    String[] society=new String[30];

    double row_num=0.0;
    String rows="0.0";

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_template);
        test = (TextView) findViewById(R.id.test);
        rows="5";
        row_num=5;
        try {
            rows = getIntent().getStringExtra("entries");
            row_num = Double.parseDouble(rows);
        }catch (NullPointerException e){
            e.printStackTrace();
        }
        readExcelData();
    }

    public void readExcelData() {
            try {
                String path_xlsx = getIntent().getStringExtra("path");
                InputStream inputfile = getContentResolver().openInputStream(Uri.parse(path_xlsx));

                XSSFWorkbook workbook = new XSSFWorkbook(inputfile);
                XSSFSheet sheet = workbook.getSheetAt(0);
                Iterator<Row> rowIterator = sheet.iterator();
                int count = 0;
                String temp;

                while (rowIterator.hasNext() && count < row_num+1) {
                    Row row = rowIterator.next();
                    Iterator<Cell> cx = row.cellIterator();
                    temp=cx.next().getStringCellValue();
                    a[count]=temp;

                    temp=cx.next().getStringCellValue();
                    b[count]=temp;

                    temp=cx.next().getStringCellValue();
                    c[count]=temp;

                    temp=cx.next().getStringCellValue();
                    d[count]=temp;

                    temp=cx.next().getStringCellValue();
                    e[count]=temp;

                    count++;
                }
                inputfile.close();
                matchColumn();
            } catch (FileNotFoundException e) {
                test.setText("FilenotFound");
                e.printStackTrace();
            } catch (IOException e) {
                test.setText("IOException");
                e.printStackTrace();
            }
    }

    public void matchColumn() {
        int i;
        switch (a[0].toLowerCase()) {
            case "name":
                for (i = 1; i <= row_num; i++) {
                    name[i-1] = a[i];
                }
                break;
            case "college":
                for (i=1; i<=row_num; i++){
                    college[i-1]=a[i];
                }
                break;
            case "course":
                for (i=1; i<=row_num; i++){
                    course[i-1]=a[i];
                }
                break;
            case "society":
                for (i=1; i<=row_num; i++){
                    society[i-1]=a[i];
                }
                break;
            case "position":
                for (i=1; i<=row_num; i++){
                    position[i-1]=a[i];
                }
                break;
        }

        switch (b[0].toLowerCase()) {
            case "name":
                for (i = 1; i <= row_num; i++) {
                    name[i-1] = b[i];
                }
                break;
            case "college":
                for (i=1; i<=row_num; i++){
                    college[i-1]=b[i];
                }
                break;
            case "course":
                for (i=1; i<=row_num; i++){
                    course[i-1]=b[i];
                }
                break;
            case "society":
                for (i=1; i<=row_num; i++){
                    society[i-1]=b[i];
                }
                break;
            case "position":
                for (i=1; i<=row_num; i++){
                    position[i-1]=b[i];
                }
                break;
        }

        switch (c[0].toLowerCase()) {
            case "name":
                for (i = 1; i <= row_num; i++) {
                    name[i-1] = c[i];
                }
                break;
            case "college":
                for (i=1; i<=row_num; i++){
                    college[i-1]=c[i];
                }
                break;
            case "course":
                for (i=1; i<=row_num; i++){
                    course[i-1]=c[i];
                }
                break;
            case "society":
                for (i=1; i<=row_num; i++){
                    society[i-1]=c[i];
                }
                break;
            case "position":
                for (i=1; i<=row_num; i++){
                    position[i-1]=c[i];
                }
                break;
        }

        switch (d[0].toLowerCase()) {
            case "name":
                for (i = 1; i <= row_num; i++) {
                    name[i-1] = d[i];
                }
                break;
            case "college":
                for (i=1; i<=row_num; i++){
                    college[i-1]=d[i];
                }
                break;
            case "course":
                for (i=1; i<=row_num; i++){
                    course[i-1]=d[i];
                }
                break;
            case "society":
                for (i=1; i<=row_num; i++){
                    society[i-1]=d[i];
                }
                break;
            case "position":
                for (i=1; i<=row_num; i++){
                    position[i-1]=d[i];
                }
                break;
        }

        switch (e[0].toLowerCase()) {
            case "name":
                for (i = 1; i <= row_num; i++) {
                    name[i-1] = e[i];
                }
                break;
            case "college":
                for (i=1; i<=row_num; i++){
                    college[i-1]=e[i];
                }
                break;
            case "course":
                for (i=1; i<=row_num; i++){
                    course[i-1]=e[i];
                }
                break;
            case "society":
                for (i=1; i<=row_num; i++){
                    society[i-1]=e[i];
                }
                break;
            case "position":
                for (i=1; i<=row_num; i++){
                    position[i-1]=e[i];
                }
                break;
        }
        //searchSlideAsync();
    }
    private void searchSlideAsync() {
        new Thread(this::searchSlide).start();
    }

    private void searchSlide() {
        try {
            AssetManager assManager = getAssets();
            InputStream is = assManager.open("templates.pptx");
            Log.e("powerpoint","1"+getFilesDir());
            OutputStream os = null;
            File outFile = new File(getFilesDir(), "templates.pptx");
            outFile.mkdirs();
            os = new FileOutputStream(outFile);
            copyFile(is, os);
            Log.e("powerpoint","2");
            XMLSlideShow ppt = new XMLSlideShow(is);
            Log.e("powerpoint","3");
            XSLFSlide mainslide = ppt.getSlides().get(0);

            for (int i=0; i<row_num; i++){
                XSLFSlide newSlide = ppt.createSlide();
                XSLFSlideLayout src_sl = mainslide.getSlideLayout();
                XSLFSlideMaster src_sm = mainslide.getSlideMaster();
                XSLFSlideLayout new_sl = newSlide.getSlideLayout();
                XSLFSlideMaster new_sm = newSlide.getSlideMaster();

                // copy source layout to the new layout
                new_sl.importContent(src_sl);
                // copy source master to the new master
                new_sm.importContent(src_sm);
                newSlide.importContent(mainslide);
                /*
                XSLFSlideLayout layout = mainslide.getSlideLayout();
                XSLFSlide newSlide = ppt.createSlide(layout);
                newSlide.importContent(mainslide);*/
                ppt.write(os);
            }
            int count=0;

            List<XSLFSlide> slides=ppt.getSlides();
            for (XSLFSlide slide : slides) {
                List<XSLFShape> shapes = slide.getShapes();
                for (XSLFShape shape : shapes) {
                    if (count>0 && shape instanceof XSLFTextBox) {
                        String text = ((XSLFTextBox) shape).getText();
                        text = text.replace("<Name>", name[count-1] );
                        text = text.replace("<Course>", course[count-1] );
                        text = text.replace("<College>", college[count-1] );
                        text = text.replace("<Position>", position[count-1] );
                        text = text.replace("<Society>", society[count-1] );
                        ((XSLFTextBox) shape).setText(text);
                        ppt.write(os);
                    }
                    count++;
                }
                ppt.write(os);
            }
            Log.d("TemplateActivity", "shapes " + slides.size());
            if(os!=null) {
                try {
                    os.flush();
                    os.close();
                    os = null;
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if(is!=null){
                try{
                    is.close();
                    is=null;
                }catch (IOException e){
                    e.printStackTrace();
                }
        } }catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void copyFile(InputStream in, OutputStream out) throws IOException {
        byte[] buffer = new byte[1024];
        int read;
        while((read = in.read(buffer)) != -1){
            out.write(buffer, 0, read);
        }
    }
}