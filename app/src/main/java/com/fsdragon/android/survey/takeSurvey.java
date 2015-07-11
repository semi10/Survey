package com.fsdragon.android.survey;

import android.app.Activity;
import android.content.Context;
import android.content.Intent;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.SeekBar;
import android.widget.TextView;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class takeSurvey extends Activity implements View.OnClickListener {
    List<Integer> answerList = new ArrayList<Integer>();
    List<String> questionList = new ArrayList<String>();
    int currentQuestionID = 0;
    int participantPosition = 0;
    String participantID = null;
    TextView question, progress;
    Button next;
    SeekBar seekBar;
    Integer survey_ID;
    Integer seekBarValue = 0;

    @Override
    public void onClick(View v) {

    }


    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.take_survey);

        question = (TextView) findViewById(R.id.question_tv);
        progress = (TextView) findViewById(R.id.progress_tv);
        next = (Button) findViewById(R.id.next_bt);
        seekBar = (SeekBar) findViewById(R.id.seekBar);

        Intent intent = getIntent();
        survey_ID = getIntent().getIntExtra("Survey_ID", 255);
        //Toast.makeText(takeSurvey.this, Integer.toString(survey_ID), Toast.LENGTH_SHORT).show();

        getQuestions(this, survey_ID, questionList);
        participantPosition = getParticipantPosition(this, survey_ID);
        question.setText(questionList.get(currentQuestionID));
        currentQuestionID++;
        progress.setText(currentQuestionID + "/" + questionList.size());

        showIDDialog();

        next.setOnClickListener(new View.OnClickListener(){
            public void onClick(View v) {
                currentQuestionID++;
                answerList.add(seekBarValue);

                if(currentQuestionID <= questionList.size()){
                    question.setText(questionList.get(currentQuestionID - 1));
                    seekBar.setProgress(0);
                    progress.setText(currentQuestionID + "/" + questionList.size());
                }
                else {
                    writeAnswer(takeSurvey.this, survey_ID, participantID, participantPosition, answerList);
                    showThxDialog();
                }
            }
        });


        seekBar.setOnSeekBarChangeListener(new SeekBar.OnSeekBarChangeListener() {

            @Override
            public void onProgressChanged(SeekBar seekBar, int progresValue, boolean fromUser) {
                seekBarValue = progresValue;
            }

            @Override
            public void onStartTrackingTouch(SeekBar seekBar) {

            }

            @Override
            public void onStopTrackingTouch(SeekBar seekBar) {

            }
        });
    }

    private void showThxDialog(){
        MyAlert myAlert = new MyAlert();
        myAlert.show(getFragmentManager(), "My Alert");
    }

    public void showIDDialog(){
        GiveID giveId = new GiveID();
        giveId.show(getFragmentManager(), "Give ID");
    }

    public void giveID(String ID){
        participantID = ID;
    }

    private static boolean writeAnswer(Context context, Integer survey_ID, String participantID, int participantPosition, List<Integer> answerList){

        boolean success = false;

        //check if available and not read only
        if (!storageAvailable()) return false;

        try {
            //Creating Input Stream
            File docsFolder = new File(Environment.getExternalStorageDirectory() + "/Documents");
            docsFolder.mkdirs();
            File file = new File(docsFolder, "Survey.xls");
            FileInputStream myInput = new FileInputStream(file);

            //Create a POIFFileSystem object
            POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);

            //Create a workbook using the File System
            HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);

            HSSFSheet mySheet = myWorkBook.getSheetAt(survey_ID);

            Cell cell = null;
            Row row = null;

            if (mySheet.getFirstRowNum() != 0) row = mySheet.createRow(0);

            //Create cell Iterator
            Iterator rowIter = mySheet.rowIterator();

            while(rowIter.hasNext()){
                HSSFRow myRow = (HSSFRow) rowIter.next();
                Iterator cellIter = myRow.cellIterator();
                while (cellIter.hasNext()) {
                    HSSFCell myCell = (HSSFCell) cellIter.next();
                    Log.d("ExcelLog", "Cell Value: " + myCell.toString());
                    // Toast.makeText(context, "cell Value: " + myCell.toString(), Toast.LENGTH_SHORT).show();
                }
            }

            row = mySheet.getRow(0);
            cell = row.createCell(participantPosition);
            mySheet.setColumnWidth(participantPosition, (1000));
            cell.setCellValue(participantID);


            for(int i = 1; i <= answerList.size(); i++){
                row = mySheet.getRow(i);
                cell = row.createCell(participantPosition);
                cell.setCellValue(answerList.get(i - 1).toString());
            }

            FileOutputStream os = null;

            try{
                os = new FileOutputStream(file);
                myWorkBook.write(os);
                Log.w("FileUtils", "Writing file " + file);
                os.close();
                success = true;
            } catch (IOException e){
                os.close();
                Log.w("FileUtils", "Error writing " + file, e);
            } catch (Exception e){
                os.close();
                Log.w("FileUtils", "Failed to save file", e);
            } finally {
                try{
                    if (null != os)
                        os.close();
                }   catch (Exception ex){}
            }


        }catch (Exception e){e.printStackTrace();}

        Cell c = null;



        return success;
    }

    private static void getQuestions(Context context,Integer survey_ID, List<String> questionList){
        //check if available and not read only
        if (!storageAvailable()) return;

        try {
            //Creating Input Stream
            File docsFolder = new File(Environment.getExternalStorageDirectory() + "/Documents");
            docsFolder.mkdirs();
            File file = new File(docsFolder, "Survey.xls");
            FileInputStream myInput = new FileInputStream(file);

            //Create a POIFFileSystem object
            POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);

            //Create a workbook using the File System
            HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);

            //Get the firs sheet from workbook
            HSSFSheet mySheet = myWorkBook.getSheetAt(survey_ID);
            mySheet.setSelected(true);

            HSSFRow row;
            if (mySheet.getFirstRowNum() != 0) row = mySheet.createRow(0);

            //Create cell Iterator
            Iterator rowIter = mySheet.rowIterator();
            rowIter.next();

            while(rowIter.hasNext()){
                row = (HSSFRow) rowIter.next();
                String cellData = row.getCell(0).toString();
                if(!cellData.isEmpty())
                    questionList.add(cellData);
            }

        }catch (Exception e){e.printStackTrace();}
    }

    private static Integer getParticipantPosition(Context context, Integer survey_ID){
        //check if available and not read only
        if (!storageAvailable()) return null;
        Integer ID = 1;

        try {
            //Creating Input Stream
            File docsFolder = new File(Environment.getExternalStorageDirectory() + "/Documents");
            docsFolder.mkdirs();
            File file = new File(docsFolder, "Survey.xls");
            FileInputStream myInput = new FileInputStream(file);

            //Create a POIFFileSystem object
            POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);

            //Create a workbook using the File System
            HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);

            //Get the firs sheet from workbook
            HSSFSheet mySheet = myWorkBook.getSheetAt(survey_ID);

            //Create cell Iterator
            Row row = null;
            Cell cell;
            if (mySheet.getFirstRowNum() != 0) row = mySheet.createRow(0);

            row = mySheet.getRow(0);
            cell = row.createCell(0);
           // Toast.makeText(context , String.valueOf(row.getLastCellNum()), Toast.LENGTH_SHORT).show();
            ID = Integer.valueOf(row.getLastCellNum());

        }catch (Exception e){e.printStackTrace();}
        return ID;
    }

    public static boolean storageAvailable(){
        //check if available and not read only
        String SD_state = Environment.getExternalStorageState();
        if (!Environment.MEDIA_MOUNTED.equals(SD_state) || Environment.MEDIA_MOUNTED_READ_ONLY.equals(SD_state)){
            Log.e("Storage", "Storage not available or read only");
            return false;
        }
        else return true;
    }
}
