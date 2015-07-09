package com.fsdragon.android.survey;

import android.app.Activity;
import android.content.Context;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.widget.AdapterView;
import android.widget.ArrayAdapter;
import android.widget.ListView;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;


public class MainActivity extends Activity implements AdapterView.OnItemClickListener , View.OnClickListener {
    ListView l;
   // Button newSurvey;
    List<String> surveyList = new ArrayList<String>();


    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

    }

    @Override
    protected void onResume() {
        super.onResume();

        //check if available and not read only
        if (!storageAvailable()){
            Toast.makeText(this, "No SD card Found", Toast.LENGTH_LONG).show();
            finish();
        }

        File docsFolder = new File(Environment.getExternalStorageDirectory() + "/Documents");
        docsFolder.mkdirs();
        File file = new File(docsFolder, "Survey.xls");

        if (!file.exists()){
            Toast.makeText(this, "Create New Survey", Toast.LENGTH_LONG).show();
            Intent intent = new Intent(this, createSurvey.class);
            startActivity(intent);
        }
        else {
            getSheetsName(MainActivity.this, "Survey.xls", surveyList);
            l = (ListView) findViewById(R.id.listView);
            ArrayAdapter<String> adapter = new ArrayAdapter<String>(this, R.layout.fill_parent_list, surveyList.toArray(new String[surveyList.size()]));
            l.setAdapter(adapter);
            l.setOnItemClickListener(this);
        }
    }

    @Override
    public void onClick(View v) {
        // saveExcelFile(MainActivity.this, "Survey.xls");
        readExcelFile(MainActivity.this, "Survey.xls");
    }

    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        // Inflate the menu; this adds items to the action bar if it is present.
        getMenuInflater().inflate(R.menu.menu_main, menu);
        return true;
    }

    @Override
    public boolean onOptionsItemSelected(MenuItem item) {
        // Handle action bar item clicks here. The action bar will
        // automatically handle clicks on the Home/Up button, so long
        // as you specify a parent activity in AndroidManifest.xml.
        int id = item.getItemId();

        //noinspection SimplifiableIfStatement
        if(id == R.id.action_newSurvey){
            Intent intent = new Intent(this, createSurvey.class);
            startActivity(intent);
            return true;
        } else if(id == R.id.action_About){
            showDialog();
            return true;
        } else if (id == R.id.action_share){
            Intent intent = null, chooser = null;

            File docsFolder = new File(Environment.getExternalStorageDirectory() + "/Documents");
            docsFolder.mkdirs();
            File file = new File(docsFolder, "Survey.xls");
            Uri xlsUri = null;
            xlsUri = Uri.parse("file:///" + file);
            intent = new Intent(Intent.ACTION_SEND);
            intent.setType("application/excel");
            intent.putExtra(Intent.EXTRA_STREAM, xlsUri);
            chooser = Intent.createChooser(intent, "My Survey");
            startActivity(chooser);
            return true;
        }

        return super.onOptionsItemSelected(item);
    }

    private void showDialog(){
        About about = new About();
        about.show(getFragmentManager(), "About");
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

    private static void readExcelFile(Context context, String fileName){
        //check if available and not read only
        if (!storageAvailable()) return ;

        try {
            //Creating Input Stream
            File docsFolder = new File(Environment.getExternalStorageDirectory() + "/Documents");
            docsFolder.mkdirs();
            File file = new File(docsFolder, fileName);
            FileInputStream myInput = new FileInputStream(file);

            //Create a POIFFileSystem object
            POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);

            //Create a workbook using the File System
            HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);

            //Get the firs sheet from workbook
            HSSFSheet mySheet = myWorkBook.getSheetAt(0);

            //String numberOfSheets = String.valueOf(myWorkBook.getNumberOfSheets());
            Toast.makeText(context, String.valueOf(myWorkBook.getNumberOfSheets()), Toast.LENGTH_SHORT).show();
            Toast.makeText(context, mySheet.getSheetName(), Toast.LENGTH_SHORT).show();

            //Create celll Iterator
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
            myInput.close();
        }catch (Exception e){e.printStackTrace();}
    }

    private static void getSheetsName(Context context, String fileName, List<String> surveyList){

        //check if available and not read only
        if (!storageAvailable()) return;

        try {
            //Creating Input Stream
            File docsFolder = new File(Environment.getExternalStorageDirectory() + "/Documents");
            docsFolder.mkdirs();
            File file = new File(docsFolder, fileName);
            FileInputStream myInput = new FileInputStream(file);

            //Create a POIFFileSystem object
            POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);

            //Create a workbook using the File System
            HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);


            surveyList.clear();
            for(int i = 0; i < myWorkBook.getNumberOfSheets(); i++){

                //Sheet from workbook
                HSSFSheet mySheet = myWorkBook.getSheetAt(i);

                int numberOfParticipants = 0;
                Row row = null;
                Cell cell;
                if (mySheet.getFirstRowNum() != 0) row = mySheet.createRow(0);

                row = mySheet.getRow(0);
                cell = row.createCell(0);

                numberOfParticipants = Integer.valueOf(row.getLastCellNum()) - 1;
                surveyList.add(myWorkBook.getSheetAt(i).getSheetName() + " (" +  numberOfParticipants + ") ");
            }

        }catch (Exception e){e.printStackTrace();}
    }

    @Override
    public void onItemClick(AdapterView<?> parent, View view, int position, long id) {
        Intent intent = null;
        intent = new Intent(this, takeSurvey.class);
        intent.putExtra("Survey_ID", position);
        startActivity(intent);
    }
}
