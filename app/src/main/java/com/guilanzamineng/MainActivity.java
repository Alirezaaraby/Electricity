package com.guilanzamineng;

import android.Manifest;
import android.app.Activity;
import android.app.Dialog;
import android.content.DialogInterface;
import android.content.Intent;
import android.content.SharedPreferences;
import android.content.pm.PackageManager;
import android.os.AsyncTask;
import android.os.Bundle;
import android.os.Environment;
import android.text.method.LinkMovementMethod;
import android.text.util.Linkify;
import android.util.Log;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.view.Window;
import android.widget.Button;
import android.widget.TextView;
import android.widget.Toast;

import androidx.appcompat.app.AlertDialog;
import androidx.appcompat.app.AppCompatActivity;
import androidx.appcompat.widget.Toolbar;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;
import androidx.viewpager.widget.ViewPager;

import com.github.clans.fab.FloatingActionButton;
import com.google.android.material.snackbar.Snackbar;
import com.google.android.material.tabs.TabLayout;
import com.guilanzamineng.main.SectionsPagerAdapter;

import org.json.JSONException;
import org.json.JSONObject;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import dmax.dialog.SpotsDialog;

public class MainActivity extends AppCompatActivity {
    String Version = "";
    String Tittle_txt = "";
    String Description_txt = "";
    String Button_txt = "";

    String versionName;

    private String TAG = MainActivity.class.getSimpleName();
    Dialog pDialog;

    private static String url = "https://gist.githubusercontent.com/Alirezaaraby/a460665919e3ef47488b292525989809/raw/Electricity.json";

    public boolean Internet;

    private FloatingActionButton fab1;
    private FloatingActionButton fab2;
    private FloatingActionButton fab3;

    SharedPreferences Excel_Pref = null;
    SharedPreferences PDF_Pref = null;

//    SearchView sh;

    AlertDialog.Builder builder;
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        versionName = BuildConfig.VERSION_NAME;

        Excel_Pref = getSharedPreferences("com.guilanzamineng.app", MODE_PRIVATE);
        PDF_Pref = getSharedPreferences("com.guilanzamineng.app", MODE_PRIVATE);

        SectionsPagerAdapter sectionsPagerAdapter = new SectionsPagerAdapter(this, getSupportFragmentManager());
        ViewPager viewPager = findViewById(R.id.view_pager);
        viewPager.setAdapter(sectionsPagerAdapter);
        TabLayout tabs = findViewById(R.id.tabs);
        tabs.setupWithViewPager(viewPager);

        Toolbar toolbar = findViewById(R.id.toolbar);
        setSupportActionBar(toolbar);

        fab1 = findViewById(R.id.excel);
        fab2 = findViewById(R.id.pdf);
        fab3 = findViewById(R.id.setting);

        fab1.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                File folder = new File(Environment.getExternalStorageDirectory() + "/Electricity/Excel");
                if (Excel_Pref.getBoolean("firstclick", true)) {
                    if (checkPermissions()){
                        if (folder.exists()){
//                            CoordinatorLayout parent = findViewById(R.id.constraintLayout);
//                            Snackbar.make(parent , "Electricity File Is Already Exist Please Make Sure It's a Junk File And Delete That And Try Again",Snackbar.LENGTH_LONG);
                            Toast.makeText(MainActivity.this, "Electricity/Excel File Is Already Exist Please Make Sure It's a Junk File And Delete That And Try Again", Toast.LENGTH_LONG).show();
                        }
                        else {
                            createFolder();
                            Excel_Pref.edit().putBoolean("firstclick", false).apply();
                            Intent intent = new Intent(MainActivity.this,Excel_Creator.class);
                            startActivity(intent);
                            finish();
                        }
                    }
                    else{
//                        CoordinatorLayout parent = findViewById(R.id.constraintLayout);
//                        Snackbar.make(parent , "You Should Grant Permission",Snackbar.LENGTH_LONG);
                        Toast.makeText(MainActivity.this, "You Should Grant Permission", Toast.LENGTH_SHORT).show();
                    }
                }
                else{
                    Intent intent = new Intent(MainActivity.this,Excel_Creator.class);
                    startActivity(intent);
                }
            }
        });
        fab2.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                File folder = new File(Environment.getExternalStorageDirectory() + "/Electricity/PDF");
                if (PDF_Pref.getBoolean("firstclick", true)) {
                    if (checkPermissions()){
                        if (folder.exists()){
//                            CoordinatorLayout parent = findViewById(R.id.constraintLayout);
//                            Snackbar.make(parent , "Electricity File Is Already Exist Please Make Sure It's a Junk File And Delete That And Try Again",Snackbar.LENGTH_LONG);
                            Toast.makeText(MainActivity.this, "Electricity/Excel File Is Already Exist Please Make Sure It's a Junk File And Delete That And Try Again", Toast.LENGTH_LONG).show();
                        }
                        else {
                            createFolder();
                            PDF_Pref.edit().putBoolean("firstclick", false).apply();
                            Intent intent = new Intent(MainActivity.this,Excel_Creator.class);
                            startActivity(intent);
                            finish();
                        }
                    }
                    else{
//                        CoordinatorLayout parent = findViewById(R.id.constraintLayout);
//                        Snackbar.make(parent , "You Should Grant Permission",Snackbar.LENGTH_LONG);
                        Toast.makeText(MainActivity.this, "You Should Grant Permission", Toast.LENGTH_SHORT).show();
                    }
                }
                else{
                    Intent intent = new Intent(MainActivity.this,PDF_Creator.class);
                    startActivity(intent);
                    finish();
//                    Toast.makeText(MainActivity.this, "PDF", Toast.LENGTH_SHORT).show();
                }
            }
        });
        fab3.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intent = new Intent(MainActivity.this,Add.class);
                startActivity(intent);
//                Toast.makeText(MainActivity.this, "Add", Toast.LENGTH_SHORT).show();
            }
        });
    }

    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        getMenuInflater().inflate(R.menu.menu_main, menu);
        return true;
    }

    @Override
    public boolean onOptionsItemSelected(MenuItem item) {
        int id = item.getItemId();

        if (id == R.id.action_Setting) {
            return true;
        }
        if (id == R.id.About_About) {
            AlertDialog alertDialog;
            alertDialog = new AlertDialog.Builder(this)
                    .setTitle("About Creator:")
                    .setMessage("Codded By AlirezaAraby For More Information And Questions Email To:"+'\n'+"alirezaaraby5@gmail.com")
                    .setPositiveButton("Ok", new DialogInterface.OnClickListener() {
                        @Override
                        public void onClick(DialogInterface dialogInterface, int i) {

                        }
                    })
                    .show();

        }
        if (id == R.id.About_CheckLatestVersion) {
            new GetVersion().execute();
        }

        return super.onOptionsItemSelected(item);
    }

    private class GetVersion extends AsyncTask<Void, Void, Void> {

        @Override
        protected void onPreExecute() {
            super.onPreExecute();

            pDialog = new SpotsDialog.Builder().setContext(MainActivity.this).build();
            pDialog.show();
        }

        @Override
        protected Void doInBackground(final Void... arg0) {
            JSONObject JsonMain = null;
            HttpHandler sh = new HttpHandler();

            String jsonStr = sh.makeServiceCall(url);

            Log.e(TAG, "Response from url: " + jsonStr);

            if (jsonStr != null) {
                try {
                    JSONObject jsonObj = new JSONObject(jsonStr);

                    JsonMain = jsonObj.getJSONObject("Header");
                    Tittle_txt = JsonMain.getString("Version_Tittle");
                    Description_txt = JsonMain.getString("Version_Description");
                    Version = JsonMain.getString("Version");
                    Button_txt = JsonMain.getString("Version_Button_Text");
                    Internet = true;

                }

                catch (final JSONException e) {
                    Log.e(TAG, "Json parsing error: " + e.getMessage());
                    runOnUiThread(new Runnable() {
                        @Override
                        public void run() {
                            View parentLayout = findViewById(R.id.coordinator_layout);
                            Snackbar.make(parentLayout,"Json parsing error", Snackbar.LENGTH_LONG)
                                    .setAction("Action", null).show();
                            Internet = false;
                        }
                    });
                }
            }

            else {
                Log.e(TAG, "Couldn't get json from server.");
                runOnUiThread(new Runnable() {
                    @Override
                    public void run() {
                        View parentLayout = findViewById(R.id.coordinator_layout);
                        Snackbar.make(parentLayout,"No Internet Connection", Snackbar.LENGTH_LONG)
                                .setAction("Action", null).show();
                        Internet = false;
                    }
                });
            }
            return null;
        }

        @Override
        protected void onPostExecute(Void result) {

            super.onPostExecute(result);

            if (Internet){
                pDialog.dismiss();
                if (!Version.equals(versionName)) {
                    showDialog(MainActivity.this,"");
                }
                else{
                    View parentLayout = findViewById(R.id.coordinator_layout);
                    Snackbar.make(parentLayout,"Your Version Is Up To Date", Snackbar.LENGTH_LONG)
                            .setAction("Action", null).show();
                }
            }
            else {
                pDialog.dismiss();
            }

        }
    }

    public void showDialog(Activity activity,String msg){
        if (Internet) {
            final Dialog dialog = new Dialog(activity);
            dialog.requestWindowFeature(Window.FEATURE_NO_TITLE);
            dialog.setCancelable(false);
            dialog.setContentView(R.layout.versiondialog);

            TextView Tittle = dialog.findViewById(R.id.Tittle);
            TextView Des = dialog.findViewById(R.id.Des);
            Tittle.setText(Tittle_txt);
            Des.setText(Description_txt);

            Des.setMovementMethod(LinkMovementMethod.getInstance());

            Linkify.addLinks(Des, Linkify.WEB_URLS);

            Button dialogButton = dialog.findViewById(R.id.buttonOk);
            dialogButton.setText(Button_txt);
            dialogButton.setOnClickListener(new View.OnClickListener() {
                @Override
                public void onClick(View v) {
                    dialog.dismiss();
                }
            });

            dialog.show();
        }
    }
    public static final int MULTIPLE_PERMISSIONS = 10; // code you want.

    String[] permissions= new String[]{
            Manifest.permission.WRITE_EXTERNAL_STORAGE,
            Manifest.permission.WRITE_EXTERNAL_STORAGE
    };

    private  boolean checkPermissions() {
        int result;
        List<String> listPermissionsNeeded = new ArrayList<>();
        for (String p:permissions) {
            result = ContextCompat.checkSelfPermission(MainActivity.this,p);
            if (result != PackageManager.PERMISSION_GRANTED) {
                listPermissionsNeeded.add(p);
            }
        }
        if (!listPermissionsNeeded.isEmpty()) {
            ActivityCompat.requestPermissions(this, listPermissionsNeeded.toArray(new String[listPermissionsNeeded.size()]),MULTIPLE_PERMISSIONS );
            return false;
        }
        return true;
    }
    public File Directory;
    private void createFolder()
    {
        File folder2 = new File(Environment.getExternalStorageDirectory() + "/Electricity");
        if (!folder2.exists()) {
            folder2.mkdir();
            Directory= folder2.getAbsoluteFile();
        }

        else{
//            CoordinatorLayout parent = findViewById(R.id.Parent_Coordinator_Excel);
//            Snackbar.make(parent , "Electricity Folder Is Already Exist" ,  Snackbar.LENGTH_SHORT).show();
            Toast.makeText(this, "Electricity Folder Is Already Exist", Toast.LENGTH_LONG).show();
        }


        File folder = new File(Environment.getExternalStorageDirectory() + "/Electricity/Excel");
        boolean success = true;
        if (!folder.exists()) {
            success = folder.mkdir();
        }

        File folder3 = new File(Environment.getExternalStorageDirectory() + "/Electricity/PDF");
        boolean success2 = true;
        if (!folder3.exists()) {
            success2 = folder3.mkdir();
        }
    }
}