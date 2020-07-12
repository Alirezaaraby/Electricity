package com.guilanzamineng;

import android.content.Intent;
import android.content.SharedPreferences;
import android.os.AsyncTask;
import android.os.Bundle;
import android.os.Handler;
import android.util.Log;
import android.view.View;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;

import com.agrawalsuneet.dotsloader.loaders.LazyLoader;
import com.google.android.material.snackbar.Snackbar;

import org.json.JSONException;
import org.json.JSONObject;

public class SplashScreen extends AppCompatActivity {

    LazyLoader lazyLoader;

    private static String url = "https://gist.githubusercontent.com/Alirezaaraby/a460665919e3ef47488b292525989809/raw/Electricity.json";

    SharedPreferences Enable_Application = null;
    public boolean Internet;
    private String TAG = SplashScreen.class.getSimpleName();
    String Enable;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_splash_screen);
        lazyLoader = findViewById(R.id.Container_Progress);

        Enable_Application = getSharedPreferences("com.guilanzamineng.app", MODE_PRIVATE);

        new GetEnable().execute();
    }

    private class GetEnable extends AsyncTask<Void, Void, Void> {

        @Override
        protected void onPreExecute() {
            super.onPreExecute();
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
                    Enable = JsonMain.getString("Enable");
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
                if (!Enable.equals("true")) {
                    Enable_Application.edit().putBoolean("Enable", false).apply();
                    Toast.makeText(SplashScreen.this, "This Application Is Not Available For You Right Now", Toast.LENGTH_SHORT).show();
                    finish();
                    System.exit(0);
                }
                else{
                    Enable_Application.edit().putBoolean("Enable", true).apply();
                    Intent intent = new Intent(SplashScreen.this, MainActivity.class);
                    startActivity(intent);
                }
            }
            else {
                new Handler().postDelayed(new Runnable() {
                    public void run() {
                        if (Enable_Application.getBoolean("Enable", true)) {
                            Intent intent = new Intent(SplashScreen.this, MainActivity.class);
                            startActivity(intent);
                            lazyLoader.setVisibility(View.GONE);
                            finish();
                        }
                        else{
                            Toast.makeText(SplashScreen.this, "This Application Is Not Available For You Right Now", Toast.LENGTH_SHORT).show();
                            finish();
                            System.exit(0);
                        }
                    }
                }, 3449);
            }
        }
    }
}
