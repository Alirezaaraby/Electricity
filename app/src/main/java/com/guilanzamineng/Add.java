package com.guilanzamineng;

import android.content.Intent;
import android.os.Bundle;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.widget.AutoCompleteTextView;
import android.widget.Button;
import android.widget.Toast;

import androidx.annotation.NonNull;
import androidx.appcompat.app.AppCompatActivity;

import com.google.android.material.floatingactionbutton.FloatingActionButton;
import com.google.android.material.textfield.TextInputLayout;

import java.util.ArrayList;

public class Add extends AppCompatActivity {
    private Menu menu;
    @Override
    public boolean onCreateOptionsMenu(final Menu menu) {
        getMenuInflater().inflate(R.menu.excel_menu, menu);
        menu.findItem(R.id.Show_Excel_Menu_Visible_Fab).setVisible(false);
        this.menu = menu;
        return true;
    }

    @Override
    public boolean onOptionsItemSelected(@NonNull final MenuItem item) {

        Intent intent = new Intent(Add.this , MainActivity.class);
        startActivity(intent);
        finish();
        return true;
    }
    TextInputLayout SUBSTATION_Text_Input_Layout;
    TextInputLayout BAY_NO_Text_Input_Layout;
    TextInputLayout PANEL_NU_Text_Input_Layout;
    TextInputLayout METER_MANUFACTURER_Text_Input_Layout;
    TextInputLayout CT_RATIO_Text_Input_Layout;
    TextInputLayout PT_RATIO_Text_Input_Layout;
    TextInputLayout CT_RATIO_CLASS__Text_Input_Layout;
    TextInputLayout PT_RATIO_CLASS__Text_Input_Layout;
    TinyDB tinydb;
    AutoCompleteTextView SUBSTATION_Text_Input_Layout_Add;
    AutoCompleteTextView BAY_NO_Text_Input_Layout_Add;
    AutoCompleteTextView PANEL_NU_Text_Input_Layout_Add;
    AutoCompleteTextView METER_MANUFACTURER_Text_Input_Layout_Add;
    AutoCompleteTextView CT_RATIO_Text_Input_Layout_Add;
    AutoCompleteTextView PT_RATIO_Text_Input_Layout_Add;
    AutoCompleteTextView CT_RATIO_CLASS__Text_Input_Layout_Add;
    AutoCompleteTextView PT_RATIO_CLASS__Text_Input_Layout_Add;

    ArrayList<String> SubStation_list;

    ArrayList<String> BAY_NO_list;

    ArrayList<String> PANEL_NU_list;

    ArrayList<String> METER_MANUFACTURER_list;

    ArrayList<String> CT_RATIO_list;

    ArrayList<String> PT_RATIO_list;

    ArrayList<String> CT_RATIO_class_list;

    ArrayList<String> PT_RATIO_class_list;
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_add);
        try {
            getSupportActionBar().setDisplayHomeAsUpEnabled(true);
            getSupportActionBar().setDisplayShowHomeEnabled(true);
        }
        catch (Exception e){
            Toast.makeText(this, "An Undefined Error Occurred", Toast.LENGTH_SHORT).show();
        }
        SUBSTATION_Text_Input_Layout = findViewById(R.id.SUBSTATION_Add);
        BAY_NO_Text_Input_Layout = findViewById(R.id.BAY_NO_Add);
        PANEL_NU_Text_Input_Layout = findViewById(R.id.PANEL_NO_Add);
        METER_MANUFACTURER_Text_Input_Layout = findViewById(R.id.METER_MANUFACTURER_Add);
        CT_RATIO_Text_Input_Layout = findViewById(R.id.CT_RATIO_Add);
        CT_RATIO_CLASS__Text_Input_Layout = findViewById(R.id.CT_RATIO_CLASS_Add);
        PT_RATIO_Text_Input_Layout = findViewById(R.id.PT_RATIO_Add);
        PT_RATIO_CLASS__Text_Input_Layout = findViewById(R.id.PT_RATIO_CLASS_Add);
        FloatingActionButton fab = findViewById(R.id.Clear_All_Add);
        fab.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                SUBSTATION_Text_Input_Layout_Add = findViewById(R.id.SUBSTATION__Add);
                BAY_NO_Text_Input_Layout_Add = findViewById(R.id.BAY_NO__Add);
                PANEL_NU_Text_Input_Layout_Add = findViewById(R.id.PANEL_NO__Add);
                METER_MANUFACTURER_Text_Input_Layout_Add = findViewById(R.id.METER_MANUFACTURER__Add);
                CT_RATIO_Text_Input_Layout_Add = findViewById(R.id.CT_RATIO__Add);
                CT_RATIO_CLASS__Text_Input_Layout_Add = findViewById(R.id.CT_RATIO_CLASS__Add);
                PT_RATIO_Text_Input_Layout_Add = findViewById(R.id.PT_RATIO__Add);
                PT_RATIO_CLASS__Text_Input_Layout_Add = findViewById(R.id.PT_RATIO_CLASS__Add);
                SUBSTATION_Text_Input_Layout_Add.setText("");
                BAY_NO_Text_Input_Layout_Add.setText("");
                PANEL_NU_Text_Input_Layout_Add.setText("");
                METER_MANUFACTURER_Text_Input_Layout_Add.setText("");
                CT_RATIO_Text_Input_Layout_Add.setText("");
                CT_RATIO_CLASS__Text_Input_Layout_Add.setText("");
                PT_RATIO_Text_Input_Layout_Add.setText("");
                PT_RATIO_CLASS__Text_Input_Layout_Add.setText("");
            }
        });
        Button btn = findViewById(R.id.save_add);
        btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                tinydb = new TinyDB(Add.this);
                if (!CT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !PT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !SUBSTATION_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !BAY_NO_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !METER_MANUFACTURER_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !PANEL_NU_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !PT_RATIO_CLASS__Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !PT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !CT_RATIO_CLASS__Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !CT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty()) {
                    try {
                        try {
                            String[] CT = CT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().split("/");
                            String[] PT = PT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().split("/");
                            String i = CT[0] + CT[1];
                            String i1 = PT[0] + PT[1];

                            SubStation_list = new ArrayList<>();
                            SubStation_list.add(SUBSTATION_Text_Input_Layout.getEditText().getText().toString().trim());

                            ArrayList<String> newList = new ArrayList<String>() {
                                {
                                    addAll(tinydb.getListString("SUBSTATION"));
                                    addAll(SubStation_list);
                                }
                            };

                            BAY_NO_list = new ArrayList<>();
                            BAY_NO_list.add(BAY_NO_Text_Input_Layout.getEditText().getText().toString().trim());
                            ArrayList<String> newList2 = new ArrayList<String>() {
                                {
                                    addAll(tinydb.getListString("BAY_NO"));
                                    addAll(BAY_NO_list);
                                }
                            };

                            PANEL_NU_list = new ArrayList<>();
                            PANEL_NU_list.add(PANEL_NU_Text_Input_Layout.getEditText().getText().toString().trim());
                            ArrayList<String> newList3 = new ArrayList<String>() {
                                {
                                    addAll(tinydb.getListString("PANEL_NU"));
                                    addAll(PANEL_NU_list);
                                }
                            };

                            METER_MANUFACTURER_list = new ArrayList<>();
                            METER_MANUFACTURER_list.add(METER_MANUFACTURER_Text_Input_Layout.getEditText().getText().toString().trim());
                            ArrayList<String> newList4 = new ArrayList<String>() {
                                {
                                    addAll(tinydb.getListString("METTER_MU"));
                                    addAll(METER_MANUFACTURER_list);
                                }
                            };

                            CT_RATIO_list = new ArrayList<>();
                            CT_RATIO_list.add(CT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim());
                            ArrayList<String> newList5 = new ArrayList<String>() {
                                {
                                    addAll(tinydb.getListString("CT_RAITIO"));
                                    addAll(CT_RATIO_list);
                                }
                            };

                            PT_RATIO_list = new ArrayList<>();
                            PT_RATIO_list.add(PT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim());
                            ArrayList<String> newList6 = new ArrayList<String>() {
                                {
                                    addAll(tinydb.getListString("PT_RAITIO"));
                                    addAll(PT_RATIO_list);
                                }
                            };

                            CT_RATIO_class_list = new ArrayList<>();
                            CT_RATIO_class_list.add(CT_RATIO_CLASS__Text_Input_Layout.getEditText().getText().toString().trim());
                            ArrayList<String> newList7 = new ArrayList<String>() {
                                {
                                    addAll(tinydb.getListString("CT_RAITIO_CLASS"));
                                    addAll(CT_RATIO_class_list);
                                }
                            };

                            PT_RATIO_class_list = new ArrayList<>();
                            PT_RATIO_class_list.add(PT_RATIO_CLASS__Text_Input_Layout.getEditText().getText().toString().trim());
                            ArrayList<String> newList8 = new ArrayList<String>() {
                                {
                                    addAll(tinydb.getListString("PT_RAITIO_CLASS"));
                                    addAll(PT_RATIO_class_list);
                                }
                            };


                            tinydb.putListString("SUBSTATION", newList);
                            tinydb.putListString("BAY_NO", newList2);
                            tinydb.putListString("PANEL_NU", newList3);
                            tinydb.putListString("METTER_MU", newList4);

                            tinydb.putListString("CT_RAITIO", newList5);
                            tinydb.putListString("PT_RAITIO", newList6);
                            tinydb.putListString("CT_RAITIO_CLASS", newList7);
                            tinydb.putListString("PT_RAITIO_CLASS", newList8);

                            Toast.makeText(Add.this, "Saved", Toast.LENGTH_SHORT).show();
                        }
                        catch (Exception e){
                            Toast.makeText(Add.this, "Format For CT And PT RATIO Is Invalid", Toast.LENGTH_SHORT).show();
                        }
                    }
                    catch (Exception e) {
                        Toast.makeText(Add.this, "Not Saved", Toast.LENGTH_SHORT).show();
                    }
                }
                else{
                    if (CT_RATIO_Text_Input_Layout.getEditText().getText().toString().isEmpty())
                        CT_RATIO_Text_Input_Layout.setError("This Field Is Required");
                    if (PT_RATIO_Text_Input_Layout.getEditText().getText().toString().isEmpty())
                        PT_RATIO_Text_Input_Layout.setError("This Field Is Required");
                    if (METER_MANUFACTURER_Text_Input_Layout.getEditText().getText().toString().isEmpty())
                        METER_MANUFACTURER_Text_Input_Layout.setError("This Field Is Required");
                    if (PT_RATIO_CLASS__Text_Input_Layout.getEditText().getText().toString().isEmpty())
                        PT_RATIO_CLASS__Text_Input_Layout.setError("This Field Is Required");
                    if (CT_RATIO_CLASS__Text_Input_Layout.getEditText().getText().toString().isEmpty())
                        CT_RATIO_CLASS__Text_Input_Layout.setError("This Field Is Required");
                    if (SUBSTATION_Text_Input_Layout.getEditText().getText().toString().isEmpty())
                        SUBSTATION_Text_Input_Layout.setError("This Field Is Required");
                    if (BAY_NO_Text_Input_Layout.getEditText().getText().toString().isEmpty())
                        BAY_NO_Text_Input_Layout.setError("This Field Is Required");
                }
            }
        });
    }
}
