package com.guilanzamineng;

import android.content.Intent;
import android.os.Bundle;
import android.os.Environment;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.view.ViewTreeObserver;
import android.widget.ArrayAdapter;
import android.widget.AutoCompleteTextView;
import android.widget.Button;
import android.widget.CheckBox;
import android.widget.ScrollView;
import android.widget.Toast;

import androidx.annotation.NonNull;
import androidx.appcompat.app.AppCompatActivity;

import com.google.android.material.floatingactionbutton.FloatingActionButton;
import com.google.android.material.snackbar.Snackbar;
import com.google.android.material.textfield.TextInputEditText;
import com.google.android.material.textfield.TextInputLayout;
import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.GrayColor;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

import java.io.File;
import java.io.FileOutputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Locale;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class Excel_Creator extends AppCompatActivity {

    FloatingActionButton fab;

    File directory, sd, file;
    WritableWorkbook workbook;

    TextInputLayout SUBSTATION_Text_Input_Layout;
    TextInputLayout BAY_NO_Text_Input_Layout;
    TextInputLayout PANEL_NU_Text_Input_Layout;
    TextInputLayout METER_MANUFACTURER_Text_Input_Layout;
    TextInputLayout CT_RATIO_Text_Input_Layout;
    TextInputLayout PT_RATIO_Text_Input_Layout;
    TextInputLayout CT_RATIO_CLASS__Text_Input_Layout;
    TextInputLayout PT_RATIO_CLASS__Text_Input_Layout;
    TextInputLayout METER_SCALE_RANGE_Text_Input_Layout;
    TextInputLayout METER_MODEL__SERIAL_NO_Text_Input_Layout;
    TextInputLayout REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout;
    TextInputLayout REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout;
    TextInputLayout REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout;
    TextInputLayout I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout;
    TextInputLayout I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout;
    TextInputLayout I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout;
    TextInputLayout REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout;
    TextInputLayout REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout;
    TextInputLayout REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout;
    TextInputLayout V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout;
    TextInputLayout V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout;
    TextInputLayout V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout;
    TextInputLayout ACTIVE_POWER;
    TextInputLayout REACTIVE_POWER;
    TextInputLayout COS_PHI;
    TextInputLayout COMMENT;
    TextInputLayout TEST_BY;
    TextInputLayout DATE;

    AutoCompleteTextView SUBSTATION_Text_Input_Layout_;
    AutoCompleteTextView BAY_NO_Text_Input_Layout_;
    AutoCompleteTextView PANEL_NU_Text_Input_Layout_;
    AutoCompleteTextView METER_MANUFACTURER_Text_Input_Layout_;
    AutoCompleteTextView CT_RATIO_Text_Input_Layout_;
    AutoCompleteTextView PT_RATIO_Text_Input_Layout_;
    AutoCompleteTextView CT_RATIO_CLASS__Text_Input_Layout_;
    AutoCompleteTextView PT_RATIO_CLASS__Text_Input_Layout_;
    TextInputEditText METER_SCALE_RANGE_Text_Input_Layout_;
    TextInputEditText METER_MODEL__SERIAL_NO_Text_Input_Layout_;
    TextInputEditText REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout_;
    TextInputEditText REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout_;
    TextInputEditText REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout_;
    TextInputEditText I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout_;
    TextInputEditText I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout_;
    TextInputEditText I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout_;
    TextInputEditText REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout_;
    TextInputEditText REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout_;
    TextInputEditText REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout_;
    TextInputEditText V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout_;
    TextInputEditText V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout_;
    TextInputEditText V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout_;
    TextInputEditText ACTIVE_POWER_;
    TextInputEditText REACTIVE_POWER_;
    TextInputEditText COS_PHI_;
    TextInputEditText COMMENT_;
    TextInputEditText TEST_BY_;
    TextInputEditText DATE_;

    FloatingActionButton scrolltotop;
    ScrollView scrollViewEventDetails;

    MenuItem visibility;

    private static final String TAG = "APP";

    CheckBox ch;

    private Menu menu;

    TinyDB tinydb;
    ArrayList<String> SubStation_list;
    AutoCompleteTextView autoCompleteTextView_SUBSTATION;

    ArrayList<String> BAY_NO_list;
    AutoCompleteTextView autoCompleteTextView_BAY_NO;

    ArrayList<String> PANEL_NU_list;
    AutoCompleteTextView autoCompleteTextView_PANEL_NU;

    ArrayList<String> METER_MANUFACTURER_list;
    AutoCompleteTextView autoCompleteTextView_METER_MANUFACTURER;

    ArrayList<String> CT_RATIO_list;
    AutoCompleteTextView autoCompleteTextView_CT_RATIO;

    ArrayList<String> PT_RATIO_list;
    AutoCompleteTextView autoCompleteTextView_PT_RATIO;

    ArrayList<String> CT_RATIO_class_list;
    AutoCompleteTextView autoCompleteTextView_CT_RATIO_class;

    ArrayList<String> PT_RATIO_class_list;
    AutoCompleteTextView autoCompleteTextView_PT_RATIO_class;

//    Boolean Saved;

    @Override
    public boolean onCreateOptionsMenu(final Menu menu) {
        getMenuInflater().inflate(R.menu.excel_menu, menu);
        menu.findItem(R.id.Show_Excel_Menu_Visible_Fab).setVisible(false);
        this.menu = menu;
        return true;
    }

    @Override
    public boolean onOptionsItemSelected(@NonNull final MenuItem item) {
        int id = item.getItemId();
        if (id == R.id.Show_Excel_Menu_Visible_Fab) {
            fab.setVisibility(View.VISIBLE);
            item.setVisible(false);
        }
        else{
            Intent intent = new Intent(Excel_Creator.this , MainActivity.class);
            startActivity(intent);
            finish();
        }
        return true;
    }

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_excel_creator);

        ch = findViewById(R.id.checkBox);

        visibility = findViewById(R.id.Show_Excel_Menu_Visible_Fab);

//        Saved = false;

        try {
            getSupportActionBar().setDisplayHomeAsUpEnabled(true);
            getSupportActionBar().setDisplayShowHomeEnabled(true);
        }
        catch (Exception e){
            Toast.makeText(this, "An Undefined Error Occurred", Toast.LENGTH_SHORT).show();
        }

        SUBSTATION_Text_Input_Layout = findViewById(R.id.SUBSTATION);
        BAY_NO_Text_Input_Layout = findViewById(R.id.BAY_NO);
        PANEL_NU_Text_Input_Layout = findViewById(R.id.PANEL_NO);
        METER_MANUFACTURER_Text_Input_Layout = findViewById(R.id.METER_MANUFACTURER);
        CT_RATIO_Text_Input_Layout = findViewById(R.id.CT_RATIO);
        CT_RATIO_CLASS__Text_Input_Layout = findViewById(R.id.CT_RATIO_CLASS);
        PT_RATIO_Text_Input_Layout = findViewById(R.id.PT_RATIO);
        PT_RATIO_CLASS__Text_Input_Layout = findViewById(R.id.PT_RATIO_CLASS);
        METER_SCALE_RANGE_Text_Input_Layout = findViewById(R.id.METER_SCALE_RANGE);
        METER_MODEL__SERIAL_NO_Text_Input_Layout = findViewById(R.id.METER_MODEL__SERIAL_NO);
        REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout = findViewById(R.id.REAL_VALUEI_LOAD_PHASE_A);
        REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout = findViewById(R.id.REAL_VALUEI_LOAD_PH_B);
        REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout = findViewById(R.id.REAL_VALUEI_LOAD_PH_C);
        I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout = findViewById(R.id.I_ONLOAD_by_clamp_PH_A);
        I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout = findViewById(R.id.I_ONLOAD_by_clamp_PH_B);
        I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout = findViewById(R.id.I_ONLOAD_by_clamp_PH_c);
        REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout = findViewById(R.id.REAL_VALUE_V_LOAD_PH_A);
        REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout = findViewById(R.id.REAL_VALUE_V_LOAD_PH_B);
        REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout = findViewById(R.id.REAL_VALUE_V_LOAD_PH_C);
        V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout = findViewById(R.id.V_ONLOAD_BY_V_M_V_LOAD_PH_A);
        V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout = findViewById(R.id.V_ONLOAD_BY_V_M_V_LOAD_PH_B);
        V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout = findViewById(R.id.V_ONLOAD_BY_V_M_V_LOAD_PH_C);
        ACTIVE_POWER = findViewById(R.id.ACTIVE_POWER);
        REACTIVE_POWER = findViewById(R.id.REACTIVE_POWER);
        COS_PHI = findViewById(R.id.COS_PHI);
        COMMENT = findViewById(R.id.COMMENT);
        TEST_BY = findViewById(R.id.TEST_BY);
        DATE = findViewById(R.id.DATE);

        SubStation_list = new ArrayList<String>();
        BAY_NO_list = new ArrayList<String>();
        PANEL_NU_list = new ArrayList<String>();
        METER_MANUFACTURER_list = new ArrayList<String>();
        CT_RATIO_list = new ArrayList<String>();
        PT_RATIO_list = new ArrayList<String>();
        CT_RATIO_class_list = new ArrayList<String>();
        PT_RATIO_class_list = new ArrayList<String>();

        tinydb = new TinyDB(Excel_Creator.this);

        autoCompleteTextView_SUBSTATION = findViewById(R.id.SUBSTATION______);
        autoCompleteTextView_BAY_NO = findViewById(R.id.BAY_NO______);
        autoCompleteTextView_PANEL_NU = findViewById(R.id.PANEL_NO______);
        autoCompleteTextView_METER_MANUFACTURER = findViewById(R.id.METER_MANUFACTURER______);

        autoCompleteTextView_CT_RATIO = findViewById(R.id.CT_RATIO______);
        autoCompleteTextView_PT_RATIO = findViewById(R.id.PT_RATIO______);
        autoCompleteTextView_CT_RATIO_class = findViewById(R.id.CT_RATIO_CLASS______);
        autoCompleteTextView_PT_RATIO_class = findViewById(R.id.PT_RATIO_CLASS______);

        ArrayAdapter<String> myAdapter_SUB = new ArrayAdapter<String>(Excel_Creator.this, android.R.layout.simple_dropdown_item_1line, tinydb.getListString("SUBSTATION"));
        autoCompleteTextView_SUBSTATION.setAdapter(myAdapter_SUB);

        ArrayAdapter<String> myAdapter_Bay = new ArrayAdapter<String>(Excel_Creator.this, android.R.layout.simple_dropdown_item_1line, tinydb.getListString("BAY_NO"));
        autoCompleteTextView_BAY_NO.setAdapter(myAdapter_Bay);

        ArrayAdapter<String> myAdapter_Panel = new ArrayAdapter<String>(Excel_Creator.this, android.R.layout.simple_dropdown_item_1line, tinydb.getListString("PANEL_NU"));
        autoCompleteTextView_PANEL_NU.setAdapter(myAdapter_Panel);

        ArrayAdapter<String> myAdapter_METTER_MU = new ArrayAdapter<String>(Excel_Creator.this, android.R.layout.simple_dropdown_item_1line, tinydb.getListString("METTER_MU"));
        autoCompleteTextView_METER_MANUFACTURER.setAdapter(myAdapter_METTER_MU);

        ArrayAdapter<String> myAdapter_CT = new ArrayAdapter<String>(Excel_Creator.this, android.R.layout.simple_dropdown_item_1line, tinydb.getListString("CT_RAITIO"));
        autoCompleteTextView_CT_RATIO.setAdapter(myAdapter_CT);

        ArrayAdapter<String> myAdapter_PT = new ArrayAdapter<String>(Excel_Creator.this, android.R.layout.simple_dropdown_item_1line, tinydb.getListString("PT_RAITIO"));
        autoCompleteTextView_PT_RATIO.setAdapter(myAdapter_PT);

        ArrayAdapter<String> myAdapter_CT_CLASS = new ArrayAdapter<String>(Excel_Creator.this, android.R.layout.simple_dropdown_item_1line, tinydb.getListString("CT_RAITIO_CLASS"));
        autoCompleteTextView_CT_RATIO_class.setAdapter(myAdapter_CT_CLASS);

        ArrayAdapter<String> myAdapter_PT_CLASS = new ArrayAdapter<String>(Excel_Creator.this, android.R.layout.simple_dropdown_item_1line, tinydb.getListString("PT_RAITIO_CLASS"));
        autoCompleteTextView_PT_RATIO_class.setAdapter(myAdapter_PT_CLASS);

        SUBSTATION_Text_Input_Layout_ = findViewById(R.id.SUBSTATION______);
        BAY_NO_Text_Input_Layout_ = findViewById(R.id.BAY_NO______);
        PANEL_NU_Text_Input_Layout_ = findViewById(R.id.PANEL_NO______);
        METER_MANUFACTURER_Text_Input_Layout_ = findViewById(R.id.METER_MANUFACTURER______);
        CT_RATIO_Text_Input_Layout_ = findViewById(R.id.CT_RATIO______);
        CT_RATIO_CLASS__Text_Input_Layout_ = findViewById(R.id.CT_RATIO_CLASS______);
        PT_RATIO_Text_Input_Layout_ = findViewById(R.id.PT_RATIO______);
        PT_RATIO_CLASS__Text_Input_Layout_ = findViewById(R.id.PT_RATIO_CLASS______);

        METER_SCALE_RANGE_Text_Input_Layout_ = findViewById(R.id.METER_SCALE_RANGE_);
        METER_MODEL__SERIAL_NO_Text_Input_Layout_ = findViewById(R.id.METER_MODEL__SERIAL_NO_);
        REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout_ = findViewById(R.id.REAL_VALUEI_LOAD_PHASE_A_);
        REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout_ = findViewById(R.id.REAL_VALUEI_LOAD_PH_B_);
        REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout_ = findViewById(R.id.REAL_VALUEI_LOAD_PH_C_);
        I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout_ = findViewById(R.id.I_ONLOAD_by_clamp_PH_A_);
        I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout_ = findViewById(R.id.I_ONLOAD_by_clamp_PH_B_);
        I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout_ = findViewById(R.id.I_ONLOAD_by_clamp_PH_C_);
        REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout_ = findViewById(R.id.REAL_VALUE_V_LOAD_PH_A_);
        REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout_ = findViewById(R.id.REAL_VALUE_V_LOAD_PH_B_);
        REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout_ = findViewById(R.id.REAL_VALUE_V_LOAD_PH_C_);
        V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout_ = findViewById(R.id.V_ONLOAD_BY_V_M_V_LOAD_PH_A_);
        V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout_ = findViewById(R.id.V_ONLOAD_BY_V_M_V_LOAD_PH_B_);
        V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout_ = findViewById(R.id.V_ONLOAD_BY_V_M_V_LOAD_PH_C_);
        ACTIVE_POWER_ = findViewById(R.id.ACTIVE_POWER_);
        REACTIVE_POWER_ = findViewById(R.id.REACTIVE_POWER_);
        COS_PHI_ = findViewById(R.id.COS_PHI_);
        COMMENT_ = findViewById(R.id.COMMENT_);
        TEST_BY_ = findViewById(R.id.TEST_BY_);
        DATE_ = findViewById(R.id.DATE_);

        scrollViewEventDetails = findViewById(R.id.scrollView);

        scrolltotop = findViewById(R.id.top);
        scrollViewEventDetails.getViewTreeObserver().addOnScrollChangedListener(new ViewTreeObserver.OnScrollChangedListener() {
            @Override
            public void onScrollChanged() {
                if (4100 >= scrollViewEventDetails.getScrollY()){
                    scrolltotop.setVisibility(View.VISIBLE);
                }
                if (4100 <= scrollViewEventDetails.getScrollY()){
                    scrolltotop.setVisibility(View.INVISIBLE);
                }
            }
        });

    scrolltotop.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                if (4100 >= scrollViewEventDetails.getScrollY()){
                    scrollViewEventDetails.fullScroll(View.FOCUS_DOWN);
                }
            }
        });

        fab = findViewById(R.id.clear_all);
        fab.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                SUBSTATION_Text_Input_Layout_.setText("");
                BAY_NO_Text_Input_Layout_.setText("");
                PANEL_NU_Text_Input_Layout_.setText("");
                METER_MANUFACTURER_Text_Input_Layout_.setText("");
                CT_RATIO_Text_Input_Layout_.setText("");
                CT_RATIO_CLASS__Text_Input_Layout_.setText("");
                PT_RATIO_Text_Input_Layout_.setText("");
                PT_RATIO_CLASS__Text_Input_Layout_.setText("");
                METER_SCALE_RANGE_Text_Input_Layout_.setText("");
                METER_MODEL__SERIAL_NO_Text_Input_Layout_.setText("");
                REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout_.setText("");
                REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout_.setText("");
                REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout_.setText("");
                I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout_.setText("");
                I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout_.setText("");
                I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout_.setText("");
                REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout_.setText("");
                REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout_.setText("");
                REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout_.setText("");
                V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout_.setText("");
                V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout_.setText("");
                V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout_.setText("");
                ACTIVE_POWER_.setText("");
                REACTIVE_POWER_.setText("");
                COS_PHI_.setText("");
                COMMENT_.setText("");
                TEST_BY_.setText("");
                DATE_.setText("");
                Snackbar.make(view, "Edit Texts Are Empty Now", Snackbar.LENGTH_SHORT)
                        .setAction("Action", null).show();
            }
        });
        fab.setOnLongClickListener(new View.OnLongClickListener() {
            @Override
            public boolean onLongClick(View v) {
                fab.setVisibility(View.GONE);

                menu.findItem(R.id.Show_Excel_Menu_Visible_Fab).setVisible(true);

                Snackbar.make(v, "Clear Button Is Invisible Now", Snackbar.LENGTH_LONG)
                        .setAction("Action", null).show();

                return true;
            }
        });
        Button Save = findViewById(R.id.outlinedButton);
        Save.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                SaveTheFile();
            }
        });
    }

    public void createExcelSheet() {
        if (!SUBSTATION_Text_Input_Layout.getEditText().getText().toString().trim().equals("")&& !BAY_NO_Text_Input_Layout.getEditText().getText().toString().trim().equals("") && !DATE.getEditText().getText().toString().trim().equals("") ) {
            String csvFile = SUBSTATION_Text_Input_Layout.getEditText().getText().toString() +"_"+ BAY_NO_Text_Input_Layout.getEditText().getText().toString() +"_"+ "metter" +"_"+ DATE.getEditText().getText().toString() + ".xls";
            sd = Environment.getExternalStorageDirectory();
            directory = new File(sd.getAbsolutePath() + "/Electricity/Excel");
            file = new File(directory, csvFile);
            WorkbookSettings wbSettings = new WorkbookSettings();
            wbSettings.setLocale(new Locale("en", "EN"));
            try {
//                if (Saved==true) {
//                    Saved = false;
                    String[] CT = CT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().split("/");
                    String[] PT = PT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().split("/");

                    int i = Integer.parseInt(CT[0]) / Integer.parseInt(CT[1]);
                    int i1 = Integer.parseInt(PT[0]) / Integer.parseInt(PT[1]);

                    workbook = Workbook.createWorkbook(file, wbSettings);
                    createFirstSheet();

                    workbook.write();
                    workbook.close();

                    SubStation_list.add(SUBSTATION_Text_Input_Layout.getEditText().getText().toString().trim());
                    ArrayList<String> newList = new ArrayList<String>() {
                        {
                            addAll(tinydb.getListString("SUBSTATION"));
                            addAll(SubStation_list);
                        }
                    };

                    BAY_NO_list.add(BAY_NO_Text_Input_Layout.getEditText().getText().toString().trim());
                    ArrayList<String> newList2 = new ArrayList<String>() {
                        {
                            addAll(tinydb.getListString("BAY_NO"));
                            addAll(BAY_NO_list);
                        }
                    };

                    PANEL_NU_list.add(PANEL_NU_Text_Input_Layout.getEditText().getText().toString().trim());
                    ArrayList<String> newList3 = new ArrayList<String>() {
                        {
                            addAll(tinydb.getListString("PANEL_NU"));
                            addAll(PANEL_NU_list);
                        }
                    };

                    METER_MANUFACTURER_list.add(METER_MANUFACTURER_Text_Input_Layout.getEditText().getText().toString().trim());
                    ArrayList<String> newList4 = new ArrayList<String>() {
                        {
                            addAll(tinydb.getListString("METTER_MU"));
                            addAll(METER_MANUFACTURER_list);
                        }
                    };

                    CT_RATIO_list.add(CT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim());
                    ArrayList<String> newList5 = new ArrayList<String>() {
                        {
                            addAll(tinydb.getListString("CT_RAITIO"));
                            addAll(CT_RATIO_list);
                        }
                    };

                    PT_RATIO_list.add(PT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim());
                    ArrayList<String> newList6 = new ArrayList<String>() {
                        {
                            addAll(tinydb.getListString("PT_RAITIO"));
                            addAll(PT_RATIO_list);
                        }
                    };

                    CT_RATIO_class_list.add(CT_RATIO_CLASS__Text_Input_Layout.getEditText().getText().toString().trim());
                    ArrayList<String> newList7 = new ArrayList<String>() {
                        {
                            addAll(tinydb.getListString("CT_RAITIO_CLASS"));
                            addAll(CT_RATIO_class_list);
                        }
                    };

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

                    ArrayAdapter<String> myAdapter_Sub = new ArrayAdapter<String>(Excel_Creator.this, android.R.layout.simple_dropdown_item_1line, tinydb.getListString("SUBSTATION"));
                    autoCompleteTextView_SUBSTATION.setAdapter(myAdapter_Sub);

                    ArrayAdapter<String> myAdapter_BAY = new ArrayAdapter<String>(Excel_Creator.this, android.R.layout.simple_dropdown_item_1line, tinydb.getListString("BAY_NO"));
                    autoCompleteTextView_BAY_NO.setAdapter(myAdapter_BAY);

                    ArrayAdapter<String> myAdapter_PNAEL = new ArrayAdapter<String>(Excel_Creator.this, android.R.layout.simple_dropdown_item_1line, tinydb.getListString("PANEL_NU"));
                    autoCompleteTextView_PANEL_NU.setAdapter(myAdapter_PNAEL);

                    ArrayAdapter<String> myAdapter_MU = new ArrayAdapter<String>(Excel_Creator.this, android.R.layout.simple_dropdown_item_1line, tinydb.getListString("METTER_MU"));
                    autoCompleteTextView_METER_MANUFACTURER.setAdapter(myAdapter_MU);

                    ArrayAdapter<String> myAdapter_CT = new ArrayAdapter<String>(Excel_Creator.this, android.R.layout.simple_dropdown_item_1line, tinydb.getListString("CT_RAITIO"));
                    autoCompleteTextView_CT_RATIO.setAdapter(myAdapter_CT);

                    ArrayAdapter<String> myAdapter_PT = new ArrayAdapter<String>(Excel_Creator.this, android.R.layout.simple_dropdown_item_1line, tinydb.getListString("PT_RAITIO"));
                    autoCompleteTextView_PT_RATIO.setAdapter(myAdapter_PT);

                    ArrayAdapter<String> myAdapter_CT_CLASS = new ArrayAdapter<String>(Excel_Creator.this, android.R.layout.simple_dropdown_item_1line, tinydb.getListString("CT_RAITIO_CLASS"));
                    autoCompleteTextView_CT_RATIO_class.setAdapter(myAdapter_CT_CLASS);

                    ArrayAdapter<String> myAdapter_PT_CLASS = new ArrayAdapter<String>(Excel_Creator.this, android.R.layout.simple_dropdown_item_1line, tinydb.getListString("PT_RAITIO_CLASS"));
                    autoCompleteTextView_PT_RATIO_class.setAdapter(myAdapter_PT_CLASS);
//                    Saved = true;
                    Toast.makeText(this, "File Saved :)", Toast.LENGTH_SHORT).show();
//                }
//                else {
//                    Toast.makeText(this, "File Not Saved :(", Toast.LENGTH_SHORT).show();
//                }
            }
            catch (Exception e) {
                e.printStackTrace();
                Toast.makeText(this, "File Not Saved :(", Toast.LENGTH_SHORT).show();
//                SUBSTATION_Text_Input_Layout.setError(null);
//                BAY_NO_Text_Input_Layout.setError("");
//                DATE.setError("");
            }
        }
        else {
//            if (SUBSTATION_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                SUBSTATION_Text_Input_Layout.setError("This Field Is Required");
//            if (BAY_NO_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                BAY_NO_Text_Input_Layout.setError("This Field Is Required");
//            if (DATE.getEditText().getText().toString().trim().isEmpty())
//                DATE.setError("This Field Is Required");
            Toast.makeText(this, "You Should Fill SUBSTATION DATE AND BAY NO Fields For Logical Reasons", Toast.LENGTH_SHORT).show();
        }
    }
    public void createFirstSheet() {
            if (!CT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !PT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty()){
                    try {
                        String[] CT = CT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().split("/");
                        String[] PT = PT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().split("/");
                        //i= /col start , i1=row statrt, i2=col end, i3=row end
                        //Excel sheet name. 0 (number)represents first sheet
                        WritableSheet sheet = workbook.createSheet("METER", 0);
                        // column and row title
                        WritableCellFormat first_row = new WritableCellFormat();
                        Label label = new Label(0, 0, "GUILAN REGIONAL ELECTRIC COMPANY", first_row);
                        first_row.setAlignment(jxl.format.Alignment.CENTRE);
                        first_row.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                        sheet.mergeCells(0, 0, 7, 0);
                        sheet.addCell(label);
                        //?//
                        WritableCellFormat second_row = new WritableCellFormat();
                        Label label1 = new Label(0, 1, "METERING TEST SHEET", second_row);
                        second_row.setAlignment(jxl.format.Alignment.CENTRE);
                        second_row.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                        sheet.mergeCells(0, 1, 7, 1);
                        sheet.addCell(label1);
                        //?//
                        WritableCellFormat border_one = new WritableCellFormat();
                        Label label1_border = new Label(0, 2, "SUBSTATION", border_one);
                        border_one.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                        sheet.mergeCells(0, 2, 1, 2);
                        sheet.addCell(label1_border);
                        //|
                        WritableCellFormat border_one_edit_text = new WritableCellFormat();
                        Label label1_border_edit_text = new Label(2, 2, SUBSTATION_Text_Input_Layout.getEditText().getText().toString(), border_one_edit_text);
                        border_one_edit_text.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                        sheet.mergeCells(2, 2, 3, 2);
                        sheet.addCell(label1_border_edit_text);
                        //|
//                        if (Integer.parseInt(CT[0]) == (int) Integer.parseInt(CT[0]) && Integer.parseInt(CT[1]) == (int) Integer.parseInt(CT[1]) && Integer.parseInt(PT[0]) == (int) Integer.parseInt(PT[0]) && Integer.parseInt(PT[1]) == (int) Integer.parseInt(PT[1])) {
                            //
                            int J3 = Integer.parseInt(CT[0]) / Integer.parseInt(CT[1]);
                            int J4 = Integer.parseInt(PT[0]) / Integer.parseInt(PT[1]);
                            // Error 1
                            double B15 = Double.parseDouble(I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout.getEditText().getText().toString().trim());
                            double B11 = Double.parseDouble(REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout.getEditText().getText().toString().trim());
                            double D15 = Double.parseDouble(I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout.getEditText().getText().toString().trim());
                            double D11 = Double.parseDouble(REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout.getEditText().getText().toString().trim());
                            double F15 = Double.parseDouble(I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout.getEditText().getText().toString().trim());
                            double F11 = Double.parseDouble(REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout.getEditText().getText().toString().trim());
                            //Error 2
                            double F28 = Double.parseDouble(V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim());
                            double F25 = Double.parseDouble(REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim());
                            double D28 = Double.parseDouble(V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout.getEditText().getText().toString().trim());
                            double D25 = Double.parseDouble(REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim());
                            double B28 = Double.parseDouble(V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout.getEditText().getText().toString().trim());
                            double B25 = Double.parseDouble(REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim());
                            //
                            double error_2_1 = Math.abs(((B28*J4-B25)/B25))*100;
                            double error_2_2 = Math.abs((((D28)*J4)-D25)/D25)*100;
                            double error_2_3 = Math.abs((((F28)*J4)-F25)/F25)*100;
                            //
                            double error_1_1 = Math.abs(((B15*J3-B11)/B11))*100;
                            double error_1_2 = Math.abs((((D15)*J3)-D11)/D11)*100;
                            double error_1_3 = Math.abs((((F15)*J3)-F11)/F11)*100;

                            WritableCellFormat border_two = new WritableCellFormat();
                            Label label2_border = new Label(4, 2, "CT RATIO", border_two);
                            border_two.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                            sheet.addCell(label2_border);
                            //|
                            WritableCellFormat border_two_edit_text = new WritableCellFormat();
                            Label label2_border_edit_text = new Label(5, 2, CT_RATIO_Text_Input_Layout.getEditText().getText().toString(), border_two_edit_text);
                            border_two_edit_text.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                            sheet.addCell(label2_border_edit_text);
                            //|
                            WritableCellFormat border_three = new WritableCellFormat();
                            Label label3_border = new Label(6, 2, "CLASS", border_three);
                            border_three.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                            sheet.addCell(label3_border);
                            //|
                            WritableCellFormat border_three_edit_text = new WritableCellFormat();
                            Label label3_border_edit_text = new Label(7, 2, CT_RATIO_CLASS__Text_Input_Layout.getEditText().getText().toString(), border_three_edit_text);
                            border_three_edit_text.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                            sheet.addCell(label3_border_edit_text);
                            //|
                            //?//
                            WritableCellFormat border_four = new WritableCellFormat();
                            Label label4_border = new Label(0, 3, "BAY NO", border_four);
                            border_four.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(0, 3, 1, 3);
                            sheet.addCell(label4_border);
                            //|
                            WritableCellFormat border_four_edit_text = new WritableCellFormat();
                            Label label4_border_edit_text = new Label(2, 3, BAY_NO_Text_Input_Layout.getEditText().getText().toString(), border_four_edit_text);
                            border_four_edit_text.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(2, 3, 3, 3);
                            sheet.addCell(label4_border_edit_text);
                            //|
                            WritableCellFormat border_five = new WritableCellFormat();
                            Label label5_border = new Label(4, 3, "PT RATIO", border_five);
                            border_five.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                            sheet.addCell(label5_border);
                            //|
                            WritableCellFormat border_five_edit_text = new WritableCellFormat();
                            Label label5_border_edit_text = new Label(5, 3, PT_RATIO_Text_Input_Layout.getEditText().getText().toString(), border_five_edit_text);
                            border_five_edit_text.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(2, 2, 3, 2);
                            sheet.addCell(label5_border_edit_text);
                            //|
                            WritableCellFormat border_six = new WritableCellFormat();
                            Label label6_border = new Label(6, 3, "CLASS", border_six);
                            border_six.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                            sheet.addCell(label6_border);
                            //|
                            WritableCellFormat border_six_edit_text = new WritableCellFormat();
                            Label label6_border_edit_text = new Label(7, 3, PT_RATIO_CLASS__Text_Input_Layout.getEditText().getText().toString(), border_six_edit_text);
                            border_six_edit_text.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                            sheet.addCell(label6_border_edit_text);
                            //|
                            //?//
                            sheet.mergeCells(0, 4, 1, 4);
                            WritableCellFormat border_seven = new WritableCellFormat();
                            Label label7_border = new Label(0, 4, "PANEL NO", border_seven);
                            border_seven.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(0, 4, 1, 4);
                            sheet.addCell(label7_border);
                            //|
                            WritableCellFormat border_seven_edit_text = new WritableCellFormat();
                            Label label7_border_edit_text = new Label(2, 4, PANEL_NU_Text_Input_Layout.getEditText().getText().toString(), border_seven_edit_text);
                            border_seven_edit_text.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(2, 4, 3, 4);
                            sheet.addCell(label7_border_edit_text);
                            //|
                            sheet.mergeCells(4, 4, 5, 4);
                            WritableCellFormat border_eight = new WritableCellFormat();
                            Label label8_border = new Label(4, 4, "METER SCALE RANGE", border_eight);
                            border_eight.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                            sheet.addCell(label8_border);
                            //?//
                            //|
                            WritableCellFormat border_eight_edit_text = new WritableCellFormat();
                            Label label8_border_edit_text = new Label(6, 4, METER_SCALE_RANGE_Text_Input_Layout.getEditText().getText().toString(), border_eight_edit_text);
                            border_eight_edit_text.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(6, 4, 7, 4);
                            sheet.addCell(label8_border_edit_text);
                            //|
                            sheet.mergeCells(0, 5, 1, 5);
                            WritableCellFormat border_nine = new WritableCellFormat();
                            Label label9_border = new Label(0, 5, "METER MANUFACTURER", border_nine);
                            border_nine.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                            sheet.addCell(label9_border);
                            //|
                            WritableCellFormat border_nine_edit_text = new WritableCellFormat();
                            Label label9_border_edit_text = new Label(2, 5, METER_MANUFACTURER_Text_Input_Layout.getEditText().getText().toString(), border_nine_edit_text);
                            border_nine_edit_text.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(2, 5, 3, 5);
                            sheet.addCell(label9_border_edit_text);
                            //|
                            sheet.mergeCells(4, 5, 5, 5);
                            WritableCellFormat border_ten = new WritableCellFormat();
                            Label label10_border = new Label(4, 5, "METER MODEL--SERIAL NO", border_ten);
                            border_ten.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                            sheet.addCell(label10_border);
                            //|
                            WritableCellFormat border_ten_edit_text = new WritableCellFormat();
                            Label label10_border_edit_text = new Label(6, 5, METER_MODEL__SERIAL_NO_Text_Input_Layout.getEditText().getText().toString(), border_ten_edit_text);
                            border_ten_edit_text.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(6, 5, 7, 5);
                            sheet.addCell(label10_border_edit_text);
                            //|
                            //?////?//
                            WritableCellFormat six_row = new WritableCellFormat();
                            six_row.setAlignment(jxl.format.Alignment.CENTRE);
                            six_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label3 = new Label(0, 6, "REAL VALUE-I LOAD", six_row);
                            six_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            six_row.setBorder(Border.TOP, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(0, 6, 0, 8);
                            sheet.addCell(label3);
                            //
                            WritableCellFormat seven_row = new WritableCellFormat();
                            seven_row.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label4 = new Label(1, 6, "PHASE A", seven_row);
                            seven_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            seven_row.setBorder(Border.TOP, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(1, 6, 2, 6);
                            sheet.addCell(label4);
                            //|
                            WritableCellFormat border_eleven_edit_text = new WritableCellFormat();
                            border_eleven_edit_text.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label11_border_edit_text = new Label(1, 7, REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout.getEditText().getText().toString(), border_eleven_edit_text);
                            border_eleven_edit_text.setBorder(Border.ALL, BorderLineStyle.THIN);
                            border_eleven_edit_text.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            sheet.mergeCells(1, 7, 2, 8);
                            sheet.addCell(label11_border_edit_text);
                            //|
                            WritableCellFormat eight_row = new WritableCellFormat();
                            eight_row.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label5 = new Label(3, 6, "PH.B", eight_row);
                            eight_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            eight_row.setBorder(Border.TOP, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(3, 6, 4, 6);
                            sheet.addCell(label5);
                            //|
                            WritableCellFormat border_twelve_edit_text = new WritableCellFormat();
                            border_twelve_edit_text.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label12_border_edit_text = new Label(3, 7, REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout.getEditText().getText().toString(), border_twelve_edit_text);
                            border_twelve_edit_text.setBorder(Border.ALL, BorderLineStyle.THIN);
                            border_twelve_edit_text.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            sheet.mergeCells(3, 7, 4, 8);
                            sheet.addCell(label12_border_edit_text);
                            //|
                            WritableCellFormat nineth_row = new WritableCellFormat();
                            nineth_row.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label6 = new Label(5, 6, "PH.C", nineth_row);
                            nineth_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            nineth_row.setBorder(Border.TOP, BorderLineStyle.MEDIUM);
                            nineth_row.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(5, 6, 7, 6);
                            sheet.addCell(label6);
                            //|
                            WritableCellFormat border_thertheen_edit_text = new WritableCellFormat();
                            border_thertheen_edit_text.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label13_border_edit_text = new Label(5, 7, REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout.getEditText().getText().toString(), border_thertheen_edit_text);
                            border_thertheen_edit_text.setBorder(Border.ALL, BorderLineStyle.THIN);
                            border_thertheen_edit_text.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
                            border_thertheen_edit_text.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            sheet.mergeCells(5, 7, 7, 8);
                            sheet.addCell(label13_border_edit_text);
                            //|
                            //?//
                            WritableCellFormat ten_row = new WritableCellFormat();
                            ten_row.setAlignment(jxl.format.Alignment.CENTRE);
                            ten_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label7 = new Label(0, 9, "I ONLOAD by clamp", ten_row);
                            ten_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(0, 9, 0, 12);
                            sheet.addCell(label7);
                            //
                            WritableCellFormat eleven_row = new WritableCellFormat();
                            eleven_row.setAlignment(jxl.format.Alignment.CENTRE);
                            eleven_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label8 = new Label(1, 9, "PH.A", eleven_row);
                            eleven_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(1, 9, 2, 10);
                            sheet.addCell(label8);
                            //|
                            WritableCellFormat border_fourteen_edit_text = new WritableCellFormat();
                            border_fourteen_edit_text.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label14_border_edit_text = new Label(1, 11, I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout.getEditText().getText().toString(), border_fourteen_edit_text);
                            border_fourteen_edit_text.setBorder(Border.ALL, BorderLineStyle.THIN);
                            border_fourteen_edit_text.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            sheet.mergeCells(1, 11, 2, 12);
                            sheet.addCell(label14_border_edit_text);
                            //|
                            WritableCellFormat twoelve_row = new WritableCellFormat();
                            twoelve_row.setAlignment(jxl.format.Alignment.CENTRE);
                            twoelve_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label9 = new Label(3, 9, "PH.B", twoelve_row);
                            twoelve_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(3, 9, 4, 10);
                            sheet.addCell(label9);
                            //|
                            WritableCellFormat border_fifteen_edit_text = new WritableCellFormat();
                            border_fifteen_edit_text.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label15_border_edit_text = new Label(3, 11, I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout.getEditText().getText().toString(), border_fifteen_edit_text);
                            border_fifteen_edit_text.setBorder(Border.ALL, BorderLineStyle.THIN);
                            border_fifteen_edit_text.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            sheet.mergeCells(3, 11, 4, 12);
                            sheet.addCell(label15_border_edit_text);
                            //|
                            WritableCellFormat thirteen_row = new WritableCellFormat();
                            thirteen_row.setAlignment(jxl.format.Alignment.CENTRE);
                            thirteen_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label10 = new Label(5, 9, "PH.C", thirteen_row);
                            thirteen_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            thirteen_row.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(5, 9, 7, 10);
                            sheet.addCell(label10);
                            //|
                            WritableCellFormat border_sixteen_edit_text = new WritableCellFormat();
                            border_sixteen_edit_text.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label16_border_edit_text = new Label(5, 11, I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout.getEditText().getText().toString(), border_sixteen_edit_text);
                            border_sixteen_edit_text.setBorder(Border.ALL, BorderLineStyle.THIN);
                            border_sixteen_edit_text.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
                            border_sixteen_edit_text.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            sheet.mergeCells(5, 11, 7, 12);
                            sheet.addCell(label16_border_edit_text);
                            //|
                            WritableCellFormat fourteen_row = new WritableCellFormat();
                            fourteen_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label11 = new Label(0, 13, "%ERROR", fourteen_row);
                            fourteen_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(0, 13, 0, 14);
                            sheet.addCell(label11);
                            //|
                            //=ABS(((B15*J3-B11)/B11))*100
                            WritableCellFormat border_seventeen_edit_text = new WritableCellFormat();
                            border_seventeen_edit_text.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label17_border_edit_text = new Label(1, 13, new DecimalFormat("##.###").format(error_1_1), border_seventeen_edit_text);
                            border_seventeen_edit_text.setBorder(Border.ALL, BorderLineStyle.THIN);
                            border_seventeen_edit_text.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            sheet.mergeCells(1, 13, 2, 14);
                            sheet.addCell(label17_border_edit_text);
                            //|
                            //=ABS((((D15)*J3)-D11)/D11)*100
                            WritableCellFormat border_seventeen_sub_edit_text = new WritableCellFormat();
                            border_seventeen_sub_edit_text.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label17_sub_border_edit_text = new Label(3, 13, new DecimalFormat("##.###").format(error_1_2), border_seventeen_sub_edit_text);
                            border_seventeen_sub_edit_text.setBorder(Border.ALL, BorderLineStyle.THIN);
                            border_seventeen_sub_edit_text.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            sheet.mergeCells(3, 13, 4, 14);
                            sheet.addCell(label17_sub_border_edit_text);
                            //|
                            //=ABS((((F15)*J3)-F11)/F11)*100
                            WritableCellFormat border_seventeen_sub3_edit_text = new WritableCellFormat();
                            border_seventeen_sub3_edit_text.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label17_border_sub_3_edit_text = new Label(5, 13, new DecimalFormat("##.###").format(error_1_3), border_seventeen_sub3_edit_text);
                            border_seventeen_sub3_edit_text.setBorder(Border.ALL, BorderLineStyle.THIN);
                            border_seventeen_sub3_edit_text.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
                            border_seventeen_sub3_edit_text.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            sheet.mergeCells(5, 13, 7, 14);
                            sheet.addCell(label17_border_sub_3_edit_text);
                            //|
                            WritableCellFormat fifteen_row = new WritableCellFormat();
                            fifteen_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label12 = new Label(0, 15, "%ERORE IN RANGE", fifteen_row);
                            fifteen_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(0, 15, 0, 16);
                            sheet.addCell(label12);
                            //|
                            WritableCellFormat border_eightteen_edit_text = new WritableCellFormat();
                            border_eightteen_edit_text.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label18_border_edit_text = new Label(1, 15, "", border_eightteen_edit_text);
                            border_eightteen_edit_text.setBorder(Border.ALL, BorderLineStyle.THIN);
                            border_eightteen_edit_text.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            sheet.mergeCells(1, 15, 2, 16);
                            sheet.addCell(label18_border_edit_text);
                            //|
                            WritableCellFormat border_eightteen_sub_edit_text = new WritableCellFormat();
                            border_eightteen_sub_edit_text.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label18_border_sub_edit_text = new Label(3, 15, "", border_eightteen_sub_edit_text);
                            border_eightteen_sub_edit_text.setBorder(Border.ALL, BorderLineStyle.THIN);
                            border_eightteen_sub_edit_text.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            sheet.mergeCells(3, 15, 4, 16);
                            sheet.addCell(label18_border_sub_edit_text);
                            //|
                            WritableCellFormat border_eightteen_sub_2_edit_text = new WritableCellFormat();
                            border_eightteen_sub_2_edit_text.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label18_border_sub2_edit_text = new Label(5, 15, "", border_eightteen_sub_2_edit_text);
                            border_eightteen_sub_2_edit_text.setBorder(Border.ALL, BorderLineStyle.THIN);
                            border_eightteen_sub_2_edit_text.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
                            border_eightteen_sub_2_edit_text.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            sheet.mergeCells(5, 15, 7, 16);
                            sheet.addCell(label18_border_sub2_edit_text);
                            //|
                            WritableCellFormat sixteen_row = new WritableCellFormat();
                            sixteen_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label13 = new Label(0, 17, "%EROR OUT OF RANGE", sixteen_row);
                            sixteen_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(0, 17, 0, 18);
                            sheet.addCell(label13);
                            //|
                            WritableCellFormat border_nineteen_sub1_edit_text = new WritableCellFormat();
                            border_nineteen_sub1_edit_text.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label19_border_sub1_edit_text = new Label(1, 17, "", border_nineteen_sub1_edit_text);
                            border_nineteen_sub1_edit_text.setBorder(Border.ALL, BorderLineStyle.THIN);
                            border_nineteen_sub1_edit_text.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            sheet.mergeCells(1, 17, 2, 18);
                            sheet.addCell(label19_border_sub1_edit_text);
                            //|
                            WritableCellFormat border_nineteen_sub2_edit_text = new WritableCellFormat();
                            border_nineteen_sub2_edit_text.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label19_border_sub2_edit_text = new Label(3, 17, "", border_nineteen_sub2_edit_text);
                            border_nineteen_sub2_edit_text.setBorder(Border.ALL, BorderLineStyle.THIN);
                            border_nineteen_sub2_edit_text.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            sheet.mergeCells(3, 17, 4, 18);
                            sheet.addCell(label19_border_sub2_edit_text);
                            //|
                            WritableCellFormat border_nineteen_sub3_edit_text = new WritableCellFormat();
                            border_nineteen_sub3_edit_text.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label19_border_sub3_edit_text = new Label(5, 17, "", border_nineteen_sub3_edit_text);
                            border_nineteen_sub3_edit_text.setBorder(Border.ALL, BorderLineStyle.THIN);
                            border_nineteen_sub3_edit_text.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
                            border_nineteen_sub3_edit_text.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            sheet.mergeCells(5, 17, 7, 18);
                            sheet.addCell(label19_border_sub3_edit_text);
                            //?////?//
                            WritableCellFormat seventeen_row = new WritableCellFormat();
                            seventeen_row.setAlignment(jxl.format.Alignment.CENTRE);
                            seventeen_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label14 = new Label(0, 19, "REAL VALUE-V LOAD", seventeen_row);
                            seventeen_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            seventeen_row.setBorder(Border.TOP, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(0, 19, 0, 21);
                            sheet.addCell(label14);
                            //
                            WritableCellFormat eighteen_row = new WritableCellFormat();
                            eighteen_row.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label15 = new Label(1, 19, "PH A", eighteen_row);
                            eighteen_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            eighteen_row.setBorder(Border.TOP, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(1, 19, 2, 19);
                            sheet.addCell(label15);
                            //|
                            WritableCellFormat border_twenty_sub1_edit_text = new WritableCellFormat();
                            border_twenty_sub1_edit_text.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label20_border_sub1_edit_text = new Label(1, 20, REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout.getEditText().getText().toString(), border_twenty_sub1_edit_text);
                            border_twenty_sub1_edit_text.setBorder(Border.ALL, BorderLineStyle.THIN);
                            border_twenty_sub1_edit_text.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            sheet.mergeCells(1, 20, 2, 21);
                            sheet.addCell(label20_border_sub1_edit_text);
                            //|
                            WritableCellFormat ninteen_row = new WritableCellFormat();
                            ninteen_row.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label16 = new Label(3, 19, "PH.B", ninteen_row);
                            ninteen_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            ninteen_row.setBorder(Border.TOP, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(3, 19, 4, 19);
                            sheet.addCell(label16);
                            //|
                            WritableCellFormat border_twenty_sub2_edit_text = new WritableCellFormat();
                            border_twenty_sub2_edit_text.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label20_border_sub2_edit_text = new Label(3, 20, REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout.getEditText().getText().toString(), border_twenty_sub2_edit_text);
                            border_twenty_sub2_edit_text.setBorder(Border.ALL, BorderLineStyle.THIN);
                            border_twenty_sub2_edit_text.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            sheet.mergeCells(3, 20, 4, 21);
                            sheet.addCell(label20_border_sub2_edit_text);
                            //|
                            WritableCellFormat twenty_row = new WritableCellFormat();
                            twenty_row.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label17 = new Label(5, 19, "PH. C", twenty_row);
                            twenty_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            twenty_row.setBorder(Border.TOP, BorderLineStyle.MEDIUM);
                            twenty_row.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(5, 19, 7, 19);
                            sheet.addCell(label17);
                            //|
                            WritableCellFormat border_twenty_sub3_edit_text = new WritableCellFormat();
                            border_twenty_sub3_edit_text.setAlignment(jxl.format.Alignment.CENTRE);
                            Label label20_border_sub3_edit_text = new Label(5, 20, REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString(), border_twenty_sub3_edit_text);
                            border_twenty_sub3_edit_text.setBorder(Border.ALL, BorderLineStyle.THIN);
                            border_twenty_sub3_edit_text.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
                            border_twenty_sub3_edit_text.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            sheet.mergeCells(5, 20, 7, 21);
                            sheet.addCell(label20_border_sub3_edit_text);
                            //?//
                            WritableCellFormat twenty_one_row = new WritableCellFormat();
                            twenty_one_row.setAlignment(jxl.format.Alignment.CENTRE);
                            twenty_one_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label18 = new Label(0, 22, "V ONLOAD BY V.M-V LOAD", twenty_one_row);
                            twenty_one_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(0, 22, 0, 25);
                            sheet.addCell(label18);
                            //
                            WritableCellFormat twenty_two_row = new WritableCellFormat();
                            twenty_two_row.setAlignment(jxl.format.Alignment.CENTRE);
                            twenty_two_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label19 = new Label(1, 22, "PH.A", twenty_two_row);
                            twenty_two_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(1, 22, 2, 23);
                            sheet.addCell(label19);
                            //|
                            WritableCellFormat twenty_two_row_sub1 = new WritableCellFormat();
                            twenty_two_row_sub1.setAlignment(jxl.format.Alignment.CENTRE);
                            twenty_two_row_sub1.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label19_sub = new Label(1, 24, V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout.getEditText().getText().toString(), twenty_two_row_sub1);
                            twenty_two_row_sub1.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(1, 24, 2, 25);
                            sheet.addCell(label19_sub);
                            //|
                            WritableCellFormat twenty_three_row = new WritableCellFormat();
                            twenty_three_row.setAlignment(jxl.format.Alignment.CENTRE);
                            twenty_three_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label20 = new Label(3, 22, "PH.B", twenty_three_row);
                            twenty_three_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(3, 22, 4, 23);
                            sheet.addCell(label20);
                            //|
                            WritableCellFormat twenty_two_row_sub2 = new WritableCellFormat();
                            twenty_two_row_sub2.setAlignment(jxl.format.Alignment.CENTRE);
                            twenty_two_row_sub2.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label19_sub2 = new Label(3, 24, V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout.getEditText().getText().toString(), twenty_two_row_sub2);
                            twenty_two_row_sub2.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(3, 24, 4, 25);
                            sheet.addCell(label19_sub2);
                            //|
                            WritableCellFormat twenty_four_row = new WritableCellFormat();
                            twenty_four_row.setAlignment(jxl.format.Alignment.CENTRE);
                            twenty_four_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label21 = new Label(5, 22, "PH.C", twenty_four_row);
                            twenty_four_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            twenty_four_row.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(5, 22, 7, 23);
                            sheet.addCell(label21);
                            //|
                            WritableCellFormat twenty_two_row_sub3 = new WritableCellFormat();
                            twenty_two_row_sub3.setAlignment(jxl.format.Alignment.CENTRE);
                            twenty_two_row_sub3.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label19_sub3 = new Label(5, 24, V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString(), twenty_two_row_sub3);
                            twenty_two_row_sub3.setBorder(Border.ALL, BorderLineStyle.THIN);
                            twenty_two_row_sub3.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(5, 24, 7, 25);
                            sheet.addCell(label19_sub3);
                            //|
                            WritableCellFormat twenty_five_row = new WritableCellFormat();
                            twenty_five_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label22 = new Label(0, 26, "%ERROR", twenty_five_row);
                            twenty_five_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(0, 26, 0, 27);
                            sheet.addCell(label22);
                            //|
                            //=ABS(((B28*J4-B25)/B25))*100
                            WritableCellFormat twenty_five_row_sub1 = new WritableCellFormat();
                            twenty_five_row_sub1.setAlignment(jxl.format.Alignment.CENTRE);
                            twenty_five_row_sub1.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label25_sub1 = new Label(1, 26, new DecimalFormat("##.###").format(error_2_1), twenty_five_row_sub1);
                            twenty_five_row_sub1.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(1, 26, 2, 27);
                            sheet.addCell(label25_sub1);
                            //|
                            //
                            //=ABS((((D28)*J4)-D25)/D25)*100
                            WritableCellFormat twenty_five_row_sub2 = new WritableCellFormat();
                            twenty_five_row_sub2.setAlignment(jxl.format.Alignment.CENTRE);
                            twenty_five_row_sub2.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label25_sub2 = new Label(3, 26, new DecimalFormat("##.###").format(error_2_2) , twenty_five_row_sub2);
                            twenty_five_row_sub2.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(3, 26, 4, 27);
                            sheet.addCell(label25_sub2);
                            //|
                            //=ABS((((F28)*J4)-F25)/F25)*100
                            WritableCellFormat twenty_five_row_sub3 = new WritableCellFormat();
                            twenty_five_row_sub3.setAlignment(jxl.format.Alignment.CENTRE);
                            twenty_five_row_sub3.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label25_sub3 = new Label(5, 26,new DecimalFormat("##.###").format(error_2_3), twenty_five_row_sub3);
                            twenty_five_row_sub3.setBorder(Border.ALL, BorderLineStyle.THIN);
                            twenty_five_row_sub3.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(5, 26, 7, 27);
                            sheet.addCell(label25_sub3);
                            //|
                            WritableCellFormat twenty_six_row = new WritableCellFormat();
                            twenty_six_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label23 = new Label(0, 28, "%ERORE IN RANGE", twenty_six_row);
                            twenty_six_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(0, 28, 0, 29);
                            sheet.addCell(label23);
                            //|
                            WritableCellFormat twenty_six_row_sub1 = new WritableCellFormat();
                            twenty_six_row_sub1.setAlignment(jxl.format.Alignment.CENTRE);
                            twenty_six_row_sub1.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label26_sub1 = new Label(1, 28, "", twenty_six_row_sub1);
                            twenty_six_row_sub1.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(1, 28, 2, 29);
                            sheet.addCell(label26_sub1);
                            //|
                            WritableCellFormat twenty_six_row_sub2 = new WritableCellFormat();
                            twenty_six_row_sub2.setAlignment(jxl.format.Alignment.CENTRE);
                            twenty_six_row_sub2.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label26_sub2 = new Label(3, 28, "", twenty_six_row_sub2);
                            twenty_six_row_sub2.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(3, 28, 4, 29);
                            sheet.addCell(label26_sub2);
                            //|
                            WritableCellFormat twenty_five_six_row_sub3 = new WritableCellFormat();
                            twenty_five_six_row_sub3.setAlignment(jxl.format.Alignment.CENTRE);
                            twenty_five_six_row_sub3.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label26_sub3 = new Label(5, 28, "", twenty_five_six_row_sub3);
                            twenty_five_six_row_sub3.setBorder(Border.ALL, BorderLineStyle.THIN);
                            twenty_five_six_row_sub3.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(5, 28, 7, 29);
                            sheet.addCell(label26_sub3);
                            //|
                            WritableCellFormat twenty_seven_row = new WritableCellFormat();
                            twenty_seven_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label24 = new Label(0, 30, "%ERROR OUT OF RANGE", twenty_seven_row);
                            twenty_seven_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                            twenty_seven_row.setBorder(Border.BOTTOM, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(0, 30, 0, 31);
                            sheet.addCell(label24);
                            //|
                            WritableCellFormat twenty_seven_row_sub1 = new WritableCellFormat();
                            twenty_seven_row_sub1.setAlignment(jxl.format.Alignment.CENTRE);
                            twenty_seven_row_sub1.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label27_sub1 = new Label(1, 30, "", twenty_seven_row_sub1);
                            twenty_seven_row_sub1.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(1, 30, 2, 31);
                            sheet.addCell(label27_sub1);
                            //|
                            WritableCellFormat twenty_seven_row_sub2 = new WritableCellFormat();
                            twenty_seven_row_sub2.setAlignment(jxl.format.Alignment.CENTRE);
                            twenty_seven_row_sub2.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label27_sub2 = new Label(3, 30, "", twenty_seven_row_sub2);
                            twenty_seven_row_sub2.setBorder(Border.ALL, BorderLineStyle.THIN);
                            sheet.mergeCells(3, 30, 4, 31);
                            sheet.addCell(label27_sub2);
                            //|
                            WritableCellFormat twenty_seven_row_sub3 = new WritableCellFormat();
                            twenty_seven_row_sub3.setAlignment(jxl.format.Alignment.CENTRE);
                            twenty_seven_row_sub3.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                            Label label27_sub3 = new Label(5, 30, "", twenty_seven_row_sub3);
                            twenty_seven_row_sub3.setBorder(Border.ALL, BorderLineStyle.THIN);
                            twenty_seven_row_sub3.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
                            sheet.mergeCells(5, 30, 7, 31);
                            sheet.addCell(label27_sub3);
//                        }
                        //?//
                        WritableCellFormat twenty_eight_row = new WritableCellFormat();
                        twenty_eight_row.setAlignment(jxl.format.Alignment.CENTRE);
                        twenty_eight_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                        twenty_eight_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                        Label label25 = new Label(0, 32, "ACTIVE POWER", twenty_eight_row);
                        sheet.mergeCells(0, 32, 0, 33);
                        sheet.addCell(label25);
                        //|
                        WritableCellFormat twenty_nine_row_sub = new WritableCellFormat();
                        twenty_nine_row_sub.setAlignment(jxl.format.Alignment.CENTRE);
                        twenty_nine_row_sub.setBorder(Border.ALL, BorderLineStyle.THIN);
                        twenty_nine_row_sub.setBorder(Border.TOP, BorderLineStyle.MEDIUM);
                        twenty_nine_row_sub.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                        Label label29_sub = new Label(1, 32, ACTIVE_POWER.getEditText().getText().toString(), twenty_nine_row_sub);
                        sheet.mergeCells(1, 32, 2, 33);
                        sheet.addCell(label29_sub);
                        //|
                        WritableCellFormat Thirty_row_sub = new WritableCellFormat();
                        Thirty_row_sub.setAlignment(jxl.format.Alignment.CENTRE);
                        Thirty_row_sub.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                        Thirty_row_sub.setBorder(Border.ALL, BorderLineStyle.THIN);
                        Label label26 = new Label(0, 34, "COS PHI", Thirty_row_sub);
                        sheet.mergeCells(0, 34, 0, 35);
                        sheet.addCell(label26);
                        //|
                        WritableCellFormat Thirty_One_row_sub = new WritableCellFormat();
                        Thirty_One_row_sub.setAlignment(jxl.format.Alignment.CENTRE);
                        Thirty_One_row_sub.setBorder(Border.ALL, BorderLineStyle.THIN);
                        Thirty_One_row_sub.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                        Label label31_sub = new Label(1, 34, COS_PHI.getEditText().getText().toString(), Thirty_One_row_sub);
                        sheet.mergeCells(1, 34, 2, 35);
                        sheet.addCell(label31_sub);
                        //|
                        WritableCellFormat thirty_row = new WritableCellFormat();
                        thirty_row.setAlignment(jxl.format.Alignment.CENTRE);
                        thirty_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                        thirty_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                        thirty_row.setBorder(Border.BOTTOM, BorderLineStyle.MEDIUM);
                        Label label27 = new Label(0, 36, "TEST BY", thirty_row);
                        sheet.mergeCells(0, 36, 0, 37);
                        sheet.addCell(label27);
                        //|
                        WritableCellFormat Thirty_Two_row_sub = new WritableCellFormat();
                        Thirty_Two_row_sub.setAlignment(jxl.format.Alignment.CENTRE);
                        Thirty_Two_row_sub.setBorder(Border.ALL, BorderLineStyle.THIN);
                        Thirty_Two_row_sub.setBorder(Border.BOTTOM, BorderLineStyle.MEDIUM);
                        Thirty_Two_row_sub.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                        Label label32_sub = new Label(1, 36, TEST_BY.getEditText().getText().toString(), Thirty_Two_row_sub);
                        sheet.mergeCells(1, 36, 2, 37);
                        sheet.addCell(label32_sub);
                        ////
                        WritableCellFormat thirty_one_row = new WritableCellFormat();
                        thirty_one_row.setAlignment(jxl.format.Alignment.CENTRE);
                        thirty_one_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                        thirty_one_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                        thirty_one_row.setBorder(Border.TOP, BorderLineStyle.MEDIUM);
                        Label label28 = new Label(3, 32, "REACTIVE POWER ", thirty_one_row);
                        sheet.mergeCells(3, 32, 4, 33);
                        sheet.addCell(label28);
                        //|
                        WritableCellFormat Thirty_Three_row_sub = new WritableCellFormat();
                        Thirty_Three_row_sub.setAlignment(jxl.format.Alignment.CENTRE);
                        Thirty_Three_row_sub.setBorder(Border.ALL, BorderLineStyle.THIN);
                        Thirty_Three_row_sub.setBorder(Border.TOP, BorderLineStyle.MEDIUM);
                        Thirty_Three_row_sub.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
                        Thirty_Three_row_sub.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                        Label label33_sub = new Label(5, 32, REACTIVE_POWER.getEditText().getText().toString(), Thirty_Three_row_sub);
                        sheet.mergeCells(5, 32, 7, 33);
                        sheet.addCell(label33_sub);
                        ////
                        WritableCellFormat thirty_two_row = new WritableCellFormat();
                        thirty_two_row.setAlignment(jxl.format.Alignment.CENTRE);
                        thirty_two_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                        Label label29 = new Label(3, 34, "COMMENT", thirty_two_row);
                        thirty_two_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                        sheet.mergeCells(3, 34, 4, 35);
                        sheet.addCell(label29);
                        //|
                        WritableCellFormat Thirty_Four_row_sub = new WritableCellFormat();
                        Thirty_Four_row_sub.setAlignment(jxl.format.Alignment.CENTRE);
                        Thirty_Four_row_sub.setBorder(Border.ALL, BorderLineStyle.THIN);
                        Thirty_Four_row_sub.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
                        Thirty_Four_row_sub.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                        Label label34_sub = new Label(5, 34, COMMENT.getEditText().getText().toString(), Thirty_Four_row_sub);
                        sheet.mergeCells(5, 34, 7, 35);
                        sheet.addCell(label34_sub);
                        ////
                        WritableCellFormat thirty_three_row = new WritableCellFormat();
                        thirty_three_row.setAlignment(jxl.format.Alignment.CENTRE);
                        thirty_three_row.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                        thirty_three_row.setBorder(Border.ALL, BorderLineStyle.THIN);
                        thirty_three_row.setBorder(Border.BOTTOM, BorderLineStyle.MEDIUM);
                        Label label30 = new Label(3, 36, "DATE", thirty_three_row);
                        sheet.mergeCells(3, 36, 4, 37);
                        sheet.addCell(label30);
                        //|
                        WritableCellFormat Thirty_Five_row_sub = new WritableCellFormat();
                        Thirty_Five_row_sub.setAlignment(jxl.format.Alignment.CENTRE);
                        Thirty_Five_row_sub.setBorder(Border.ALL, BorderLineStyle.THIN);
                        Thirty_Five_row_sub.setBorder(Border.BOTTOM, BorderLineStyle.MEDIUM);
                        Thirty_Five_row_sub.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
                        Thirty_Five_row_sub.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
                        Label label35_sub = new Label(5, 36, DATE.getEditText().getText().toString(), Thirty_Five_row_sub);
                        sheet.mergeCells(5, 36, 7, 37);
                        sheet.addCell(label35_sub);
                        ////
                        addCellView(sheet, 0, 24);
                        addCellView(sheet, 4, 20);


                        if (ch.isChecked()){
                            Document document = new Document();
                            String mFileName = SUBSTATION_Text_Input_Layout.getEditText().getText().toString() + "_" + BAY_NO_Text_Input_Layout.getEditText().getText().toString() +"_"+ "metter" +"_"+ DATE.getEditText().getText().toString();
                            String mFilePath = Environment.getExternalStorageDirectory().getPath() + "/Electricity/PDF/"+ mFileName + ".pdf";

//                            int J3 = Integer.parseInt(CT[0]) / Integer.parseInt(CT[1]);
//                            int J4 = Integer.parseInt(PT[0]) / Integer.parseInt(PT[1]);
                            String i = CT[0] + CT[1];
                            String i1 = PT[0] + PT[1];
                            // Error 1
//                            double B15 = Double.parseDouble(I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout.getEditText().getText().toString().trim());
//                            double B11 = Double.parseDouble(REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout.getEditText().getText().toString().trim());
//                            double D15 = Double.parseDouble(I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout.getEditText().getText().toString().trim());
//                            double D11 = Double.parseDouble(REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout.getEditText().getText().toString().trim());
//                            double F15 = Double.parseDouble(I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout.getEditText().getText().toString().trim());
//                            double F11 = Double.parseDouble(REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout.getEditText().getText().toString().trim());
                            //Error 2
//                            double F28 = Double.parseDouble(V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim());
//                            double F25 = Double.parseDouble(REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim());
//                            double D28 = Double.parseDouble(V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout.getEditText().getText().toString().trim());
//                            double D25 = Double.parseDouble(REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim());
//                            double B28 = Double.parseDouble(V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout.getEditText().getText().toString().trim());
//                            double B25 = Double.parseDouble(REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim());
                            //
//                            double error_2_1 = Math.abs(((B28*J4-B25)/B25))*100;
//                            double error_2_2 = Math.abs((((D28)*J4)-D25)/D25)*100;
//                            double error_2_3 = Math.abs((((F28)*J4)-F25)/F25)*100;
                            //
//                            double error_1_1 = Math.abs(((B15*J3-B11)/B11))*100;
//                            double error_1_2 = Math.abs((((D15)*J3)-D11)/D11)*100;
//                            double error_1_3 = Math.abs((((F15)*J3)-F11)/F11)*100;
//                            try {
                                PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(mFilePath));
                                document.open();

                                PdfPTable table = new PdfPTable(8);
                                table.setWidthPercentage(100);
                                table.getDefaultCell().setUseAscender(true);
                                table.getDefaultCell().setUseDescender(true);

                                float[] columnWidths = new float[]{25f, 8f, 10f, 10f, 17f, 10f, 10f,10f};
                                table.setWidths(columnWidths);

                                Font f = new Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL, GrayColor.BLACK);
                                Font f_1 = new Font(Font.FontFamily.HELVETICA, 8, Font.NORMAL, GrayColor.BLACK);
                                PdfPCell cell1 = new PdfPCell(new Phrase("GUILAN REGIONAL ELECTRIC COMPANY", f));
                                cell1.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell1.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell1.setColspan(8);
                                cell1.setPaddingBottom(4);
                                cell1.setBorderWidth(2);
                                table.addCell(cell1);

                                table.setTableEvent(new BorderEvent());


                                PdfPCell cell2 = new PdfPCell(new Phrase("METERING TEST SHEET", f));
                                cell2.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell2.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell2.setColspan(8);
                                cell2.setPaddingBottom(4);
                                cell2.setBorderWidth(2);
                                table.addCell(cell2);

                                //Row 3
                                PdfPCell cell3 = new PdfPCell(new Phrase("SUBSTATION",f));
                                cell3.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell3.setColspan(2);
                                cell3.setBorderWidth(2);
                                cell3.setFixedHeight(30f);
                                table.addCell(cell3);

                                PdfPCell cell4 = new PdfPCell(new Phrase(SUBSTATION_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell4.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell4.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell4.setColspan(2);
                                cell4.setBorderWidth(2);
                                table.addCell(cell4);

                                PdfPCell cell5 = new PdfPCell(new Phrase("CT RATIO",f));
                                cell5.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell5.setColspan(1);
                                cell5.setBorderWidth(2);
                                table.addCell(cell5);

                                PdfPCell cell6 = new PdfPCell(new Phrase(CT_RATIO_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell6.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell6.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell6.setColspan(1);
                                cell6.setBorderWidth(2);
                                table.addCell(cell6);

                                PdfPCell cell7 = new PdfPCell(new Phrase("CLASS",f));
                                cell7.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell7.setColspan(1);
                                cell7.setBorderWidth(2);
                                table.addCell(cell7);

                                PdfPCell cell8 = new PdfPCell(new Phrase(CT_RATIO_CLASS__Text_Input_Layout.getEditText().getText().toString(),f));
                                cell8.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell8.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell8.setColspan(1);
                                cell8.setBorderWidth(2);
                                table.addCell(cell8);

                                //Row 4
                                PdfPCell cell9 = new PdfPCell(new Phrase("BAY NO",f));
                                cell9.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell9.setColspan(2);
                                cell9.setBorderWidth(2);
                                cell9.setFixedHeight(30f);
                                table.addCell(cell9);

                                PdfPCell cell10 = new PdfPCell(new Phrase(BAY_NO_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell10.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell10.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell10.setColspan(2);
                                cell10.setBorderWidth(2);
                                table.addCell(cell10);

                                PdfPCell cell11 = new PdfPCell(new Phrase("PT RATIO",f));
                                cell11.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell11.setBorderWidth(2);
                                table.addCell(cell11);

                                PdfPCell cell12 = new PdfPCell(new Phrase(PT_RATIO_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell12.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell12.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell12.setBorderWidth(2);
                                table.addCell(cell12);

                                PdfPCell cell13 = new PdfPCell(new Phrase("CLASS",f));
                                cell13.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell13.setBorderWidth(2);
                                table.addCell(cell13);

                                PdfPCell cell14 = new PdfPCell(new Phrase(PT_RATIO_CLASS__Text_Input_Layout.getEditText().getText().toString(),f));
                                cell14.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell14.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell14.setBorderWidth(2);
                                table.addCell(cell14);

                                //Row 5
                                PdfPCell cell15 = new PdfPCell(new Phrase("PANEL NO",f));
                                cell15.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell15.setBorderWidth(2);
                                cell15.setColspan(2);
                                cell15.setFixedHeight(30f);
                                table.addCell(cell15);

                                PdfPCell cell16 = new PdfPCell(new Phrase(PANEL_NU_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell16.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell16.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell16.setBorderWidth(2);
                                cell16.setColspan(2);
                                table.addCell(cell16);

                                PdfPCell cell17 = new PdfPCell(new Phrase("METER SCALE RANGE",f));
                                cell17.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell17.setBorderWidth(2);
                                cell17.setColspan(2);
                                table.addCell(cell17);

                                PdfPCell cell18 = new PdfPCell(new Phrase(METER_SCALE_RANGE_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell18.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell18.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell18.setBorderWidth(2);
                                cell18.setColspan(2);
                                table.addCell(cell18);

                                //Row 5
                                PdfPCell cell19 = new PdfPCell(new Phrase("METER MANUFACTURER ",f));
                                cell19.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell19.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell19.setFixedHeight(30f);
                                cell19.setBorderWidth(2);
                                cell19.setColspan(2);
                                table.addCell(cell19);


                                PdfPCell cell20 = new PdfPCell(new Phrase(METER_MANUFACTURER_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell20.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell20.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell20.setBorderWidth(2);
                                cell20.setColspan(2);
                                table.addCell(cell20);

                                PdfPCell cell21 = new PdfPCell(new Phrase("METER MODEL--SERIAL NO",f));
                                cell21.setBorderWidth(2);
                                cell21.setColspan(2);
                                cell21.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell21.setHorizontalAlignment(Element.ALIGN_CENTER);
                                table.addCell(cell21);

                                PdfPCell cell22 = new PdfPCell(new Phrase(METER_MODEL__SERIAL_NO_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell22.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell22.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell22.setBorderWidth(2);
                                cell22.setColspan(2);
                                table.addCell(cell22);

                                // Row 6
                                PdfPCell cell23 = new PdfPCell(new Phrase("REAL VALUE-I LOAD",f));
                                cell23.setFixedHeight(70f);
                                cell23.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell23.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell23.setColspan(1);
                                cell23.setRowspan(3);
                                table.addCell(cell23);

                                PdfPCell cell24 = new PdfPCell(new Phrase("PHASE A",f));
                                cell24.setFixedHeight(35f);
                                cell24.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell24.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell24.setColspan(2);
                                cell24.setRowspan(1);
                                table.addCell(cell24);

                                PdfPCell cell25 = new PdfPCell(new Phrase("PH.B",f));
                                cell25.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell25.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell25.setColspan(2);
                                cell25.setRowspan(1);
                                table.addCell(cell25);

                                PdfPCell cell26 = new PdfPCell(new Phrase("PH.C",f));
                                cell26.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell26.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell26.setColspan(3);
                                cell26.setRowspan(1);
                                table.addCell(cell26);
                                //
                                PdfPCell cell27 = new PdfPCell(new Phrase(REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell27.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell27.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell27.setColspan(2);
                                cell27.setRowspan(2);
                                table.addCell(cell27);

                                PdfPCell cell28 = new PdfPCell(new Phrase(REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell28.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell28.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell28.setColspan(2);
                                cell28.setRowspan(2);
                                table.addCell(cell28);

                                PdfPCell cell29 = new PdfPCell(new Phrase(REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell29.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell29.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell29.setColspan(3);
                                cell29.setRowspan(2);
                                table.addCell(cell29);
                                //Row 9
                                PdfPCell cell30 = new PdfPCell(new Phrase("I ONLOAD by clamp",f));
                                cell30.setFixedHeight(70f);
                                cell30.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell30.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell30.setColspan(1);
                                cell30.setRowspan(4);
                                table.addCell(cell30);

                                PdfPCell cell31 = new PdfPCell(new Phrase("PH.A",f));
                                cell31.setFixedHeight(35f);
                                cell31.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell31.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell31.setColspan(2);
                                cell31.setRowspan(2);
                                table.addCell(cell31);

                                PdfPCell cell32 = new PdfPCell(new Phrase("PH.B",f));
                                cell32.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell32.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell32.setColspan(2);
                                cell32.setRowspan(2);
                                table.addCell(cell32);

                                PdfPCell cell33 = new PdfPCell(new Phrase("PH.C",f));
                                cell33.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell33.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell33.setColspan(3);
                                cell33.setRowspan(2);
                                table.addCell(cell33);

                                //

                                PdfPCell cell34 = new PdfPCell(new Phrase(I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell34.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell34.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell34.setColspan(2);
                                cell34.setRowspan(2);
                                table.addCell(cell34);

                                PdfPCell cell35 = new PdfPCell(new Phrase(I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell35.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell35.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell35.setColspan(2);
                                cell35.setRowspan(2);
                                table.addCell(cell35);

                                PdfPCell cell36 = new PdfPCell(new Phrase(I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell36.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell36.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell36.setColspan(3);
                                cell36.setRowspan(2);
                                table.addCell(cell36);

                                //Row 13

                                PdfPCell cell37 = new PdfPCell(new Phrase("%ERROR",f));
                                cell37.setFixedHeight(35f);
                                cell37.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell37.setColspan(1);
                                cell37.setRowspan(2);
                                table.addCell(cell37);

                                PdfPCell cell38 = new PdfPCell(new Phrase(String.valueOf(error_1_1),f));
                                cell38.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell38.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell38.setColspan(2);
                                cell38.setRowspan(2);
                                table.addCell(cell38);

                                PdfPCell cell39 = new PdfPCell(new Phrase(String.valueOf(error_1_2),f));
                                cell39.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell39.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell39.setColspan(2);
                                cell39.setRowspan(2);
                                table.addCell(cell39);

                                PdfPCell cell40 = new PdfPCell(new Phrase(String.valueOf(error_1_3),f));
                                cell40.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell40.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell40.setColspan(3);
                                cell40.setRowspan(2);
                                table.addCell(cell40);

                                PdfPCell cell41 = new PdfPCell(new Phrase("%ERROR IN RANGE",f));
                                cell41.setFixedHeight(35f);
                                cell41.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell41.setColspan(1);
                                cell41.setRowspan(2);
                                table.addCell(cell41);

                                PdfPCell cell42 = new PdfPCell(new Phrase("",f));
                                cell42.setColspan(2);
                                cell42.setRowspan(2);
                                table.addCell(cell42);

                                PdfPCell cell43 = new PdfPCell(new Phrase("",f));
                                cell43.setColspan(2);
                                cell43.setRowspan(2);
                                table.addCell(cell43);

                                PdfPCell cell44 = new PdfPCell(new Phrase("",f));
                                cell44.setColspan(3);
                                cell44.setRowspan(2);
                                table.addCell(cell44);
                                //

                                PdfPCell cell45 = new PdfPCell(new Phrase("%ERROR OUT OF RANGE",f));
                                cell45.setFixedHeight(35f);
                                cell45.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell45.setColspan(1);
                                cell45.setRowspan(2);
                                table.addCell(cell45);

                                PdfPCell cell46 = new PdfPCell(new Phrase("",f));
                                cell46.setColspan(2);
                                cell46.setRowspan(2);
                                table.addCell(cell46);

                                PdfPCell cell47 = new PdfPCell(new Phrase("",f));
                                cell47.setColspan(2);
                                cell47.setRowspan(2);
                                table.addCell(cell47);

                                PdfPCell cell48 = new PdfPCell(new Phrase("",f));
                                cell48.setColspan(3);
                                cell48.setRowspan(2);
                                table.addCell(cell48);

                                // Row 6
                                PdfPCell cell49 = new PdfPCell(new Phrase("REAL VALUE-V LOAD",f));
                                cell49.setFixedHeight(70f);
                                cell49.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell49.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell49.setBorderWidthTop(2);
                                cell49.setColspan(1);
                                cell49.setRowspan(3);
                                table.addCell(cell49);

                                PdfPCell cell50 = new PdfPCell(new Phrase("PH.A",f));
                                cell50.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell50.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell50.setFixedHeight(35f);
                                cell50.setBorderWidthTop(2);
                                cell50.setColspan(2);
                                cell50.setRowspan(1);
                                table.addCell(cell50);

                                PdfPCell cell51 = new PdfPCell(new Phrase("PH.B",f));
                                cell51.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell51.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell51.setBorderWidthTop(2);
                                cell51.setColspan(2);
                                cell51.setRowspan(1);
                                table.addCell(cell51);

                                PdfPCell cell52 = new PdfPCell(new Phrase("PH.C",f));
                                cell52.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell52.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell52.setBorderWidthTop(2);
                                cell52.setColspan(3);
                                cell52.setRowspan(1);
                                table.addCell(cell52);
                                //
                                PdfPCell cell53 = new PdfPCell(new Phrase(REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell53.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell53.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell53.setColspan(2);
                                cell53.setRowspan(2);
                                table.addCell(cell53);

                                PdfPCell cell54 = new PdfPCell(new Phrase(REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell54.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell54.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell54.setColspan(2);
                                cell54.setRowspan(2);
                                table.addCell(cell54);

                                PdfPCell cell55 = new PdfPCell(new Phrase(REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell55.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell55.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell55.setColspan(3);
                                cell55.setRowspan(2);
                                table.addCell(cell55);
                                //Row 9
                                PdfPCell cell56 = new PdfPCell(new Phrase("V ONLOAD BY V.M-V LOAD",f_1));
                                cell56.setFixedHeight(70f);
                                cell56.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell56.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell56.setColspan(1);
                                cell56.setRowspan(4);
                                table.addCell(cell56);

                                PdfPCell cell57 = new PdfPCell(new Phrase("PH.A",f));
                                cell57.setFixedHeight(35f);
                                cell57.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell57.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell57.setColspan(2);
                                cell57.setRowspan(2);
                                table.addCell(cell57);

                                PdfPCell cell58 = new PdfPCell(new Phrase("PH.B",f));
                                cell58.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell58.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell58.setColspan(2);
                                cell58.setRowspan(2);
                                table.addCell(cell58);

                                PdfPCell cell59 = new PdfPCell(new Phrase("PH.C",f));
                                cell59.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell59.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell59.setColspan(3);
                                cell59.setRowspan(2);
                                table.addCell(cell59);

                                PdfPCell cell60 = new PdfPCell(new Phrase(V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell60.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell60.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell60.setColspan(2);
                                cell60.setRowspan(2);
                                table.addCell(cell60);

                                PdfPCell cell61 = new PdfPCell(new Phrase(REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell61.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell61.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell61.setColspan(2);
                                cell61.setRowspan(2);
                                table.addCell(cell61);

                                PdfPCell cell62 = new PdfPCell(new Phrase(REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell62.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell62.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell62.setColspan(3);
                                cell62.setRowspan(2);
                                table.addCell(cell62);
                                //Row 13

                                PdfPCell cell63 = new PdfPCell(new Phrase("%ERROR",f));
                                cell63.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell63.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell63.setFixedHeight(35f);
                                cell63.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell63.setColspan(1);
                                cell63.setRowspan(2);
                                table.addCell(cell63);

                                PdfPCell cell64 = new PdfPCell(new Phrase(String.valueOf(error_2_1),f));
                                cell64.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell64.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell64.setColspan(2);
                                cell64.setRowspan(2);
                                table.addCell(cell64);

                                PdfPCell cell65 = new PdfPCell(new Phrase(String.valueOf(error_2_2),f));
                                cell65.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell65.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell65.setColspan(2);
                                cell65.setRowspan(2);
                                table.addCell(cell65);

                                PdfPCell cell66 = new PdfPCell(new Phrase(String.valueOf(error_2_3),f));
                                cell66.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell66.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell66.setColspan(3);
                                cell66.setRowspan(2);
                                table.addCell(cell66);

                                PdfPCell cell67 = new PdfPCell(new Phrase("%ERROR IN RANGE",f));
                                cell67.setFixedHeight(35f);
                                cell67.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell67.setColspan(1);
                                cell67.setRowspan(2);
                                table.addCell(cell67);

                                PdfPCell cell68 = new PdfPCell(new Phrase("",f));
                                cell68.setColspan(2);
                                cell68.setRowspan(2);
                                table.addCell(cell68);

                                PdfPCell cell69 = new PdfPCell(new Phrase("",f));
                                cell69.setColspan(2);
                                cell69.setRowspan(2);
                                table.addCell(cell69);

                                PdfPCell cell70 = new PdfPCell(new Phrase("",f));
                                cell70.setColspan(3);
                                cell70.setRowspan(2);
                                table.addCell(cell70);

                                //


                                PdfPCell cell71 = new PdfPCell(new Phrase("%ERROR OUT OF RANGE",f));
                                cell71.setFixedHeight(35f);
                                cell71.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell71.setColspan(1);
                                cell71.setRowspan(2);
                                table.addCell(cell71);

                                PdfPCell cell72 = new PdfPCell(new Phrase("",f));
                                cell72.setColspan(2);
                                cell72.setRowspan(2);
                                table.addCell(cell72);

                                PdfPCell cell73 = new PdfPCell(new Phrase("",f));
                                cell73.setColspan(2);
                                cell73.setRowspan(2);
                                table.addCell(cell73);

                                PdfPCell cell74 = new PdfPCell(new Phrase("",f));
                                cell74.setColspan(3);
                                cell74.setRowspan(2);
                                table.addCell(cell74);

                                //

                                PdfPCell cell75 = new PdfPCell(new Phrase("ACTIVE POWER",f));
                                cell75.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell75.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell75.setBorderWidth(2);
                                cell75.setColspan(1);
                                cell75.setRowspan(2);
                                table.addCell(cell75);

                                PdfPCell cell76 = new PdfPCell(new Phrase(ACTIVE_POWER.getEditText().getText().toString(),f));
                                cell76.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell76.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell76.setBorderWidth(2);
                                cell76.setColspan(2);
                                cell76.setRowspan(2);
                                table.addCell(cell76);

                                PdfPCell cell77 = new PdfPCell(new Phrase("REACTIVE POWER",f));
                                cell77.setFixedHeight(40f);
                                cell77.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell77.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell77.setBorderWidth(2);
                                cell77.setColspan(2);
                                cell77.setRowspan(2);
                                table.addCell(cell77);

                                PdfPCell cell78 = new PdfPCell(new Phrase(REACTIVE_POWER.getEditText().getText().toString(),f));
                                cell78.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell78.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell78.setBorderWidth(2);
                                cell78.setColspan(3);
                                cell78.setRowspan(2);
                                table.addCell(cell78);

                                //

                                PdfPCell cell79 = new PdfPCell(new Phrase("COS PHI",f));
                                cell79.setFixedHeight(40f);
                                cell79.setBorderWidth(2);
                                cell79.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell79.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell79.setColspan(1);
                                cell79.setRowspan(2);
                                table.addCell(cell79);

                                PdfPCell cell80 = new PdfPCell(new Phrase(COS_PHI.getEditText().getText().toString(),f));
                                cell80.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell80.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell80.setBorderWidth(2);
                                cell80.setColspan(2);
                                cell80.setRowspan(2);
                                table.addCell(cell80);

                                PdfPCell cell81 = new PdfPCell(new Phrase("COMMENT",f));
                                cell81.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell81.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell81.setBorderWidth(2);
                                cell81.setColspan(2);
                                cell81.setRowspan(2);
                                table.addCell(cell81);

                                PdfPCell cell82 = new PdfPCell(new Phrase(COMMENT.getEditText().getText().toString(),f));
                                cell82.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell82.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell82.setBorderWidth(2);
                                cell82.setColspan(3);
                                cell82.setRowspan(2);
                                table.addCell(cell82);

                                //

                                PdfPCell cell83 = new PdfPCell(new Phrase("TEST BY",f));
                                cell83.setFixedHeight(40f);
                                cell83.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell83.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell83.setBorderWidth(2);
                                cell83.setColspan(1);
                                cell83.setRowspan(3);
                                table.addCell(cell83);

                                PdfPCell cell84 = new PdfPCell(new Phrase(TEST_BY.getEditText().getText().toString(),f));
                                cell84.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell84.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell84.setBorderWidth(2);
                                cell84.setColspan(2);
                                cell84.setRowspan(3);
                                table.addCell(cell84);

                                PdfPCell cell85 = new PdfPCell(new Phrase("DATE",f));
                                cell85.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell85.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell85.setBorderWidth(2);
                                cell85.setColspan(2);
                                cell85.setRowspan(3);
                                table.addCell(cell85);

                                PdfPCell cell86 = new PdfPCell(new Phrase(DATE.getEditText().getText().toString(),f));
                                cell86.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell86.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell86.setBorderWidth(2);
                                cell86.setColspan(3);
                                cell86.setRowspan(3);
                                table.addCell(cell86);

                                HeaderFooterPageEvent event = new HeaderFooterPageEvent();
                                writer.setPageEvent(event);

                                document.add(table);
                                document.close();
//                                Saved = true;
                                Toast.makeText(Excel_Creator.this, "Saved To" + mFilePath, Toast.LENGTH_SHORT).show();
//                            }
//                            catch (Exception e) {
//                                Toast.makeText(Excel_Creator.this, "Not Saved", Toast.LENGTH_SHORT).show();
//                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Toast.makeText(Excel_Creator.this, "Format Doesn't Support", Toast.LENGTH_SHORT).show();
                    }
                }
            else{
//                if (CT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    CT_RATIO_Text_Input_Layout.setError("This Field Is Required");
//                if (PT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    PT_RATIO_Text_Input_Layout.setError("This Field Is Required");
//                if (REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout.setError("This Field Is Required");
//                if (REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout.setError("This Field Is Required");
//                if (REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout.setError("This Field Is Required");
//                if (REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout.setError("This Field Is Required");
//                if (REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout.setError("This Field Is Required");
//                if (REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.setError("This Field Is Required");
//                if (I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout.setError("This Field Is Required");
//                if (I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout.setError("This Field Is Required");
//                if (I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout.setError("This Field Is Required");
//                if (V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout.setError("This Field Is Required");
//                if (V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout.setError("This Field Is Required");
//                if (V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout.setError("This Field Is Required");
                Toast.makeText(this, "Please Fill The Required Fields", Toast.LENGTH_SHORT).show();
            }

    }

    public void SaveTheFile() {
        createExcelSheet();
    }

    private static void addCellView(WritableSheet sheet, int col, int widthInChars) throws WriteException {
        WritableCellFormat cellFormat = new WritableCellFormat();

        CellView cellView = new CellView();
        cellView.setSize(widthInChars * 256);
        cellView.setFormat(cellFormat);
        sheet.setColumnView(col, cellView);
    }
}