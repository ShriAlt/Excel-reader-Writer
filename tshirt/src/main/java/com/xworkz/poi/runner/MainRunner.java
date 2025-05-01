package com.xworkz.poi.runner;

import com.xworkz.poi.internal.Bottle;
import com.xworkz.poi.internal.Pillow;

import java.io.IOException;

public class MainRunner {

    public static void main(String[] args) throws IOException {
//        Bottle bottle=new Bottle();
//        bottle.readExcel();

        Pillow pillow=new Pillow();
        pillow.readSheet();
        pillow.excelWrite();
    }
}
