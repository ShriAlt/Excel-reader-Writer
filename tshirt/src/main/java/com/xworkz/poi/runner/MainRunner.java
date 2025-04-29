package com.xworkz.poi.runner;

import com.xworkz.poi.internal.Bottle;

import java.io.IOException;

public class MainRunner {

    public static void main(String[] args) throws IOException {
        Bottle bottle=new Bottle();
        bottle.readExcel();
    }
}
