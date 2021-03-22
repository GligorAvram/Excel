package com.gligor.excel.excel.logic;

import org.apache.commons.io.FileUtils;

import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.util.Timer;
import java.util.TimerTask;

public class DeleteFile extends Thread {

    //paths of the files to be deleted
    private String fileToDelete;

    public DeleteFile(String path){
        this.fileToDelete = path;
    }

    public void run(){
        deleteFile();
    }

    private void deleteFile() {
        try {
            //10 min
            Thread.sleep(60000);
            FileUtils.forceDelete(new File(fileToDelete));
        }
        catch (IOException | InterruptedException e) {
            e.printStackTrace();
        }
    }
}

