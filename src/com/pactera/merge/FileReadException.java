package com.pactera.merge;

public class FileReadException extends Exception {
    public FileReadException(){
        super();
    }
    public FileReadException(String msg){
        super(msg);
    }
    public FileReadException(String msg,Throwable cause){
        super(msg,cause);
    }
    public FileReadException(Throwable cause){
        super(cause);
    }


}
