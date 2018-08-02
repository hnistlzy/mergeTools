package com.pactera.contrast;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;

public interface SelectBehavior {
    void selectWayXlsx(File file,String str) throws IOException;
    void selectWayXls(File file,String str);
}
