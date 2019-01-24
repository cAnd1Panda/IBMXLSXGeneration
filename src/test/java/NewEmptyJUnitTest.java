/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

import com.ingos.exeltemplatelib.MSFactory;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import org.junit.Before;
import org.junit.Test;
import teamworks.TWList;
import teamworks.TWObject;
import teamworks.TWObjectFactory;

/**
 *
 * @author aarapov
 */
public class NewEmptyJUnitTest {

    public NewEmptyJUnitTest() {
    }
    MSFactory factory;

    @Before
    public void setUp() {
        factory = new MSFactory();
    }

    // TODO add test methods here.
    // The methods must be annotated with annotation @Test. For example:
    //
    @Test
    public void hello() throws IOException {
        ArrayList<HashMap<String, String>> str = new ArrayList<HashMap<String, String>>();
        HashMap<String, String> keyMap = new HashMap<String, String>();
        keyMap.put("userName", "русский");
        str.add(keyMap);

        keyMap = new HashMap<String, String>();
        keyMap.put("userName", "ваня");
        keyMap.put("startDate", new SimpleDateFormat("dd/MM/yyyy").format(new Date()));
        str.add(keyMap);
        
        
        // String base64 = Base64.getEncoder().encodeToString(factory.formXLSDocument(str, "se.xls").toByteArray());
        String base64 = new sun.misc.BASE64Encoder().encode(factory.formXLSDocument(str).toByteArray());
        byte[] data = (new sun.misc.BASE64Decoder()).decodeBuffer(base64);
        FileOutputStream out = new FileOutputStream("test.xls");
        out.write(data);
        out.close();
    }
    
}
