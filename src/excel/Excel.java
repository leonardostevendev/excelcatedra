package excel;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.*;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

public final class Excel{
    int line = 0;
    FileInputStream input_document;
    private String URL = "";
    HSSFWorkbook my_xls_workbook;
    Scanner entrada = new Scanner(System.in);
    List<String> correos = new ArrayList();
    Map<String, Integer> repetidos = new HashMap<String, Integer>();

    String[] URLS = {"D:\\Info\\Bucaramanga\\BD1CM.xls",
        "D:\\Info\\Bucaramanga\\BE1CM.xls",
        "D:\\Info\\Bucaramanga\\BE2CM.xls",
        "D:\\Info\\Monteria\\DM1CM.xls",
        "D:\\Info\\Monteria\\EM1CM.xls",
        "D:\\Info\\Monteria\\EM2CM.xls",
        "D:\\Info\\Pasto\\PD1CM.xls",
        "D:\\Info\\Pasto\\PE1CM.xls",
        "D:\\Info\\Valledupar\\VD1CM.xls",
        "D:\\Info\\Valledupar\\VE1CM.xls",
        "D:\\Info\\Valledupar\\VE2CM.xls"};

    public Excel2() {
        correos.add("admin@pegui.edu.co");
        correos.add("alexandermedina@pegui.edu.co");
        correos.add("alexburbano@pegui.edu.co");
        correos.add("analondono@pegui.edu.co");
        correos.add("arlethsanchez@pegui.edu.co");
        correos.add("aydaflorez@pegui.edu.co");
        correos.add("ayuda@pegui.edu.co");
        correos.add("benjaminstein@pegui.edu.co");
        correos.add("camilovelandia@pegui.edu.co");
        correos.add("claudianaranjo@pegui.edu.co");
        correos.add("claudiaparrado@pegui.edu.co");
        correos.add("comunicaciones@pegui.edu.co");
        correos.add("constanzahernandez@pegui.edu.co");
        correos.add("convocatorias@pegui.edu.co");
        correos.add("dianadelgado@pegui.edu.co");
        correos.add("djamelkadi@pegui.edu.co");
        correos.add("gloriadiaz@pegui.edu.co");
        correos.add("gloriaortiz@pegui.edu.co");
        correos.add("hernandovesga@pegui.edu.co");
        correos.add("hugolondono@pegui.edu.co");
        correos.add("janmunoz@pegui.edu.co");
        correos.add("jhadercano@pegui.edu.co");
        correos.add("johannalinares@pegui.edu.co");
        correos.add("josegutierrez@pegui.edu.co");
        correos.add("katherinemontero@pegui.edu.co");
        correos.add("katherinerios@pegui.edu.co");
        correos.add("lauraayala@pegui.edu.co");
        correos.add("lidabejarano@pegui.edu.co");
        correos.add("linacordero@pegui.edu.co");
        correos.add("linamantilla@pegui.edu.co");
        correos.add("mariacortes@pegui.edu.co");
        correos.add("mariafranco@pegui.edu.co");
        correos.add("mariagaona@pegui.edu.co");
        correos.add("mauricioparra@pegui.edu.co");
        correos.add("natalialondono@pegui.edu.co");
        correos.add("olgamogollon@pegui.edu.co");
        correos.add("paulabravo@pegui.edu.co");
        correos.add("paulodiaz@pegui.edu.co");
        correos.add("peguiedu@pegui.edu.co");
        correos.add("rafaelgarcia@pegui.edu.co");
        correos.add("rafaelmurgas@pegui.edu.co");
        correos.add("santiagovillarraga@pegui.edu.co");
        correos.add("telmaherazo@pegui.edu.co");        
        editSheet();
      

    }

    public HSSFSheet obtSheet(String url) {
        try {
            //Read the spreadsheet that needs to be updated
            input_document = new FileInputStream(new File(url));
            //Access the workbook            
            my_xls_workbook = new HSSFWorkbook(input_document);
            //Access the worksheet, so that we can update / modify it.
            return my_xls_workbook.getSheetAt(0);

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    public void editSheet() {

        for (int j = 0; j < URLS.length; j++) {
            try {                
                URL = URLS[j];
                HSSFSheet sheetToProcess = obtSheet(URL);
                // declare a Cell object

                Iterator<Row> rowIterator = sheetToProcess.iterator();
                int control = 0;
                while (rowIterator.hasNext()) {
                    line++;
                    Row row = rowIterator.next();
                    StringBuffer correo = new StringBuffer();
                    if (control == 0) {
                        control = 1;
                    } else {
                        String ini = primerNombre(row.getCell(2).getStringCellValue());
                        String apellido = apellido(row.getCell(3).getStringCellValue());
                        correo.append(ini.toLowerCase()).append(apellido.toLowerCase()).append("@pegui.edu.co");
                        if (correos.indexOf(correo.toString()) >= 0) {
                            if (repetidos.get(correo.toString()) != null) {
                                int consecutivo = ((repetidos.get(correo.toString())) + 1);
                                repetidos.replace(correo.toString(), (repetidos.get(correo.toString())) + 1);
                                correo.insert((correo.toString().indexOf("@")), String.valueOf(consecutivo));
                            } else {
                                repetidos.put(correo.toString(), 1);
                                correo.insert(correo.toString().indexOf("@"), "1");
                                correos.add(correo.toString());
                            }
                        } else {
                            correos.add(correo.toString());
                        }
                        Cell mail = row.getCell(7);
                        mail.setCellValue(correo.toString());
                    }
                    System.out.println(URLS[0] + "  "+line);
                }

                FileOutputStream output = new FileOutputStream(new File(URL));
                my_xls_workbook.write(output);
                output.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    public String primerNombre(String nombre) {
        String[] nombres = nombre.split(" ");
        return nombres[0].toLowerCase().trim().replace("ñ", "n");
    }

    public String apellido(String apellido) {
        String[] apellidos = apellido.split(" ");
        return apellidos[0].toLowerCase().trim().replace("ñ", "n");
    }

    public void mostrarRepetidos() {
        for (Map.Entry<String, Integer> entry : repetidos.entrySet()) {
            String key = entry.getKey();
            int value = entry.getValue();
            System.out.println(key + " " + value);
        }
    }

    public static void main(String[] args) throws Exception {

        Excel2 procesar = new Excel2();
    }

}
