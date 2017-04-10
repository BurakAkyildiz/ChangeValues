/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package changevalues;

import java.io.File;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.List;
import java.util.function.Consumer;
import java.util.stream.Stream;

/**
 *
 * @author b.akyildiz
 */
public class ChangeValues {

    
    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        
            
            
            
        File basePath = new File("./");
        File files[] = basePath.listFiles();


        Stream.of(files).forEach(new Consumer<File>() {

            @Override
            public void accept(File t) {

                if( t.getName().toLowerCase().startsWith("vbsvalues") == true )
                {
                    try {
                        System.out.println(t.getAbsolutePath());
                        List<String> settingLines = Files.readAllLines(Paths.get(t.getAbsolutePath()));
                        File vbsFile = new File(settingLines.get( settingLines.size() - 1 ));
                        if(vbsFile.canRead())
                        {
                            List<String> vbsLines = Files.readAllLines(Paths.get(vbsFile.getAbsolutePath()));
                            settingLines.stream().forEach(new Consumer<String>() {

                                @Override
                                public void accept(String t) {
                                    System.out.println(t);
                                    try {
                                        int settingLineIndex = settingLines.indexOf(t);
                                        
                                        if( settingLineIndex  < settingLines.size() - 1 )
                                        {
                                            String params[] = t.split(",");
                                            int lineIndex = Integer.parseInt(params[0]);
                                            
                                            String textLine = vbsLines.get(lineIndex);
                                            String result = changeValue(textLine, Integer.parseInt(params[1]), params[2]);
                                            if( result != null )
                                                vbsLines.set(lineIndex, result);

                                        }


                                    } catch (Exception e) {
                                        e.printStackTrace();
                                    }
                                }
                            });




                            Files.write(Paths.get(vbsFile.getAbsolutePath()), vbsLines, StandardOpenOption.TRUNCATE_EXISTING);
                        }                            
                    } catch (Exception e) {
                        e.printStackTrace();
                        System.err.println(e);
                    }
                }
            }
        });
        
        
        
    }
    
    
    public static String changeValue(String line, int startIndex, String newValue)
    {
        
        //System.out.println(line);
        
        if(line == null || startIndex < 0 || newValue == null)
            return null;
        
        String oldValuePart = line.substring(startIndex);
        
        int firstIndex = oldValuePart.indexOf("\"");
        int lastIndex = oldValuePart.indexOf("\"", firstIndex + 1);
        
        String firstPart = line.substring(0, startIndex+firstIndex);
        String lastPart = line.substring(startIndex+lastIndex+1);
        //System.out.println(firstPart+" \n"+lastPart);
        
        String newLine = firstPart +"\""+ newValue +"\""+ lastPart;
        
        
        return newLine;
        
    }
    
}
