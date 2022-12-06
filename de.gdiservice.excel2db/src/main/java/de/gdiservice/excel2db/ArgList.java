package de.gdiservice.excel2db;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

import de.gdiservice.accessdb.AccessRead;

public class ArgList {
	
	Map<String, String> argMap = new HashMap<>();
	
	public ArgList(String[] args) {
		if (args!=null) {
			for (int i=0; i<args.length; i++) {
				String[] sA = args[i].split("=");
				if (sA.length==2) {
					argMap.put(sA[0], sA[1]);
				}
			}
		}
	}
	
	public String get(String argName) {
		return argMap.get(argName);
	}
	
	public static void main(String[] args) {

	    try {
            ArgList argList = new ArgList(args);

            String datei = argList.get("datei");


            if (datei==null) { 
                System.out.println("Parameter \"datei\" wurde nicht angegeben.");
            }

            if (datei != null) {
                File f = new File(datei);

                System.out.println("Datei: \""+datei+"\"");
                System.out.println(f.getAbsoluteFile());
                System.out.println(f.getCanonicalPath());

                if (!f.exists()) {
                    System.out.println("Datei \""+datei+"\" existiert nicht");
                } else {
                    System.out.println("Datei \""+datei+"\" existiert");                    
                }

                if (!Files.isReadable(Paths.get(f.getCanonicalPath()))) {
                    System.out.println("Datei \""+datei+"\" isReadable=false");                
                } else {
                    System.out.println("Datei \""+datei+"\" isReadable=true");
                }
            }
        } catch (IOException e) {
           e.printStackTrace();
        }
	}
	
}