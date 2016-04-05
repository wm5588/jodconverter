//
// JODConverter - Java OpenDocument Converter
// Copyright 2004-2012 Mirko Nasato and contributors
//
// JODConverter is Open Source software, you can redistribute it and/or
// modify it under either (at your option) of the following licenses
//
// 1. The GNU Lesser General Public License v3 (or later)
//    -> http://www.gnu.org/licenses/lgpl-3.0.txt
// 2. The Apache License, Version 2.0
//    -> http://www.apache.org/licenses/LICENSE-2.0.txt
//
package org.artofsolving.jodconverter.office;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.artofsolving.jodconverter.util.PlatformUtils;

import com.sun.star.beans.PropertyValue;
import com.sun.star.uno.UnoRuntime;

public class OfficeUtils {

    public static final String SERVICE_DESKTOP = "com.sun.star.frame.Desktop";

    private OfficeUtils() {
        throw new AssertionError("utility class must not be instantiated");
    }

    public static <T> T cast(Class<T> type, Object object) {
        return (T) UnoRuntime.queryInterface(type, object);
    }

    public static PropertyValue property(String name, Object value) {
        PropertyValue propertyValue = new PropertyValue();
        propertyValue.Name = name;
        propertyValue.Value = value;
        return propertyValue;
    }

    @SuppressWarnings("unchecked")
    public static PropertyValue[] toUnoProperties(Map<String,?> properties) {
        PropertyValue[] propertyValues = new PropertyValue[properties.size()];
        int i = 0;
        for (Map.Entry<String,?> entry : properties.entrySet()) {
            Object value = entry.getValue();
            if (value instanceof Map) {
                Map<String,Object> subProperties = (Map<String,Object>) value;
                value = toUnoProperties(subProperties);
            }
            propertyValues[i++] = property((String) entry.getKey(), value);
        }
        return propertyValues;
    }

    public static String toUrl(File file) {
        String path = file.toURI().getRawPath();
        String url = path.startsWith("//") ? "file:" + path : "file://" + path;
        return url.endsWith("/") ? url.substring(0, url.length() - 1) : url;
    }

    public static File getDefaultOfficeHome() {
        if (System.getProperty("office.home") != null) {
            return new File(System.getProperty("office.home"));
        }
        if (PlatformUtils.isWindows()) {
            // %ProgramFiles(x86)% on 64-bit machines; %ProgramFiles% on 32-bit ones
            String programFiles = System.getenv("ProgramFiles(x86)");
            if (programFiles == null) {
                programFiles = System.getenv("ProgramFiles");
            }
            return findOfficeHome(
                programFiles + File.separator + "OpenOffice.org 3",
                programFiles + File.separator + "LibreOffice 3"
            );
        } else if (PlatformUtils.isMac()) {
            return findOfficeHome(
                "/Applications/OpenOffice.org.app/Contents",
                "/Applications/LibreOffice.app/Contents"
            );
        } else {
            // Linux or other *nix variants
            return findOfficeHome(
                      linuxOfficeHomes(50)
            );
        }
    }
    
    private static String[] linuxOfficeHomes(int officeMaxVersions){
    	String path1 = "/usr/lib/openoffice";
    	String path2 = "/opt/libreoffice";
    	int size = (officeMaxVersions-4)*10;
    	int from = 3;
    	List<String> targetPaths = new ArrayList<String>(size*2+4);
    	{
    		targetPaths.add("/opt/openoffice.org3");
    		targetPaths.add("/opt/libreoffice");
    		targetPaths.add("/usr/lib/openoffice");
    		targetPaths.add("/usr/lib/libreoffice");
    	}
    	StringBuilder sb = new StringBuilder();
    	for (int i = 0; i < size; i++) {
    		int mod = i%10;
    		if(mod == 0){
    			from += 1;
    		}
    		sb.delete(0, sb.length());
    		targetPaths.add(sb.append(path1).append(from).append('.').append(mod).toString());
    		sb.delete(0, sb.length());
    		targetPaths.add(sb.append(path2).append(from).append('.').append(mod).toString());
    		sb.delete(0, sb.length());
		}
    	String[] outArr = new String[targetPaths.size()];
    	return targetPaths.toArray(outArr);
    }
    
    private static File findOfficeHome(String... knownPaths) {
        for (String path : knownPaths) {
            File home = new File(path);
            if (getOfficeExecutable(home).isFile()) {
                return home;
            }
        }
        return null;
    }

    public static File getOfficeExecutable(File officeHome) {
        if (PlatformUtils.isMac()) {
            return new File(officeHome, "MacOS/soffice.bin");
        } else {
            return new File(officeHome, "program/soffice.bin");
        }
    }

}
