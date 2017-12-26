package fileOperation;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalTime;
import java.time.temporal.ChronoUnit;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.logging.Logger;
import java.util.stream.Collectors;

public class Root {

    public static Logger logger = Logger.getLogger("Root");
    static String path;
    Map<String, List<String>> map = new HashMap<>();
    static List<Root> roots = new ArrayList<>();
    static String target;

    public static void main(String[] args) {
        path = "B:\\project";
        target = "B:\\project";
        List<File> files = filterFiles(path);
        if (files.size() == 0) {
            logger.info("No files found.");
        }
        store(files);
    }

    private static void store(List<File> files) {
        files.forEach(file -> {
            try {
                display(Files.readAllLines(file.toPath()));
            } catch (IOException e) {
                e.printStackTrace();
            } catch (ParseException e) {
                e.printStackTrace();
            }
        });
    }

    public static LocalTime convert(String time) {
//        String converted = convertToTimeFormat(time);
        LocalTime localTime = LocalTime.parse(time);
        return localTime;
    }

    private static String convertToTimeFormat(String time) {
        if (time.length() < 6) {
            time = "0" + time;
        }
        return time.substring(0, 2) + ":" + time.substring(2, 4) + ":" + time.substring(4);
    }

    private static void display(List<String> list) throws ParseException {
        final Map<String, List<String>>[] f = new HashMap[]{null};
        roots = list.stream().map(l -> {
            Root root = new Root();
            if (l.split("-").length < 2) {
                return null;
            }
            String key = l.split("-")[0].trim();
            String value = l.split("-")[1].trim();

            if (key.equals("STARTTIME") || key.equals("ENDTIME")){
                    value = convertToTimeFormat(value);
            }

            if (key.equals("DATE")){
                value = convertToDateFormat(value);
            }
            if (key.equals(Constants.MACHINEID))
                f[0] = root.getMap();
            List<String> g = new ArrayList<>();
            g.add(value);
            f[0].putIfAbsent(key, g);
            String finalValue = value;
            f[0].computeIfPresent(key, (k, v) -> {
                if (!v.contains(finalValue)) {
                    v.add(finalValue);
                }
                return v;
            });
            return root;
        }).filter(each -> each != null && !each.getMap().isEmpty()).collect(Collectors.toList());
        final String[] e1 = {null};
        final String[] s2 = {null};
        roots = roots.stream().map(eachRoot -> {
            Map<String, List<String>> smap = eachRoot.getMap();

            if (e1[0] != null && smap.keySet().contains("STARTTIME")){
                s2[0] = smap.get("STARTTIME").get(0);
            }
            if (e1[0] != null && s2[0] !=null){
                LocalTime start2 = convert(s2[0]);
                LocalTime end1 = convert(e1[0]);
                long ct = ChronoUnit.SECONDS.between(end1, start2);
                List<String> l = new ArrayList<>();
                l.add(splitToComponentTimes(ct));
                smap.put("DELAYTIME", l);
                e1[0] = null;
            }
            if (smap.keySet().contains("ENDTIME") && e1[0] == null){
                e1[0] = smap.get("ENDTIME").get(0);
            }
            if (smap.keySet().contains("STARTTIME") && smap.keySet().contains("ENDTIME")) {
                LocalTime startTime = convert(smap.get("STARTTIME").get(0));
                LocalTime endTime = convert(smap.get("ENDTIME").get(0));
                long ct = ChronoUnit.SECONDS.between(startTime, endTime);
                List<String> l = new ArrayList<>();
                l.add(splitToComponentTimes(ct));
                smap.put("CYCLETIME", l);
            }
            eachRoot.setMap(smap);
            return eachRoot;
        }).collect(Collectors.toList());
        creatingExcel(roots);
    }

    private static String convertToDateFormat(String value) {

        String converted = value.substring(0,4) + "/" + value.substring(4,6) + "/" + value.substring(6);
        return converted;
    }

    public static String splitToComponentTimes(long longVal)
    {
        int hours = (int) longVal / 3600;
        int remainder = (int) longVal - hours * 3600;
        int mins = remainder / 60;
        remainder = remainder - mins * 60;
        int secs = remainder;

        String ints = hours + ":" +mins +":"+ secs;
        return ints;
    }

    private static void creatingExcel(List<Root> roots) throws ParseException {

        File f = checkFile(target);
        XSSFWorkbook workbook = new XSSFWorkbook();
        final XSSFSheet[] s = {null};
        String date1 = null;
        for (Root root : roots) {
            Map<String, List<String>> map = root.getMap();
            if (map.values().isEmpty()) {
                continue;
            }
            if(date1 == null){
               date1 =  map.get("DATE").get(0);
            }
            Row h;
            AtomicInteger numberOfRows;
            if (map.containsKey(Constants.MACHINEID)) {
                s[0] = workbook.getSheet(map.get(Constants.MACHINEID).get(0));
                if (s[0] == null) {
                    s[0] = workbook.createSheet(map.get(Constants.MACHINEID).get(0));
                    System.out.println("created sheet " + s[0].getSheetName());
                }

            }
//			if( map.containsKey(Constants.KEYWORDS) )
//			{
            System.out.println("using " + s[0].getSheetName());
            if (s[0].getRow(0) == null) {
                h = s[0].createRow(0);
                AtomicInteger i = new AtomicInteger(0);
                Row finalH = h;
                map.keySet().forEach(header -> {
                    if (!header.equals(Constants.KEYWORDS) && !header.equals(Constants.MACHINEID)) {
                        finalH.createCell(i.getAndIncrement()).setCellValue(header);
                    }
                });
                //Arrays.asList(map.get(Constants.KEYWORDS).get(0).split("/")).forEach(lon -> finalH.createCell(i.getAndIncrement()).setCellValue(lon));
                System.out.println("created new header");
            } else {
                h = s[0].getRow(0);
                AtomicInteger numberOfCells = new AtomicInteger(h.getLastCellNum());

                Row finalH = h;
                int it = 0;
                int n = numberOfCells.get();
                List<String> values = new ArrayList<>();
                while (it <= n - 1) {
                    values.add(h.getCell(it).getStringCellValue());
                    it++;
                }
                map.keySet().forEach(keyword -> {
                    if (!keyword.equals(Constants.KEYWORDS) && !keyword.equals(Constants.MACHINEID)) {
                        if (!values.contains(keyword)) {
                            finalH.createCell(numberOfCells.getAndIncrement()).setCellValue(keyword);

                        }
                    }
                });
//							Arrays.asList(map.get(Constants.KEYWORDS).get(0).split("/")).forEach(keyword -> {
//						if( !values.contains(keyword) )
//						{
//							finalH.createCell(numberOfCells.getAndIncrement()).setCellValue(keyword);
//
//						}
//					});
            }
//			}
            int index;

            numberOfRows = new AtomicInteger(s[0].getLastRowNum());
            if (date1 != null && !date1.equals(map.get("DATE").get(0))){
                Date d1 = new SimpleDateFormat("yyyy/mm/dd").parse(date1);
                Date d2 = new SimpleDateFormat("yyyy/mm/dd").parse(map.get("DATE").get(0));
                if (d2.after(d1))
                    numberOfRows.addAndGet(3);
                date1 = map.get("DATE").get(0);
            }
            h = s[0].createRow(numberOfRows.incrementAndGet());
            Iterator<Cell> itr = s[0].getRow(0).cellIterator();
            while (itr.hasNext()) {
                Cell cell = itr.next();
                if (map.containsKey(cell.getStringCellValue())) {
                    index = cell.getColumnIndex();
                    Cell cel = h.createCell(index);
                    cel.setCellValue(map.get(cell.getStringCellValue()).get(0));
                }
            }
        }
        try (FileOutputStream outputStream = new FileOutputStream(f)) {
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static File checkFile(String target) {
        try {
            if (Files.list(Paths.get(target)).anyMatch(file -> file.getFileName().toString().contains("final")))
                return Paths.get(target + "/final.xlsx").toFile();
            return Files.createFile(Paths.get(target + "/final.xlsx")).toFile();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    private static List<File> filterFiles(String path) {
        try {
            return Files.list(Paths.get(path)).map(Path::toFile).filter(file -> file.getName().contains(Constants.BHPRO))
                    .collect(Collectors.toList());
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    public Map<String, List<String>> getMap() {
        return map;
    }

    public void setMap(Map<String, List<String>> map) {
        this.map = map;
    }

}
