package fileOperation;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

public class Root {

	static String path;
	HashMap<String, List<String>> map = new HashMap<>();
	static List<Root> roots = new ArrayList<>();
	static String target;

	public static void main( String[] args ) throws IOException
	{
		path = "/home/raksha/Downloads";
		target = "/home/raksha/Downloads";
		List<File> files = filterFiles(path);
		store(files);
	}

	private static void store( List<File> files )
	{
		files.forEach(file -> {
			try
			{
				display(Files.readAllLines(file.toPath()));
			}
			catch( IOException e )
			{
				e.printStackTrace();
			}
		});
	}

	private static void display( List<String> list ) throws IOException
	{
		final HashMap<String, List<String>>[] f = new HashMap[] { null };
		roots = list.stream().map(l -> {
			Root root = new Root();
			String key = l.split("-")[0].trim();
			String value = l.split("-")[1].trim();

			if( key.equals("MACHINEID") )
				f[0] = root.getMap();

			f[0].putIfAbsent(key, Arrays.asList(value));
			f[0].computeIfPresent(key, ( k, v ) -> {
				if( !v.contains(value) )
				{
					v.add(value);
				}
				return v;
			});

			return root;
		}).collect(Collectors.toList());

		creatingExcel(roots);
	}

	private static void creatingExcel( List<Root> roots ) throws IOException
	{

		File f = checkFile(target);
		XSSFWorkbook workbook = new XSSFWorkbook();
		List<String> sheets = new ArrayList<>();
		final XSSFSheet[] s = { null };
		for( Root root : roots )
		{
			HashMap<String, List<String>> m = root.getMap();
			if( m.values().isEmpty() )
			{
				continue;
			}

			Row h;
			AtomicInteger numberOfRows;
			if( m.containsKey("MACHINEID") )
			{
				s[0] = workbook.getSheet(m.get("MACHINEID").get(0));
				if( s[0] == null )
				{
					s[0] = workbook.createSheet(m.get("MACHINEID").get(0));
					System.out.println("created sheet " + s[0].getSheetName());
				}

			}
			if( m.containsKey("kEYWORDS") )
			{
				System.out.println("using " + s[0].getSheetName());
				if( s[0].getRow(0) == null )
				{
					h = s[0].createRow(0);
					AtomicInteger i = new AtomicInteger(0);
					Row finalH = h;
					Arrays.asList(m.get("kEYWORDS").get(0).split("/")).forEach(lon -> {
						finalH.createCell(i.getAndIncrement()).setCellValue(lon);
					});
					System.out.println("created new header");
				}
				else
				{
					h = s[0].getRow(0);
					AtomicInteger numberOfCells = new AtomicInteger(h.getLastCellNum());
					Row finalH = h;
					int it = 0;
					int n = numberOfCells.get();
					List<String> values = new ArrayList<>();
					while( it <= n - 1 )
					{
						values.add(h.getCell(it).getStringCellValue());
						it++;
					}
					System.out.println(values.toString());
					Arrays.asList(m.get("kEYWORDS").get(0).split("/")).forEach(lon -> {
						if( !values.contains(lon) )
						{
							finalH.createCell(numberOfCells.getAndIncrement()).setCellValue(lon);

						}
					});
				}

			}
			int index;

			numberOfRows = new AtomicInteger(s[0].getLastRowNum());
			h = s[0].createRow(numberOfRows.incrementAndGet());
			Iterator<Cell> itr = s[0].getRow(0).cellIterator();
			while( itr.hasNext() )
			{
				Cell cell = itr.next();
				if( m.containsKey(cell.getStringCellValue()) )
				{
					index = cell.getColumnIndex();
					Cell cel = h.createCell(index);
					cel.setCellValue(m.get(cell.getStringCellValue()).get(0));
				}
			}
		}
		try( FileOutputStream outputStream = new FileOutputStream(f) )
		{
			workbook.write(outputStream);
		}
	}

	private static File checkFile( String target ) throws IOException
	{
		if( Files.list(Paths.get(target)).anyMatch(file -> file.getFileName().toString().contains("final")) )
			return Paths.get(target + "/final.xlsx").toFile();
		return Files.createFile(Paths.get(target + "/final.xlsx")).toFile();
	}

	private static List<File> filterFiles( String path ) throws IOException
	{
		return Files.list(Paths.get(path)).map(path1 -> path1.toFile()).filter(file -> file.getName().contains("BHPRO"))
				.collect(Collectors.toList());
	}

	public HashMap<String, List<String>> getMap()
	{
		return map;
	}

	public void setMap( HashMap<String, List<String>> map )
	{
		this.map = map;
	}
}
