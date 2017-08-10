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

	public static void main( String[] args )
	{
		path = args[0];
		target = args[1];
		List<File> files = filterFiles(path);
		if(files.size() == 0){
			logger.info("No files found.");
		}
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

	private static void display( List<String> list )
	{
		final Map<String, List<String>>[] f = new HashMap[] { null };
		roots = list.stream().map(l -> {
			Root root = new Root();
			String key = l.split("-")[0].trim();
			String value = l.split("-")[1].trim();

			if( key.equals(Constants.MACHINEID) )
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

	private static void creatingExcel( List<Root> roots )
	{

		File f = checkFile(target);
		XSSFWorkbook workbook = new XSSFWorkbook();
		final XSSFSheet[] s = { null };
		for( Root root : roots )
		{
			Map<String, List<String>> map = root.getMap();
			if( map.values().isEmpty() )
			{
				continue;
			}

			Row h;
			AtomicInteger numberOfRows;
			if( map.containsKey(Constants.MACHINEID) )
			{
				s[0] = workbook.getSheet(map.get(Constants.MACHINEID).get(0));
				if( s[0] == null )
				{
					s[0] = workbook.createSheet(map.get(Constants.MACHINEID).get(0));
					System.out.println("created sheet " + s[0].getSheetName());
				}

			}
			if( map.containsKey(Constants.KEYWORDS) )
			{
				System.out.println("using " + s[0].getSheetName());
				if( s[0].getRow(0) == null )
				{
					h = s[0].createRow(0);
					AtomicInteger i = new AtomicInteger(0);
					Row finalH = h;
					Arrays.asList(map.get(Constants.KEYWORDS).get(0).split("/")).forEach(lon -> finalH.createCell(i.getAndIncrement()).setCellValue(lon));
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
					Arrays.asList(map.get(Constants.KEYWORDS).get(0).split("/")).forEach(keyword -> {
						if( !values.contains(keyword) )
						{
							finalH.createCell(numberOfCells.getAndIncrement()).setCellValue(keyword);

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
				if( map.containsKey(cell.getStringCellValue()) )
				{
					index = cell.getColumnIndex();
					Cell cel = h.createCell(index);
					cel.setCellValue(map.get(cell.getStringCellValue()).get(0));
				}
			}
		}
		try( FileOutputStream outputStream = new FileOutputStream(f) )
		{
			workbook.write(outputStream);
		}
		catch( IOException e )
		{
			e.printStackTrace();
		}
	}

	private static File checkFile( String target )
	{
		try
		{
			if( Files.list(Paths.get(target)).anyMatch(file -> file.getFileName().toString().contains("final")) )
				return Paths.get(target + "/final.xlsx").toFile();
			return Files.createFile(Paths.get(target + "/final.xlsx")).toFile();
		}
		catch( IOException e )
		{
			e.printStackTrace();
		}
		return null;
	}

	private static List<File> filterFiles( String path )
	{
		try
		{
			return Files.list(Paths.get(path)).map(Path::toFile).filter(file -> file.getName().contains(Constants.BHPRO))
					.collect(Collectors.toList());
		}
		catch( IOException e )
		{
			e.printStackTrace();
		}
		return null;
	}

	public Map<String, List<String>> getMap()
	{
		return map;
	}

	public void setMap( Map<String, List<String>> map )
	{
		this.map = map;
	}

}
