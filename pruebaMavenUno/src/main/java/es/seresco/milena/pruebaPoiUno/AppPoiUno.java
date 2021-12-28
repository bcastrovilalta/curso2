package es.seresco.milena.pruebaPoiUno;

import java.io.FileOutputStream;

import org.apache.logging.log4j.*;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.*;

public class AppPoiUno 
{
	private static final Logger logger = LogManager.getLogger(AppPoiUno.class);
	
	private static XSSFFont crearFuente(XSSFWorkbook libroExcel, String tipo, int altura, boolean bold, boolean italic, byte[] rgb)
	{
		XSSFFont fuente = libroExcel.createFont();
		
		fuente.setFontName(tipo);
		fuente.setFontHeightInPoints((short) altura);
		fuente.setBold(bold);	
		fuente.setItalic(italic);
		fuente.setColor(new XSSFColor(rgb, null));
		
		return(fuente);
	}
		
	public static void main( String[] args )
	{
		try
		{
			logger.info("Inicio AppPoiUno");
			
			final XSSFWorkbook libro = new XSSFWorkbook();	
			//Creamos hoja de datos
			XSSFSheet hojaDatos = libro.createSheet(WorkbookUtil.createSafeSheetName("Datos"));
			
			//Dentro de la hoja de datos creamos la fila de cabeceras.
			//Creamos la primera fila, que será l aque tenga indice 0. 
			final XSSFRow fila = hojaDatos.createRow((short) 0);
			
			//Dentro de la fila creamos las columnas, es decir las cedas que forman las columnas.
			//final XSSFCell celda = fila.createCell(0);
			// o tambien
			fila.createCell(0).setCellValue("Cliente");
			fila.createCell(1).setCellValue("Concepto");
			fila.createCell(2).setCellValue("Cantidad");
			fila.createCell(3).setCellValue("Precio");
			fila.createCell(4).setCellValue("Importe");
			
			//Damos un estilo a las celdas de la cabecera.
			XSSFCellStyle estiloCabecera = libro.createCellStyle();
			//Creamos una fuente
			final byte[] rgb_blanco = { (byte) 251, (byte) 251, (byte) 251 };
			XSSFFont fuente_cabecera = crearFuente(libro,"Arial",9,true,false,rgb_blanco);
			//Al estilo cabecera le ponemos la fuente_cabecera. Y el color de relleno.						
			estiloCabecera.setFont(fuente_cabecera);
			final byte[] rgb_azul_indigo = { (byte) 0, (byte) 65, (byte) 106 };
			estiloCabecera.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			estiloCabecera.setFillForegroundColor(new XSSFColor(rgb_azul_indigo, null));
							
			//A cada celda de la fila de cabeceras le ponemos el estilo.
			fila.getCell(0).setCellStyle(estiloCabecera);
			fila.getCell(1).setCellStyle(estiloCabecera);
			fila.getCell(2).setCellStyle(estiloCabecera);
			fila.getCell(3).setCellStyle(estiloCabecera);
			fila.getCell(4).setCellStyle(estiloCabecera);
			
			
			final XSSFRow fila1 = hojaDatos.createRow((short) 1);
			
			fila1.createCell(0).setCellValue("Federico");
			fila1.createCell(1).setCellValue("Patatas");
			fila1.createCell(2).setCellValue(20);
			fila1.createCell(3).setCellValue(5);
			fila1.createCell(4).setCellValue(100);
			
			final XSSFRow fila2 = hojaDatos.createRow((short) 2);
			
			fila2.createCell(0).setCellValue("Lorenzo");
			fila2.createCell(1).setCellValue("Lirones");
			fila2.createCell(2).setCellValue(10);
			fila2.createCell(3).setCellValue(77);
			fila2.createCell(4).setCellValue(770);
			
			final XSSFRow fila3 = hojaDatos.createRow((short) 3);
			
			fila3.createCell(0).setCellValue("Lorenzo");
			fila3.createCell(1).setCellValue("Mochuelos");
			fila3.createCell(2).setCellValue(2);
			fila3.createCell(3).setCellValue(15);
			fila3.createCell(4).setCellValue(30);
			
			final XSSFRow fila4 = hojaDatos.createRow((short) 4);
			
			fila4.createCell(0).setCellValue("Federico");
			fila4.createCell(1).setCellValue("Patatas");
			fila4.createCell(2).setCellValue(3);
			fila4.createCell(3).setCellValue(7);
			fila4.createCell(4).setCellValue(21);			
			
			final XSSFRow fila5 = hojaDatos.createRow((short) 5);
			
			fila5.createCell(0).setCellValue("Lorenzo");
			fila5.createCell(1).setCellValue("Patatas");
			fila5.createCell(2).setCellValue(6);
			fila5.createCell(3).setCellValue(7);
			fila5.createCell(4).setCellValue(42);	
			
			final XSSFRow fila6 = hojaDatos.createRow((short) 6);
			
			fila6.createCell(0).setCellValue("Fabala");
			fila6.createCell(1).setCellValue("Mochuelos");
			fila6.createCell(2).setCellValue(4);
			fila6.createCell(3).setCellValue(3);
			fila6.createCell(4).setCellValue(12);	
			
			hojaDatos.setAutoFilter(new CellRangeAddress(0,0,0,4));
			
			//Creamos una hoja para la tabla dinámica.
			final XSSFSheet hojaTablaDin = libro.createSheet(WorkbookUtil.createSafeSheetName("Tabla dinamica"));
			
			final int primeraFila = 0;
			final int ultimaFila = hojaDatos.getLastRowNum();
			 
			final int primeraColumna = hojaDatos.getRow(primeraFila).getFirstCellNum();
			final int ultimaColumna = hojaDatos.getRow(primeraFila).getLastCellNum(); // devuelve la última + 1
			
			// convierte numeros de fila y columna a nombres excel: A1E6
			final CellReference celdaSuperiorIzquierda = new CellReference(hojaDatos.getSheetName(), primeraFila, primeraColumna, true, true);
			final CellReference celdaInferiorDerecha = new CellReference(hojaDatos.getSheetName(), ultimaFila, ultimaColumna - 1, true, true);

			// área de datos de la tabla dinámica
			final AreaReference aref = new AreaReference(celdaSuperiorIzquierda, celdaInferiorDerecha, SpreadsheetVersion.EXCEL2007);
	
			// posicion donde se va a insertar la tabla dinámica
			final CellReference posicion = new CellReference(0, 0);
	
			// crea la tabla dinámica
			final XSSFPivotTable pivotTable = hojaTablaDin.createPivotTable(aref, posicion,hojaDatos);
			
			//Establecemos el rowlabel, que será el nombre
			pivotTable.addRowLabel(0);
			//Establecemos el collabel, que será el concepto
			pivotTable.addColLabel(1);
			
			//Dentro de la columna concepto iran:
			pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 2, "Suma Cantidad");
			pivotTable.addColumnLabel(DataConsolidateFunction.AVERAGE,3,"Promedio Precio");
			pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 4, "Suma Importe");
			
			//Por defecto no pone correctamentee el nombre del rowlabel ni el del colLabel. Los ponemos.
			pivotTable.getCTPivotTableDefinition().setColHeaderCaption("Concepto");
		    pivotTable.getCTPivotTableDefinition().setRowHeaderCaption("Nombre");
		    
		    //Podemos quitar las columnas de totales
		    //pivotTable.getCTPivotTableDefinition().setColGrandTotals(false);
		    //Podemos quitar las filas de totales
		    //pivotTable.getCTPivotTableDefinition().setRowGrandTotals(false);
			
		    //Creamos e fichero donde guardar el libro excel.
			final FileOutputStream out = new FileOutputStream("c:/temp/curso.xlsx");
			//Grabamos.
			libro.write(out);
			out.close();
			libro.close();
			
		}
		catch(Exception e)
		{
			logger.error("Exception {}", e.getMessage());
		}
	}
}
