package es.seresco.milena.pruebaMavenUno;

import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;

/**
 * Hello world!
 *
 */
public class App 
{
	private static final Logger logger = LogManager.getLogger(App.class);
	
	public static String devuelvePatata()
	{
		return("Patata");
	}
	
	public static String devuelvePiloro()
	{
		return("Nariz");
	}
	
    public static void main( String[] args )
    {
        System.out.println( "Hello World!" );
                
        logger.info("Traza desde logger");
        
        logger.info("Patata devuelve {}", devuelvePatata());
        
        logger.info("Piloro devuelve {}", devuelvePiloro());
        
    }
}
