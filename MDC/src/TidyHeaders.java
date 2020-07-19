import java.nio.file.*;
import java.nio.charset.*;
import java.util.*;
import java.util.stream.*;
class TidyHeaders{
static Path completepath = Paths.get("U:\\H--ay\\T--MDC\\LG-------ssing\\Com----e\\");
static Charset standardcharset = StandardCharsets.UTF_8;
public static void main(String args[]){
Stream<Path> completefiles = Stream.of(completepath);
try { 
completefiles = Files.walk(completepath);}
catch (Exception e){
System.out.println("An error occurred when finding complete files.");}
try {
completefiles.forEach(c  -> replacecommas(c));}
catch (Exception e){
System.out.println("An error occurred when starting to process complete files.");}
}
public static void replacecommas(Path c){
	try {
String content = new String(Files.readAllBytes(c), standardcharset);
content = content.replaceAll(",,,,,,,,,,,,,,,,,,,,,,,,,,,", "");
	Files.write(c, content.getBytes(standardcharset));}
	catch (Exception e){
	System.out.println("An error occurred when removing commas.");}
}
}