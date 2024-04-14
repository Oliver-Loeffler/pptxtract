package net.raumzeitfalle.pptxtract;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.PrintStream;
import java.nio.file.Files;
import java.nio.file.InvalidPathException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Enumeration;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.concurrent.Callable;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;

import picocli.CommandLine;
import picocli.CommandLine.Command;
import picocli.CommandLine.Option;
import picocli.CommandLine.Parameters;

/*
 * <a target="_blank" href="https://icons8.com/icon/dUePbPSb2D7d/microsoft-powerpoint-2019">Microsoft PowerPoint 2019</a> icon by <a target="_blank" href="https://icons8.com">Icons8</a>
 * <a target="_blank" href="https://icons8.com/icon/t7ev7yZzFFFy/explosion">Explosion</a> icon by <a target="_blank" href="https://icons8.com">Icons8</a>
 * 
 */
@Command(name="pptxtract",
         version = "0.0.15",
         mixinStandardHelpOptions = true)
public class PPTXtract implements Callable<Integer>{

    @Parameters(description = "PowerPoint FILE(S) where paths to embedded or linked documents shall be extracted. "
                            + "Must be of PowerPoint 2007 format (or later versions). Older *.ppt files must be "
                            + "converted into PowerPoint 2007 (or newer) format before use.",
                paramLabel = "FILE",
                arity = "1")
    List<String> sourceFiles = new ArrayList<>();
    
    @Option(names = {"-x", "--extract-embeddings"},
            description = "When set, embedded files such as *.docx, *.xlsx or other *.pptx files will be extracted.",
            defaultValue = "false")
    Boolean extractEmbeddings;
    
    @Option(names = {"-i", "--extract-images"},
            description = "When defined, imagfe files such as *.png, *.wmf or *.emf will be extracted as well.",
            defaultValue = "false")
    Boolean extractMedia; 
    
    @Option(names = {"-o"},
            description = "When extracting embedded files, this option will force overwriting existing files.",
            defaultValue = "false")
    Boolean overwriteExistingFiles; 
    
    private final PrintStream stdout = System.out;
    
    private final PrintStream stderr = System.err;
    
    private final Set<String> processedFiles = new HashSet<>();
    
    private final Set<String> extractableMediaExtensions = Set.of("png", "bmp", "dib", "wmf", "emf", "jpg");
    
    private final Set<String> extractableFileNameExtensions = merge(extractableMediaExtensions, "xlsx","xls", 
                                                                    "docx","doc", "pptx","ppt",
                                                                     "pdf", "txt","csv","ini",
                                                                     "tiff", "tif", "wav");
    
    private static Set<String> merge(Set<String> set, String ... other) {
        Set<String> merged = new HashSet<>(set);
        for (String s : other) {
            merged.add(s);
        }
        return Collections.unmodifiableSet(merged); 
    }
    
    @Override
    public Integer call() throws Exception {
        int returnCode = 0;
        for (String sourceFile : sourceFiles) {            
            int result = extractFromFile(sourceFile);
            processedFiles.add(sourceFile);
            if (result > returnCode) {
                returnCode = result;
            }
        }
        return returnCode;
    }

    private int extractFromFile(String fileName) {
        String sourceFile = fileName.strip();
        if (processedFiles.contains(sourceFile)) {
            return 0;
        }
        if (sourceFile.startsWith("\"")) {
            sourceFile = sourceFile.substring(1);
        }
        if (sourceFile.endsWith("\"")) {
            sourceFile = sourceFile.substring(0, sourceFile.length()-1);
        }
        
        if (sourceFile.endsWith(".ppt")) {
            stderr.println("Only *.pptx files are supported! "
                             + "Please convert *.ppt files into *.pptx using MS PowerPoint.");
            return 1;
        } else if (!sourceFile.endsWith(".pptx")) {
            stderr.println("Only *.pptx files are supported!");
            return 1;
        }
        
        int exitCode = 0;
        Path source = Path.of(sourceFile).toAbsolutePath().normalize();
        String extractableMediaFileExtensions = extractableMediaExtensions.stream().collect(Collectors.joining("|"));
        String extractableExtensions = extractableFileNameExtensions.stream().collect(Collectors.joining("|"));
        try (ZipFile zipFile = new ZipFile(source.toFile())) {
            Enumeration<? extends ZipEntry> zipEntries = zipFile.entries();
            while(zipEntries.hasMoreElements()) {
                ZipEntry entry = zipEntries.nextElement();
                String entryName = entry.getName();
                if (entryName.toLowerCase().matches("^ppt[/]embeddings[/].*[.]("+extractableExtensions+")$")) {
                    if (extractEmbeddings && !entry.isDirectory()) {
                        exitCode = handleEmbeddedFile(sourceFile, exitCode, source, zipFile, entry, entryName, "embedded");
                    }
                }
                if (entryName.toLowerCase().matches("^ppt[/]media[/].*[.]("+extractableMediaFileExtensions+")$")) {
                    if (extractMedia && !entry.isDirectory()) {
                        exitCode = handleEmbeddedFile(sourceFile, exitCode, source, zipFile, entry, entryName, "media");
                    }
                }
                if (entryName.matches("^ppt[/]slides[/]_rels[/]slide\\d+[.]xml[.]rels$")) {
                    XMLInputFactory xmlInputFactory = XMLInputFactory.newInstance();
                    try (InputStream is = zipFile.getInputStream(entry)) {
                        XMLEventReader reader = xmlInputFactory.createXMLEventReader(is);
                        while (reader.hasNext()) {
                            XMLEvent nextEvent = reader.nextEvent();
                            String embeddedFile = null;
                            if (nextEvent.isStartElement()) {
                                StartElement startElement = nextEvent.asStartElement();
                                String elementName = startElement.getName().getLocalPart();
                                if (elementName.equals("Relationship")) {
                                    Attribute targetMode = startElement.getAttributeByName(new QName("TargetMode"));
                                    Attribute target = startElement.getAttributeByName(new QName("Target"));
                                    if (targetMode != null && target != null && "External".equals(targetMode.getValue())) {
                                        embeddedFile = target.getValue();
                                        try {
                                            Path file = Path.of(embeddedFile.replace("file:///", ""))
                                                    .toAbsolutePath()
                                                    .normalize();
                                            stdout.println(sourceFile+";"+file);
                                        } catch (InvalidPathException pathError) {
                                            stdout.println(sourceFile+";"+embeddedFile);
                                        }
                                    }
                                }
                            }
                        }
                    } catch (IOException ioError) {
                        stderr.println("Failed to read OOXML file.");
                        if (exitCode < 2) {
                            exitCode = 2;
                        }
                        return 2;
                    } catch (XMLStreamException e) {
                        stderr.println("Unsupported OOXML/PPTX file format found.");
                        if (exitCode < 3) {
                            exitCode = 3;
                        }
                    }
                }            
            }
        } catch (IOException ioError) {
            stderr.println("Failed to extract embedded/linked files from given OOXML/PPTX file.");
            if (exitCode < 4) {
                exitCode = 4;
            }
        }
        return exitCode;
    }

    private int handleEmbeddedFile(String sourceFile, int exitCode, Path source, ZipFile zipFile, ZipEntry entry, String entryName, String type) {
        String targetName = Path.of(entryName).getFileName().toString();
        Path localTarget = source.getParent().resolve(targetName);
        int copyCount = 1;
        while(!overwriteExistingFiles && Files.exists(localTarget)) {
            var split = targetName.lastIndexOf('.');
            copyCount++;
            var counter = "("+copyCount+")";
            if (split > 0) {
                var head = targetName.substring(0, split);
                var tail = targetName.substring(split);
                localTarget = source.getParent().resolve(head+counter+tail);
            } else {
                localTarget = source.getParent().resolve(targetName+counter);
            }
        }
        try (InputStream is = zipFile.getInputStream(entry);
             OutputStream os = new FileOutputStream(localTarget.toFile())) {
            byte[] buffer = new byte[8 * 1024];
            int bytesRead;
            while ((bytesRead = is.read(buffer)) != -1) {
                os.write(buffer, 0, bytesRead);
            }
            stdout.println(sourceFile+";"+localTarget.toAbsolutePath()+";("+type+")");
        } catch (IOException error) {
            stderr.println("Failed to extract "+type+" file: " + targetName);
            if (exitCode < 5) {
                exitCode = 5;
            }
        }
        return exitCode;
    }
        
    public static void main(String[] args) {
        CommandLine commandLine = new CommandLine(new PPTXtract());
        String[] paramsAndOptions = readStdinWhenAvailable(args);
        int exitCode = commandLine.execute(paramsAndOptions);
        System.exit(exitCode);
    }

    private static String[] readStdinWhenAvailable(String[] args) {
        List<String> stdInItems = new ArrayList<>();
        try (InputStreamReader isr = new InputStreamReader(System.in);
             BufferedReader reader = new BufferedReader(isr)) {
             if (System.in.available() > 0) {
                String line = reader.readLine();
                if (line != null) {
                    stdInItems.add(line.strip());
                }
                while ((line = reader.readLine()) != null) {
                    stdInItems.add(line.strip());
                }
             }
        } catch (Exception e) {
            System.err.println("Error while reading from stdin.");
        }
        String[] params = new String[0];
        if (!stdInItems.isEmpty()) {
            params = stdInItems.toArray(new String[stdInItems.size()]);
        }
        String[] paramsAndOptions = new String[params.length+args.length];
        System.arraycopy(params,0,paramsAndOptions,0,params.length);
        System.arraycopy(args, 0, paramsAndOptions, params.length, args.length);
        return paramsAndOptions;
    }
}
